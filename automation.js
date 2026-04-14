/**
 * NEO Quiz -- Automation Module (Reusable Class with EventEmitter)
 *
 * Refactored from fill-questions.js into a reusable Node.js module.
 * Emits progress/log events for real-time UI streaming.
 * OS-aware keyboard shortcuts (Meta on macOS, Control on Windows/Linux).
 *
 * Usage:
 *   const { QuizAutomation } = require('./automation');
 *   const bot = new QuizAutomation({ screenshotsDir: './screenshots' });
 *   bot.on('progress', data => console.log(data));
 *   bot.on('log', msg => console.log(msg));
 *   await bot.launch();
 *   await bot.processBatch([{ quizData, formUrl }]);
 */

const { chromium } = require('playwright');
const { EventEmitter } = require('events');
const fs = require('fs');
const path = require('path');
const os = require('os');

// ─── OS-Aware Keyboard Modifier ─────────────────────────────────────────────

const IS_MAC = os.platform() === 'darwin';
const MOD = IS_MAC ? 'Meta' : 'Control';
const SELECT_ALL = `${MOD}+A`;
const DUPLICATE_SHORTCUT = `${MOD}+d`;

// ─── Sample Question Names ──────────────────────────────────────────────────

const SAMPLE_TF = 'Sample Two Option';
const SAMPLE_4OPT = 'Sample Four Option';

// ─── Timing Constants ───────────────────────────────────────────────────────

const TIMING = {
  shortWait: 1000,
  mediumWait: 2000,
  longWait: 4000,
  pageLoadWait: 8000,
  typeDelay: 10,
};

// ─── Helper ─────────────────────────────────────────────────────────────────

function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

// ─── QuizAutomation Class ───────────────────────────────────────────────────

class QuizAutomation extends EventEmitter {
  /**
   * @param {Object} opts
   * @param {string} [opts.screenshotsDir] - Directory for screenshots
   * @param {string} [opts.authPath]       - Path to auth-state.json
   */
  constructor(opts = {}) {
    super();
    this.screenshotsDir = opts.screenshotsDir || path.join(__dirname, 'screenshots');
    this.authPath = opts.authPath || path.join(__dirname, 'auth-state.json');
    this.browser = null;
    this.context = null;
    this.page = null;
  }

  // ── Logging helpers ─────────────────────────────────────────────────────

  _log(msg) {
    const stamped = `[${new Date().toLocaleTimeString()}] ${msg}`;
    this.emit('log', stamped);
  }

  _progress(data) {
    this.emit('progress', data);
  }

  // ── Dismiss overlays / popups ────────────────────────────────────────

  /**
   * Dismiss any modal overlays that block interaction with the form.
   * Microsoft Forms can show feedback dialogs ("How's your experience?"),
   * Copilot banners, tooltips, etc. that have a backdrop intercepting all
   * pointer events. This method detects and closes them.
   */
  async _dismissOverlays() {
    try {
      // 1. Microsoft in-app feedback dialog (the main culprit)
      //    Has role="alertdialog" with id containing "ocv-inapp-feedback"
      const feedbackDialog = this.page.locator('[id*="ocv-inapp-feedback"]');
      if (await feedbackDialog.first().isVisible({ timeout: 500 }).catch(() => false)) {
        this._log('  Dismissing feedback dialog...');

        // Try close/dismiss/X buttons inside the dialog
        const closeSelectors = [
          '[id*="ocv-inapp-feedback"] button[aria-label*="Close" i]',
          '[id*="ocv-inapp-feedback"] button[aria-label*="Dismiss" i]',
          '[id*="ocv-inapp-feedback"] button[aria-label*="No" i]',
          '[id*="ocv-inapp-feedback"] button:has-text("Close")',
          '[id*="ocv-inapp-feedback"] button:has-text("Not now")',
          '[id*="ocv-inapp-feedback"] button:has-text("Dismiss")',
          '[id*="ocv-inapp-feedback"] button:has-text("No thanks")',
          '[id*="ocv-inapp-feedback"] button:has-text("Cancel")',
        ];

        let dismissed = false;
        for (const sel of closeSelectors) {
          const btn = this.page.locator(sel).first();
          if (await btn.isVisible({ timeout: 500 }).catch(() => false)) {
            await btn.click({ force: true });
            await wait(500);
            dismissed = true;
            this._log('  Feedback dialog dismissed (button)');
            break;
          }
        }

        // Fallback: press Escape to close the modal
        if (!dismissed) {
          await this.page.keyboard.press('Escape');
          await wait(500);
          this._log('  Feedback dialog dismissed (Escape)');
        }
      }

      // 2. Any other modal dialogs / overlays with backdrops
      const modalBackdrop = this.page.locator('.fui-DialogSurface__backdrop:visible');
      if (await modalBackdrop.first().isVisible({ timeout: 300 }).catch(() => false)) {
        await this.page.keyboard.press('Escape');
        await wait(500);
        this._log('  Dismissed modal overlay (Escape)');
      }

      // 3. Copilot / tooltip banners
      const copilotDismiss = this.page.locator(
        'button[aria-label*="Dismiss" i]:visible, button[aria-label*="Close" i]:visible',
      );
      const copilotCount = await copilotDismiss.count().catch(() => 0);
      for (let i = 0; i < copilotCount; i++) {
        const btn = copilotDismiss.nth(i);
        const label = await btn.getAttribute('aria-label').catch(() => '');
        // Only dismiss if it looks like an overlay/banner button, not a form element
        if (/dismiss|close/i.test(label) && !/delete|remove/i.test(label)) {
          const isInDialog = await btn.evaluate(
            el => !!el.closest('[role="alertdialog"], [role="dialog"], [class*="banner"], [class*="tooltip"]'),
          ).catch(() => false);
          if (isInDialog) {
            await btn.click({ force: true }).catch(() => {});
            await wait(300);
            this._log('  Dismissed banner/tooltip');
          }
        }
      }
    } catch (_) {
      // Non-critical — if overlay dismissal fails, we continue anyway
    }
  }

  // ── Screenshot helper ───────────────────────────────────────────────────

  async _screenshot(folder, name) {
    if (!this.page) return;
    const dir = path.join(this.screenshotsDir, folder);
    if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
    const file = `${name}.png`;
    await this.page.screenshot({ path: path.join(dir, file), fullPage: false });
    this._log(`  [screenshot] ${folder}/${file}`);
  }

  // ── Browser lifecycle ───────────────────────────────────────────────────

  /**
   * Launch a Chromium browser with persistent auth state if available.
   */
  async launch() {
    this._log('Launching browser...');
    this.browser = await chromium.launch({
      headless: false,
      slowMo: 100,
      args: ['--start-maximized'],
    });

    const hasAuth =
      fs.existsSync(this.authPath) && fs.statSync(this.authPath).size > 10;

    this.context = await this.browser.newContext({
      viewport: { width: 1400, height: 900 },
      storageState: hasAuth ? this.authPath : undefined,
    });

    this.page = await this.context.newPage();
    this._log('Browser launched');
  }

  /**
   * Close the browser and clean up.
   */
  async close() {
    if (this.browser) {
      await this.browser.close();
      this.browser = null;
      this.context = null;
      this.page = null;
      this._log('Browser closed');
    }
  }

  // ── Auth ────────────────────────────────────────────────────────────────

  async _saveAuth() {
    try {
      await this.context.storageState({ path: this.authPath });
    } catch (_) {
      /* ignore */
    }
  }

  async _handleSignIn() {
    const url = this.page.url();
    if (!url.includes('login') && !url.includes('microsoftonline')) return;

    this._log('  Sign-in required -- please sign in in the browser...');
    this._progress({ type: 'auth_required' });

    await this.page.waitForURL(
      u =>
        !u.toString().includes('login') &&
        !u.toString().includes('microsoftonline'),
      { timeout: 180000 },
    );
    await wait(TIMING.longWait);
    await this._saveAuth();
    this._log('  Sign-in complete, auth state saved');
  }

  // ── Navigation ──────────────────────────────────────────────────────────

  async _scrollToSampleQuestions() {
    this._log('  Scrolling to sample questions...');
    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(TIMING.shortWait);

    // Look for either sample question or a "Self Assessment" section header
    for (let attempt = 0; attempt < 40; attempt++) {
      for (const target of [SAMPLE_TF, SAMPLE_4OPT, 'Self Assessment']) {
        const el = this.page.locator(`text="${target}"`).first();
        const visible = await el.isVisible({ timeout: 300 }).catch(() => false);
        if (visible) {
          await el.scrollIntoViewIfNeeded();
          await wait(TIMING.shortWait);
          this._log(`  Found "${target}" on page`);
          return true;
        }
      }
      await this.page.mouse.wheel(0, 400);
      await wait(400);
    }
    this._log('  ERROR: Could not find sample questions on page');
    return false;
  }

  // ── Tab-Navigation Helpers ──────────────────────────────────────────────

  /**
   * Press Tab until landing on an editable field.
   * Returns info about the focused element or null.
   */
  async _tabToNextEditable(maxTabs) {
    if (maxTabs === undefined) maxTabs = 15;
    for (let i = 0; i < maxTabs; i++) {
      await this.page.keyboard.press('Tab');
      await wait(250);

      const info = await this.page.evaluate(() => {
        const el = document.activeElement;
        if (!el || el === document.body) return null;
        const tag = el.tagName.toLowerCase();
        const isEditable =
          el.contentEditable === 'true' ||
          tag === 'input' ||
          tag === 'textarea' ||
          el.getAttribute('role') === 'textbox';
        if (!isEditable) return null;
        return {
          tag,
          value: (el.value || el.textContent || '').substring(0, 100),
          placeholder: el.getAttribute('placeholder') || '',
          ariaLabel: el.getAttribute('aria-label') || '',
        };
      });

      if (info) return info;
    }
    return null;
  }

  /**
   * Select-all and type replacement text into the currently focused field.
   */
  async _clearAndType(text) {
    await this.page.keyboard.press(SELECT_ALL);
    await wait(150);
    await this.page.keyboard.type(text, { delay: TIMING.typeDelay });
    await wait(300);
  }

  // ── Form Title Replacement ──────────────────────────────────────────────

  async _replaceFormTitle(formName) {
    this._log('  Replacing form title...');

    // Exact sequence from Playwright codegen recording:
    // 1. Click the title card button (role="button" containing "Sample Title")
    // 2. Click the title text inside #desktop-scroller to enter edit mode
    // 3. Ctrl+A to select all, then .fill() with the new title

    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(TIMING.shortWait);

    for (let attempt = 0; attempt < 3; attempt++) {
      this._log(`  Title replacement attempt ${attempt + 1}/3...`);
      await this.page.evaluate(() => window.scrollTo(0, 0));
      await wait(TIMING.shortWait);

      try {
        // Step 1: Click the title card button to select it
        // Codegen recorded: getByRole('button', { name: 'Sample Title (Copy) Rewrite' })
        const titleBtn = this.page.getByRole('button').filter({ hasText: /Sample Title/ });
        if (await titleBtn.first().isVisible({ timeout: 3000 }).catch(() => false)) {
          await titleBtn.first().click();
          this._log('  Clicked title card button');
          await wait(TIMING.shortWait);
        } else {
          this._log('  Title card button not found, trying fallback...');
          // Fallback: click any visible element containing "Sample Title"
          const fallback = this.page.locator(':visible:has-text("Sample Title")').first();
          if (await fallback.isVisible({ timeout: 2000 }).catch(() => false)) {
            await fallback.click();
            await wait(TIMING.shortWait);
          } else {
            this._log('  ERROR: Could not find form title element');
            return false;
          }
        }

        // Step 2: Click the title text within #desktop-scroller to enter edit mode
        // Codegen recorded: locator('#desktop-scroller').getByText('Sample Title (Copy)')
        const titleText = this.page.locator('#desktop-scroller').getByText(/Sample Title/);
        if (await titleText.isVisible({ timeout: 3000 }).catch(() => false)) {
          await titleText.click();
          this._log('  Clicked title text in #desktop-scroller');
          await wait(300);

          // Step 3: Select all and fill with the new title
          // Codegen recorded: .press('ControlOrMeta+a') then .fill('...')
          await titleText.press(`${MOD}+a`);
          await wait(150);
          await titleText.fill(formName);
          this._log(`  Filled title with: "${formName}"`);
          await wait(300);
        } else {
          // Fallback: look for contenteditable containing "Sample Title"
          this._log('  #desktop-scroller text not found, trying contenteditable...');
          const ceElements = this.page.locator('[contenteditable="true"]:visible');
          const ceCount = await ceElements.count();
          let found = false;
          for (let i = 0; i < ceCount; i++) {
            const text = await ceElements.nth(i).textContent().catch(() => '');
            if (text.includes('Sample Title')) {
              await ceElements.nth(i).click();
              await wait(300);
              await this._clearAndType(formName);
              found = true;
              break;
            }
          }
          if (!found) {
            this._log('  Could not enter title edit mode this attempt');
            await this.page.mouse.click(700, 500);
            await wait(TIMING.shortWait);
            continue;
          }
        }

        // Click outside to commit
        await this.page.mouse.click(700, 500);
        await wait(TIMING.shortWait);

        // Verify the title was actually replaced
        const verified = await this._verifyTitleReplaced(formName);
        if (verified) {
          this._log(`  Title verified: "${formName}"`);
          return true;
        }

        this._log('  Title verification failed, will retry...');
        await this.page.mouse.click(700, 500);
        await wait(TIMING.shortWait);
      } catch (err) {
        this._log(`  Title attempt ${attempt + 1} error: ${err.message}`);
        await this.page.mouse.click(700, 500).catch(() => {});
        await wait(TIMING.shortWait);
      }
    }

    this._log(`  ERROR: Title replacement failed — could not verify "${formName}" on page`);
    return false;
  }

  async _verifyTitleReplaced(expectedTitle) {
    // Check if the expected title text appears anywhere on the page
    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(500);

    const prefix = expectedTitle.substring(0, 20);
    const found = this.page.locator(`:visible:has-text("${prefix}")`).first();
    return await found.isVisible({ timeout: 3000 }).catch(() => false);
  }

  // ── Sample-Question Operations ──────────────────────────────────────────

  async _clickSample(sampleName) {
    await this._dismissOverlays();

    const el = this.page.locator(`text="${sampleName}"`).first();

    for (let i = 0; i < 10; i++) {
      if (await el.isVisible({ timeout: 1000 }).catch(() => false)) break;
      await this._dismissOverlays();
      await this.page.mouse.wheel(0, 300);
      await wait(400);
    }

    if (!(await el.isVisible({ timeout: 3000 }).catch(() => false))) {
      throw new Error(`Cannot find sample question: "${sampleName}"`);
    }

    await el.scrollIntoViewIfNeeded();
    await wait(500);
    await el.click();
    await wait(TIMING.mediumWait);
    this._log(`  Selected sample: "${sampleName}"`);
  }

  async _duplicateSelected() {
    // Primary: exact selector from codegen recording
    // Codegen recorded: getByRole('button', { name: 'Copy question' })
    const copyBtn = this.page.getByRole('button', { name: 'Copy question' });
    if (await copyBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
      await copyBtn.click();
      await wait(TIMING.mediumWait);
      this._log('  Duplicated (Copy question button)');
      return true;
    }

    // Fallback 1: aria-label based selectors
    const iconSelectors = [
      '[aria-label*="Duplicate" i]',
      '[aria-label*="Copy" i]:not([aria-label*="Copilot" i])',
      '[title*="Duplicate" i]',
      '[title*="Copy" i]:not([title*="Copilot" i])',
    ];

    for (const sel of iconSelectors) {
      const btn = this.page.locator(sel).first();
      if (await btn.isVisible({ timeout: 1500 }).catch(() => false)) {
        await btn.click();
        await wait(TIMING.mediumWait);
        this._log('  Duplicated (icon button fallback)');
        return true;
      }
    }

    // Fallback 2: Keyboard shortcut
    await this.page.keyboard.press(DUPLICATE_SHORTCUT);
    await wait(TIMING.mediumWait);
    this._log(`  Duplicated (${DUPLICATE_SHORTCUT} fallback)`);
    return true;
  }

  // ── Fill a Duplicated Question ──────────────────────────────────────────

  async _fillDuplicatedQuestion(questionData, sampleName, ssFolder) {
    await this._dismissOverlays();
    const q = questionData;
    const sampleOptCount = sampleName === SAMPLE_TF ? 2 : 4;
    const targetOptCount = q.options.length;

    const isTrueFalse =
      q.options.length === 2 &&
      q.options[0].toLowerCase() === 'true' &&
      q.options[1].toLowerCase() === 'false';

    // ── 1. Adjust option count (multi-option only) ──────────────────────

    if (!isTrueFalse) {
      if (targetOptCount > sampleOptCount) {
        const toAdd = targetOptCount - sampleOptCount;
        this._log(`  Adding ${toAdd} option(s)...`);
        for (let i = 0; i < toAdd; i++) {
          const addBtn = this.page.locator(
            'button:has-text("Add option"), :text("+ Add option")',
          ).last();
          if (await addBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
            await addBtn.click();
            await wait(800);
          } else {
            this._log('  WARNING: Could not find "+ Add option" button');
          }
        }
      } else if (targetOptCount < sampleOptCount) {
        const toRemove = sampleOptCount - targetOptCount;
        this._log(`  Removing ${toRemove} extra option(s)...`);
        for (let i = 0; i < toRemove; i++) {
          const delBtn = this.page.locator(
            '[aria-label*="Delete option" i]:visible, ' +
              '[aria-label*="Remove option" i]:visible, ' +
              '[title*="Delete option" i]:visible',
          ).last();
          if (await delBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
            await delBtn.click();
            await wait(800);
          } else {
            this._log('  WARNING: Could not find option delete button');
          }
        }
      }
    } else {
      this._log('  True/False question -- skipping option adjustment');
    }

    // ── 2. Replace question text ────────────────────────────────────────
    // Codegen recording: getByRole('textbox', { name: /Question title \d+/ }).fill(text)
    // The duplicated question's textbox contains the sample name text.

    this._log('  Replacing question text...');
    let textReplaced = false;
    const samplePrefix = sampleName.substring(0, 12);

    // Primary: find a textbox role element containing the sample text (from codegen)
    const allTextboxes = this.page.getByRole('textbox');
    const tbCount = await allTextboxes.count();
    for (let i = tbCount - 1; i >= 0; i--) {
      const text = await allTextboxes.nth(i).textContent().catch(() => '');
      const val = await allTextboxes.nth(i).inputValue().catch(() => '');
      if ((text && text.includes(samplePrefix)) || (val && val.includes(samplePrefix))) {
        await allTextboxes.nth(i).click();
        await wait(300);
        await allTextboxes.nth(i).press(`${MOD}+a`);
        await wait(150);
        await allTextboxes.nth(i).fill(q.text);
        textReplaced = true;
        this._log(`  Question text replaced (textbox role #${i})`);
        break;
      }
    }

    // Fallback A: visible contenteditable elements
    if (!textReplaced) {
      const ceElements = this.page.locator('[contenteditable="true"]:visible');
      const ceCount = await ceElements.count();
      for (let i = ceCount - 1; i >= 0; i--) {
        const text = await ceElements.nth(i).textContent().catch(() => '');
        if (text.includes(samplePrefix)) {
          await ceElements.nth(i).click();
          await wait(300);
          await this._clearAndType(q.text);
          textReplaced = true;
          this._log(`  Question text replaced (contenteditable #${i})`);
          break;
        }
      }
    }

    if (!textReplaced) {
      this._log('  ERROR: Could not find or replace question text');
      await this._screenshot(ssFolder, 'question_text_failed');
      return false;
    }

    // ── 3. Replace option texts (multi-option only) ─────────────────────
    // Codegen recording shows each option is in a listitem with a field
    // labeled 'Choice Option Text Please'. After duplication, the default
    // option texts are "Option 1", "Option 2", etc. We target each by its
    // default text, then .fill() with the real option text.

    if (!isTrueFalse) {
      this._log(`  Filling ${q.options.length} option(s) via direct targeting...`);
      const defaultOptionNames = ['Option 1', 'Option 2', 'Option 3', 'Option 4', 'Option 5', 'Option 6'];

      for (let i = 0; i < q.options.length; i++) {
        const defaultName = defaultOptionNames[i] || `Option ${i + 1}`;
        let optionFilled = false;

        // Primary: find the listitem containing the default option name,
        // then fill its "Choice Option Text Please" field
        const optItem = this.page.getByRole('listitem').filter({ hasText: defaultName });
        const optField = optItem.getByLabel('Choice Option Text Please');

        // Use .last() in case multiple listitems match (e.g. sample + duplicate)
        if (await optField.last().isVisible({ timeout: 2000 }).catch(() => false)) {
          await optField.last().click();
          await wait(200);
          await optField.last().fill(q.options[i]);
          await wait(300);
          optionFilled = true;
          this._log(`    Option ${i + 1}: "${q.options[i].substring(0, 40)}${q.options[i].length > 40 ? '...' : ''}" (direct targeting)`);
        }

        // Fallback: Tab-based navigation
        if (!optionFilled) {
          this._log(`    Option ${i + 1}: direct targeting failed, trying Tab...`);
          const fieldInfo = await this._tabToNextEditable();
          if (fieldInfo) {
            await this._clearAndType(q.options[i]);
            optionFilled = true;
            this._log(`    Option ${i + 1}: "${q.options[i].substring(0, 40)}..." (Tab fallback, was: "${fieldInfo.value.substring(0, 30)}")`);
          }
        }

        if (!optionFilled) {
          this._log(`  ERROR: Could not fill option ${i + 1}`);
          await this._screenshot(ssFolder, `option_${i + 1}_failed`);
        }
      }
    } else {
      this._log('  True/False question -- options already correct, skipping');
    }

    // ── 4. Mark correct answer ──────────────────────────────────────────
    // Codegen recording confirms: after filling options, the "Correct answer"
    // buttons are ALREADY visible in each option's listitem. Do NOT press
    // Escape or click outside — that collapses the question card and hides
    // the buttons. Just find the correct option's listitem and click its button.
    //
    // IMPORTANT: For TF questions we skip option filling, so the cursor is
    // still in the question text field. We must Tab out of it first —
    // otherwise the button click doesn't register. The codegen recording
    // also shows Tab presses between filling text and marking the answer.
    //
    // Pattern from recording:
    //   page.getByRole('listitem').filter({ hasText: optionText }).getByLabel('Correct answer').click()

    if (!q.needs_review) {
      const correctOptionText = q.options[q.correct_answer_index];
      this._log(`  Marking correct answer: option ${q.correct_answer_index + 1} ("${correctOptionText.substring(0, 30)}")`);

      // Tab out of any active text field so the correct answer button is clickable.
      // This is critical for TF questions where option filling is skipped.
      await this.page.keyboard.press('Tab');
      await wait(300);

      await this._dismissOverlays();
      let answerMarked = false;

      for (let attempt = 0; attempt < 4; attempt++) {
        if (attempt > 0) {
          this._log(`  Correct-answer marking retry (attempt ${attempt + 1})...`);
        }

        // On first attempt: buttons should already be visible (question card is open).
        // On retries: re-expand the question card by clicking it.
        if (attempt > 0) {
          await this._dismissOverlays();
          // Re-expand: find and click the question card button
          const qSearchText = q.text.substring(0, 30).replace(/["""'']/g, '');
          const qCardBtn = this.page.getByRole('button').filter({ hasText: qSearchText });
          if (await qCardBtn.first().isVisible({ timeout: 2000 }).catch(() => false)) {
            await qCardBtn.first().scrollIntoViewIfNeeded().catch(() => {});
            await wait(300);
            await qCardBtn.first().click();
            await wait(TIMING.shortWait);
          }
        }

        // Primary: find the listitem containing the correct option text,
        // then click its "Correct answer" button.
        // Use .last() because the same option text (e.g. "True", "False")
        // may exist in BOTH the sample question and the duplicate we just filled.
        // The duplicate is always below the sample, so .last() gets the right one.
        const optionItem = this.page.getByRole('listitem').filter({ hasText: correctOptionText });
        const correctBtn = optionItem.last().getByLabel('Correct answer');

        if (await correctBtn.isVisible({ timeout: 3000 }).catch(() => false)) {
          await correctBtn.scrollIntoViewIfNeeded().catch(() => {});
          await wait(200);
          await correctBtn.click();
          await wait(500);
          this._log(`  Marked "${correctOptionText.substring(0, 30)}" as correct (listitem approach, attempt ${attempt + 1})`);
          answerMarked = true;
          break;
        }

        // Fallback: index-based — find all visible "Correct answer" buttons
        const allCorrectBtns = this.page.locator(
          '[aria-label*="Correct answer" i]:visible, ' +
            '[title*="Correct answer" i]:visible',
        );
        const btnCount = await allCorrectBtns.count();
        if (btnCount > 0 && q.correct_answer_index < btnCount) {
          const targetBtn = allCorrectBtns.nth(q.correct_answer_index);
          await targetBtn.scrollIntoViewIfNeeded().catch(() => {});
          await wait(200);
          await targetBtn.click();
          await wait(500);
          this._log(`  Marked option ${q.correct_answer_index + 1} as correct (index fallback, ${btnCount} buttons, attempt ${attempt + 1})`);
          answerMarked = true;
          break;
        }

        this._log(`  No correct-answer buttons found (attempt ${attempt + 1}, listitem matches: ${await optionItem.count().catch(() => 0)}, fallback buttons: ${btnCount})`);
      }

      if (!answerMarked) {
        this._log(`  ERROR: Could not mark correct answer after 4 attempts`);
        await this._screenshot(ssFolder, 'correct_answer_failed');
        return false;
      }
    } else {
      this._log('  Skipping correct answer (needs review)');
    }

    return true;
  }

  // ── Delete Sample Questions ─────────────────────────────────────────────

  async _deleteSampleQuestion(sampleText) {
    this._log(`  Deleting sample: "${sampleText}"`);

    const el = this.page.locator(`text="${sampleText}"`).first();

    // Scroll up to find it
    for (let i = 0; i < 15; i++) {
      if (await el.isVisible({ timeout: 800 }).catch(() => false)) break;
      await this.page.mouse.wheel(0, -300);
      await wait(400);
    }

    if (!(await el.isVisible({ timeout: 3000 }).catch(() => false))) {
      this._log(`  "${sampleText}" not found (may already be deleted or scrolled away)`);
      return;
    }

    await el.scrollIntoViewIfNeeded();
    await wait(500);
    await el.click();
    await wait(TIMING.mediumWait);

    // Find the delete button in the question toolbar
    const delSelectors = [
      '[aria-label*="Delete question" i]',
      '[aria-label*="Delete" i]:not([aria-label*="option" i])',
      '[title*="Delete question" i]',
      '[title*="Delete" i]:not([title*="option" i])',
    ];

    for (const sel of delSelectors) {
      const btn = this.page.locator(sel).first();
      if (await btn.isVisible({ timeout: 2000 }).catch(() => false)) {
        await btn.click();
        await wait(TIMING.shortWait);

        // Handle confirmation dialog
        const confirmBtn = this.page.locator(
          'button:has-text("OK"), button:has-text("Yes"), button:has-text("Delete")',
        ).first();
        if (await confirmBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
          await confirmBtn.click();
          await wait(TIMING.shortWait);
        }

        this._log(`  Deleted "${sampleText}"`);
        return;
      }
    }

    // Fallback: "..." menu -> Delete
    const moreBtn = this.page.locator(
      '[aria-label*="More" i]:not([aria-label*="Learn more" i])',
    ).first();
    if (await moreBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
      await moreBtn.click();
      await wait(TIMING.shortWait);
      const delItem = this.page.locator(
        '[role="menuitem"]:has-text("Delete")',
      ).first();
      if (await delItem.isVisible({ timeout: 2000 }).catch(() => false)) {
        await delItem.click();
        await wait(TIMING.shortWait);

        const confirmBtn = this.page.locator(
          'button:has-text("OK"), button:has-text("Yes"), button:has-text("Delete")',
        ).first();
        if (await confirmBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
          await confirmBtn.click();
          await wait(TIMING.shortWait);
        }

        this._log(`  Deleted "${sampleText}" (via ... menu)`);
        return;
      }
      await this.page.keyboard.press('Escape');
      await wait(300);
    }

    this._log(`  WARNING: Could not delete "${sampleText}"`);
  }

  // ── Template Duplication ────────────────────────────────────────────────

  /**
   * Navigate to a shared form template URL, click "Duplicate it",
   * wait for the new form editor to load, and return the new edit URL.
   * @param {string} templateUrl - The shared template URL
   * @param {string} formName    - For logging/progress events
   * @returns {string} The edit URL of the newly duplicated form
   */
  async _duplicateTemplate(templateUrl, formName) {
    this._log(`  Duplicating template for "${formName}"...`);
    this._progress({ type: 'template_start', formName });

    // Navigate to the shared template page
    await this.page.goto(templateUrl, {
      waitUntil: 'domcontentloaded',
      timeout: 60000,
    });
    await wait(TIMING.pageLoadWait);
    await this._handleSignIn();
    await wait(TIMING.longWait);

    // Look for the "Duplicate it" button
    const dupSelectors = [
      'button:has-text("Duplicate it")',
      'button:has-text("Duplicate It")',
      'a:has-text("Duplicate it")',
      '[aria-label*="Duplicate" i]',
    ];

    let clicked = false;
    for (const sel of dupSelectors) {
      const btn = this.page.locator(sel).first();
      if (await btn.isVisible({ timeout: 5000 }).catch(() => false)) {
        await btn.click();
        clicked = true;
        this._log('  Clicked "Duplicate it" button');
        break;
      }
    }

    if (!clicked) {
      throw new Error('Could not find "Duplicate it" button on template page');
    }

    // Wait for the form editor to load
    this._log('  Waiting for duplicated form to load...');
    await this.page.waitForURL(
      url => {
        const u = url.toString();
        return u.includes('DesignPage') && u.includes('subpage=design');
      },
      { timeout: 90000 }
    );
    await wait(TIMING.pageLoadWait);

    // Dismiss any startup overlays / Copilot banners / feedback dialogs
    for (let d = 0; d < 3; d++) {
      await this.page.keyboard.press('Escape');
      await wait(300);
    }
    await this._dismissOverlays();
    await wait(TIMING.shortWait);

    const newUrl = this.page.url();
    this._log(`  Template duplicated. New form URL: ${newUrl}`);
    this._progress({ type: 'template_done', formName, newUrl });

    return newUrl;
  }

  // ── Process One Form (public) ───────────────────────────────────────────

  /**
   * Process a single form: navigate, fill questions, clean up samples.
   * @param {Object} quizData - Parsed quiz JSON with form_name, questions[]
   * @param {string} formUrl  - Microsoft Forms edit URL
   * @returns {Object} { success: boolean, error?: string }
   */
  async processForm(quizData, formUrl) {
    const formName = quizData.form_name;
    const ssFolder = `form_${formName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)}`;

    this._progress({ type: 'form_start', formName, formIndex: 0, totalForms: 1 });

    try {
      // Navigate to the form
      await this.page.goto(formUrl, {
        waitUntil: 'domcontentloaded',
        timeout: 60000,
      });
      await wait(TIMING.pageLoadWait);
      await this._handleSignIn();
      await wait(TIMING.longWait);

      // Dismiss startup overlays / tooltips / feedback dialogs
      for (let d = 0; d < 3; d++) {
        await this.page.keyboard.press('Escape');
        await wait(300);
      }
      await this._dismissOverlays();
      await wait(TIMING.shortWait);

      await this._screenshot(ssFolder, '00_form_loaded');

      // Core processing
      await this._processFormInternal(quizData, ssFolder);

      this._progress({ type: 'form_done', formName, success: true });
      this._log(`DONE: ${formName}`);
      return { success: true };
    } catch (err) {
      this._progress({ type: 'form_error', formName, error: err.message });
      this._log(`FAILED: ${formName} -- ${err.message}`);
      await this._screenshot(ssFolder, 'FAILED').catch(() => {});
      return { success: false, error: err.message };
    }
  }

  /**
   * Internal form processing (title, questions, cleanup).
   */
  async _processFormInternal(quizData, ssFolder) {
    const formName = quizData.form_name;
    this._log(`Processing: ${formName} (${quizData.questions.length} questions)`);

    // 1. Replace the form title
    const titleReplaced = await this._replaceFormTitle(formName);
    if (titleReplaced) {
      this._progress({ type: 'title_replaced', formName });
    } else {
      this._log('  ERROR: Title replacement failed');
      this._progress({ type: 'title_error', formName, error: 'Title replacement failed — could not verify new title on page' });
    }
    await this._screenshot(ssFolder, '00_title_replaced');

    // 2. Navigate to sample questions
    const found = await this._scrollToSampleQuestions();
    if (!found) throw new Error('Could not find sample questions on page');
    this._progress({ type: 'section_found', formName });
    await this._screenshot(ssFolder, '01_questions_section');

    // 3. Duplicate + fill each question
    const totalQuestions = quizData.questions.length;
    let firstFourOptDone = false;

    for (let i = 0; i < totalQuestions; i++) {
      const q = quizData.questions[i];
      const qLabel = `Q${String(i + 1).padStart(2, '0')}`;
      const sampleName = q.options.length <= 2 ? SAMPLE_TF : SAMPLE_4OPT;
      const questionNum = i + 1;

      this._progress({
        type: 'question_start',
        formName,
        question: questionNum,
        totalQuestions,
        questionText: q.text.substring(0, 80),
      });

      this._log(
        `\n  --- ${qLabel}/${totalQuestions}: ` +
          `"${q.text.substring(0, 50)}..." (${q.options.length} opts -> "${sampleName}") ---`,
      );

      try {
        // a. Click the sample to select it
        await this._clickSample(sampleName);
        await this._screenshot(ssFolder, `${qLabel}_a_selected`);

        // b. Duplicate it
        const duped = await this._duplicateSelected();
        if (!duped) {
          this._log(`  ERROR: Duplication failed for ${qLabel}`);
          await this._screenshot(ssFolder, `${qLabel}_b_dup_failed`);
          this._progress({
            type: 'question_error',
            formName,
            question: questionNum,
            error: 'Duplication failed',
          });
          continue;
        }
        await this._screenshot(ssFolder, `${qLabel}_b_duplicated`);

        // c. Fill the duplicate with actual quiz data
        const filled = await this._fillDuplicatedQuestion(q, sampleName, ssFolder);
        await this._screenshot(ssFolder, `${qLabel}_c_filled`);

        if (!filled) {
          this._log(`  ERROR: ${qLabel} failed — question may be incomplete`);
          this._progress({
            type: 'question_error',
            formName,
            question: questionNum,
            error: 'Question fill or correct answer marking failed',
          });
        }

        // d. Click outside to commit changes
        await this.page.mouse.click(700, 150);
        await wait(TIMING.mediumWait);

        // e. Extra wait for the first 4-option question to ensure auto-save
        if (sampleName === SAMPLE_4OPT && !firstFourOptDone) {
          this._log('  First 4-option question -- extra wait for auto-save');
          await wait(TIMING.longWait);
          firstFourOptDone = true;
        }

        // f. Scroll down slightly so next duplicate appears in view
        await this.page.mouse.wheel(0, 200);
        await wait(500);

        this._progress({
          type: 'question_done',
          formName,
          question: questionNum,
          totalQuestions,
        });
        this._log(`  ${qLabel} complete`);
      } catch (err) {
        this._progress({
          type: 'question_error',
          formName,
          question: questionNum,
          error: err.message,
        });
        this._log(`  ERROR on ${qLabel}: ${err.message}`);
        await this._screenshot(ssFolder, `${qLabel}_error`);
        await this.page.mouse.click(700, 150).catch(() => {});
        await wait(TIMING.shortWait);
      }
    }

    // 4. Delete the two sample questions
    await this._dismissOverlays();
    this._log('\n  --- Cleaning up sample questions ---');
    this._progress({ type: 'cleanup_start', formName });

    await this._scrollToSampleQuestions();
    await wait(TIMING.shortWait);

    await this._deleteSampleQuestion(SAMPLE_TF);
    await wait(TIMING.shortWait);
    await this._deleteSampleQuestion(SAMPLE_4OPT);
    await wait(TIMING.shortWait);

    await this._screenshot(ssFolder, '99_cleanup_done');

    // 5. Quality check — verify title, questions, and correct answers
    this._log('\n  --- Running quality check ---');
    this._progress({ type: 'verification_start', formName });
    const qcResults = await this._qualityCheck(quizData, ssFolder);
    this._progress({ type: 'verification_done', formName, results: qcResults });

    if (qcResults.issues.length > 0) {
      this._log(`  Quality check: ${qcResults.issues.length} issue(s) found`);
      for (const issue of qcResults.issues) {
        this._log(`    - ${issue}`);
      }
    } else {
      this._log('  Quality check: all checks passed');
    }

    await this._screenshot(ssFolder, '99_complete');
    this._log('  Form processing complete');
  }

  // ── Quality Check ──────────────────────────────────────────────────────

  async _qualityCheck(quizData, ssFolder) {
    const issues = [];
    const formName = quizData.form_name;

    // 1. Verify title
    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(TIMING.shortWait);

    const titlePrefix = formName.substring(0, 25);
    const titleFound = this.page.locator(`:visible:has-text("${titlePrefix}")`).first();
    const titleOk = await titleFound.isVisible({ timeout: 3000 }).catch(() => false);

    if (!titleOk) {
      issues.push(`Title not found: expected "${formName}" but could not find it on the page`);
    }

    // Also check that "Sample Title" is NOT still present (means replacement failed)
    const sampleStillThere = this.page.locator(':visible:has-text("Sample Title")').first();
    const sampleVisible = await sampleStillThere.isVisible({ timeout: 2000 }).catch(() => false);
    if (sampleVisible) {
      issues.push('Title still shows "Sample Title" — replacement did not take effect');
    }

    // 2. Verify each question is present on the page
    for (let i = 0; i < quizData.questions.length; i++) {
      const q = quizData.questions[i];
      const qNum = i + 1;
      // Escape special chars (quotes, etc.) that break :has-text() selectors
      const qPrefix = q.text.substring(0, 30).replace(/["""'''`\\]/g, '');

      // Scroll down to find the question
      let qFound = false;
      for (let scroll = 0; scroll < 15; scroll++) {
        const el = this.page.locator(`:visible:has-text("${qPrefix}")`).first();
        if (await el.isVisible({ timeout: 500 }).catch(() => false)) {
          qFound = true;
          break;
        }
        await this.page.mouse.wheel(0, 300);
        await wait(300);
      }

      if (!qFound) {
        issues.push(`Q${qNum}: Question text not found on page — "${q.text.substring(0, 60)}..."`);
        continue;
      }

      // Note: correct-answer marking verification is handled during processing
      // (the automation logs success/failure when clicking the "Correct answer"
      // button). DOM-based detection of the marked state is unreliable because
      // Microsoft Forms uses internal state that isn't exposed via standard
      // aria attributes or visible DOM changes that Playwright can detect.

      // Click outside to deselect
      await this.page.mouse.click(700, 150);
      await wait(500);
    }

    // 3. Verify sample questions were deleted
    for (const sampleName of [SAMPLE_TF, SAMPLE_4OPT]) {
      const sampleEl = this.page.locator(`text="${sampleName}"`).first();
      const stillThere = await sampleEl.isVisible({ timeout: 2000 }).catch(() => false);
      if (stillThere) {
        issues.push(`Sample question "${sampleName}" was not deleted`);
      }
    }

    await this._screenshot(ssFolder, '99_quality_check');
    return { passed: issues.length === 0, issues };
  }

  // ── Process Batch (public) ──────────────────────────────────────────────

  /**
   * Process multiple forms sequentially, then auto-close the browser.
   * @param {Array<{quizData: Object, formUrl: string}>} items
   * @param {Object} [opts]
   * @param {string} [opts.templateUrl] - If provided, each form is created by
   *   duplicating this template first (instead of using a pre-existing form URL).
   * @returns {Array<{formName: string, success: boolean, error?: string, newFormUrl?: string}>}
   */
  async processBatch(items, opts) {
    const templateUrl = (opts && opts.templateUrl) || null;
    const total = items.length;
    this._progress({ type: 'batch_start', total, useTemplate: !!templateUrl });
    this._log(`=== NEO Quiz Automation: ${total} form(s) ===`);
    if (templateUrl) {
      this._log(`Template mode: each form will be duplicated from the shared template`);
    }

    // Ensure browser is launched
    if (!this.page) {
      await this.launch();
    }

    const results = [];

    for (let i = 0; i < items.length; i++) {
      const { quizData, formUrl, form_url } = items[i];
      let url = formUrl || form_url || null;
      const formName = quizData.form_name;
      const ssFolder = `form_${String(i + 1).padStart(2, '0')}_${formName.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)}`;

      this._progress({
        type: 'form_start',
        formName,
        formIndex: i,
        totalForms: total,
      });

      this._log(`\n${'='.repeat(60)}`);
      this._log(`[${i + 1}/${total}] ${formName}`);
      this._log('='.repeat(60));

      try {
        // Template mode: duplicate the template to get a fresh form URL
        if (templateUrl && !url) {
          url = await this._duplicateTemplate(templateUrl, formName);
          await this._screenshot(ssFolder, '00_template_duplicated');
        }

        if (!url) {
          throw new Error(`No form URL provided for "${formName}"`);
        }

        // If we just duplicated, we're already on the editor page — skip navigation
        const currentUrl = this.page.url();
        const alreadyOnForm = url === currentUrl;

        if (!alreadyOnForm) {
          await this.page.goto(url, {
            waitUntil: 'domcontentloaded',
            timeout: 60000,
          });
          await wait(TIMING.pageLoadWait);
          await this._handleSignIn();
          await wait(TIMING.longWait);

          // Dismiss startup overlays / tooltips / feedback dialogs
          for (let d = 0; d < 3; d++) {
            await this.page.keyboard.press('Escape');
            await wait(300);
          }
          await this._dismissOverlays();
          await wait(TIMING.shortWait);
        }

        await this._screenshot(ssFolder, '00_form_loaded');

        // Core processing
        await this._processFormInternal(quizData, ssFolder);

        this._progress({ type: 'form_done', formName, success: true, newFormUrl: url });
        this._log(`DONE: ${formName}`);
        results.push({ formName, success: true, newFormUrl: url });
      } catch (err) {
        this._progress({ type: 'form_error', formName, error: err.message });
        this._log(`FAILED: ${formName} -- ${err.message}`);
        await this._screenshot(ssFolder, 'FAILED').catch(() => {});
        results.push({ formName, success: false, error: err.message, newFormUrl: url || null });
      }
    }

    this._progress({ type: 'batch_done', results });
    this._log(`\n${'='.repeat(60)}`);
    this._log('ALL FORMS COMPLETE');
    this._log('='.repeat(60));

    // Auto-close browser after batch
    await wait(3000);
    await this.close();

    return results;
  }
}

// ─── Exports ────────────────────────────────────────────────────────────────

module.exports = { QuizAutomation, SAMPLE_TF, SAMPLE_4OPT };
