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

  async _scrollToSelfAssessment() {
    this._log('  Scrolling to Self Assessment section...');
    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(TIMING.shortWait);

    for (let attempt = 0; attempt < 40; attempt++) {
      const el = this.page.locator('text="Self Assessment"').first();
      const visible = await el.isVisible({ timeout: 500 }).catch(() => false);
      if (visible) {
        await el.scrollIntoViewIfNeeded();
        await wait(TIMING.shortWait);
        this._log('  Found Self Assessment section');
        return true;
      }
      await this.page.mouse.wheel(0, 400);
      await wait(400);
    }
    this._log('  ERROR: Could not find Self Assessment section');
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

    await this.page.evaluate(() => window.scrollTo(0, 0));
    await wait(TIMING.shortWait);

    // Strategy A: contenteditable containing "Sample Title"
    const ceElements = this.page.locator('[contenteditable="true"]:visible');
    const ceCount = await ceElements.count();
    for (let i = 0; i < ceCount; i++) {
      const text = await ceElements.nth(i).textContent().catch(() => '');
      if (text.includes('Sample Title')) {
        await ceElements.nth(i).click();
        await wait(300);
        await this._clearAndType(formName);
        this._log(`  Title replaced (contenteditable): "${formName}"`);
        await this.page.keyboard.press('Escape');
        await wait(TIMING.shortWait);
        return true;
      }
    }

    // Strategy B: heading containing "Sample Title"
    const headings = this.page.locator('h1:visible, h2:visible, [role="heading"]:visible');
    const hCount = await headings.count();
    for (let i = 0; i < hCount; i++) {
      const text = await headings.nth(i).textContent().catch(() => '');
      if (text.includes('Sample Title')) {
        await headings.nth(i).click();
        await wait(500);
        await this._clearAndType(formName);
        this._log(`  Title replaced (heading click): "${formName}"`);
        await this.page.keyboard.press('Escape');
        await wait(TIMING.shortWait);
        return true;
      }
    }

    // Strategy C: any visible text containing "Sample Title" (handles "Sample Title (Copy)" etc.)
    const sampleTitleEl = this.page.locator(':visible:has-text("Sample Title")').first();
    if (await sampleTitleEl.isVisible({ timeout: 3000 }).catch(() => false)) {
      await sampleTitleEl.click();
      await wait(500);
      await this._clearAndType(formName);
      this._log(`  Title replaced (text match): "${formName}"`);
      await this.page.keyboard.press('Escape');
      await wait(TIMING.shortWait);
      return true;
    }

    // Strategy D: input with value containing "Sample Title"
    const inputs = this.page.locator('input:visible, textarea:visible');
    const inputCount = await inputs.count();
    for (let i = 0; i < inputCount; i++) {
      const val = await inputs.nth(i).inputValue().catch(() => '');
      if (val.includes('Sample Title')) {
        await inputs.nth(i).click();
        await wait(300);
        await this._clearAndType(formName);
        this._log(`  Title replaced (input): "${formName}"`);
        await this.page.keyboard.press('Escape');
        await wait(TIMING.shortWait);
        return true;
      }
    }

    this._log('  WARNING: Could not find form title to replace');
    return false;
  }

  // ── Sample-Question Operations ──────────────────────────────────────────

  async _clickSample(sampleName) {
    const el = this.page.locator(`text="${sampleName}"`).first();

    for (let i = 0; i < 10; i++) {
      if (await el.isVisible({ timeout: 1000 }).catch(() => false)) break;
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
    // Strategy 1: Copy / Duplicate icon button
    const iconSelectors = [
      '[aria-label*="Duplicate" i]',
      '[aria-label*="Copy question" i]',
      '[aria-label*="Copy" i]:not([aria-label*="Copilot" i])',
      '[title*="Duplicate" i]',
      '[title*="Copy" i]:not([title*="Copilot" i])',
    ];

    for (const sel of iconSelectors) {
      const btn = this.page.locator(sel).first();
      if (await btn.isVisible({ timeout: 1500 }).catch(() => false)) {
        await btn.click();
        await wait(TIMING.mediumWait);
        this._log('  Duplicated (icon button)');
        return true;
      }
    }

    // Strategy 2: "..." menu -> Duplicate
    const moreBtn = this.page.locator(
      '[aria-label*="More options" i], [aria-label*="More actions" i]',
    ).first();
    if (await moreBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
      await moreBtn.click();
      await wait(TIMING.shortWait);
      const dupItem = this.page.locator(
        '[role="menuitem"]:has-text("Duplicate"), [role="menuitem"]:has-text("Copy")',
      ).first();
      if (await dupItem.isVisible({ timeout: 2000 }).catch(() => false)) {
        await dupItem.click();
        await wait(TIMING.mediumWait);
        this._log('  Duplicated (... menu)');
        return true;
      }
      await this.page.keyboard.press('Escape');
      await wait(500);
    }

    // Strategy 3: Keyboard shortcut
    await this.page.keyboard.press(DUPLICATE_SHORTCUT);
    await wait(TIMING.mediumWait);
    this._log(`  Duplicated (${DUPLICATE_SHORTCUT} fallback)`);
    return true;
  }

  // ── Fill a Duplicated Question ──────────────────────────────────────────

  async _fillDuplicatedQuestion(questionData, sampleName, ssFolder) {
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

    this._log('  Replacing question text...');
    let textReplaced = false;
    const samplePrefix = sampleName.substring(0, 12);

    // Approach A: visible contenteditable elements
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

    // Approach B: visible input / textarea elements
    if (!textReplaced) {
      const inputs = this.page.locator('input:visible, textarea:visible');
      const inputCount = await inputs.count();
      for (let i = inputCount - 1; i >= 0; i--) {
        const val = await inputs.nth(i).inputValue().catch(() => '');
        if (val.includes(samplePrefix)) {
          await inputs.nth(i).click();
          await wait(300);
          await this._clearAndType(q.text);
          textReplaced = true;
          this._log(`  Question text replaced (input #${i})`);
          break;
        }
      }
    }

    // Approach C: visible elements with role="textbox"
    if (!textReplaced) {
      const textboxes = this.page.locator('[role="textbox"]:visible');
      const tbCount = await textboxes.count();
      for (let i = tbCount - 1; i >= 0; i--) {
        const text = await textboxes.nth(i).textContent().catch(() => '');
        if (text.includes(samplePrefix)) {
          await textboxes.nth(i).click();
          await wait(300);
          await this._clearAndType(q.text);
          textReplaced = true;
          this._log(`  Question text replaced (textbox #${i})`);
          break;
        }
      }
    }

    // Approach D: Tab through editable fields looking for sample text
    if (!textReplaced) {
      this._log('  Trying Tab-based search for question text field...');
      for (let t = 0; t < 20; t++) {
        const info = await this._tabToNextEditable(1);
        if (info && info.value.includes(samplePrefix)) {
          await this._clearAndType(q.text);
          textReplaced = true;
          this._log('  Question text replaced (Tab search)');
          break;
        }
      }
    }

    if (!textReplaced) {
      this._log('  ERROR: Could not find or replace question text');
      await this._screenshot(ssFolder, 'question_text_failed');
      return false;
    }

    // ── 3. Replace option texts via Tab (multi-option only) ─────────────

    if (!isTrueFalse) {
      this._log(`  Filling ${q.options.length} option(s) via Tab...`);

      for (let i = 0; i < q.options.length; i++) {
        const fieldInfo = await this._tabToNextEditable();
        if (!fieldInfo) {
          this._log(`  ERROR: Could not Tab to option ${i + 1}`);
          await this._screenshot(ssFolder, `option_${i + 1}_tab_failed`);
          break;
        }
        await this._clearAndType(q.options[i]);
        this._log(
          `    Option ${i + 1}: "${q.options[i].substring(0, 40)}${q.options[i].length > 40 ? '...' : ''}"` +
            ` (was: "${fieldInfo.value.substring(0, 30)}")`,
        );
      }
    } else {
      this._log('  True/False question -- options already correct, skipping');
    }

    // ── 4. Mark correct answer ──────────────────────────────────────────

    if (!q.needs_review) {
      this._log(`  Marking correct answer: option ${q.correct_answer_index + 1}`);

      const correctBtns = this.page.locator(
        '[aria-label*="Correct answer" i]:visible, ' +
          '[title*="Correct answer" i]:visible',
      );
      const btnCount = await correctBtns.count();

      if (btnCount > 0 && q.correct_answer_index < btnCount) {
        await correctBtns.nth(q.correct_answer_index).click();
        await wait(500);
        this._log(
          `  Marked option ${q.correct_answer_index + 1} as correct (${btnCount} buttons visible)`,
        );
      } else {
        this._log(
          `  WARNING: Found ${btnCount} correct-answer buttons, need index ${q.correct_answer_index}`,
        );
        await this._screenshot(ssFolder, 'correct_answer_failed');
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

      // Dismiss startup overlays / tooltips
      for (let d = 0; d < 3; d++) {
        await this.page.keyboard.press('Escape');
        await wait(300);
      }
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
    }
    await this._screenshot(ssFolder, '00_title_replaced');

    // 2. Navigate to Self Assessment
    const found = await this._scrollToSelfAssessment();
    if (!found) throw new Error('Self Assessment section not found');
    this._progress({ type: 'section_found', formName });
    await this._screenshot(ssFolder, '01_self_assessment');

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
          this._log(`  WARNING: ${qLabel} may be incomplete`);
        }

        // d. Click outside (on "Self Assessment" header) to commit changes
        const sectionHeader = this.page.locator('text="Self Assessment"').first();
        if (await sectionHeader.isVisible({ timeout: 2000 }).catch(() => false)) {
          await sectionHeader.click();
        } else {
          // Fallback: click on an empty area at the top of the page
          await this.page.mouse.click(700, 150);
        }
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
    this._log('\n  --- Cleaning up sample questions ---');
    this._progress({ type: 'cleanup_start', formName });

    await this._scrollToSelfAssessment();
    await wait(TIMING.shortWait);

    await this._deleteSampleQuestion(SAMPLE_TF);
    await wait(TIMING.shortWait);
    await this._deleteSampleQuestion('Sample True or Flase'); // alternate spelling (old template)
    await wait(TIMING.shortWait);
    await this._deleteSampleQuestion('Sample True or False'); // alternate spelling
    await wait(TIMING.shortWait);
    await this._deleteSampleQuestion(SAMPLE_4OPT);
    await wait(TIMING.shortWait);
    await this._deleteSampleQuestion('Sample four option'); // alternate casing (old template)
    await wait(TIMING.shortWait);

    await this._screenshot(ssFolder, '99_complete');
    this._log('  Form processing complete');
  }

  // ── Process Batch (public) ──────────────────────────────────────────────

  /**
   * Process multiple forms sequentially, then auto-close the browser.
   * @param {Array<{quizData: Object, formUrl: string}>} items
   * @returns {Array<{formName: string, success: boolean, error?: string}>}
   */
  async processBatch(items) {
    const total = items.length;
    this._progress({ type: 'batch_start', total });
    this._log(`=== NEO Quiz Automation: ${total} form(s) ===`);

    // Ensure browser is launched
    if (!this.page) {
      await this.launch();
    }

    const results = [];

    for (let i = 0; i < items.length; i++) {
      const { quizData, formUrl, form_url } = items[i];
      const url = formUrl || form_url;
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
        // Navigate to the form
        if (!url) {
          throw new Error(`No form URL provided for "${formName}"`);
        }

        await this.page.goto(url, {
          waitUntil: 'domcontentloaded',
          timeout: 60000,
        });
        await wait(TIMING.pageLoadWait);
        await this._handleSignIn();
        await wait(TIMING.longWait);

        // Dismiss startup overlays / tooltips
        for (let d = 0; d < 3; d++) {
          await this.page.keyboard.press('Escape');
          await wait(300);
        }
        await wait(TIMING.shortWait);

        await this._screenshot(ssFolder, '00_form_loaded');

        // Core processing
        await this._processFormInternal(quizData, ssFolder);

        this._progress({ type: 'form_done', formName, success: true });
        this._log(`DONE: ${formName}`);
        results.push({ formName, success: true });
      } catch (err) {
        this._progress({ type: 'form_error', formName, error: err.message });
        this._log(`FAILED: ${formName} -- ${err.message}`);
        await this._screenshot(ssFolder, 'FAILED').catch(() => {});
        results.push({ formName, success: false, error: err.message });
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
