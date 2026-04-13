/**
 * NEO Quiz -- Fill Questions Script (v2: Scoped Selectors + Tab Navigation)
 *
 * Fixes from v1:
 *   - Question text replacement scopes to the ACTIVE card (no global text search)
 *   - Option filling uses Tab navigation within the active card (no global indexing)
 *   - Correct answer marking uses :visible selectors (only expanded card matches)
 *   - Does NOT modify points (already set to 1 on samples)
 *   - Deletes both sample questions after all quiz questions are added
 *
 * Usage:
 *   node fill-questions.js --test       # First form in batch.json only
 *   node fill-questions.js              # All forms in batch.json
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

// ─── Configuration ───────────────────────────────────────────────────────────

const CONFIG = {
  quizJsonDir: path.join(__dirname, 'quiz_json'),
  manifestPath: path.join(__dirname, 'quiz_json', 'quiz_manifest.json'),
  batchPath: path.join(__dirname, 'batch.json'),
  screenshotsDir: path.join(__dirname, 'screenshots'),
  authPath: path.join(__dirname, 'auth-state.json'),
  shortWait: 1000,
  mediumWait: 2000,
  longWait: 4000,
  pageLoadWait: 8000,
  typeDelay: 10,
};

const SAMPLE_TF = 'Sample Two Option';
const SAMPLE_4OPT = 'Sample Four Option';

// ─── Utilities ───────────────────────────────────────────────────────────────

function log(msg) {
  console.log(`[${new Date().toLocaleTimeString()}] ${msg}`);
}

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function screenshot(page, folder, name) {
  const dir = path.join(CONFIG.screenshotsDir, folder);
  if (!fs.existsSync(dir)) fs.mkdirSync(dir, { recursive: true });
  const file = `${name}.png`;
  await page.screenshot({ path: path.join(dir, file), fullPage: false });
  log(`  [screenshot] ${folder}/${file}`);
}

function findQuizData(formName) {
  const manifest = JSON.parse(fs.readFileSync(CONFIG.manifestPath, 'utf-8'));
  const entry =
    manifest.forms.find(f => f.form_name === formName) ||
    manifest.forms.find(
      f =>
        f.form_name.toLowerCase().includes(formName.toLowerCase()) ||
        formName.toLowerCase().includes(f.form_name.toLowerCase()),
    );
  if (!entry) return null;
  return JSON.parse(
    fs.readFileSync(path.join(CONFIG.quizJsonDir, entry.filename), 'utf-8'),
  );
}

// ─── Auth ────────────────────────────────────────────────────────────────────

async function handleSignIn(page, saveAuth) {
  const url = page.url();
  if (!url.includes('login') && !url.includes('microsoftonline')) return;
  log('  Sign-in required -- please sign in in the browser...');
  await page.waitForURL(
    u =>
      !u.toString().includes('login') &&
      !u.toString().includes('microsoftonline'),
    { timeout: 180000 },
  );
  await wait(CONFIG.longWait);
  await saveAuth();
  log('  Sign-in complete, auth state saved');
}

// ─── Navigation ──────────────────────────────────────────────────────────────

async function scrollToSelfAssessment(page) {
  log('  Scrolling to Self Assessment section...');
  await page.evaluate(() => window.scrollTo(0, 0));
  await wait(CONFIG.shortWait);

  for (let attempt = 0; attempt < 40; attempt++) {
    const el = page.locator('text="Self Assessment"').first();
    const visible = await el.isVisible({ timeout: 500 }).catch(() => false);
    if (visible) {
      await el.scrollIntoViewIfNeeded();
      await wait(CONFIG.shortWait);
      log('  Found Self Assessment section');
      return true;
    }
    await page.mouse.wheel(0, 400);
    await wait(400);
  }
  log('  ERROR: Could not find Self Assessment section');
  return false;
}

// ─── Tab-Navigation Helpers ──────────────────────────────────────────────────

/**
 * Press Tab until landing on an editable field (input, textarea,
 * contenteditable, or role="textbox"). Returns info about the focused
 * element or null if none found after maxTabs presses.
 */
async function tabToNextEditable(page, maxTabs) {
  if (maxTabs === undefined) maxTabs = 15;
  for (let i = 0; i < maxTabs; i++) {
    await page.keyboard.press('Tab');
    await wait(250);

    const info = await page.evaluate(() => {
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
async function clearAndType(page, text) {
  await page.keyboard.press('Meta+A');
  await wait(150);
  await page.keyboard.type(text, { delay: CONFIG.typeDelay });
  await wait(300);
}

// ─── Form Title Replacement ──────────────────────────────────────────────────

/**
 * Replace the form title ("Sample Title (2 Points)" or similar) with the
 * actual form name from the quiz JSON.  The title is at the top of the page
 * and becomes editable when clicked.
 */
async function replaceFormTitle(page, formName) {
  log('  Replacing form title...');

  // Scroll to top where the title lives
  await page.evaluate(() => window.scrollTo(0, 0));
  await wait(CONFIG.shortWait);

  // Strategy A: find a visible contenteditable or input containing "Sample Title"
  const ceElements = page.locator('[contenteditable="true"]:visible');
  const ceCount = await ceElements.count();
  for (let i = 0; i < ceCount; i++) {
    const text = await ceElements.nth(i).textContent().catch(() => '');
    if (text.includes('Sample Title')) {
      await ceElements.nth(i).click();
      await wait(300);
      await clearAndType(page, formName);
      log(`  Title replaced (contenteditable): "${formName}"`);
      await page.keyboard.press('Escape');
      await wait(CONFIG.shortWait);
      return true;
    }
  }

  // Strategy B: find heading / large text containing "Sample Title"
  const headings = page.locator('h1:visible, h2:visible, [role="heading"]:visible');
  const hCount = await headings.count();
  for (let i = 0; i < hCount; i++) {
    const text = await headings.nth(i).textContent().catch(() => '');
    if (text.includes('Sample Title')) {
      await headings.nth(i).click();
      await wait(500);
      // After clicking the heading, an editable field may appear
      await clearAndType(page, formName);
      log(`  Title replaced (heading click): "${formName}"`);
      await page.keyboard.press('Escape');
      await wait(CONFIG.shortWait);
      return true;
    }
  }

  // Strategy C: look for any visible text containing "Sample Title" (handles "(Copy)" etc.)
  const sampleTitleEl = page.locator(':visible:has-text("Sample Title")').first();
  if (await sampleTitleEl.isVisible({ timeout: 3000 }).catch(() => false)) {
    await sampleTitleEl.click();
    await wait(500);
    await clearAndType(page, formName);
    log(`  Title replaced (text match): "${formName}"`);
    await page.keyboard.press('Escape');
    await wait(CONFIG.shortWait);
    return true;
  }

  // Strategy D: find input with value containing "Sample Title"
  const inputs = page.locator('input:visible, textarea:visible');
  const inputCount = await inputs.count();
  for (let i = 0; i < inputCount; i++) {
    const val = await inputs.nth(i).inputValue().catch(() => '');
    if (val.includes('Sample Title')) {
      await inputs.nth(i).click();
      await wait(300);
      await clearAndType(page, formName);
      log(`  Title replaced (input): "${formName}"`);
      await page.keyboard.press('Escape');
      await wait(CONFIG.shortWait);
      return true;
    }
  }

  log('  WARNING: Could not find form title to replace');
  return false;
}

// ─── Sample-Question Operations ──────────────────────────────────────────────

/**
 * Click a sample question to select it. Scrolls to find it if needed.
 */
async function clickSample(page, sampleName) {
  const el = page.locator(`text="${sampleName}"`).first();

  for (let i = 0; i < 10; i++) {
    if (await el.isVisible({ timeout: 1000 }).catch(() => false)) break;
    await page.mouse.wheel(0, 300);
    await wait(400);
  }

  if (!(await el.isVisible({ timeout: 3000 }).catch(() => false))) {
    throw new Error(`Cannot find sample question: "${sampleName}"`);
  }

  await el.scrollIntoViewIfNeeded();
  await wait(500);
  await el.click();
  await wait(CONFIG.mediumWait);
  log(`  Selected sample: "${sampleName}"`);
}

/**
 * Duplicate the currently selected question.
 * Tries: copy icon -> "..." menu -> keyboard shortcut.
 */
async function duplicateSelected(page) {
  // Strategy 1: Copy / Duplicate icon button
  const iconSelectors = [
    '[aria-label*="Duplicate" i]',
    '[aria-label*="Copy question" i]',
    '[aria-label*="Copy" i]:not([aria-label*="Copilot" i])',
    '[title*="Duplicate" i]',
    '[title*="Copy" i]:not([title*="Copilot" i])',
  ];

  for (const sel of iconSelectors) {
    const btn = page.locator(sel).first();
    if (await btn.isVisible({ timeout: 1500 }).catch(() => false)) {
      await btn.click();
      await wait(CONFIG.mediumWait);
      log('  Duplicated (icon button)');
      return true;
    }
  }

  // Strategy 2: "..." menu -> Duplicate
  const moreBtn = page.locator(
    '[aria-label*="More options" i], [aria-label*="More actions" i]',
  ).first();
  if (await moreBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
    await moreBtn.click();
    await wait(CONFIG.shortWait);
    const dupItem = page.locator(
      '[role="menuitem"]:has-text("Duplicate"), [role="menuitem"]:has-text("Copy")',
    ).first();
    if (await dupItem.isVisible({ timeout: 2000 }).catch(() => false)) {
      await dupItem.click();
      await wait(CONFIG.mediumWait);
      log('  Duplicated (... menu)');
      return true;
    }
    await page.keyboard.press('Escape');
    await wait(500);
  }

  // Strategy 3: Keyboard shortcut
  await page.keyboard.press('Meta+d');
  await wait(CONFIG.mediumWait);
  log('  Duplicated (Cmd+D fallback)');
  return true;
}

// ─── Fill a Duplicated Question ──────────────────────────────────────────────

/**
 * After duplication the duplicate card is auto-selected/expanded.
 * This function:
 *   1. Adjusts option count (add/remove) so Tab order is correct
 *   2. Replaces question text (finds the editable field with sample text)
 *   3. Tabs through option fields and replaces each
 *   4. Marks the correct answer via :visible checkmark buttons
 *
 * Key insight: collapsed cards do NOT show editable inputs or correct-answer
 * buttons, so :visible selectors naturally scope to the active card.
 */
async function fillDuplicatedQuestion(page, questionData, sampleName, ssFolder) {
  const q = questionData;
  const sampleOptCount = sampleName === SAMPLE_TF ? 2 : 4;
  const targetOptCount = q.options.length;

  // True/False questions: the sample already has "True" and "False" as
  // options, so we only need to replace the question text -- skip option
  // filling entirely.
  const isTrueFalse =
    q.options.length === 2 &&
    q.options[0].toLowerCase() === 'true' &&
    q.options[1].toLowerCase() === 'false';

  // ── 1. Adjust option count (multi-option only) ────────────────────────

  if (!isTrueFalse) {
    if (targetOptCount > sampleOptCount) {
      const toAdd = targetOptCount - sampleOptCount;
      log(`  Adding ${toAdd} option(s)...`);
      for (let i = 0; i < toAdd; i++) {
        const addBtn = page.locator(
          'button:has-text("Add option"), :text("+ Add option")',
        ).last();
        if (await addBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
          await addBtn.click();
          await wait(800);
        } else {
          log('  WARNING: Could not find "+ Add option" button');
        }
      }
    } else if (targetOptCount < sampleOptCount) {
      const toRemove = sampleOptCount - targetOptCount;
      log(`  Removing ${toRemove} extra option(s)...`);
      for (let i = 0; i < toRemove; i++) {
        const delBtn = page.locator(
          '[aria-label*="Delete option" i]:visible, ' +
            '[aria-label*="Remove option" i]:visible, ' +
            '[title*="Delete option" i]:visible',
        ).last();
        if (await delBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
          await delBtn.click();
          await wait(800);
        } else {
          log('  WARNING: Could not find option delete button');
        }
      }
    }
  } else {
    log('  True/False question -- skipping option adjustment');
  }

  // ── 2. Replace question text ───────────────────────────────────────────
  //
  // The duplicate's question-text field still contains the sample name.
  // Only the active (expanded) card exposes editable fields, so searching
  // visible editable elements for the sample text targets the duplicate --
  // not the collapsed original.

  log('  Replacing question text...');
  let textReplaced = false;
  const samplePrefix = sampleName.substring(0, 12); // "Sample True " | "Sample four "

  // Approach A: visible contenteditable elements
  const ceElements = page.locator('[contenteditable="true"]:visible');
  const ceCount = await ceElements.count();
  for (let i = ceCount - 1; i >= 0; i--) {
    const text = await ceElements.nth(i).textContent().catch(() => '');
    if (text.includes(samplePrefix)) {
      await ceElements.nth(i).click();
      await wait(300);
      await clearAndType(page, q.text);
      textReplaced = true;
      log(`  Question text replaced (contenteditable #${i})`);
      break;
    }
  }

  // Approach B: visible input / textarea elements
  if (!textReplaced) {
    const inputs = page.locator('input:visible, textarea:visible');
    const inputCount = await inputs.count();
    for (let i = inputCount - 1; i >= 0; i--) {
      const val = await inputs.nth(i).inputValue().catch(() => '');
      if (val.includes(samplePrefix)) {
        await inputs.nth(i).click();
        await wait(300);
        await clearAndType(page, q.text);
        textReplaced = true;
        log(`  Question text replaced (input #${i})`);
        break;
      }
    }
  }

  // Approach C: visible elements with role="textbox"
  if (!textReplaced) {
    const textboxes = page.locator('[role="textbox"]:visible');
    const tbCount = await textboxes.count();
    for (let i = tbCount - 1; i >= 0; i--) {
      const text = await textboxes.nth(i).textContent().catch(() => '');
      if (text.includes(samplePrefix)) {
        await textboxes.nth(i).click();
        await wait(300);
        await clearAndType(page, q.text);
        textReplaced = true;
        log(`  Question text replaced (textbox #${i})`);
        break;
      }
    }
  }

  // Approach D: Tab through editable fields looking for sample text
  if (!textReplaced) {
    log('  Trying Tab-based search for question text field...');
    for (let t = 0; t < 20; t++) {
      const info = await tabToNextEditable(page, 1);
      if (info && info.value.includes(samplePrefix)) {
        await clearAndType(page, q.text);
        textReplaced = true;
        log('  Question text replaced (Tab search)');
        break;
      }
    }
  }

  if (!textReplaced) {
    log('  ERROR: Could not find or replace question text');
    await screenshot(page, ssFolder, 'question_text_failed');
    return false;
  }

  // ── 3. Replace option texts via Tab (multi-option only) ─────────────
  //
  // For True/False: options are already "True" and "False" in the sample,
  // so we skip this step entirely.
  //
  // For multi-option: after replacing the question text the cursor is in
  // the question-text field.  Pressing Tab moves to the first option
  // field, then the next, etc.  tabToNextEditable() skips non-editable
  // focusable elements (toolbar buttons, etc.) between fields.

  if (!isTrueFalse) {
    log(`  Filling ${q.options.length} option(s) via Tab...`);

    for (let i = 0; i < q.options.length; i++) {
      const fieldInfo = await tabToNextEditable(page);
      if (!fieldInfo) {
        log(`  ERROR: Could not Tab to option ${i + 1}`);
        await screenshot(page, ssFolder, `option_${i + 1}_tab_failed`);
        break;
      }
      await clearAndType(page, q.options[i]);
      log(
        `    Option ${i + 1}: "${q.options[i].substring(0, 40)}${q.options[i].length > 40 ? '...' : ''}"` +
          ` (was: "${fieldInfo.value.substring(0, 30)}")`,
      );
    }
  } else {
    log('  True/False question -- options already correct, skipping');
  }

  // ── 4. Mark correct answer ─────────────────────────────────────────────
  //
  // Each option has a circular checkmark icon (aria-label or title
  // containing "Correct answer") that is only visible in the expanded
  // card.  We find all visible ones and click the one at the correct
  // index.

  if (!q.needs_review) {
    log(`  Marking correct answer: option ${q.correct_answer_index + 1}`);

    const correctBtns = page.locator(
      '[aria-label*="Correct answer" i]:visible, ' +
        '[title*="Correct answer" i]:visible',
    );
    const btnCount = await correctBtns.count();

    if (btnCount > 0 && q.correct_answer_index < btnCount) {
      await correctBtns.nth(q.correct_answer_index).click();
      await wait(500);
      log(
        `  Marked option ${q.correct_answer_index + 1} as correct (${btnCount} buttons visible)`,
      );
    } else {
      log(
        `  WARNING: Found ${btnCount} correct-answer buttons, need index ${q.correct_answer_index}`,
      );
      await screenshot(page, ssFolder, 'correct_answer_failed');
    }
  } else {
    log('  Skipping correct answer (needs review)');
  }

  return true;
}

// ─── Delete Sample Questions ─────────────────────────────────────────────────

async function deleteSampleQuestion(page, sampleText) {
  log(`  Deleting sample: "${sampleText}"`);

  const el = page.locator(`text="${sampleText}"`).first();

  // Scroll up to find it (samples are above the newly added questions)
  for (let i = 0; i < 15; i++) {
    if (await el.isVisible({ timeout: 800 }).catch(() => false)) break;
    await page.mouse.wheel(0, -300);
    await wait(400);
  }

  if (!(await el.isVisible({ timeout: 3000 }).catch(() => false))) {
    log(`  "${sampleText}" not found (may already be deleted or scrolled away)`);
    return;
  }

  await el.scrollIntoViewIfNeeded();
  await wait(500);
  await el.click();
  await wait(CONFIG.mediumWait);

  // Find the delete button in the question toolbar
  const delSelectors = [
    '[aria-label*="Delete question" i]',
    '[aria-label*="Delete" i]:not([aria-label*="option" i])',
    '[title*="Delete question" i]',
    '[title*="Delete" i]:not([title*="option" i])',
  ];

  for (const sel of delSelectors) {
    const btn = page.locator(sel).first();
    if (await btn.isVisible({ timeout: 2000 }).catch(() => false)) {
      await btn.click();
      await wait(CONFIG.shortWait);

      // Handle confirmation dialog if it appears
      const confirmBtn = page.locator(
        'button:has-text("OK"), button:has-text("Yes"), button:has-text("Delete")',
      ).first();
      if (await confirmBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
        await confirmBtn.click();
        await wait(CONFIG.shortWait);
      }

      log(`  Deleted "${sampleText}"`);
      return;
    }
  }

  // Fallback: "..." menu -> Delete
  const moreBtn = page.locator(
    '[aria-label*="More" i]:not([aria-label*="Learn more" i])',
  ).first();
  if (await moreBtn.isVisible({ timeout: 1500 }).catch(() => false)) {
    await moreBtn.click();
    await wait(CONFIG.shortWait);
    const delItem = page.locator(
      '[role="menuitem"]:has-text("Delete")',
    ).first();
    if (await delItem.isVisible({ timeout: 2000 }).catch(() => false)) {
      await delItem.click();
      await wait(CONFIG.shortWait);

      const confirmBtn = page.locator(
        'button:has-text("OK"), button:has-text("Yes"), button:has-text("Delete")',
      ).first();
      if (await confirmBtn.isVisible({ timeout: 2000 }).catch(() => false)) {
        await confirmBtn.click();
        await wait(CONFIG.shortWait);
      }

      log(`  Deleted "${sampleText}" (via ... menu)`);
      return;
    }
    await page.keyboard.press('Escape');
    await wait(300);
  }

  log(`  WARNING: Could not delete "${sampleText}"`);
}

// ─── Process One Form ────────────────────────────────────────────────────────

async function processForm(page, quizData, ssFolder) {
  log(`Processing: ${quizData.form_name} (${quizData.questions.length} questions)`);

  // 1. Replace the form title
  await replaceFormTitle(page, quizData.form_name);
  await screenshot(page, ssFolder, '00_title_replaced');

  // 2. Navigate to Self Assessment
  const found = await scrollToSelfAssessment(page);
  if (!found) throw new Error('Self Assessment section not found');
  await screenshot(page, ssFolder, '01_self_assessment');

  // 3. Duplicate + fill each question
  let firstFourOptDone = false;
  for (let i = 0; i < quizData.questions.length; i++) {
    const q = quizData.questions[i];
    const qLabel = `Q${String(i + 1).padStart(2, '0')}`;
    const sampleName = q.options.length <= 2 ? SAMPLE_TF : SAMPLE_4OPT;

    log(
      `\n  --- ${qLabel}/${quizData.questions.length}: ` +
        `"${q.text.substring(0, 50)}..." (${q.options.length} opts -> "${sampleName}") ---`,
    );

    try {
      // a. Click the sample to select it
      await clickSample(page, sampleName);
      await screenshot(page, ssFolder, `${qLabel}_a_selected`);

      // b. Duplicate it
      const duped = await duplicateSelected(page);
      if (!duped) {
        log(`  ERROR: Duplication failed for ${qLabel}`);
        await screenshot(page, ssFolder, `${qLabel}_b_dup_failed`);
        continue;
      }
      await screenshot(page, ssFolder, `${qLabel}_b_duplicated`);

      // c. Fill the duplicate with actual quiz data
      const filled = await fillDuplicatedQuestion(page, q, sampleName, ssFolder);
      await screenshot(page, ssFolder, `${qLabel}_c_filled`);

      if (!filled) {
        log(`  WARNING: ${qLabel} may be incomplete`);
      }

      // d. Deselect the question by clicking outside the card
      //    (clicking outside commits changes; Escape may cancel them)
      const sectionHeader = page.locator('text="Self Assessment"').first();
      if (await sectionHeader.isVisible({ timeout: 2000 }).catch(() => false)) {
        await sectionHeader.click();
      } else {
        // Fallback: click on an empty area at the top of the page
        await page.mouse.click(700, 150);
      }
      await wait(CONFIG.mediumWait);

      // e. Extra wait for the first 4-option question to ensure auto-save
      if (sampleName === SAMPLE_4OPT && !firstFourOptDone) {
        log('  First 4-option question -- extra wait for auto-save');
        await wait(CONFIG.longWait);
        firstFourOptDone = true;
      }

      // f. Scroll down slightly so next duplicate appears in view
      await page.mouse.wheel(0, 200);
      await wait(500);

      log(`  ${qLabel} complete`);
    } catch (err) {
      log(`  ERROR on ${qLabel}: ${err.message}`);
      await screenshot(page, ssFolder, `${qLabel}_error`);
      await page.mouse.click(700, 150).catch(() => {});
      await wait(CONFIG.shortWait);
    }
  }

  // 4. Delete the two sample questions
  log('\n  --- Cleaning up sample questions ---');
  await scrollToSelfAssessment(page);
  await wait(CONFIG.shortWait);

  await deleteSampleQuestion(page, SAMPLE_TF);
  await wait(CONFIG.shortWait);
  await deleteSampleQuestion(page, 'Sample True or Flase'); // alternate spelling (old template)
  await wait(CONFIG.shortWait);
  await deleteSampleQuestion(page, 'Sample True or False'); // alternate spelling
  await wait(CONFIG.shortWait);
  await deleteSampleQuestion(page, SAMPLE_4OPT);
  await wait(CONFIG.shortWait);
  await deleteSampleQuestion(page, 'Sample four option'); // alternate casing (old template)
  await wait(CONFIG.shortWait);

  await screenshot(page, ssFolder, '99_complete');
  log('  Form processing complete');
}

// ─── Main ────────────────────────────────────────────────────────────────────

async function main() {
  const args = process.argv.slice(2);
  const testMode = args.includes('--test');

  const batch = JSON.parse(fs.readFileSync(CONFIG.batchPath, 'utf-8'));
  const toProcess = testMode ? batch.slice(0, 1) : batch;

  log('=== NEO Quiz Fill Questions (v2) ===');
  log(
    `Mode: ${testMode ? 'TEST (first form only)' : `FULL (${toProcess.length} forms)`}`,
  );

  // Validate all entries up front
  const validated = [];
  for (const entry of toProcess) {
    const qd = findQuizData(entry.form_name);
    if (!qd) {
      log(`ERROR: No quiz data found for "${entry.form_name}"`);
      process.exit(1);
    }
    if (!entry.form_url || entry.form_url.includes('PASTE')) {
      log(`ERROR: Invalid URL for "${entry.form_name}"`);
      process.exit(1);
    }
    validated.push({ entry, quizData: qd });
    log(`  "${entry.form_name}" -> ${qd.question_count} questions`);
  }

  // Launch browser
  log('Launching browser...');
  const browser = await chromium.launch({
    headless: false,
    slowMo: 100,
    args: ['--start-maximized'],
  });

  const hasAuth =
    fs.existsSync(CONFIG.authPath) && fs.statSync(CONFIG.authPath).size > 10;
  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
    storageState: hasAuth ? CONFIG.authPath : undefined,
  });

  const saveAuth = async () => {
    try {
      await context.storageState({ path: CONFIG.authPath });
    } catch (_) {
      /* ignore */
    }
  };

  const page = await context.newPage();

  for (let i = 0; i < validated.length; i++) {
    const { entry, quizData } = validated[i];
    const ssFolder = `form_${String(i + 1).padStart(2, '0')}_${entry.form_name.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 40)}`;

    log(`\n${'='.repeat(60)}`);
    log(`[${i + 1}/${validated.length}] ${entry.form_name}`);
    log('='.repeat(60));

    try {
      await page.goto(entry.form_url, {
        waitUntil: 'domcontentloaded',
        timeout: 60000,
      });
      await wait(CONFIG.pageLoadWait);
      await handleSignIn(page, saveAuth);
      await wait(CONFIG.longWait);

      // Dismiss any startup overlays / tooltips
      for (let d = 0; d < 3; d++) {
        await page.keyboard.press('Escape');
        await wait(300);
      }
      await wait(CONFIG.shortWait);

      await screenshot(page, ssFolder, '00_form_loaded');

      await processForm(page, quizData, ssFolder);

      log(`DONE: ${entry.form_name}`);
    } catch (err) {
      log(`FAILED: ${entry.form_name} -- ${err.message}`);
      await screenshot(page, ssFolder, 'FAILED').catch(() => {});
    }
  }

  log(`\n${'='.repeat(60)}`);
  log('ALL FORMS COMPLETE');
  log('='.repeat(60));

  await wait(3000);
  await browser.close();
  log('Browser closed. Done.');
}

main().catch(err => {
  console.error('Fatal error:', err);
  process.exit(1);
});
