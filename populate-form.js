/**
 * NEO Quiz — Populate Form Script v2
 *
 * Workflow:
 *   1. You manually duplicate the template form and paste the edit URL into batch.json
 *   2. Run: node populate-form.js --test    (one form)
 *      Run: node populate-form.js           (all forms in batch.json)
 *
 * batch.json format:
 *   [{ "form_name": "F2S - S&OP RESA & Advanced ATP", "form_url": "https://..." }]
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const CONFIG = {
  quizJsonDir: path.join(__dirname, 'quiz_json'),
  manifestPath: path.join(__dirname, 'quiz_json', 'quiz_manifest.json'),
  batchPath: path.join(__dirname, 'batch.json'),
  resultsPath: path.join(__dirname, 'results.json'),
  screenshotsDir: path.join(__dirname, 'screenshots'),
  authPath: path.join(__dirname, 'auth-state.json'),

  // Direct link to Forms Tracker sheet
  excelUrl: 'https://oxygy.sharepoint.com/:x:/s/OXYGY_General-AInewsletterdesk/IQAVzcP6X111QZOXiBdO9u7tATJVKdHjF1z2C2pE9U4DcsA?e=ZnDqsg&nav=MTVfezAwMDAwMDAwLTAwMDEtMDAwMC0wMTAwLTAwMDAwMDAwMDAwMH0',

  shortWait: 1500,
  mediumWait: 3000,
  longWait: 5000,
  pageLoadWait: 10000,
};

// ─── Helpers ──────────────────────────────────────────────────────────────────
function log(msg) {
  console.log(`[${new Date().toLocaleTimeString()}] ${msg}`);
}

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function ss(page, name) {
  if (!fs.existsSync(CONFIG.screenshotsDir)) fs.mkdirSync(CONFIG.screenshotsDir, { recursive: true });
  const filename = `${name}_${Date.now()}.png`;
  await page.screenshot({ path: path.join(CONFIG.screenshotsDir, filename), fullPage: false });
  log(`  [ss] ${filename}`);
}

async function handleSignIn(page, saveAuth, timeout = 180000) {
  const url = page.url();
  if (!url.includes('login') && !url.includes('microsoftonline')) return;
  log('  Sign-in required — please sign in in the browser window...');
  await page.waitForURL(
    u => !u.toString().includes('login') && !u.toString().includes('microsoftonline'),
    { timeout }
  );
  await wait(CONFIG.longWait);
  await saveAuth();
  log('  Sign-in complete');
}

function findQuizData(formName) {
  const manifest = JSON.parse(fs.readFileSync(CONFIG.manifestPath, 'utf-8'));
  const entry = manifest.forms.find(f => f.form_name === formName);
  if (!entry) {
    const partial = manifest.forms.find(f =>
      f.form_name.toLowerCase().includes(formName.toLowerCase()) ||
      formName.toLowerCase().includes(f.form_name.toLowerCase())
    );
    if (partial) {
      log(`  Partial match: "${partial.form_name}"`);
      return JSON.parse(fs.readFileSync(path.join(CONFIG.quizJsonDir, partial.filename), 'utf-8'));
    }
    return null;
  }
  return JSON.parse(fs.readFileSync(path.join(CONFIG.quizJsonDir, entry.filename), 'utf-8'));
}

// Click on a neutral/empty area to deselect anything
async function clickNeutralArea(page) {
  // Click on the far-right empty space of the editor, above the content
  try {
    await page.mouse.click(100, 500);
    await wait(500);
    await page.keyboard.press('Escape');
    await wait(500);
    await page.keyboard.press('Escape');
    await wait(500);
  } catch (e) { /* */ }
}

// ─── STEP 1: Open the form editor ────────────────────────────────────────────
async function openFormEditor(page, formUrl, saveAuth) {
  log('  STEP 1: Opening form editor...');
  await page.goto(formUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await wait(CONFIG.pageLoadWait);
  await handleSignIn(page, saveAuth);
  await wait(CONFIG.longWait);

  log(`  URL: ${page.url()}`);

  // Close any banners/panels (Copilot suggestions, Collaboration panel, etc.)
  for (let attempt = 0; attempt < 3; attempt++) {
    try {
      const closeBtns = page.locator('[aria-label="Close"], [aria-label="Dismiss"]');
      for (let i = 0; i < await closeBtns.count(); i++) {
        try {
          if (await closeBtns.nth(i).isVisible({ timeout: 800 })) {
            await closeBtns.nth(i).click();
            await wait(400);
          }
        } catch (e) { /* */ }
      }
    } catch (e) { /* */ }
    await wait(300);
  }

  // Also try closing the Copilot "Scanning/Checking" banner by clicking its X
  try {
    const bannerClose = page.locator('button:near(:text("Copilot"), 200)').first();
    // Just press Escape a couple times to dismiss any overlays
    await page.keyboard.press('Escape');
    await wait(300);
    await page.keyboard.press('Escape');
    await wait(300);
  } catch (e) { /* */ }

  await wait(CONFIG.shortWait);
  await ss(page, '01_form_opened');
  log('  STEP 1: Form editor loaded');
}

// ─── STEP 2: Rename the form ──────────────────────────────────────────────────
async function renameForm(page, newName) {
  log(`  STEP 2: Renaming form to "${newName}"`);

  // Scroll to top first
  await page.evaluate(() => window.scrollTo(0, 0));
  await wait(CONFIG.shortWait);

  // The title text is in the left preview panel. Double-click to edit it.
  let renamed = false;

  // Approach 1: Double-click the title text in the preview
  try {
    const titleText = page.locator('text=/LUSA NEO.*Feedback Form|End User Training/').first();
    if (await titleText.isVisible({ timeout: 3000 })) {
      await titleText.dblclick();
      await wait(CONFIG.shortWait);

      // Select all and replace
      await page.keyboard.press('Meta+A');
      await wait(200);
      await page.keyboard.type(newName, { delay: 15 });
      await wait(CONFIG.shortWait);

      // CRITICAL: Click outside the title to deselect it completely
      await page.keyboard.press('Escape');
      await wait(500);
      await clickNeutralArea(page);
      await wait(CONFIG.mediumWait);

      renamed = true;
      log('  Title renamed via preview panel');
    }
  } catch (e) {
    log(`  Preview rename failed: ${e.message}`);
  }

  // Approach 2: Click the header bar title
  if (!renamed) {
    try {
      const headerText = page.locator('text=/Feedback Form.*Copy/').first();
      if (await headerText.isVisible({ timeout: 3000 })) {
        await headerText.click();
        await wait(CONFIG.shortWait);
        await page.keyboard.press('Meta+A');
        await wait(200);
        await page.keyboard.type(newName, { delay: 15 });
        await wait(CONFIG.shortWait);

        // Deselect
        await page.keyboard.press('Escape');
        await wait(500);
        await clickNeutralArea(page);
        await wait(CONFIG.mediumWait);

        renamed = true;
        log('  Title renamed via header');
      }
    } catch (e) {
      log(`  Header rename failed: ${e.message}`);
    }
  }

  if (!renamed) {
    log('  WARNING: Could not rename title');
  }

  // Verify title is deselected — click neutral area again
  await clickNeutralArea(page);
  await wait(CONFIG.shortWait);

  await ss(page, '02_after_rename');
  return renamed;
}

// ─── STEP 3: Add questions to Self Assessment section ─────────────────────────
async function addQuestions(page, questions) {
  log(`  STEP 3: Adding ${questions.length} questions...`);
  const reviewItems = [];

  // 3a: Scroll to find "Self Assessment" section header
  log('  Scrolling to Self Assessment section...');
  let found = false;

  for (let attempt = 0; attempt < 30; attempt++) {
    const saVisible = await page.locator('text="Self Assessment"').first()
      .isVisible({ timeout: 600 }).catch(() => false);
    if (saVisible) {
      found = true;
      break;
    }
    await page.mouse.wheel(0, 350);
    await wait(500);
  }

  if (!found) {
    await ss(page, 'no_self_assessment');
    throw new Error('Could not find Self Assessment section');
  }

  // Scroll the Self Assessment header into view
  await page.locator('text="Self Assessment"').first().scrollIntoViewIfNeeded();
  await wait(CONFIG.shortWait);
  log('  Found Self Assessment section');
  await ss(page, '03_self_assessment_found');

  // 3b: Now scroll down a bit MORE to find the "Insert new question" button
  // that is positioned AFTER the Self Assessment section and BEFORE Section 4.
  // We need to scroll past the Self Assessment content to see this button.
  await page.mouse.wheel(0, 400);
  await wait(CONFIG.shortWait);

  // Now find the section boundary. We need the "Insert new question" that's
  // NEAR "Section 4" or near the end of Self Assessment content.
  // Strategy: find ALL "Insert new question" buttons, then pick the one
  // that's closest to (just above) "Section 4"

  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    const qLabel = `Q${i + 1}/${questions.length}`;
    log(`  ${qLabel}: "${q.text.substring(0, 55)}..."`);

    try {
      // Find the correct "Insert new question" button.
      // We want the one that appears near/after Self Assessment.
      // After the first question is added, new "Insert new question" buttons
      // will appear below each added question.
      //
      // Strategy: get all "Insert new question" elements, find the one
      // closest to the bottom of Self Assessment / before Section 4

      await wait(CONFIG.shortWait);

      // Scroll to make sure we can see the insert button near Self Assessment
      // On first question, scroll to just below Self Assessment
      // On subsequent questions, the latest "Insert new question" should be visible

      const insertBtns = page.locator('text="Insert new question"');
      const btnCount = await insertBtns.count();
      log(`  Found ${btnCount} "Insert new question" buttons`);

      if (btnCount === 0) {
        log(`  ERROR: No "Insert new question" button found for ${qLabel}`);
        await ss(page, `no_insert_btn_${i + 1}`);
        break;
      }

      // Click the LAST visible "Insert new question" button.
      // Reasoning: the buttons appear between sections. The last one visible
      // after scrolling to Self Assessment should be the one at the bottom
      // of the Self Assessment section, right where we want new questions.
      let clicked = false;
      for (let b = btnCount - 1; b >= 0; b--) {
        try {
          const btn = insertBtns.nth(b);
          if (await btn.isVisible({ timeout: 1000 })) {
            await btn.scrollIntoViewIfNeeded();
            await wait(300);
            await btn.click();
            clicked = true;
            log(`  Clicked "Insert new question" button #${b} (of ${btnCount})`);
            break;
          }
        } catch (e) { /* try next */ }
      }

      if (!clicked) {
        // Scroll down more and try again
        await page.mouse.wheel(0, 400);
        await wait(1000);
        const lastBtn = page.locator('text="Insert new question"').last();
        if (await lastBtn.isVisible({ timeout: 3000 })) {
          await lastBtn.click();
          clicked = true;
          log('  Clicked last "Insert new question" after scrolling');
        }
      }

      if (!clicked) {
        // Try "Add new question" as alternative text
        const altBtn = page.locator('text="Add new question"').last();
        if (await altBtn.isVisible({ timeout: 2000 })) {
          await altBtn.click();
          clicked = true;
          log('  Clicked "Add new question"');
        }
      }

      if (!clicked) {
        log(`  ERROR: Could not click insert button for ${qLabel}`);
        await ss(page, `insert_failed_${i + 1}`);
        continue;
      }

      await wait(CONFIG.mediumWait);
      await ss(page, `04_type_picker_q${i + 1}`);

      // 3c: Click "Choice" from the question type picker
      // The type picker shows: Choice, Text, Rating, Date, Ranking, Likert, etc.
      log('  Clicking "Choice" question type...');
      const choiceBtn = page.locator('text="Choice"').first();
      if (await choiceBtn.isVisible({ timeout: 5000 })) {
        await choiceBtn.click();
        await wait(CONFIG.mediumWait);
        log('  Selected "Choice" type');
      } else {
        log('  WARNING: Could not find "Choice" button');
        await ss(page, `no_choice_btn_${i + 1}`);
        continue;
      }

      await ss(page, `05_choice_selected_q${i + 1}`);

      // 3d: Type the question text
      // After selecting "Choice", a new question card should appear with:
      //   - A question text area (placeholder like "Question" or empty)
      //   - Option 1, Option 2 fields
      // The question text area should be focused or be the first editable area
      // inside the newly created question.
      log('  Typing question text...');

      // Try to find the question text input specifically
      // Look for the most recently created question's text field
      let questionTextTyped = false;

      // Approach 1: Find input with "Question" placeholder
      const qInputs = page.locator('[placeholder*="Question" i]');
      const qInputCount = await qInputs.count();
      if (qInputCount > 0) {
        // Click the LAST one (most recently added question)
        const qInput = qInputs.last();
        if (await qInput.isVisible({ timeout: 3000 })) {
          await qInput.click();
          await wait(300);
          await page.keyboard.type(q.text, { delay: 8 });
          await wait(CONFIG.shortWait);
          questionTextTyped = true;
          log('  Typed question via placeholder input');
        }
      }

      // Approach 2: Look for aria-label containing "Question"
      if (!questionTextTyped) {
        const qAria = page.locator('[aria-label*="Question" i]:not([aria-label*="type" i]):not([aria-label*="option" i])');
        const ariaCount = await qAria.count();
        if (ariaCount > 0) {
          const lastQ = qAria.last();
          if (await lastQ.isVisible({ timeout: 2000 })) {
            await lastQ.click();
            await wait(300);
            await page.keyboard.type(q.text, { delay: 8 });
            await wait(CONFIG.shortWait);
            questionTextTyped = true;
            log('  Typed question via aria-label input');
          }
        }
      }

      // Approach 3: Look for contenteditable div that's empty/new
      if (!questionTextTyped) {
        const editables = page.locator('[contenteditable="true"]:visible');
        const editCount = await editables.count();
        if (editCount > 0) {
          // The last contenteditable is likely the new question
          const lastEdit = editables.last();
          await lastEdit.click();
          await wait(300);
          await page.keyboard.type(q.text, { delay: 8 });
          await wait(CONFIG.shortWait);
          questionTextTyped = true;
          log('  Typed question via last contenteditable');
        }
      }

      if (!questionTextTyped) {
        log('  WARNING: Could not find question text input, trying Tab...');
        // Press Tab to move to the question field
        await page.keyboard.press('Tab');
        await wait(500);
        await page.keyboard.type(q.text, { delay: 8 });
        await wait(CONFIG.shortWait);
      }

      await ss(page, `06_question_text_q${i + 1}`);

      // 3e: Add options
      log(`  Adding ${q.options.length} options...`);

      // Find option input fields
      const optionInputs = page.locator('[placeholder*="Option" i], [aria-label*="Option" i]');
      let optCount = await optionInputs.count();
      log(`  Found ${optCount} option fields on page`);

      // We need to fill the options for THIS question, not previous ones.
      // Since we're adding questions sequentially, the options for the newest
      // question will be at the END of the list.
      // Default new Choice question has "Option 1" and "Option 2" fields.

      // First, add more option fields if we need more than 2
      let addAttempts = 0;
      // Count how many we need to add
      const baseOptionCount = 2; // Default options for a new Choice question
      const extraNeeded = Math.max(0, q.options.length - baseOptionCount);

      for (let a = 0; a < extraNeeded && addAttempts < 15; a++) {
        try {
          const addOptBtns = page.locator('text=/[Aa]dd option/');
          const addCount = await addOptBtns.count();
          if (addCount > 0) {
            // Click the last "Add option" (belongs to the latest question)
            const lastAdd = addOptBtns.last();
            if (await lastAdd.isVisible({ timeout: 2000 })) {
              await lastAdd.click();
              await wait(800);
              addAttempts++;
            } else {
              break;
            }
          } else {
            break;
          }
        } catch (e) {
          break;
        }
      }

      // Re-count option fields
      const allOptions = page.locator('[placeholder*="Option" i], [aria-label*="Option" i]');
      const totalOptCount = await allOptions.count();
      log(`  Total option fields on page: ${totalOptCount}`);

      // Fill in the options for THIS question.
      // They are the LAST N option fields on the page.
      const startIdx = totalOptCount - Math.max(q.options.length, baseOptionCount + addAttempts);
      const optStartIdx = Math.max(0, totalOptCount - q.options.length - (addAttempts > 0 ? 0 : (baseOptionCount - q.options.length)));

      // Simpler approach: fill the last q.options.length option fields
      const fillStart = totalOptCount - Math.max(q.options.length, baseOptionCount);
      log(`  Filling options starting from index ${Math.max(0, fillStart)}`);

      for (let j = 0; j < q.options.length; j++) {
        const optIdx = Math.max(0, totalOptCount - q.options.length) + j;
        // But we may have more slots than options if we over-added; try from the end
        const actualIdx = totalOptCount - q.options.length + j;

        if (actualIdx < 0 || actualIdx >= totalOptCount) {
          log(`  Warning: Option index ${actualIdx} out of range`);
          continue;
        }

        try {
          const opt = allOptions.nth(actualIdx);
          if (await opt.isVisible({ timeout: 2000 })) {
            await opt.click();
            await wait(200);
            // Select any existing text
            await page.keyboard.press('Meta+A');
            await wait(100);
            // Type the option text
            await page.keyboard.type(q.options[j], { delay: 8 });
            await wait(400);
            log(`  Option ${j + 1}: "${q.options[j].substring(0, 40)}"`);
          } else {
            log(`  Warning: Option field ${actualIdx} not visible`);
          }
        } catch (e) {
          log(`  Warning: Error filling option ${j + 1}: ${e.message}`);
        }
      }

      // 3f: Mark correct answer
      if (q.needs_review) {
        log('  Skipping correct answer (needs review)');
        reviewItems.push({ questionNumber: q.number, text: q.text.substring(0, 60) });
      } else {
        try {
          // Find correct answer toggles — these are typically near each option
          const correctBtns = page.locator('[aria-label*="correct" i], [title*="correct" i]');
          const correctCount = await correctBtns.count();
          log(`  Found ${correctCount} correct-answer toggles`);

          if (correctCount > 0) {
            // The toggles for THIS question are the last ones on the page
            const toggleIdx = correctCount - q.options.length + q.correct_answer_index;
            if (toggleIdx >= 0 && toggleIdx < correctCount) {
              await correctBtns.nth(toggleIdx).click();
              await wait(500);
              log(`  Marked option ${q.correct_answer_index + 1} as correct`);
            }
          }
        } catch (e) {
          log(`  Warning: Could not mark correct answer: ${e.message}`);
        }
      }

      // 3g: Deselect the question — click neutral area
      await clickNeutralArea(page);
      await wait(CONFIG.shortWait);

      // Scroll down to keep the "Insert new question" visible for next question
      await page.mouse.wheel(0, 250);
      await wait(500);

      await ss(page, `07_question_done_q${i + 1}`);
      log(`  ${qLabel} done`);

    } catch (err) {
      log(`  ERROR on ${qLabel}: ${err.message}`);
      await ss(page, `error_q${i + 1}`);
      await clickNeutralArea(page);
      await wait(CONFIG.shortWait);
    }
  }

  await ss(page, '08_all_questions_done');
  log(`  STEP 3: Finished adding ${questions.length} questions`);
  return reviewItems;
}

// ─── STEP 4: Update Excel Tracker ────────────────────────────────────────────
async function updateExcelTracker(excelPage, formIndex, saveAuth) {
  log('  STEP 4: Updating Excel tracker...');
  await excelPage.bringToFront();
  await wait(CONFIG.mediumWait);

  // The Excel URL already points directly to the Forms Tracker sheet.
  // formIndex is 1-based, header is row 1, so data row = formIndex + 1
  const targetRow = formIndex + 1;

  log(`  Target: F${targetRow} = Complete, G${targetRow} = today's date`);

  try {
    // First, try to use the Name Box to navigate to F{row}
    // The Name Box is the cell address input at top-left of the Excel editor
    const nameBoxSelectors = [
      '#NameBox',
      '[id*="NameBox"]',
      '[id*="namebox"]',
      '[aria-label*="Name Box" i]',
      'input[aria-label*="cell" i]',
      '#m_excelWebRenderer_ewaCtl_NameBox',
    ];

    let nameBoxFound = false;
    for (const sel of nameBoxSelectors) {
      try {
        const nb = excelPage.locator(sel).first();
        if (await nb.isVisible({ timeout: 2000 })) {
          // Click the Name Box
          await nb.click();
          await wait(500);
          // Select existing text
          await excelPage.keyboard.press('Meta+A');
          await wait(200);
          // Type cell address
          await excelPage.keyboard.type(`F${targetRow}`, { delay: 30 });
          await excelPage.keyboard.press('Enter');
          await wait(CONFIG.shortWait);
          nameBoxFound = true;
          log(`  Navigated to F${targetRow} via Name Box (${sel})`);
          break;
        }
      } catch (e) { /* try next */ }
    }

    if (!nameBoxFound) {
      log('  Name Box not found, using Ctrl+Home + arrow keys...');
      // Navigate from A1
      await excelPage.keyboard.press('Control+Home');
      await wait(CONFIG.shortWait);

      // Go to column F (press Right 5 times from A)
      for (let c = 0; c < 5; c++) {
        await excelPage.keyboard.press('ArrowRight');
        await wait(80);
      }
      // Go to target row (press Down from row 1)
      for (let r = 1; r < targetRow; r++) {
        await excelPage.keyboard.press('ArrowDown');
        await wait(80);
      }
      await wait(CONFIG.shortWait);
      log(`  Navigated to F${targetRow} via arrow keys`);
    }

    // Type "Complete" in column F
    await excelPage.keyboard.type('Complete', { delay: 15 });
    await excelPage.keyboard.press('Tab'); // Move to column G
    await wait(500);

    // Type today's date in column G
    const today = new Date().toLocaleDateString('en-US');
    await excelPage.keyboard.type(today, { delay: 15 });
    await excelPage.keyboard.press('Enter');
    await wait(CONFIG.shortWait);

    log(`  Updated: F${targetRow}=Complete, G${targetRow}=${today}`);
  } catch (e) {
    log(`  ERROR updating Excel: ${e.message}`);
    await ss(excelPage, 'excel_error');
  }

  await ss(excelPage, '09_excel_updated');
  log('  STEP 4: Excel done');
}

// ─── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  const args = process.argv.slice(2);
  const testMode = args.includes('--test');

  // Read batch
  if (!fs.existsSync(CONFIG.batchPath)) {
    log('ERROR: batch.json not found.');
    process.exit(1);
  }

  const batch = JSON.parse(fs.readFileSync(CONFIG.batchPath, 'utf-8'));
  const toProcess = testMode ? batch.slice(0, 1) : batch;

  log(`Batch: ${toProcess.length} form(s)${testMode ? ' (TEST MODE)' : ''}`);

  const manifest = JSON.parse(fs.readFileSync(CONFIG.manifestPath, 'utf-8'));

  // Validate
  for (const entry of toProcess) {
    const quizData = findQuizData(entry.form_name);
    if (!quizData) {
      log(`ERROR: No quiz data for "${entry.form_name}"`);
      process.exit(1);
    }
    if (!entry.form_url || entry.form_url.includes('PASTE')) {
      log(`ERROR: Missing form URL for "${entry.form_name}"`);
      process.exit(1);
    }
    log(`  "${entry.form_name}" → ${quizData.question_count} questions`);
  }

  // Launch browser
  log('Launching browser...');
  const browser = await chromium.launch({
    headless: false,
    slowMo: 150,
    args: ['--start-maximized'],
  });

  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
    storageState: fs.existsSync(CONFIG.authPath) && fs.statSync(CONFIG.authPath).size > 10
      ? CONFIG.authPath : undefined,
  });

  const saveAuth = async () => {
    try { await context.storageState({ path: CONFIG.authPath }); } catch (e) { /* */ }
  };

  const results = [];

  try {
    // Open Excel (direct link to Forms Tracker sheet)
    log('Opening Excel tracker (Forms Tracker sheet)...');
    const excelPage = await context.newPage();
    await excelPage.goto(CONFIG.excelUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await wait(CONFIG.pageLoadWait);
    await handleSignIn(excelPage, saveAuth);
    await wait(CONFIG.longWait);
    await saveAuth();
    await ss(excelPage, '00_excel_loaded');
    log('Excel ready');

    // Open form page
    const formPage = await context.newPage();

    for (let i = 0; i < toProcess.length; i++) {
      const entry = toProcess[i];
      const quizData = findQuizData(entry.form_name);
      const manifestEntry = manifest.forms.find(f => f.form_name === quizData.form_name);
      const formIndex = manifestEntry ? manifestEntry.index : i + 1;

      log('');
      log(`======== [${i + 1}/${toProcess.length}] ${entry.form_name} (${quizData.question_count}q) ========`);

      try {
        // Step 1: Open form
        await formPage.bringToFront();
        await openFormEditor(formPage, entry.form_url, saveAuth);

        // Step 2: Rename
        const renamed = await renameForm(formPage, quizData.form_name);

        // Step 3: Add questions
        const reviewItems = await addQuestions(formPage, quizData.questions);

        // Step 4: Update Excel
        await updateExcelTracker(excelPage, formIndex, saveAuth);

        results.push({
          form_name: entry.form_name,
          status: 'completed',
          renamed,
          questions_added: quizData.question_count,
          needs_review: reviewItems,
        });

        log(`  DONE: ${entry.form_name}`);

      } catch (err) {
        log(`  FAILED: ${err.message}`);
        await ss(formPage, `failed_${i + 1}`);
        results.push({
          form_name: entry.form_name,
          status: 'failed',
          error: err.message,
        });
      }
    }

    // Save results
    fs.writeFileSync(CONFIG.resultsPath, JSON.stringify(results, null, 2));

    log('');
    log('======== COMPLETE ========');
    const ok = results.filter(r => r.status === 'completed').length;
    const fail = results.filter(r => r.status === 'failed').length;
    log(`OK: ${ok}, Failed: ${fail}`);
    results.forEach(r => log(`  ${r.status === 'completed' ? 'OK' : 'FAIL'}: ${r.form_name}${r.error ? ' — ' + r.error : ''}`));

    log('\nCtrl+C to close browser.');
    await new Promise(() => {});

  } catch (err) {
    log(`FATAL: ${err.message}`);
    console.error(err);
    fs.writeFileSync(CONFIG.resultsPath, JSON.stringify(results, null, 2));
    await browser.close();
    process.exit(1);
  }
}

main().catch(console.error);
