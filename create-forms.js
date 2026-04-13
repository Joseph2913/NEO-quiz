/**
 * NEO Quiz Forms Automation Script v3
 *
 * Key fixes:
 * - Handles "Duplicate it" opening a NEW TAB (listens for popup/new page)
 * - Navigates to Self Assessment section by scrolling the form editor
 * - Switches Excel to "Forms Tracker" sheet tab before updating
 * - Uses Name Box for direct cell navigation in Excel
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const CONFIG = {
  quizJsonDir: path.join(__dirname, 'quiz_json'),
  manifestPath: path.join(__dirname, 'quiz_json', 'quiz_manifest.json'),
  progressPath: path.join(__dirname, 'progress.json'),
  screenshotsDir: path.join(__dirname, 'screenshots'),
  authPath: path.join(__dirname, 'auth-state.json'),

  templateUrl: 'https://forms.cloud.microsoft/Pages/ShareFormPage.aspx?id=-PwcN9hMeUuH3N6aiZ96iJL6XI4jatJEuJk0OOXdqXtUNTRCMUo5VklaVkU4MTVMUjU0UVJGSTlHTi4u&sharetoken=qdz0Z4MCrYM2sgv3XoSU',
  excelUrl: 'https://oxygy.sharepoint.com/:x:/r/sites/OXYGY_General-AInewsletterdesk/_layouts/15/Doc2.aspx?action=edit&sourcedoc=%7Bfac3cd15-5d5f-4175-9397-88174ef6eeed%7D&wdExp=TEAMS-TREATMENT&web=1',

  shortWait: 1500,
  mediumWait: 3000,
  longWait: 5000,
  pageLoadWait: 10000,
  summaryInterval: 3,
};

// ─── Progress ─────────────────────────────────────────────────────────────────
function loadProgress() {
  try {
    if (fs.existsSync(CONFIG.progressPath)) {
      return JSON.parse(fs.readFileSync(CONFIG.progressPath, 'utf-8'));
    }
  } catch (e) { /* */ }
  return { completed: [], skipped: [], needsReview: [], urls: {} };
}

function saveProgress(progress) {
  fs.writeFileSync(CONFIG.progressPath, JSON.stringify(progress, null, 2));
}

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
  log(`  [screenshot] ${filename}`);
}

async function handleSignIn(page, saveAuth, timeout = 180000) {
  const url = page.url();
  if (!url.includes('login') && !url.includes('microsoftonline')) return;
  log('Sign-in required. Please sign in in the browser...');
  await page.waitForURL(
    u => !u.toString().includes('login') && !u.toString().includes('microsoftonline'),
    { timeout }
  );
  await wait(CONFIG.longWait);
  await saveAuth();
  log('Sign-in complete');
}

// Dump visible elements for debugging
async function debugPage(page, label) {
  const info = await page.evaluate(() => {
    const btns = Array.from(document.querySelectorAll('button'))
      .map(b => b.textContent?.trim()).filter(t => t && t.length < 60).slice(0, 20);
    const inputs = Array.from(document.querySelectorAll('input, textarea, [contenteditable="true"]'))
      .map(el => ({
        tag: el.tagName, type: el.type || '', aria: el.getAttribute('aria-label') || '',
        ph: el.placeholder || '', val: (el.value || el.textContent || '').substring(0, 50),
        autoId: el.getAttribute('data-automation-id') || '',
      })).slice(0, 20);
    return { btns, inputs, url: location.href };
  });
  log(`  [debug ${label}] URL: ${info.url}`);
  log(`  [debug ${label}] Buttons: ${info.btns.join(' | ')}`);
  info.inputs.forEach((inp, i) => {
    log(`  [debug ${label}] Input[${i}]: ${inp.tag} type=${inp.type} aria="${inp.aria}" ph="${inp.ph}" autoId="${inp.autoId}" val="${inp.val}"`);
  });
}

// ─── STEP 2: Duplicate Template ───────────────────────────────────────────────
async function duplicateTemplate(context, templatePage, saveAuth) {
  log('  STEP 2: Duplicating template...');

  // Navigate to the template share page
  await templatePage.goto(CONFIG.templateUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await wait(CONFIG.pageLoadWait);
  await handleSignIn(templatePage, saveAuth);
  await wait(CONFIG.mediumWait);

  log(`  Template URL: ${templatePage.url()}`);

  // Set up listener for new tab/popup BEFORE clicking Duplicate
  const newPagePromise = context.waitForEvent('page', { timeout: 30000 });

  // Click "Duplicate it" button
  const dupBtn = templatePage.locator('button:has-text("Duplicate"), a:has-text("Duplicate")').first();
  if (!await dupBtn.isVisible({ timeout: 5000 })) {
    await ss(templatePage, 'no_duplicate_btn');
    await debugPage(templatePage, 'no_dup');
    throw new Error('Cannot find "Duplicate it" button');
  }

  await dupBtn.click();
  log('  Clicked "Duplicate it"');

  // Wait for the new tab to open (form editor)
  let editorPage;
  try {
    editorPage = await newPagePromise;
    log('  New tab opened for form editor');
    await editorPage.waitForLoadState('domcontentloaded', { timeout: 30000 });
    await wait(CONFIG.pageLoadWait);
  } catch (e) {
    // Maybe it didn't open a new tab — check if current page navigated
    log('  No new tab detected, checking current page...');
    await wait(CONFIG.pageLoadWait);
    editorPage = templatePage;
  }

  // Handle sign-in if the editor page requires it
  await handleSignIn(editorPage, saveAuth);
  await wait(CONFIG.longWait);

  log(`  Editor URL: ${editorPage.url()}`);
  await ss(editorPage, 'editor_loaded');
  await debugPage(editorPage, 'editor');

  // Close any overlay panels (Collaboration, Copilot suggestions)
  try {
    // Close collaboration/share panel if open
    const closeBtns = editorPage.locator('[aria-label="Close"], [aria-label="Dismiss"]');
    const closeCount = await closeBtns.count();
    for (let i = 0; i < closeCount; i++) {
      try {
        if (await closeBtns.nth(i).isVisible({ timeout: 1000 })) {
          await closeBtns.nth(i).click();
          await wait(500);
        }
      } catch (e) { /* */ }
    }
  } catch (e) { /* */ }

  // Close Copilot suggestion banner (the X button)
  try {
    const xBtns = editorPage.locator('button[aria-label="Close"], button[aria-label="Dismiss suggestions"]');
    for (let i = 0; i < await xBtns.count(); i++) {
      try {
        if (await xBtns.nth(i).isVisible({ timeout: 1000 })) {
          await xBtns.nth(i).click();
          await wait(300);
        }
      } catch (e) { /* */ }
    }
  } catch (e) { /* */ }

  await wait(CONFIG.shortWait);
  log('  STEP 2: Form duplicated');
  return editorPage;
}

// ─── STEP 3: Rename Form ─────────────────────────────────────────────────────
async function renameForm(page, formName) {
  log(`  STEP 3: Renaming to "${formName}"`);

  // From screenshots, the form title is the large bold text on the form preview:
  // "LUSA NEO End User Training Feedback Form (Copy)"
  // We need to click on it. It might be rendered as a contenteditable div or
  // as part of the form header in the editor.

  // Strategy: Click on the text containing "(Copy)" or the full default title
  let clicked = false;

  // Try clicking the title text directly
  const titleTexts = [
    'LUSA NEO End User Training Feedback Form (Copy)',
    'LUSA NEO End User Training Feedback Form',
    'Feedback Form (Copy)',
    'Feedback Form',
  ];

  for (const t of titleTexts) {
    try {
      const el = page.locator(`text="${t}"`).first();
      if (await el.isVisible({ timeout: 2000 })) {
        await el.click();
        await wait(CONFIG.shortWait);
        clicked = true;
        log(`  Clicked title text: "${t}"`);
        break;
      }
    } catch (e) { /* */ }
  }

  if (!clicked) {
    // Try clicking on the form header bar at the top (where the title is shown)
    // In the screenshots it's in the gray bar: "LUSA NEO End User Training Feedback Form (Copy) (Copy) · Saved..."
    try {
      const headerTitle = page.locator('[class*="formTitle" i], [class*="FormTitle"], [data-automation-id*="title" i]').first();
      if (await headerTitle.isVisible({ timeout: 2000 })) {
        await headerTitle.click();
        await wait(CONFIG.shortWait);
        clicked = true;
        log('  Clicked header title element');
      }
    } catch (e) { /* */ }
  }

  if (!clicked) {
    await ss(page, 'title_not_found');
    await debugPage(page, 'title');
    log('  WARNING: Could not click title. Will try Settings approach...');

    // Try using Settings > Form settings to change the name
    try {
      const settingsBtn = page.locator('button:has-text("Settings")').first();
      if (await settingsBtn.isVisible({ timeout: 2000 })) {
        await settingsBtn.click();
        await wait(CONFIG.mediumWait);
        await ss(page, 'settings_panel');
        await debugPage(page, 'settings');
        await page.keyboard.press('Escape');
        await wait(500);
      }
    } catch (e) { /* */ }
    return; // Continue without renaming — we'll note this
  }

  // Now the title should be editable — select all and type new name
  await wait(500);

  // Try to find an active input/contenteditable and replace text
  try {
    // Select all text
    await page.keyboard.press('Meta+A');
    await wait(200);
    // Type replacement
    await page.keyboard.type(formName, { delay: 15 });
    await wait(CONFIG.shortWait);
    // Confirm by clicking outside or pressing Tab
    await page.keyboard.press('Tab');
    await wait(CONFIG.mediumWait);
  } catch (e) {
    log(`  Warning: Error during title text replacement: ${e.message}`);
  }

  await ss(page, 'after_rename');
  log('  STEP 3: Rename attempted');
}

// ─── STEP 4: Edit Questions in Self Assessment Section ────────────────────────
async function editQuestions(page, questions, formName) {
  log(`  STEP 4: Editing ${questions.length} questions...`);
  const reviewItems = [];

  // We need to find the "Self Assessment" section in the form editor.
  // The editor shows the form as a scrollable preview. We need to scroll
  // within the form content area to reach the Self Assessment section.

  // First, scroll to find "Self Assessment" or "Self-assessment"
  log('  Looking for Self Assessment section...');
  let found = false;

  // Scroll the page to find the section
  for (let attempt = 0; attempt < 20; attempt++) {
    const visible = await page.locator('text=/[Ss]elf.?[Aa]ssess/').first().isVisible({ timeout: 1000 }).catch(() => false);
    if (visible) {
      found = true;
      await page.locator('text=/[Ss]elf.?[Aa]ssess/').first().scrollIntoViewIfNeeded();
      await wait(CONFIG.shortWait);
      log('  Found Self Assessment section');
      break;
    }
    // Scroll down in the content area
    await page.mouse.wheel(0, 400);
    await wait(800);
  }

  if (!found) {
    await ss(page, 'no_self_assessment');
    await debugPage(page, 'no_sa');
    log('  ERROR: Could not find Self Assessment section');

    // Take a full-page screenshot for diagnosis
    await page.screenshot({
      path: path.join(CONFIG.screenshotsDir, `full_page_${Date.now()}.png`),
      fullPage: true,
    });

    throw new Error('Self Assessment section not found');
  }

  await ss(page, 'self_assessment_found');

  // Now we need to click on each self-assessment question placeholder and edit it.
  // The placeholders are likely "Self-assessment question one", "two", etc.

  // Find all self-assessment placeholder texts
  const saLabels = page.locator('text=/[Ss]elf.?assessment question/');
  const saCount = await saLabels.count();
  log(`  Found ${saCount} self-assessment placeholder questions`);

  for (let i = 0; i < questions.length; i++) {
    const q = questions[i];
    log(`  Q${i + 1}/${questions.length}: "${q.text.substring(0, 50)}..."`);

    try {
      // Click on the placeholder question to enter edit mode
      if (i < saCount) {
        const label = saLabels.nth(i);
        await label.scrollIntoViewIfNeeded();
        await wait(500);
        await label.click();
        await wait(CONFIG.mediumWait);
      } else {
        // Need to add new question
        log(`  Adding new question (no placeholder for index ${i})`);
        const addBtn = page.locator('text="Insert new question", button:has-text("Add new"), button:has-text("Insert")').first();
        if (await addBtn.isVisible({ timeout: 3000 })) {
          await addBtn.click();
          await wait(CONFIG.mediumWait);
        }
      }

      await ss(page, `q${i + 1}_editing`);
      await debugPage(page, `q${i + 1}`);

      // At this point we should see the question editing interface.
      // We need to:
      // a) Change type from Rating to Choice (if it's a rating question)
      // b) Clear and type the question text
      // c) Add the options
      // d) Mark correct answer

      // For now, log what we see and take screenshots.
      // The actual selectors depend on what the editing UI looks like.

      // Try to find the question text area and replace it
      // Look for any contenteditable or text input that's currently active
      const editableFields = page.locator('[contenteditable="true"]:visible, textarea:visible, input[type="text"]:visible');
      const fieldCount = await editableFields.count();
      log(`  Found ${fieldCount} editable fields`);

      if (fieldCount > 0) {
        // The first editable field is usually the question title
        const qField = editableFields.first();
        await qField.click();
        await wait(300);
        await page.keyboard.press('Meta+A');
        await wait(100);
        await page.keyboard.type(q.text, { delay: 10 });
        await wait(CONFIG.shortWait);
        log(`  Typed question text`);
      }

      // Now handle options — this depends on the question type
      // If it's a Rating (1-5), we need to change it to Choice first
      // Then add the options

      // Look for a type selector/dropdown
      const typeBtn = page.locator('button:has-text("Rating"), button:has-text("Choice"), button:has-text("Text"), [aria-label*="question type" i]').first();
      if (await typeBtn.isVisible({ timeout: 2000 })) {
        const typeText = await typeBtn.textContent();
        if (!typeText?.includes('Choice')) {
          await typeBtn.click();
          await wait(CONFIG.shortWait);
          // Select "Choice" from dropdown
          const choiceOpt = page.locator('[role="option"]:has-text("Choice"), [role="menuitem"]:has-text("Choice"), text="Choice"').first();
          if (await choiceOpt.isVisible({ timeout: 3000 })) {
            await choiceOpt.click();
            await wait(CONFIG.mediumWait);
            log(`  Changed type to Choice`);
          }
        }
      }

      // Find option fields and fill them
      const optionFields = page.locator('[placeholder*="Option" i], [aria-label*="Option" i]');
      let optCount = await optionFields.count();
      log(`  Found ${optCount} option fields`);

      // Add more options if needed
      while (optCount < q.options.length) {
        const addOpt = page.locator('text="Add option", button:has-text("Add option"), [aria-label*="Add option" i]').first();
        if (await addOpt.isVisible({ timeout: 2000 })) {
          await addOpt.click();
          await wait(800);
          optCount = await optionFields.count();
        } else {
          break;
        }
      }

      // Fill options
      for (let j = 0; j < q.options.length && j < optCount; j++) {
        try {
          const opt = optionFields.nth(j);
          await opt.click();
          await wait(200);
          await page.keyboard.press('Meta+A');
          await wait(100);
          await page.keyboard.type(q.options[j], { delay: 8 });
          await wait(400);
        } catch (e) {
          log(`  Warning: Could not fill option ${j + 1}: ${e.message}`);
        }
      }

      // Mark correct answer
      if (q.needs_review) {
        log(`  Skipping correct answer (needs review)`);
        reviewItems.push({ formName, questionNumber: q.number, text: q.text });
      } else {
        // Try to find correct answer toggles
        const correctBtns = page.locator('[aria-label*="correct" i], [data-automation-id*="correct" i], [title*="correct" i]');
        const correctCount = await correctBtns.count();
        if (correctCount > q.correct_answer_index) {
          await correctBtns.nth(q.correct_answer_index).click();
          await wait(500);
          log(`  Marked correct: option ${q.correct_answer_index + 1}`);
        } else {
          log(`  Note: Could not find correct answer toggle (${correctCount} found)`);
        }
      }

      // Click outside to deselect question
      await page.keyboard.press('Escape');
      await wait(CONFIG.mediumWait);

    } catch (err) {
      log(`  ERROR Q${i + 1}: ${err.message}`);
      await ss(page, `error_q${i + 1}`);
      try { await page.keyboard.press('Escape'); } catch (e) { /* */ }
      await wait(CONFIG.shortWait);
    }
  }

  // Delete extra placeholder questions
  if (saCount > questions.length) {
    log(`  Deleting ${saCount - questions.length} extra placeholders...`);
    // Re-find the remaining self-assessment labels
    for (let i = saCount - 1; i >= questions.length; i--) {
      try {
        const extra = page.locator('text=/[Ss]elf.?assessment question/').nth(i);
        if (await extra.isVisible({ timeout: 2000 })) {
          await extra.click();
          await wait(CONFIG.shortWait);
          // Find delete/trash button
          const delBtn = page.locator('[aria-label*="Delete" i], [aria-label*="Remove" i], button:has-text("Delete")').first();
          if (await delBtn.isVisible({ timeout: 2000 })) {
            await delBtn.click();
            await wait(CONFIG.shortWait);
          }
        }
      } catch (e) {
        log(`  Warning: Could not delete placeholder ${i + 1}`);
      }
    }
  }

  await ss(page, 'questions_done');
  log(`  STEP 4: Questions edited`);
  return reviewItems;
}

// ─── STEP 5: Get Share URL ────────────────────────────────────────────────────
async function getShareUrl(page) {
  log('  STEP 5: Getting share URL...');

  // Scroll back to top
  await page.evaluate(() => window.scrollTo(0, 0));
  await wait(CONFIG.shortWait);

  // Click "Collect responses" button (seen in screenshot toolbar)
  const collectBtn = page.locator('button:has-text("Collect responses"), [aria-label*="Collect responses" i]').first();
  if (await collectBtn.isVisible({ timeout: 5000 })) {
    await collectBtn.click();
    await wait(CONFIG.mediumWait);
    log('  Clicked "Collect responses"');
  } else {
    // Try "Share" or "Send"
    const shareBtn = page.locator('button:has-text("Share"), button:has-text("Send")').first();
    if (await shareBtn.isVisible({ timeout: 3000 })) {
      await shareBtn.click();
      await wait(CONFIG.mediumWait);
    }
  }

  await ss(page, 'share_panel');

  // The Collaboration panel shows a URL input with a "Copy" button
  // From screenshot: input field with https://forms.cloud.microsoft... and blue "Copy" button
  let shareUrl = null;

  // Try to get URL from any input on page
  shareUrl = await page.evaluate(() => {
    const inputs = document.querySelectorAll('input');
    for (const inp of inputs) {
      if (inp.value && inp.value.includes('forms.cloud.microsoft')) {
        return inp.value;
      }
    }
    // Check text nodes for URL
    const body = document.body.innerText;
    const match = body.match(/https:\/\/forms\.cloud\.microsoft[^\s"'<>]+/);
    return match ? match[0] : null;
  });

  if (shareUrl) {
    log(`  Share URL: ${shareUrl}`);
  } else {
    log('  WARNING: Could not find share URL');
    await debugPage(page, 'share');

    // Use current page URL as fallback
    shareUrl = page.url();
    log(`  Using editor URL as fallback: ${shareUrl}`);
  }

  // Close the panel
  try { await page.keyboard.press('Escape'); } catch (e) { /* */ }
  await wait(CONFIG.shortWait);

  return shareUrl;
}

// ─── STEP 6: Update Excel ────────────────────────────────────────────────────
async function updateExcel(excelPage, formName, shareUrl, formIndex) {
  log('  STEP 6: Updating Excel tracker...');
  await excelPage.bringToFront();
  await wait(CONFIG.mediumWait);

  // FIRST: Switch to "Forms Tracker" sheet
  // From the Excel screenshot, tabs are at the bottom: "Question Bank", "Forms Tracker", "Extraction Report"
  log('  Switching to Forms Tracker sheet...');

  // Click the "Forms Tracker" tab
  const trackerTab = excelPage.locator('button:has-text("Forms Tracker"), [role="tab"]:has-text("Forms Tracker"), a:has-text("Forms Tracker")').first();
  if (await trackerTab.isVisible({ timeout: 5000 })) {
    await trackerTab.click();
    await wait(CONFIG.mediumWait);
    log('  Clicked "Forms Tracker" tab');
  } else {
    // Try finding it as just text
    const tabText = excelPage.locator('text="Forms Tracker"').first();
    if (await tabText.isVisible({ timeout: 3000 })) {
      await tabText.click();
      await wait(CONFIG.mediumWait);
      log('  Clicked "Forms Tracker" text');
    } else {
      log('  WARNING: Could not find Forms Tracker tab');
      await ss(excelPage, 'excel_no_tracker_tab');
      await debugPage(excelPage, 'excel_tabs');
    }
  }

  await ss(excelPage, 'forms_tracker_sheet');

  // Navigate to the cell using the Name Box
  // The Name Box is the cell reference input at the top-left (shows like "A1")
  // Form URL is column B, and formIndex 1 = row 2 (row 1 is header)
  const targetRow = formIndex + 1;
  log(`  Navigating to B${targetRow}...`);

  // Click the Name Box — it's typically an input showing the current cell reference
  // In Excel Online, you can click the Name Box and type a cell address
  try {
    const nameBox = excelPage.locator('#NameBox, [id*="NameBox"], [id*="namebox"], [aria-label*="Name Box" i]').first();
    if (await nameBox.isVisible({ timeout: 3000 })) {
      await nameBox.click();
      await wait(500);
      await nameBox.fill(`B${targetRow}`);
      await excelPage.keyboard.press('Enter');
      await wait(CONFIG.shortWait);
      log(`  Navigated to B${targetRow} via Name Box`);
    } else {
      // Fallback: use Ctrl+G / F5 for Go To
      log('  Name Box not found, trying Ctrl+G...');
      await excelPage.keyboard.press('Control+g');
      await wait(1000);
      // Check if a dialog appeared
      await ss(excelPage, 'goto_dialog');
    }
  } catch (e) {
    log(`  Warning: Name Box navigation failed: ${e.message}`);
  }

  // Type the share URL into cell B
  if (shareUrl) {
    await excelPage.keyboard.type(shareUrl, { delay: 5 });
    await excelPage.keyboard.press('Tab'); // Move to next cell
    await wait(500);
    log('  Typed share URL');
  }

  // Now navigate to column F (Status) for the same row
  log(`  Navigating to F${targetRow} (Status)...`);
  try {
    const nameBox = excelPage.locator('#NameBox, [id*="NameBox"], [id*="namebox"], [aria-label*="Name Box" i]').first();
    await nameBox.click();
    await wait(500);
    await nameBox.fill(`F${targetRow}`);
    await excelPage.keyboard.press('Enter');
    await wait(CONFIG.shortWait);
  } catch (e) { /* */ }

  await excelPage.keyboard.type('Complete', { delay: 10 });
  await excelPage.keyboard.press('Tab');
  await wait(500);

  // Column G (Date Created) — should already be selected after Tab from F
  const today = new Date().toLocaleDateString('en-US');
  await excelPage.keyboard.type(today, { delay: 10 });
  await excelPage.keyboard.press('Enter');
  await wait(CONFIG.shortWait);

  await ss(excelPage, `excel_updated_form${formIndex}`);
  log('  STEP 6: Excel updated');
}

// ─── Main ─────────────────────────────────────────────────────────────────────
async function main() {
  const args = process.argv.slice(2);
  const startArg = args.find(a => a.startsWith('--start='));
  const onlyArg = args.find(a => a.startsWith('--only='));
  const dryRun = args.includes('--dry-run');
  const startIndex = startArg ? parseInt(startArg.split('=')[1]) : 1;
  const onlyIndex = onlyArg ? parseInt(onlyArg.split('=')[1]) : null;

  const manifest = JSON.parse(fs.readFileSync(CONFIG.manifestPath, 'utf-8'));
  log(`Manifest: ${manifest.total_forms} forms, ${manifest.total_questions} questions`);

  const progress = loadProgress();
  log(`Progress: ${progress.completed.length} completed, ${progress.skipped.length} skipped`);

  const formsToProcess = manifest.forms.filter(f => {
    if (progress.completed.includes(f.index)) return false;
    if (onlyIndex !== null) return f.index === onlyIndex;
    return f.index >= startIndex;
  });

  log(`To process: ${formsToProcess.length} forms`);

  if (dryRun) {
    formsToProcess.forEach(f => log(`  #${f.index}: ${f.form_name} (${f.question_count}q)`));
    return;
  }

  if (formsToProcess.length === 0) {
    log('Nothing to do!');
    return;
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

  try {
    // Open template page
    const templatePage = await context.newPage();
    log('Opening template page...');
    await templatePage.goto(CONFIG.templateUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await wait(CONFIG.longWait);
    await handleSignIn(templatePage, saveAuth);
    await wait(CONFIG.mediumWait);
    await saveAuth();
    log('Template page ready');

    // Open Excel
    log('Opening Excel...');
    const excelPage = await context.newPage();
    await excelPage.goto(CONFIG.excelUrl, { waitUntil: 'domcontentloaded', timeout: 60000 });
    await wait(CONFIG.pageLoadWait);
    await handleSignIn(excelPage, saveAuth);
    await wait(CONFIG.mediumWait);
    log('Excel loaded');

    // Process forms
    let completedThisRun = 0;

    for (const form of formsToProcess) {
      const num = String(form.index).padStart(2, '0');
      log('');
      log(`======== FORM #${num}: ${form.form_name} (${form.question_count}q) ========`);

      try {
        // STEP 1: Read JSON
        const quizData = JSON.parse(fs.readFileSync(path.join(CONFIG.quizJsonDir, form.filename), 'utf-8'));
        log(`  STEP 1: Loaded ${quizData.questions.length} questions`);

        // STEP 2: Duplicate (returns the editor page — may be a new tab)
        await templatePage.bringToFront();
        const editorPage = await duplicateTemplate(context, templatePage, saveAuth);

        // STEP 3: Rename
        await renameForm(editorPage, quizData.form_name);

        // STEP 4: Questions
        const reviewItems = await editQuestions(editorPage, quizData.questions, quizData.form_name);
        if (reviewItems.length > 0) progress.needsReview.push(...reviewItems);

        // STEP 5: Share URL
        const shareUrl = await getShareUrl(editorPage);
        progress.urls[form.index] = shareUrl;

        // STEP 6: Excel
        await updateExcel(excelPage, quizData.form_name, shareUrl, form.index);

        // Close the editor tab (we're done with this form)
        if (editorPage !== templatePage) {
          await editorPage.close();
        }

        // Mark complete
        progress.completed.push(form.index);
        saveProgress(progress);
        await saveAuth();

        completedThisRun++;
        log(`  DONE! Form #${num} (${completedThisRun} this run, ${progress.completed.length} total)`);

        // Summary
        if (completedThisRun % CONFIG.summaryInterval === 0) {
          log('');
          log(`--- SUMMARY: ${progress.completed.length}/${manifest.total_forms} done, ${progress.skipped.length} skipped, ${progress.needsReview.length} need review ---`);
        }

      } catch (err) {
        log(`  FAILED: ${err.message}`);
        await ss(templatePage, `error_form_${num}`);
        progress.skipped.push({ index: form.index, name: form.form_name, error: err.message });
        saveProgress(progress);

        // Close any extra editor tabs
        const pages = context.pages();
        while (pages.length > 2) {
          await pages[pages.length - 1].close();
          pages.pop();
        }
      }
    }

    // Final report
    log('');
    log('======== RUN COMPLETE ========');
    log(`Completed: ${completedThisRun} this run, ${progress.completed.length} total`);
    log(`Skipped: ${progress.skipped.length}`);
    if (progress.skipped.length > 0) {
      progress.skipped.forEach(s => log(`  - #${s.index} ${s.name}: ${s.error}`));
    }
    if (progress.needsReview.length > 0) {
      log(`Needs review: ${progress.needsReview.length} questions`);
    }
    saveProgress(progress);
    log('\nCtrl+C to exit.');
    await new Promise(() => {});

  } catch (err) {
    log(`FATAL: ${err.message}`);
    console.error(err);
    saveProgress(progress);
    await browser.close();
    process.exit(1);
  }
}

main().catch(console.error);
