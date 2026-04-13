/**
 * Diagnostic script — explores the MS Forms editor DOM and Excel Online
 * to find the exact selectors we need. Does NOT modify anything.
 *
 * Run: node diagnose.js
 */

const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const TEMPLATE_URL = 'https://forms.cloud.microsoft/Pages/ShareFormPage.aspx?id=-PwcN9hMeUuH3N6aiZ96iJL6XI4jatJEuJk0OOXdqXtUNTRCMUo5VklaVkU4MTVMUjU0UVJGSTlHTi4u&sharetoken=qdz0Z4MCrYM2sgv3XoSU';
const EXCEL_URL = 'https://oxygy.sharepoint.com/:x:/r/sites/OXYGY_General-AInewsletterdesk/_layouts/15/Doc2.aspx?action=edit&sourcedoc=%7Bfac3cd15-5d5f-4175-9397-88174ef6eeed%7D&wdExp=TEAMS-TREATMENT&web=1';

const screenshotsDir = path.join(__dirname, 'diagnostics');
const authPath = path.join(__dirname, 'auth-state.json');
let ssCount = 0;

async function ss(page, label) {
  if (!fs.existsSync(screenshotsDir)) fs.mkdirSync(screenshotsDir, { recursive: true });
  ssCount++;
  const name = `${String(ssCount).padStart(2, '0')}_${label}.png`;
  await page.screenshot({ path: path.join(screenshotsDir, name), fullPage: false });
  console.log(`  [screenshot] ${name}`);
}

async function wait(ms) {
  return new Promise(r => setTimeout(r, ms));
}

async function dumpDOM(page, label) {
  console.log(`\n=== ${label} ===`);
  console.log(`URL: ${page.url()}`);
  console.log(`Title: ${await page.title()}`);

  const info = await page.evaluate(() => {
    function describeEl(el, depth = 0) {
      const tag = el.tagName?.toLowerCase() || '?';
      const id = el.id ? `#${el.id}` : '';
      const cls = el.className && typeof el.className === 'string' ? `.${el.className.split(' ').slice(0, 3).join('.')}` : '';
      const role = el.getAttribute?.('role') ? `[role=${el.getAttribute('role')}]` : '';
      const aria = el.getAttribute?.('aria-label') ? `[aria-label="${el.getAttribute('aria-label').substring(0, 50)}"]` : '';
      const ce = el.getAttribute?.('contenteditable') === 'true' ? '[contenteditable]' : '';
      const text = el.textContent?.trim().substring(0, 60) || '';
      const pad = '  '.repeat(depth);
      return `${pad}${tag}${id}${cls}${role}${aria}${ce} → "${text}"`;
    }

    // Get all contenteditable elements
    const editables = Array.from(document.querySelectorAll('[contenteditable="true"]'))
      .map(el => describeEl(el));

    // Get all inputs and textareas
    const inputs = Array.from(document.querySelectorAll('input, textarea'))
      .map(el => {
        const tag = el.tagName.toLowerCase();
        const type = el.type || '';
        const name = el.name || '';
        const placeholder = el.placeholder || '';
        const ariaLabel = el.getAttribute('aria-label') || '';
        const value = (el.value || '').substring(0, 80);
        const autoId = el.getAttribute('data-automation-id') || '';
        return `${tag}[type=${type}][name=${name}] placeholder="${placeholder}" aria="${ariaLabel}" autoId="${autoId}" value="${value}"`;
      });

    // Get all buttons
    const buttons = Array.from(document.querySelectorAll('button'))
      .map(el => {
        const ariaLabel = el.getAttribute('aria-label') || '';
        const text = el.textContent?.trim().substring(0, 60) || '';
        const autoId = el.getAttribute('data-automation-id') || '';
        return `button aria="${ariaLabel}" autoId="${autoId}" → "${text}"`;
      })
      .filter(t => t.includes('→ "') && !t.endsWith('→ ""'));

    // Get all sections/headings
    const headings = Array.from(document.querySelectorAll('h1, h2, h3, h4, [role="heading"]'))
      .map(el => describeEl(el));

    // Get all elements with data-automation-id
    const autoIds = Array.from(document.querySelectorAll('[data-automation-id]'))
      .map(el => {
        const autoId = el.getAttribute('data-automation-id');
        const tag = el.tagName.toLowerCase();
        const text = el.textContent?.trim().substring(0, 50) || '';
        return `${tag}[data-automation-id="${autoId}"] → "${text}"`;
      });

    // Get sheet tabs (for Excel)
    const tabs = Array.from(document.querySelectorAll('[role="tab"], [class*="sheet-tab"], [class*="SheetTab"]'))
      .map(el => {
        const text = el.textContent?.trim() || '';
        const ariaLabel = el.getAttribute('aria-label') || '';
        const cls = el.className || '';
        return `tab: "${text}" aria="${ariaLabel}" class="${cls.substring(0, 80)}"`;
      });

    // Get sections of the form
    const sections = Array.from(document.querySelectorAll('[class*="section" i], [class*="Section"]'))
      .map(el => {
        const cls = el.className?.substring(0, 60) || '';
        const text = el.textContent?.trim().substring(0, 80) || '';
        return `section class="${cls}" → "${text}"`;
      });

    return { editables, inputs, buttons: buttons.slice(0, 30), headings, autoIds: autoIds.slice(0, 40), tabs, sections: sections.slice(0, 15) };
  });

  console.log('\nContenteditable elements:');
  info.editables.forEach(e => console.log(`  ${e}`));

  console.log('\nInputs/Textareas:');
  info.inputs.forEach(e => console.log(`  ${e}`));

  console.log('\nButtons (non-empty):');
  info.buttons.forEach(e => console.log(`  ${e}`));

  console.log('\nHeadings:');
  info.headings.forEach(e => console.log(`  ${e}`));

  console.log('\ndata-automation-id elements:');
  info.autoIds.forEach(e => console.log(`  ${e}`));

  if (info.tabs.length > 0) {
    console.log('\nSheet tabs:');
    info.tabs.forEach(e => console.log(`  ${e}`));
  }

  if (info.sections.length > 0) {
    console.log('\nSections:');
    info.sections.forEach(e => console.log(`  ${e}`));
  }
}

async function main() {
  console.log('=== MS Forms & Excel Diagnostic Tool ===\n');

  const browser = await chromium.launch({
    headless: false,
    slowMo: 100,
    args: ['--start-maximized'],
  });

  const context = await browser.newContext({
    viewport: { width: 1400, height: 900 },
    storageState: fs.existsSync(authPath) && fs.statSync(authPath).size > 10
      ? authPath
      : undefined,
  });

  const saveAuth = async () => {
    try {
      await context.storageState({ path: authPath });
    } catch (e) { /* */ }
  };

  const page = await context.newPage();

  // ─── PART 1: Template page (before duplicate) ─────────────────────
  console.log('\n──── PART 1: Template Share Page ────');
  await page.goto(TEMPLATE_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await wait(8000);

  // Handle login
  const url = page.url();
  if (url.includes('login') || url.includes('microsoftonline')) {
    console.log('\n*** SIGN IN REQUIRED ***');
    console.log('Please sign in in the browser window. Waiting up to 3 minutes...\n');
    await page.waitForURL(
      u => !u.toString().includes('login') && !u.toString().includes('microsoftonline'),
      { timeout: 180000 }
    );
    await wait(8000);
    await saveAuth();
    console.log('Sign-in complete, auth saved.\n');
  }

  await ss(page, 'template_page');
  await dumpDOM(page, 'TEMPLATE SHARE PAGE');

  // ─── PART 2: Click Duplicate ──────────────────────────────────────
  console.log('\n──── PART 2: After Clicking Duplicate ────');
  const dupBtn = page.locator('button:has-text("Duplicate"), a:has-text("Duplicate"), span:has-text("Duplicate")').first();
  if (await dupBtn.isVisible({ timeout: 5000 })) {
    await dupBtn.click();
    console.log('Clicked Duplicate button');
  } else {
    console.log('ERROR: No Duplicate button found');
    await ss(page, 'no_duplicate_btn');
    await dumpDOM(page, 'NO DUPLICATE BUTTON');
  }

  await wait(15000); // Wait a long time for form editor to load

  // Handle login again if needed
  const url2 = page.url();
  if (url2.includes('login') || url2.includes('microsoftonline')) {
    console.log('\n*** SIGN IN REQUIRED (after duplicate) ***');
    await page.waitForURL(
      u => !u.toString().includes('login') && !u.toString().includes('microsoftonline'),
      { timeout: 180000 }
    );
    await wait(8000);
    await saveAuth();
  }

  await ss(page, 'form_editor_full');
  await dumpDOM(page, 'FORM EDITOR (top of page)');

  // ─── PART 3: Scroll through the form to find all sections ─────────
  console.log('\n──── PART 3: Scrolling Through Form Sections ────');
  for (let i = 0; i < 15; i++) {
    await page.mouse.wheel(0, 600);
    await wait(1500);
  }
  await ss(page, 'form_editor_scrolled_bottom');
  await dumpDOM(page, 'FORM EDITOR (after scrolling down)');

  // Scroll back to top
  for (let i = 0; i < 15; i++) {
    await page.mouse.wheel(0, -600);
    await wait(500);
  }

  // ─── PART 4: Click on a question to see editing UI ────────────────
  console.log('\n──── PART 4: Click on a Self Assessment Question ────');
  // Try to find and click on any self-assessment question
  const saText = page.locator('text=/[Ss]elf/').first();
  if (await saText.isVisible({ timeout: 5000 })) {
    await saText.scrollIntoViewIfNeeded();
    await wait(1000);
    await ss(page, 'self_assessment_visible');
    await saText.click();
    await wait(3000);
    await ss(page, 'self_assessment_clicked');
    await dumpDOM(page, 'AFTER CLICKING SELF-ASSESSMENT QUESTION');
  } else {
    console.log('Could not find "Self" text on page');
    // Try to find any clickable question area
    const questions = page.locator('[class*="question" i], [class*="Question"]');
    const qCount = await questions.count();
    console.log(`Found ${qCount} question-like elements`);
  }

  // ─── PART 5: Check the Collect Responses button ───────────────────
  console.log('\n──── PART 5: Collect Responses ────');
  // Scroll back to top first
  await page.evaluate(() => window.scrollTo(0, 0));
  await wait(1000);

  const collectBtn = page.locator('button:has-text("Collect responses"), [aria-label*="Collect responses"]').first();
  if (await collectBtn.isVisible({ timeout: 3000 })) {
    await collectBtn.click();
    await wait(3000);
    await ss(page, 'collect_responses_panel');
    await dumpDOM(page, 'COLLECT RESPONSES PANEL');
    await page.keyboard.press('Escape');
    await wait(1000);
  } else {
    console.log('Could not find Collect responses button');
  }

  // ─── PART 6: Excel Online ────────────────────────────────────────
  console.log('\n──── PART 6: Excel Online ────');
  const excelPage = await context.newPage();
  await excelPage.goto(EXCEL_URL, { waitUntil: 'domcontentloaded', timeout: 60000 });
  await wait(10000);

  // Handle login
  const excelUrl = excelPage.url();
  if (excelUrl.includes('login') || excelUrl.includes('microsoftonline')) {
    console.log('*** SIGN IN REQUIRED for Excel ***');
    await excelPage.waitForURL(
      u => !u.toString().includes('login') && !u.toString().includes('microsoftonline'),
      { timeout: 180000 }
    );
    await wait(8000);
    await saveAuth();
  }

  await ss(excelPage, 'excel_initial');
  await dumpDOM(excelPage, 'EXCEL ONLINE (initial load)');

  // Look at sheet tabs specifically
  console.log('\n──── Excel Sheet Tabs ────');
  const tabInfo = await excelPage.evaluate(() => {
    // Find anything at the bottom that looks like sheet tabs
    const allEls = document.querySelectorAll('*');
    const tabLike = [];
    for (const el of allEls) {
      const text = el.textContent?.trim() || '';
      if ((text.includes('Tracker') || text.includes('Question Bank') || text.includes('Sheet'))
          && el.children.length < 5
          && text.length < 50) {
        tabLike.push({
          tag: el.tagName,
          text,
          class: (el.className || '').substring(0, 80),
          role: el.getAttribute('role') || '',
          id: el.id || '',
          ariaLabel: el.getAttribute('aria-label') || '',
          clickable: el.onclick !== null || el.tagName === 'BUTTON' || el.tagName === 'A' || el.getAttribute('role') === 'tab',
        });
      }
    }
    return tabLike;
  });
  console.log('Tab-like elements:');
  tabInfo.forEach(t => console.log(`  ${t.tag} "${t.text}" role=${t.role} class="${t.class}" id="${t.id}" clickable=${t.clickable}`));

  // Try clicking "Forms Tracker" tab
  console.log('\n──── Trying to switch to Forms Tracker ────');
  const trackerTab = excelPage.locator('text="Forms Tracker"').first();
  if (await trackerTab.isVisible({ timeout: 3000 })) {
    await trackerTab.click();
    await wait(3000);
    await ss(excelPage, 'forms_tracker_sheet');
    await dumpDOM(excelPage, 'FORMS TRACKER SHEET');
  } else {
    console.log('Could not find "Forms Tracker" tab');
    // Try partial match
    const tracker2 = excelPage.locator('text=/[Tt]racker/').first();
    if (await tracker2.isVisible({ timeout: 3000 })) {
      await tracker2.click();
      await wait(3000);
      await ss(excelPage, 'tracker_sheet');
    }
  }

  // Check the Name Box
  console.log('\n──── Excel Name Box ────');
  const nameBoxInfo = await excelPage.evaluate(() => {
    const candidates = document.querySelectorAll('input');
    return Array.from(candidates).map(el => ({
      id: el.id,
      class: (el.className || '').substring(0, 80),
      ariaLabel: el.getAttribute('aria-label') || '',
      value: (el.value || '').substring(0, 30),
      name: el.name || '',
      placeholder: el.placeholder || '',
    })).filter(c => c.value.match(/^[A-Z]+\d+$/) || c.ariaLabel.toLowerCase().includes('name') || c.id.toLowerCase().includes('name'));
  });
  console.log('Name Box candidates:');
  nameBoxInfo.forEach(n => console.log(`  input#${n.id} class="${n.class}" aria="${n.ariaLabel}" value="${n.value}"`));

  // ─── Done ─────────────────────────────────────────────────────────
  console.log('\n\n=== DIAGNOSTICS COMPLETE ===');
  console.log(`Screenshots saved to: ${screenshotsDir}/`);
  console.log('Review the output above to identify the correct selectors.');
  console.log('\nPress Ctrl+C to close the browser.');
  await new Promise(() => {});
}

main().catch(console.error);
