/**
 * Reconnaissance: can we programmatically grab the response link from the Forms editor?
 *
 * Usage: node recon-response-link.js "<editor-url>"
 *
 * Tries three strategies in order:
 *   1. Click "Collect responses" button, then copy link from the modal
 *   2. Intercept network calls — the button may fetch the share URL from an API
 *   3. Construct the response URL from the form ID (well-known pattern)
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const AUTH_PATH = path.join(__dirname, 'auth-state-pa.json');

async function main() {
  const editorUrl = process.argv[2];
  if (!editorUrl) {
    console.error('Usage: node recon-response-link.js "<editor-url>"');
    process.exit(1);
  }

  const hasAuth = fs.existsSync(AUTH_PATH);
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasAuth ? { storageState: AUTH_PATH } : {});
  const page = await context.newPage();

  // Capture any response-link-related API calls
  const suspiciousResponses = [];
  page.on('response', async (resp) => {
    const url = resp.url();
    if (url.includes('formapi') || url.includes('share') || url.includes('collect')) {
      try {
        const body = await resp.text();
        if (body.includes('ResponsePage') || body.includes('/r/')) {
          suspiciousResponses.push({ url, body: body.substring(0, 500) });
        }
      } catch { /* ignore */ }
    }
  });

  console.log('[1] Opening editor...');
  await page.goto(editorUrl, { waitUntil: 'networkidle', timeout: 60_000 });
  await page.waitForTimeout(3000);
  await context.storageState({ path: AUTH_PATH });

  console.log('[2] Looking for "Collect responses" button...');
  // Try multiple selectors that might match the button
  const collectSelectors = [
    'button:has-text("Collect responses")',
    'button:has-text("Send form")',
    'button:has-text("Share")',
    '[aria-label*="Collect responses"]',
    '[aria-label*="Send"]',
    '[data-automation-id*="share"]',
    '[data-automation-id*="collect"]',
    '[data-automation-id*="send"]',
  ];

  let clicked = false;
  for (const sel of collectSelectors) {
    try {
      const el = await page.locator(sel).first();
      if (await el.count() > 0 && await el.isVisible({ timeout: 2000 })) {
        console.log(`    Found with selector: ${sel}`);
        await el.click();
        clicked = true;
        break;
      }
    } catch { /* try next */ }
  }

  if (!clicked) {
    console.log('    ❌ Could not find a Collect-responses-like button.');
    console.log('    Dumping all button aria-labels for inspection...');
    const labels = await page.evaluate(() => {
      return Array.from(document.querySelectorAll('button'))
        .map((b) => b.getAttribute('aria-label') || b.textContent?.trim())
        .filter((x) => x && x.length < 100)
        .slice(0, 50);
    });
    console.log(JSON.stringify(labels, null, 2));
  } else {
    await page.waitForTimeout(2500);

    console.log('[3] Looking for response link in modal / panel...');
    // The link usually appears in an input/textarea or as visible text
    const linkCandidates = await page.evaluate(() => {
      const results = [];
      // Check inputs and textareas
      document.querySelectorAll('input, textarea').forEach((el) => {
        const v = el.value || '';
        if (v.includes('forms.') && (v.includes('/r/') || v.includes('ResponsePage'))) {
          results.push({ source: 'input', value: v });
        }
      });
      // Check visible text (spans, divs)
      const allText = document.body.innerText || '';
      const m = allText.match(/https:\/\/forms\.[^\s"']+/g);
      if (m) m.forEach((url) => results.push({ source: 'text', value: url }));
      return results;
    });

    if (linkCandidates.length > 0) {
      console.log('    ✅ Found link candidates:');
      linkCandidates.forEach((c) => console.log(`    [${c.source}] ${c.value}`));
    } else {
      console.log('    ❌ No link found in DOM after clicking.');
    }
  }

  console.log('\n[4] Suspicious network responses containing response-page URLs:');
  if (suspiciousResponses.length === 0) {
    console.log('    (none)');
  } else {
    suspiciousResponses.forEach((r) => {
      console.log(`    URL: ${r.url.substring(0, 100)}`);
      console.log(`    Body excerpt: ${r.body.substring(0, 200)}\n`);
    });
  }

  console.log('\nBrowser left open for manual inspection. Close when done.');
}

main().catch((err) => { console.error('Fatal:', err); process.exit(1); });
