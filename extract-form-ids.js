/**
 * Stage 3: Form ID extractor.
 *
 * Given a Microsoft Forms response link, extracts:
 *   - the internal form ID (used in Power Automate's formId parameter)
 *   - the question IDs for every question in the form
 *
 * Usage:
 *   node extract-form-ids.js "<response-link>"
 *
 * Example:
 *   node extract-form-ids.js "https://forms.office.com/r/ABC123"
 *
 * Output: prints a JSON object { formId, questions: [{ number, id, title }] }
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const AUTH_PATH = path.join(__dirname, 'auth-state-pa.json');

async function extract(responseLink) {
  const hasAuth = fs.existsSync(AUTH_PATH);
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasAuth ? { storageState: AUTH_PATH } : {});
  const page = await context.newPage();

  // Watch every response for the Forms API payload that contains question metadata.
  const formApiPayloads = [];
  page.on('response', async (resp) => {
    const url = resp.url();
    if ((url.includes('forms.office.com') || url.includes('forms.cloud.microsoft') || url.includes('forms.microsoft.com')) && url.includes('/formapi/api/forms')) {
      try {
        const body = await resp.json();
        formApiPayloads.push({ url, body });
      } catch { /* non-JSON, ignore */ }
    }
  });

  console.log(`Navigating to response link to extract form ID: ${responseLink}`);
  await page.goto(responseLink, { waitUntil: 'domcontentloaded', timeout: 60_000 });

  // Extract formId from the URL
  let formId = null;
  try {
    const url = new URL(page.url());
    formId = url.searchParams.get('id');
  } catch { /* ignore */ }

  // Now navigate to the EDITOR page — this triggers the formapi calls with valid auth
  if (formId) {
    const editorUrl = `https://forms.office.com/Pages/DesignPageV2.aspx?origin=NeoPortalPage&subpage=design&id=${encodeURIComponent(formId)}`;
    console.log(`Navigating to editor to capture question metadata: ${editorUrl}`);
    try {
      await page.goto(editorUrl, { waitUntil: 'networkidle', timeout: 60_000 });
    } catch (e) {
      console.log('  Editor navigation issue (may still have captured enough):', e.message);
    }
    // Give any late-firing API calls a moment
    await page.waitForTimeout(3000);
  }

  // Save auth for future runs
  await context.storageState({ path: AUTH_PATH });

  // Walk captured API payloads for a questions array
  let questions = [];
  for (const { body } of formApiPayloads) {
    const found = findQuestions(body);
    if (found.length > 0) {
      questions = found;
      break;
    }
  }

  // Primary approach: query DOM for question cards with id="QuestionId_<id>"
  // and extract the matching questionTitle inside each.
  try {
    const domQuestions = await page.evaluate(() => {
      const results = [];
      document.querySelectorAll('[id^="QuestionId_"]').forEach((el) => {
        const id = el.id.replace(/^QuestionId_/, '');
        const titleEl = el.querySelector('[data-automation-id="questionTitle"]');
        let title = null;
        if (titleEl) {
          // Drop the "1." number prefix span, keep the actual question text
          const clone = titleEl.cloneNode(true);
          // Remove anything that's just a number prefix
          clone.querySelectorAll('span').forEach((s) => {
            if (/^\d+\.$/.test((s.textContent || '').trim())) s.remove();
          });
          title = (clone.innerText || clone.textContent || '').replace(/\s+/g, ' ').trim();
        }
        results.push({ id, title });
      });
      return results;
    });

    if (domQuestions.length > 0) {
      questions = domQuestions.map((q, i) => ({ number: i + 1, id: q.id, title: q.title }));
      const withTitles = questions.filter((q) => q.title).length;
      console.log(`  (DOM found ${questions.length} question cards, ${withTitles} with titles)`);
    }
  } catch (e) { console.log('  DOM extraction failed:', e.message); }

  // Fallback: if DOM didn't work (virtualization / different selector), use HTML regex
  if (questions.length === 0) {
    try {
      const html = await page.content();
      const found = extractQuestionsFromHtml(html);
      if (found.length > 0) {
        questions = found;
        console.log(`  (Fallback HTML regex extracted ${found.length} questions)`);
      }
    } catch (e) { console.log('  HTML scrape failed:', e.message); }
  }

  // Fallback 2: call the Forms metadata API directly using the authenticated session
  if (questions.length === 0 && formId) {
    console.log('  Trying direct Forms API call...');
    try {
      const apiUrl = `https://forms.office.com/formapi/api/forms('${formId}')/questions`;
      const resp = await page.request.get(apiUrl, { headers: { Accept: 'application/json' } });
      if (resp.ok()) {
        const json = await resp.json();
        const found = findQuestions(json);
        if (found.length > 0) {
          questions = found;
          console.log(`  (API returned ${found.length} questions)`);
        } else {
          console.log('  API returned 200 but no questions found in payload. Keys:', Object.keys(json).join(','));
        }
      } else {
        console.log(`  API call returned HTTP ${resp.status()}`);
      }
    } catch (e) { console.log('  API call error:', e.message); }
  }

  await browser.close();
  return { formId, questions, capturedApiCalls: formApiPayloads.length };
}

/**
 * Extract question objects from the form editor HTML.
 * The editor embeds question data as escaped JSON inside script tags.
 * We look for patterns where a question ID (rXXX...) appears near a "title" field.
 */
function extractQuestionsFromHtml(html) {
  // Forms editor often embeds JSON like: {"id":"rXXX...","questionInfo":"{...\"title\":\"...\"}","title":"..."}
  // We search for id patterns and try to find the nearest title before or after.
  const idPattern = /"([a-z][a-f0-9]{31,32})"/g;
  const seen = new Set();
  const candidates = [];
  let m;
  while ((m = idPattern.exec(html)) !== null) {
    const id = m[1];
    if (!/^r[a-f0-9]{32}$/.test(id) || seen.has(id)) continue;
    seen.add(id);
    candidates.push({ id, pos: m.index });
  }

  // For each id, look for the closest "title":"..." within a 3000-char window.
  // Unescape common JSON escapes (\" and \\).
  const results = candidates.map(({ id, pos }) => {
    const windowStart = Math.max(0, pos - 1500);
    const windowEnd = Math.min(html.length, pos + 1500);
    const slice = html.substring(windowStart, windowEnd);
    // Look for title field; handle both "title":"..." and escaped \"title\":\"...\"
    const titleMatches = [
      /"title"\s*:\s*"((?:\\.|[^"\\])+?)"/g,
      /\\"title\\"\s*:\s*\\"((?:\\\\.|[^"\\])+?)\\"/g,
    ];
    let bestTitle = null;
    let bestDist = Infinity;
    for (const pattern of titleMatches) {
      pattern.lastIndex = 0;
      let tm;
      while ((tm = pattern.exec(slice)) !== null) {
        const absPos = windowStart + tm.index;
        const dist = Math.abs(absPos - pos);
        if (dist < bestDist) {
          bestDist = dist;
          bestTitle = tm[1];
        }
      }
    }
    // Clean up the title: unescape, strip HTML tags, limit length
    let title = null;
    if (bestTitle) {
      title = bestTitle
        .replace(/\\"/g, '"')
        .replace(/\\\\/g, '\\')
        .replace(/\\u([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)))
        .replace(/<[^>]+>/g, '')
        .trim();
      if (title.length > 500) title = title.substring(0, 500);
    }
    return { id, title };
  });

  // Filter: drop entries where title is obviously irrelevant (form title, section headers, etc.)
  // and dedupe. Then number them in document order.
  return results.map((q, i) => ({ number: i + 1, id: q.id, title: q.title }));
}

/**
 * Recursively search a JSON blob for a questions array.
 * Forms API shape: { questions: [{ questionId: "...", title: "...", ... }] }
 */
function findQuestions(obj, depth = 0) {
  if (!obj || depth > 6) return [];
  if (Array.isArray(obj)) {
    // Check if this looks like a questions array
    if (obj.length > 0 && obj[0] && (obj[0].questionId || obj[0].id)) {
      return obj.map((q, i) => ({
        number: i + 1,
        id: q.questionId || q.id,
        title: q.title || q.question || null,
      }));
    }
    for (const item of obj) {
      const r = findQuestions(item, depth + 1);
      if (r.length > 0) return r;
    }
  } else if (typeof obj === 'object') {
    if (obj.questions && Array.isArray(obj.questions)) {
      const r = findQuestions(obj.questions, depth + 1);
      if (r.length > 0) return r;
    }
    for (const key of Object.keys(obj)) {
      const r = findQuestions(obj[key], depth + 1);
      if (r.length > 0) return r;
    }
  }
  return [];
}

async function main() {
  const link = process.argv[2];
  if (!link) {
    console.error('Usage: node extract-form-ids.js "<response-link>"');
    process.exit(1);
  }

  const result = await extract(link);
  console.log('\n=== RESULT ===');
  console.log(JSON.stringify(result, null, 2));
  console.log(`\nFound: formId=${result.formId ? 'YES' : 'NO'}, questions=${result.questions.length}`);
}

module.exports = { extract };

if (require.main === module) main().catch((err) => { console.error('Fatal error:', err); process.exit(1); });
