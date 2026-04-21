/**
 * Full end-to-end pipeline in a single browser session.
 *
 * Input:   an already-created Microsoft Forms editor URL + a quiz JSON file
 * Output:  a live Power Automate flow ingesting responses for that quiz
 *
 * Steps (all in one browser):
 *   1. Open the form editor
 *   2. Click "Collect responses" → capture the response link
 *   3. Extract form ID + question IDs + titles from the editor DOM
 *   4. Match form questions to quiz JSON by title similarity
 *   5. Navigate to Power Automate and capture a bearer token
 *   6. Build flow definition and POST to the API
 *   7. Report + optionally update a manifest
 *
 * Usage:
 *   node full-pipeline.js --quiz quiz_json/quiz_07.json --editor-url "<url>"
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const { buildFlowBody } = require('./create-flow');
const { matchQuestions, deriveShortCode } = require('./run-pipeline');

const AUTH_PATH = path.join(__dirname, 'auth-state-pa.json');
const ENV_ID = 'Default-371cfcf8-4cd8-4b79-87dc-de9a899f7a88';
const API_BASE = 'https://emea.api.flow.microsoft.com';
const API_VERSION = '2016-11-01';

async function captureResponseLink(page) {
  const selectors = [
    'button:has-text("Collect responses")',
    '[aria-label*="Collect responses"]',
  ];
  let clicked = false;
  for (const sel of selectors) {
    try {
      const el = page.locator(sel).first();
      if (await el.count() > 0 && await el.isVisible({ timeout: 3000 })) {
        await el.click();
        clicked = true;
        break;
      }
    } catch { /* try next */ }
  }
  if (!clicked) throw new Error('Collect responses button not found');

  await page.waitForTimeout(2500);

  const links = await page.evaluate(() => {
    const results = [];
    document.querySelectorAll('input, textarea').forEach((el) => {
      const v = el.value || '';
      if (v.includes('forms.') && (v.includes('/r/') || v.includes('ResponsePage'))) results.push(v);
    });
    return results;
  });
  if (links.length === 0) throw new Error('Response link not found after clicking Collect responses');
  // Prefer the ResponsePage version (more explicit)
  return links.find((l) => l.includes('ResponsePage')) || links[0];
}

async function extractFromEditor(page) {
  const formId = new URL(page.url()).searchParams.get('id') || new URL(page.url()).searchParams.get('FormId');
  const domQuestions = await page.evaluate(() => {
    const results = [];
    document.querySelectorAll('[id^="QuestionId_"]').forEach((el) => {
      const id = el.id.replace(/^QuestionId_/, '');
      const titleEl = el.querySelector('[data-automation-id="questionTitle"]');
      let title = null;
      if (titleEl) {
        const clone = titleEl.cloneNode(true);
        clone.querySelectorAll('span').forEach((s) => {
          if (/^\d+\.$/.test((s.textContent || '').trim())) s.remove();
        });
        title = (clone.innerText || clone.textContent || '').replace(/\s+/g, ' ').trim();
      }
      results.push({ id, title });
    });
    return results;
  });
  return { formId, questions: domQuestions.map((q, i) => ({ number: i + 1, id: q.id, title: q.title })) };
}

async function captureBearerToken(page, timeoutMs = 90_000) {
  return new Promise((resolve, reject) => {
    let settled = false;
    const timeout = setTimeout(() => {
      if (!settled) { settled = true; reject(new Error(`Token capture timed out after ${timeoutMs / 1000}s`)); }
    }, timeoutMs);
    const handler = async (req) => {
      if (settled) return;
      const url = req.url();
      if (!url.includes('api.flow.microsoft.com') && !url.includes('api.powerapps.com')) return;
      try {
        const headers = await req.allHeaders();
        const auth = headers['authorization'];
        if (auth && auth.toLowerCase().startsWith('bearer ')) {
          settled = true;
          clearTimeout(timeout);
          page.off('request', handler);
          resolve(auth.substring('Bearer '.length));
        }
      } catch { /* ignore */ }
    };
    page.on('request', handler);
  });
}

async function runFullPipeline({ quizPath, editorUrl, saveManifest = true }) {
  const quiz = JSON.parse(fs.readFileSync(quizPath, 'utf8'));
  const scorable = quiz.questions.filter((q) => !q.needs_review);
  console.log(`\nQuiz: ${quiz.form_name} (${scorable.length} scorable questions)\n`);

  const hasAuth = fs.existsSync(AUTH_PATH);
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasAuth ? { storageState: AUTH_PATH } : {});
  const page = await context.newPage();

  try {
    console.log('[1/6] Opening form editor...');
    await page.goto(editorUrl, { waitUntil: 'networkidle', timeout: 120_000 });
    await page.waitForURL(/forms\.(cloud\.microsoft|office\.com)/, { timeout: 300_000 });
    await page.waitForTimeout(3000);
    await context.storageState({ path: AUTH_PATH });

    console.log('[2/6] Capturing response link...');
    const responseLink = await captureResponseLink(page);
    console.log(`    ✅ ${responseLink}`);

    console.log('[3/6] Extracting form ID + question IDs + titles from editor DOM...');
    const extracted = await extractFromEditor(page);
    if (!extracted.formId || extracted.questions.length === 0) {
      throw new Error('Extraction returned no questions');
    }
    if (extracted.questions.length !== scorable.length) {
      throw new Error(`Form has ${extracted.questions.length} questions, quiz JSON has ${scorable.length} scorable`);
    }
    console.log(`    ✅ Form ID: ${extracted.formId.substring(0, 40)}...`);
    console.log(`    ✅ ${extracted.questions.length} question IDs with titles`);

    console.log('[4/6] Matching form questions to quiz JSON by title...');
    const matches = matchQuestions(extracted.questions, scorable);
    const failed = matches.filter((m) => !m.quizQuestion || m.score === 0);
    if (failed.length > 0) throw new Error(`${failed.length} question(s) failed to match`);
    matches.forEach((m, i) => {
      console.log(`    ✅ Form Q${i + 1} → ${m.quizQuestion.text.substring(0, 50)}... (score ${m.score})`);
    });

    console.log('[5/6] Navigating to Power Automate and capturing bearer token...');
    await page.goto(`https://make.powerautomate.com/environments/${ENV_ID}/flows`);
    await page.waitForURL(/make\.powerautomate\.com.*\/flows/, { timeout: 300_000 });
    await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});
    await context.storageState({ path: AUTH_PATH });
    const tokenPromise = captureBearerToken(page);
    await page.reload({ waitUntil: 'networkidle' }).catch(() => {});
    const token = await tokenPromise;
    console.log(`    ✅ Token captured (${token.length} chars)`);

    console.log('[6/6] Building and uploading flow...');
    const questionIds = matches.map((m) => m.formQuestion.id);
    const correctAnswers = matches.map((m) => m.quizQuestion.options[m.quizQuestion.correct_answer_index]);
    const shortCode = deriveShortCode(quiz);
    const displayName = `NEO_Quiz_${shortCode.replace(/-/g, '_')}`;

    const body = buildFlowBody({
      displayName,
      formId: extracted.formId,
      questionIds,
      correctAnswers,
      shortCode,
      topicName: quiz.form_name,
      totalPoints: scorable.length,
    });

    const url = `${API_BASE}/providers/Microsoft.ProcessSimple/environments/${ENV_ID}/flows?api-version=${API_VERSION}`;
    const resp = await page.request.fetch(url, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json' },
      data: JSON.stringify(body),
    });
    const status = resp.status();
    const text = await resp.text();
    if (status < 200 || status >= 300) {
      throw new Error(`Flow creation failed: HTTP ${status}\n${text.substring(0, 1000)}`);
    }
    const flowJson = JSON.parse(text);
    const flowId = flowJson.name;
    const flowUrl = `https://make.powerautomate.com/environments/${ENV_ID}/flows/${flowId}/details`;

    const result = {
      quiz: path.basename(quizPath),
      formName: quiz.form_name,
      editorUrl,
      responseLink,
      formId: extracted.formId,
      questionMap: matches.map((m) => ({ formQuestionId: m.formQuestion.id, quizQuestionNumber: m.quizQuestion.number, correctAnswer: m.quizQuestion.options[m.quizQuestion.correct_answer_index] })),
      flowId,
      flowUrl,
      createdAt: new Date().toISOString(),
    };

    if (saveManifest) {
      const outPath = path.join(__dirname, 'pipeline-results.json');
      let existing = {};
      if (fs.existsSync(outPath)) existing = JSON.parse(fs.readFileSync(outPath, 'utf8'));
      existing[path.basename(quizPath)] = result;
      fs.writeFileSync(outPath, JSON.stringify(existing, null, 2));
      console.log(`    ✅ Saved to pipeline-results.json`);
    }

    console.log(`\n✅ SUCCESS`);
    console.log(`   Flow ID: ${flowId}`);
    console.log(`   Flow:    ${flowUrl}`);
    console.log(`   Form:    ${responseLink}`);
    return result;
  } finally {
    await browser.close();
  }
}

async function main() {
  const args = {};
  for (let i = 2; i < process.argv.length; i += 2) {
    args[process.argv[i].replace(/^--/, '')] = process.argv[i + 1];
  }
  if (!args.quiz || !args['editor-url']) {
    console.error('Usage: node full-pipeline.js --quiz <path> --editor-url "<url>"');
    process.exit(1);
  }
  await runFullPipeline({ quizPath: args.quiz, editorUrl: args['editor-url'] });
}

module.exports = { runFullPipeline };

if (require.main === module) main().catch((err) => { console.error('\n❌ Fatal:', err.message); process.exit(1); });
