/**
 * End-to-end: from a quiz JSON + Microsoft Forms template URL, produce a live
 * Power Automate flow. One browser session. No human-in-the-loop.
 *
 * Steps:
 *   1. QuizAutomation duplicates the template and fills in questions
 *   2. On the same page, click "Collect responses" and capture the response link
 *   3. Extract form ID + question IDs + titles from the editor DOM
 *   4. Match form questions to quiz JSON by title
 *   5. Navigate to Power Automate (same browser) and capture a bearer token
 *   6. Build flow definition and POST to the API
 *   7. Save everything to pipeline-results.json
 *
 * Usage:
 *   node end-to-end.js --quiz quiz_json/quiz_71.json --template-url "<forms-template-url>"
 */

const fs = require('fs');
const path = require('path');
const { QuizAutomation } = require('./automation');
const { buildFlowBody } = require('./create-flow');
const { matchQuestions, deriveShortCode } = require('./run-pipeline');

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
    const out = [];
    document.querySelectorAll('input, textarea').forEach((el) => {
      const v = el.value || '';
      if (v.includes('forms.') && (v.includes('/r/') || v.includes('ResponsePage'))) out.push(v);
    });
    return out;
  });
  if (links.length === 0) throw new Error('Response link not found in DOM after click');

  // Close the panel so subsequent DOM scraping for questions isn't obscured
  await page.keyboard.press('Escape').catch(() => {});
  await page.waitForTimeout(500);

  return links.find((l) => l.includes('ResponsePage')) || links[0];
}

async function captureCollaborationLink(page) {
  // The "Collaborate or Duplicate" button opens a panel with a share-to-edit link
  // (a URL containing &Token=... that lets other users edit the form)
  const triggers = [
    'button:has-text("Collaborate or Duplicate")',
    '[aria-label*="Collaborate or Duplicate"]',
    '[aria-label*="Collaborate"]',
    'button:has-text("Collaborate")',
  ];
  let opened = false;
  for (const sel of triggers) {
    try {
      const el = page.locator(sel).first();
      if (await el.count() > 0 && await el.isVisible({ timeout: 2500 })) {
        await el.click();
        opened = true;
        break;
      }
    } catch { /* try next */ }
  }
  if (!opened) return null; // optional — don't fail the pipeline if not found

  await page.waitForTimeout(2000);

  // The collaboration link contains &Token= in the URL
  const links = await page.evaluate(() => {
    const out = [];
    document.querySelectorAll('input, textarea').forEach((el) => {
      const v = el.value || '';
      if (v.includes('forms.') && v.includes('Token=')) out.push(v);
    });
    return out;
  });

  await page.keyboard.press('Escape').catch(() => {});
  await page.waitForTimeout(500);

  return links.length > 0 ? links[0] : null;
}

async function extractFromEditor(page) {
  const urlObj = new URL(page.url());
  const formId = urlObj.searchParams.get('id') || urlObj.searchParams.get('FormId');
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

async function runEndToEnd({ quizPath, templateUrl, existingEditorUrl, onEvent = () => {} }) {
  const quiz = JSON.parse(fs.readFileSync(quizPath, 'utf8'));
  const scorable = quiz.questions.filter((q) => !q.needs_review);
  console.log(`\n=== END-TO-END: ${quiz.form_name} (${scorable.length} scorable questions) ===\n`);
  onEvent({ type: 'e2e_start', formName: quiz.form_name, quizFilename: path.basename(quizPath), scorable: scorable.length });

  const bot = new QuizAutomation();
  bot.on('log', (m) => { console.log('  [automation]', m); onEvent({ type: 'log', message: m }); });
  bot.on('progress', (p) => {
    onEvent({ type: 'automation_progress', ...p });
    if (p.type === 'template_done') console.log(`  [progress] Template duplicated → ${p.newUrl}`);
    else if (p.type === 'form_done') console.log(`  [progress] Form filled: ${p.formName}`);
    else if (p.type === 'form_error') console.log(`  [progress] FORM ERROR: ${p.error}`);
  });

  await bot.launch();
  const page = bot.page;

  try {
    let editorUrl;
    if (existingEditorUrl) {
      onEvent({ type: 'e2e_stage', stage: 1, total: 6, label: 'Opening existing form' });
      editorUrl = existingEditorUrl;
      await page.goto(editorUrl, { waitUntil: 'networkidle', timeout: 60_000 });
      await page.waitForTimeout(3000);
    } else {
      onEvent({ type: 'e2e_stage', stage: 1, total: 6, label: 'Creating form + filling questions' });
      const results = await bot.processBatch([{ quizData: quiz }], { templateUrl, keepOpen: true });
      if (!results[0].success) throw new Error(`Form creation failed: ${results[0].error}`);
      editorUrl = results[0].newFormUrl;
      if (page.url() !== editorUrl) {
        await page.goto(editorUrl, { waitUntil: 'networkidle', timeout: 60_000 });
        await page.waitForTimeout(2000);
      }
    }
    onEvent({ type: 'e2e_form_ready', editorUrl });

    onEvent({ type: 'e2e_stage', stage: 2, total: 6, label: 'Capturing response + collaboration links' });
    const responseLink = await captureResponseLink(page);
    onEvent({ type: 'e2e_response_link', responseLink });
    const collaborationLink = await captureCollaborationLink(page).catch(() => null);
    if (collaborationLink) onEvent({ type: 'e2e_collaboration_link', collaborationLink });

    onEvent({ type: 'e2e_stage', stage: 3, total: 6, label: 'Extracting question IDs + titles' });
    const extracted = await extractFromEditor(page);
    if (!extracted.formId || extracted.questions.length === 0) throw new Error('Extraction returned no questions');
    if (extracted.questions.length !== scorable.length) {
      throw new Error(`Form has ${extracted.questions.length} questions, quiz JSON has ${scorable.length} scorable`);
    }
    onEvent({ type: 'e2e_extracted', formId: extracted.formId, questionCount: extracted.questions.length });

    onEvent({ type: 'e2e_stage', stage: 4, total: 6, label: 'Matching form to quiz' });
    const matches = matchQuestions(extracted.questions, scorable);
    const failed = matches.filter((m) => !m.quizQuestion || m.score === 0);
    if (failed.length > 0) throw new Error(`${failed.length} question(s) failed to match`);
    onEvent({ type: 'e2e_matched', count: matches.length });

    onEvent({ type: 'e2e_stage', stage: 5, total: 6, label: 'Authenticating to Power Automate' });
    await page.goto(`https://make.powerautomate.com/environments/${ENV_ID}/flows`);
    await page.waitForURL(/make\.powerautomate\.com.*\/flows/, { timeout: 300_000 });
    await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});
    const tokenPromise = captureBearerToken(page);
    await page.reload({ waitUntil: 'networkidle' }).catch(() => {});
    const token = await tokenPromise;

    onEvent({ type: 'e2e_stage', stage: 6, total: 6, label: 'Creating Power Automate flow' });
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

    const apiUrl = `${API_BASE}/providers/Microsoft.ProcessSimple/environments/${ENV_ID}/flows?api-version=${API_VERSION}`;
    const resp = await page.request.fetch(apiUrl, {
      method: 'POST',
      headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json' },
      data: JSON.stringify(body),
    });
    const status = resp.status();
    const text = await resp.text();
    if (status < 200 || status >= 300) throw new Error(`Flow creation failed: HTTP ${status}\n${text.substring(0, 1000)}`);
    const flowJson = JSON.parse(text);
    const flowId = flowJson.name;
    const flowUrl = `https://make.powerautomate.com/environments/${ENV_ID}/flows/${flowId}/details`;

    const result = {
      quiz: path.basename(quizPath),
      formName: quiz.form_name,
      editorUrl,
      responseLink,
      collaborationLink: collaborationLink || null,
      formId: extracted.formId,
      flowId,
      flowUrl,
      questionMap: matches.map((m) => ({
        formQuestionId: m.formQuestion.id,
        quizQuestionNumber: m.quizQuestion.number,
        correctAnswer: m.quizQuestion.options[m.quizQuestion.correct_answer_index],
      })),
      createdAt: new Date().toISOString(),
    };

    // Persistence handled by caller (server.js writes to quiz-runtime-state.json)

    console.log(`    ✅ Flow ID:  ${flowId}`);
    console.log(`    ✅ Flow:     ${flowUrl}`);
    onEvent({ type: 'e2e_done', flowId, flowUrl, responseLink, editorUrl, collaborationLink: collaborationLink || null, formId: extracted.formId });
    return result;
  } catch (err) {
    onEvent({ type: 'e2e_error', message: err.message });
    throw err;
  } finally {
    await bot.close().catch(() => {});
  }
}

async function main() {
  const args = {};
  for (let i = 2; i < process.argv.length; i += 2) {
    args[process.argv[i].replace(/^--/, '')] = process.argv[i + 1];
  }
  if (!args.quiz || (!args['template-url'] && !args['editor-url'])) {
    console.error('Usage: node end-to-end.js --quiz <path> (--template-url "<url>" | --editor-url "<url>")');
    process.exit(1);
  }
  await runEndToEnd({ quizPath: args.quiz, templateUrl: args['template-url'], existingEditorUrl: args['editor-url'] });
}

module.exports = { runEndToEnd };

if (require.main === module) main().catch((err) => { console.error('\n❌ Fatal:', err.message); process.exit(1); });
