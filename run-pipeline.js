/**
 * End-to-end pipeline: extract form IDs → match questions → build & upload flow.
 *
 * Usage:
 *   node run-pipeline.js --quiz quiz_json/quiz_07.json --form-link "<response-link>"
 */

const fs = require('fs');
const { extract } = require('./extract-form-ids');
const { buildFlowBody, uploadFlow } = require('./create-flow');

function normalize(s) {
  return (s || '')
    .toLowerCase()
    .replace(/[\u2018\u2019]/g, "'")  // curly single → straight
    .replace(/[\u201C\u201D]/g, '"')  // curly double → straight
    .replace(/[^\w\s]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

/**
 * Match each form question (with title) to a quiz JSON question.
 * Strategy: normalize both sides, find the quiz question whose text is
 * a substring of (or most overlaps with) the form title.
 * Returns an array parallel to form.questions: [{ formQuestion, quizQuestion }]
 */
function matchQuestions(formQuestions, quizQuestions) {
  const normalizedQuiz = quizQuestions.map((q) => ({ q, norm: normalize(q.text) }));
  const results = [];
  const usedQuizIndices = new Set();

  for (const fq of formQuestions) {
    const formNorm = normalize(fq.title);
    let best = null;
    let bestScore = 0;
    normalizedQuiz.forEach(({ q, norm }, idx) => {
      if (usedQuizIndices.has(idx)) return;
      // Substring match wins decisively
      let score = 0;
      if (formNorm.includes(norm)) score = norm.length;
      else if (norm.includes(formNorm)) score = formNorm.length;
      else {
        // Word overlap count as tiebreaker
        const formWords = new Set(formNorm.split(' '));
        const quizWords = norm.split(' ');
        score = quizWords.filter((w) => formWords.has(w)).length;
      }
      if (score > bestScore) { bestScore = score; best = { q, idx }; }
    });

    if (!best) {
      results.push({ formQuestion: fq, quizQuestion: null, score: 0 });
    } else {
      usedQuizIndices.add(best.idx);
      results.push({ formQuestion: fq, quizQuestion: best.q, score: bestScore });
    }
  }
  return results;
}

function deriveShortCode(quiz) {
  return quiz.form_name
    .replace(/&/g, '')
    .replace(/[^a-zA-Z0-9 -]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .toUpperCase()
    .substring(0, 60);
}

async function main() {
  const args = {};
  for (let i = 2; i < process.argv.length; i += 2) {
    args[process.argv[i].replace(/^--/, '')] = process.argv[i + 1];
  }
  if (!args.quiz || !args['form-link']) {
    console.error('Usage: node run-pipeline.js --quiz <path> --form-link "<url>"');
    process.exit(1);
  }

  const quiz = JSON.parse(fs.readFileSync(args.quiz, 'utf8'));
  const scorable = quiz.questions.filter((q) => !q.needs_review);
  console.log(`Quiz: ${quiz.form_name} (${scorable.length} scorable questions)\n`);

  console.log('[1/3] Extracting form ID + question IDs + titles...');
  const extracted = await extract(args['form-link']);
  if (!extracted.formId || extracted.questions.length === 0) {
    console.error('  ❌ Extraction failed.');
    process.exit(1);
  }
  console.log(`  ✅ Form ID: ${extracted.formId.substring(0, 40)}...`);
  console.log(`  ✅ Captured ${extracted.questions.length} question IDs with titles\n`);

  if (extracted.questions.length !== scorable.length) {
    console.error(`  ❌ Mismatch: form has ${extracted.questions.length} questions, quiz JSON has ${scorable.length} scorable.`);
    process.exit(1);
  }

  console.log('[2/3] Matching form questions to quiz JSON by title similarity...');
  const matches = matchQuestions(extracted.questions, scorable);
  let matchFail = false;
  matches.forEach((m, i) => {
    const formText = m.formQuestion.title ? m.formQuestion.title.substring(0, 60) : '(no title)';
    const quizText = m.quizQuestion ? m.quizQuestion.text.substring(0, 60) : '(unmatched)';
    const status = m.quizQuestion && m.score > 0 ? '✅' : '❌';
    console.log(`  ${status} Form Q${i + 1} [${m.formQuestion.id.substring(0, 10)}...]`);
    console.log(`     Form:  "${formText}..."`);
    console.log(`     Quiz:  "${quizText}..." (score: ${m.score})`);
    if (!m.quizQuestion || m.score === 0) matchFail = true;
  });
  if (matchFail) {
    console.error('\n❌ One or more questions failed to match. Aborting.');
    process.exit(1);
  }
  console.log('\n');

  console.log('[3/3] Building and uploading flow...');
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

  const result = await uploadFlow(body);
  if (result.success) {
    console.log(`\n✅ SUCCESS`);
    console.log(`   Flow ID: ${result.flowId}`);
    console.log(`   URL:     ${result.url}`);
  } else {
    console.log(`\n❌ FAILED (HTTP ${result.status})`);
    console.log(result.body.substring(0, 1500));
    process.exit(1);
  }
}

module.exports = { matchQuestions, normalize, deriveShortCode };

if (require.main === module) main().catch((err) => { console.error('Fatal:', err); process.exit(1); });
