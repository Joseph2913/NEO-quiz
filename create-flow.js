/**
 * Stage 4: Flow builder.
 *
 * Takes a quiz JSON + form ID + question IDs, builds the Power Automate flow
 * definition from the template, and uploads it via REST API.
 *
 * Usage:
 *   node create-flow.js \
 *     --quiz quiz_json/quiz_01.json \
 *     --form-id "-PwcN9hMeUuH3N6..." \
 *     --question-ids "r0ee...,r1c7...,..." \
 *     [--short "F2S-SOP-RESA-ATP"] \
 *     [--topic "F2S - S&OP RESA & Advanced ATP"] \
 *     [--name "NEO_Quiz_F2S_SOP_RESA_ATP"]
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');

const AUTH_PATH = path.join(__dirname, 'auth-state-pa.json');
const ENV_ID = 'Default-371cfcf8-4cd8-4b79-87dc-de9a899f7a88';
const API_BASE = 'https://emea.api.flow.microsoft.com';
const API_VERSION = '2016-11-01';

const CONNECTIONS = {
  shared_microsoftforms: {
    id: '/providers/Microsoft.PowerApps/apis/shared_microsoftforms',
    connectionName: '9db533a0149545a6aa1d4eca3d54e9cd',
    source: 'Embedded',
  },
  shared_office365users: {
    id: '/providers/Microsoft.PowerApps/apis/shared_office365users',
    connectionName: 'shared-office365user-d258d9ce-90cb-4b58-aff7-dd3a91df1f39',
    source: 'Embedded',
  },
  shared_sharepointonline: {
    id: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
    connectionName: 'shared-sharepointonl-b7ece7a0-68ae-4495-9939-b6e95d05643b',
    source: 'Embedded',
  },
};

const SP_DATASET = 'https://oxygy.sharepoint.com/sites/NEOOnboarding-2023';
const SP_TABLE = '2935f221-9ac2-4e24-93c2-4c8d6a86bc52';

/**
 * Build the flow definition body from per-quiz inputs.
 *
 * @param {Object} opts
 * @param {string} opts.displayName         Flow name shown in Power Automate
 * @param {string} opts.formId              Microsoft Forms internal form ID
 * @param {string[]} opts.questionIds       Forms question IDs, in question-number order
 * @param {string[]} opts.correctAnswers    Correct answer strings, in question-number order
 * @param {string} opts.shortCode           Shorthand Form ID (e.g. "F2S-SOP-RESA-ATP")
 * @param {string} opts.topicName           Human-readable topic
 * @param {number} opts.totalPoints         Total scorable questions
 */
function buildFlowBody(opts) {
  const {
    displayName, formId, questionIds, correctAnswers,
    shortCode, topicName, totalPoints,
  } = opts;

  if (questionIds.length !== correctAnswers.length) {
    throw new Error(`Mismatch: ${questionIds.length} question IDs but ${correctAnswers.length} correct answers`);
  }

  // Build Compose Correct Answers: { A1: "...", A2: "...", ... }
  const composeCorrect = {};
  correctAnswers.forEach((ans, i) => { composeCorrect[`A${i + 1}`] = ans; });

  // Build Compose Answer Check: [{ answer: equals(...), points: 1 }, ...]
  const composeCheck = questionIds.map((qid, i) => ({
    answer: `@equals(outputs('Compose_Correct_Answers')?['A${i + 1}'], outputs('Get_response_details')?['body']?['${qid}'])`,
    points: 1,
  }));

  return {
    properties: {
      displayName,
      definition: {
        $schema: 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#',
        contentVersion: '1.0.0.0',
        parameters: {
          $connections: { defaultValue: {}, type: 'Object' },
          $authentication: { defaultValue: {}, type: 'SecureObject' },
        },
        triggers: {
          When_a_new_response_is_submitted: {
            type: 'OpenApiConnectionWebhook',
            inputs: {
              host: {
                connectionName: 'shared_microsoftforms',
                operationId: 'CreateFormWebhook',
                apiId: '/providers/Microsoft.PowerApps/apis/shared_microsoftforms',
              },
              parameters: { form_id: formId },
              authentication: "@parameters('$authentication')",
            },
            splitOn: "@triggerOutputs()?['body/value']",
          },
        },
        actions: {
          Get_response_details: {
            type: 'OpenApiConnection',
            inputs: {
              host: {
                connectionName: 'shared_microsoftforms',
                operationId: 'GetFormResponseById',
                apiId: '/providers/Microsoft.PowerApps/apis/shared_microsoftforms',
              },
              parameters: { form_id: formId, response_id: "@triggerOutputs()?['body/resourceData/responseId']" },
              authentication: "@parameters('$authentication')",
            },
            runAfter: {},
          },
          'Get_user_profile_(V2)': {
            type: 'OpenApiConnection',
            inputs: {
              host: {
                connectionName: 'shared_office365users',
                operationId: 'UserProfile_V2',
                apiId: '/providers/Microsoft.PowerApps/apis/shared_office365users',
              },
              parameters: { id: "@outputs('Get_response_details')?['body/responder']" },
              authentication: "@parameters('$authentication')",
            },
            runAfter: { Get_response_details: ['Succeeded'] },
          },
          Compose_Raw_Answers_JSON: {
            type: 'Compose',
            inputs: "@body('Get_response_details')",
            runAfter: { 'Get_user_profile_(V2)': ['Succeeded'] },
            runtimeConfiguration: { secureData: { properties: ['inputs'] } },
          },
          Compose_Correct_Answers: {
            type: 'Compose',
            inputs: composeCorrect,
            runAfter: { Compose_Raw_Answers_JSON: ['Succeeded'] },
          },
          Compose_Answer_Check: {
            type: 'Compose',
            inputs: composeCheck,
            runAfter: { Compose_Correct_Answers: ['Succeeded'] },
          },
          Filter_Correct_Answers: {
            type: 'Query',
            inputs: { from: "@outputs('Compose_Answer_Check')", where: "@equals(item()?['answer'],true)" },
            runAfter: { Compose_Answer_Check: ['Succeeded'] },
          },
          Create_item: {
            type: 'OpenApiConnection',
            inputs: {
              host: {
                connectionName: 'shared_sharepointonline',
                operationId: 'PostItem',
                apiId: '/providers/Microsoft.PowerApps/apis/shared_sharepointonline',
              },
              parameters: {
                dataset: SP_DATASET,
                table: SP_TABLE,
                'item/Title': `${shortCode}-@{triggerOutputs()?['body/resourceData/responseId']}`,
                'item/field_1': shortCode,
                'item/field_2': topicName,
                'item/field_3': "@outputs('Get_user_profile_(V2)')?['body/displayName']",
                'item/field_4': "@outputs('Get_user_profile_(V2)')?['body/mail']",
                'item/field_5': "@outputs('Get_response_details')?['body/submitDate']",
                'item/field_7': totalPoints,
                'item/field_8': "@length(body('Filter_Correct_Answers'))",
                'item/field_11': "@outputs('Compose_Raw_Answers_JSON')",
              },
              authentication: "@parameters('$authentication')",
            },
            runAfter: { Filter_Correct_Answers: ['Succeeded'] },
          },
        },
        outputs: {},
      },
      connectionReferences: CONNECTIONS,
      parameters: { $connections: { value: CONNECTIONS } },
      environment: { name: ENV_ID },
    },
  };
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

async function uploadFlow(body) {
  const hasAuth = fs.existsSync(AUTH_PATH);
  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasAuth ? { storageState: AUTH_PATH } : {});
  const page = await context.newPage();

  await page.goto(`https://make.powerautomate.com/environments/${ENV_ID}/flows`);
  await page.waitForURL(/make\.powerautomate\.com.*\/flows/, { timeout: 300_000 });
  await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});
  await context.storageState({ path: AUTH_PATH });

  const tokenPromise = captureBearerToken(page);
  await page.reload({ waitUntil: 'networkidle' }).catch(() => {});
  const token = await tokenPromise;

  const url = `${API_BASE}/providers/Microsoft.ProcessSimple/environments/${ENV_ID}/flows?api-version=${API_VERSION}`;
  const resp = await page.request.fetch(url, {
    method: 'POST',
    headers: { Authorization: `Bearer ${token}`, 'Content-Type': 'application/json', Accept: 'application/json' },
    data: JSON.stringify(body),
  });

  const status = resp.status();
  const text = await resp.text();
  await browser.close();

  if (status >= 200 && status < 300) {
    const json = JSON.parse(text);
    return { success: true, flowId: json.name, url: `https://make.powerautomate.com/environments/${ENV_ID}/flows/${json.name}/details` };
  }
  return { success: false, status, body: text };
}

function parseArgs() {
  const args = {};
  for (let i = 2; i < process.argv.length; i += 2) {
    const key = process.argv[i].replace(/^--/, '');
    args[key] = process.argv[i + 1];
  }
  return args;
}

function deriveShortCode(quiz) {
  // e.g. "F2S - S&OP RESA & Advanced ATP" → "F2S-SOP-RESA-ATP"
  return quiz.form_name
    .replace(/&/g, '')
    .replace(/[^a-zA-Z0-9 -]/g, '')
    .replace(/\s+/g, '-')
    .replace(/-+/g, '-')
    .toUpperCase()
    .substring(0, 60);
}

async function main() {
  const args = parseArgs();
  if (!args.quiz || !args['form-id'] || !args['question-ids']) {
    console.error('Usage: node create-flow.js --quiz <path> --form-id <id> --question-ids "id1,id2,..." [--short X] [--topic Y] [--name Z]');
    process.exit(1);
  }

  const quiz = JSON.parse(fs.readFileSync(args.quiz, 'utf8'));
  const questionIds = args['question-ids'].split(',').map((s) => s.trim());

  // Derive correct answers in question-number order
  const scorable = quiz.questions.filter((q) => !q.needs_review);
  if (scorable.length !== questionIds.length) {
    console.warn(`⚠️  Quiz has ${scorable.length} scorable questions but you provided ${questionIds.length} question IDs`);
  }
  const correctAnswers = scorable.map((q) => q.options[q.correct_answer_index]);

  const shortCode = args.short || deriveShortCode(quiz);
  const topicName = args.topic || quiz.form_name;
  const displayName = args.name || `NEO_Quiz_${shortCode.replace(/-/g, '_')}`;

  console.log('Building flow with:');
  console.log(`  Display name:  ${displayName}`);
  console.log(`  Short code:    ${shortCode}`);
  console.log(`  Topic:         ${topicName}`);
  console.log(`  Form ID:       ${args['form-id'].substring(0, 40)}...`);
  console.log(`  Questions:     ${questionIds.length}`);
  console.log(`  Total points:  ${scorable.length}`);

  const body = buildFlowBody({
    displayName,
    formId: args['form-id'],
    questionIds,
    correctAnswers,
    shortCode,
    topicName,
    totalPoints: scorable.length,
  });

  console.log('\nUploading via Power Automate API...');
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

module.exports = { buildFlowBody, uploadFlow };

if (require.main === module) main().catch((err) => { console.error('Fatal:', err); process.exit(1); });
