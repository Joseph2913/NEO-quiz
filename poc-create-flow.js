/**
 * POC: create a single Power Automate flow via REST API.
 *
 * What it does:
 *   1. Opens Chromium with saved auth (auth-state-pa.json) or prompts sign-in
 *   2. Navigates to make.powerautomate.com to force an authenticated request
 *   3. Intercepts network traffic to capture a bearer token for *.api.flow.microsoft.com
 *   4. Builds a near-identical clone of the template flow definition
 *   5. PUTs it to a new client-generated GUID
 *   6. Prints the outcome
 *
 * Run: node poc-create-flow.js
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const crypto = require('crypto');

const AUTH_PATH = path.join(__dirname, 'auth-state-pa.json');
const ENV_ID = 'Default-371cfcf8-4cd8-4b79-87dc-de9a899f7a88';
const API_BASE = 'https://emea.api.flow.microsoft.com';
const API_VERSION = '2016-11-01';

// Three connections already authorized under joseph.thomas@oxygyconsulting.com
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

// The template's F2S definition, verbatim. Only displayName will be changed for this POC.
function buildFlowBody(displayName) {
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
              parameters: {
                form_id: '-PwcN9hMeUuH3N6aiZ96iJL6XI4jatJEuJk0OOXdqXtUMFJFWkUyQ0I5U0ZXWFZHM1BLWE8yWlo2WC4u',
              },
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
              parameters: {
                form_id: '-PwcN9hMeUuH3N6aiZ96iJL6XI4jatJEuJk0OOXdqXtUMFJFWkUyQ0I5U0ZXWFZHM1BLWE8yWlo2WC4u',
                response_id: "@triggerOutputs()?['body/resourceData/responseId']",
              },
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
            inputs: {
              A1: 'False',
              A2: 'Reste A Réserver - Represents the unreserved portion of the confirmed quantity for future periods.',
              A3: "NEO and KISS+ S&OP departments use reservation, while WISE departments' use quotas.",
              A4: 'Capped affiliates are prevented from accessing the stock pool.',
              A5: 'When performing manual arbitration/allocation.',
              A6: 'False',
              A7: 'Capping flag is set to trigger the creation of a PAL.',
            },
            runAfter: { Compose_Raw_Answers_JSON: ['Succeeded'] },
          },
          Compose_Answer_Check: {
            type: 'Compose',
            inputs: [
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A1'], outputs('Get_response_details')?['body']?['r9e858ca7c19f4b1aaadd803b507c3782'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A2'], outputs('Get_response_details')?['body']?['ref51db54af07401b8518e473cc8067d1'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A3'], outputs('Get_response_details')?['body']?['r1827e176a1554b7d8a48363f9cfa04aa'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A4'], outputs('Get_response_details')?['body']?['r6c99902c7df342359984064dafc8bf51'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A5'], outputs('Get_response_details')?['body']?['r836c1c7af5654272bd782a3143d2d855'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A6'], outputs('Get_response_details')?['body']?['r5845a764addf4ac9b9d2dd42e044c1cf'])", points: 1 },
              { answer: "@equals(outputs('Compose_Correct_Answers')?['A7'], outputs('Get_response_details')?['body']?['r3daf342123b5421ea7bd5520a8f14677'])", points: 1 },
            ],
            runAfter: { Compose_Correct_Answers: ['Succeeded'] },
          },
          Filter_Correct_Answers: {
            type: 'Query',
            inputs: {
              from: "@outputs('Compose_Answer_Check')",
              where: "@equals(item()?['answer'],true)",
            },
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
                dataset: 'https://oxygy.sharepoint.com/sites/NEOOnboarding-2023',
                table: '2935f221-9ac2-4e24-93c2-4c8d6a86bc52',
                'item/Title': "F2S-SOP-RESA-ATP-@{triggerOutputs()?['body/resourceData/responseId']}",
                'item/field_1': 'F2S-SOP-RESA-ATP',
                'item/field_2': 'F2S - S&OP RESA & Advanced ATP',
                'item/field_3': "@outputs('Get_user_profile_(V2)')?['body/displayName']",
                'item/field_4': "@outputs('Get_user_profile_(V2)')?['body/mail']",
                'item/field_5': "@outputs('Get_response_details')?['body/submitDate']",
                'item/field_7': 7,
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
      if (!settled) {
        settled = true;
        reject(new Error(`Token capture timed out after ${timeoutMs / 1000}s`));
      }
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
      } catch (e) { /* request may have been aborted, ignore */ }
    };

    page.on('request', handler);
  });
}

async function main() {
  const hasAuth = fs.existsSync(AUTH_PATH);
  console.log(`[1/6] Launching browser (auth: ${hasAuth ? 'reusing saved' : 'first run, sign in manually'})`);

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext(hasAuth ? { storageState: AUTH_PATH } : {});
  const page = await context.newPage();

  console.log('[2/6] Navigating to Power Automate. If prompted, sign in — script will wait up to 5 min.');
  await page.goto('https://make.powerautomate.com/environments/' + ENV_ID + '/flows');

  // Wait until we see the flows list (signals sign-in complete)
  console.log('[3/6] Waiting for authenticated session...');
  try {
    await page.waitForURL(/make\.powerautomate\.com.*\/flows/, { timeout: 300_000 });
    await page.waitForLoadState('networkidle', { timeout: 30_000 }).catch(() => {});
  } catch (e) {
    console.error('Timed out waiting for sign-in.');
    await browser.close();
    process.exit(1);
  }

  // Save auth for next run
  await context.storageState({ path: AUTH_PATH });
  console.log('[4/6] Auth saved. Capturing bearer token (triggering reload to force API call)...');

  // Arm the token listener, then reload to guarantee a fresh authenticated API request fires
  const tokenPromise = captureBearerToken(page, 90_000);
  await page.reload({ waitUntil: 'networkidle' }).catch(() => {});

  const token = await tokenPromise;
  console.log(`    Token captured (${token.length} chars, first 20: ${token.substring(0, 20)}...)`);

  const displayName = `CLONE_POC_TEST_${Date.now()}`;
  const body = buildFlowBody(displayName);

  const attempts = [
    {
      label: 'POST collection',
      method: 'POST',
      url: `${API_BASE}/providers/Microsoft.ProcessSimple/environments/${ENV_ID}/flows?api-version=${API_VERSION}`,
    },
    {
      label: 'PUT new GUID (fallback)',
      method: 'PUT',
      url: `${API_BASE}/providers/Microsoft.ProcessSimple/environments/${ENV_ID}/flows/${crypto.randomUUID()}?api-version=${API_VERSION}`,
    },
  ];

  console.log(`[5/6] Attempting flow creation. Display name: ${displayName}`);

  let success = false;
  for (const attempt of attempts) {
    console.log(`    → Trying ${attempt.label}: ${attempt.method} ${attempt.url.substring(0, 120)}...`);
    const response = await page.request.fetch(attempt.url, {
      method: attempt.method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Accept: 'application/json',
      },
      data: JSON.stringify(body),
    });

    const status = response.status();
    const respText = await response.text();
    console.log(`    ← HTTP ${status}`);

    if (status >= 200 && status < 300) {
      console.log(`    ✅ SUCCESS with ${attempt.label}`);
      try {
        const json = JSON.parse(respText);
        console.log(`    Flow ID: ${json.name || '(see body)'}`);
        console.log(`    Verify at: https://make.powerautomate.com/environments/${ENV_ID}/flows/${json.name}/details`);
      } catch {
        console.log('    (response was not JSON, full body below)');
        console.log(respText.substring(0, 1000));
      }
      success = true;
      break;
    } else {
      console.log(`    ❌ Failed. Body: ${respText.substring(0, 500)}`);
    }
  }

  console.log(`[6/6] Final: ${success ? 'SUCCESS' : 'ALL ATTEMPTS FAILED'}`);

  console.log('\nBrowser left open so you can inspect My Flows. Close it manually when done.');
}

main().catch((err) => {
  console.error('Fatal error:', err);
  process.exit(1);
});
