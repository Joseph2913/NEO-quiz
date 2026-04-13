const express = require('express');
const fs = require('fs');
const path = require('path');
const { exec } = require('child_process');

const app = express();
const PORT = 3000;

const PROJECT_ROOT = __dirname;
const QUIZ_JSON_DIR = path.join(PROJECT_ROOT, 'quiz_json');
const MANIFEST_PATH = path.join(QUIZ_JSON_DIR, 'quiz_manifest.json');
const LOG_PATH = path.join(PROJECT_ROOT, 'processing-log.json');

app.use(express.json());
app.use(express.static(path.join(PROJECT_ROOT, 'public')));

// --- Helpers ---

function readProcessingLog() {
  try {
    const data = fs.readFileSync(LOG_PATH, 'utf-8');
    return JSON.parse(data);
  } catch {
    const empty = { entries: {} };
    fs.writeFileSync(LOG_PATH, JSON.stringify(empty, null, 2));
    return empty;
  }
}

function writeProcessingLog(log) {
  fs.writeFileSync(LOG_PATH, JSON.stringify(log, null, 2));
}

function readManifest() {
  return JSON.parse(fs.readFileSync(MANIFEST_PATH, 'utf-8'));
}

// --- SSE ---

let sseClients = [];

function sendSSE(data) {
  const msg = `data: ${JSON.stringify(data)}\n\n`;
  for (const res of sseClients) {
    res.write(msg);
  }
}

// --- Processing state ---

let processingRunning = false;

// --- Routes ---

// GET /api/status
app.get('/api/status', (req, res) => {
  res.json({ running: processingRunning, platform: process.platform });
});

// GET /api/quizzes
app.get('/api/quizzes', (req, res) => {
  try {
    const manifest = readManifest();
    const log = readProcessingLog();

    const quizzes = manifest.forms.map((form) => {
      const entry = log.entries[form.form_name];
      return {
        index: form.index,
        filename: form.filename,
        form_name: form.form_name,
        tribe: form.tribe,
        question_count: form.question_count,
        status: entry ? entry.status : 'not_started',
        form_url: entry ? entry.form_url || null : null,
        date: entry ? entry.date || null : null,
      };
    });

    res.json(quizzes);
  } catch (err) {
    console.error('Error in GET /api/quizzes:', err);
    res.status(500).json({ error: 'Failed to load quizzes' });
  }
});

// GET /api/quiz/:filename
app.get('/api/quiz/:filename', (req, res) => {
  try {
    const { filename } = req.params;
    const filePath = path.join(QUIZ_JSON_DIR, filename);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'Quiz file not found' });
    }

    const data = JSON.parse(fs.readFileSync(filePath, 'utf-8'));
    res.json(data);
  } catch (err) {
    console.error('Error in GET /api/quiz/:filename:', err);
    res.status(500).json({ error: 'Failed to load quiz data' });
  }
});

// PUT /api/quiz/:filename
app.put('/api/quiz/:filename', (req, res) => {
  try {
    const { filename } = req.params;

    if (!/^quiz_\d+\.json$/.test(filename)) {
      return res.status(400).json({ error: 'Invalid filename format. Must match quiz_\\d+\\.json' });
    }

    const filePath = path.join(QUIZ_JSON_DIR, filename);

    if (!fs.existsSync(filePath)) {
      return res.status(404).json({ error: 'Quiz file not found' });
    }

    fs.writeFileSync(filePath, JSON.stringify(req.body, null, 2));
    res.json({ saved: true });
  } catch (err) {
    console.error('Error in PUT /api/quiz/:filename:', err);
    res.status(500).json({ error: 'Failed to save quiz data' });
  }
});

// POST /api/process
app.post('/api/process', (req, res) => {
  try {
    if (processingRunning) {
      return res.status(409).json({ error: 'A batch is already running' });
    }

    const { items } = req.body;
    if (!items || !Array.isArray(items) || items.length === 0) {
      return res.status(400).json({ error: 'items array is required' });
    }

    processingRunning = true;
    res.json({ started: true });

    // Run processing asynchronously
    (async () => {
      try {
        const { QuizAutomation } = require('./automation.js');
        const manifest = readManifest();

        // Build items with quiz data
        const enrichedItems = items.map((item) => {
          const form = manifest.forms.find((f) => f.form_name === item.form_name);
          if (!form) return item;

          const quizData = JSON.parse(
            fs.readFileSync(path.join(QUIZ_JSON_DIR, form.filename), 'utf-8')
          );
          return { ...item, quizData };
        });

        const automation = new QuizAutomation();

        automation.on('progress', (event) => {
          sendSSE(event);
        });

        const results = await automation.processBatch(enrichedItems);

        // Update processing log
        const log = readProcessingLog();
        for (let ri = 0; ri < results.length; ri++) {
          const result = results[ri];
          const originalItem = enrichedItems[ri] || {};
          const formName = result.formName || result.form_name || (originalItem.quizData && originalItem.quizData.form_name);
          if (!formName) continue;
          log.entries[formName] = {
            status: result.success ? 'processed' : 'failed',
            form_url: originalItem.form_url || null,
            date: new Date().toISOString(),
            questions_processed: result.success ? (originalItem.quizData ? originalItem.quizData.question_count : 0) : 0,
            questions_total: originalItem.quizData ? originalItem.quizData.question_count : 0,
          };
        }
        writeProcessingLog(log);

        sendSSE({ type: 'batch_complete', results });
      } catch (err) {
        console.error('Processing error:', err);
        sendSSE({ type: 'error', message: err.message });
      } finally {
        processingRunning = false;
      }
    })();
  } catch (err) {
    console.error('Error in POST /api/process:', err);
    processingRunning = false;
    res.status(500).json({ error: 'Failed to start processing' });
  }
});

// GET /api/progress (SSE)
app.get('/api/progress', (req, res) => {
  res.writeHead(200, {
    'Content-Type': 'text/event-stream',
    'Cache-Control': 'no-cache',
    Connection: 'keep-alive',
  });

  res.write('\n');
  sseClients.push(res);

  // Heartbeat every 15 seconds
  const heartbeat = setInterval(() => {
    res.write(': heartbeat\n\n');
  }, 15000);

  req.on('close', () => {
    clearInterval(heartbeat);
    sseClients = sseClients.filter((c) => c !== res);
  });
});

// DELETE /api/log/:formName
app.delete('/api/log/:formName', (req, res) => {
  try {
    const formName = decodeURIComponent(req.params.formName);
    const log = readProcessingLog();

    if (!log.entries[formName]) {
      return res.status(404).json({ error: 'Log entry not found' });
    }

    delete log.entries[formName];
    writeProcessingLog(log);
    res.json({ deleted: true });
  } catch (err) {
    console.error('Error in DELETE /api/log/:formName:', err);
    res.status(500).json({ error: 'Failed to delete log entry' });
  }
});

// --- Start server ---

app.listen(PORT, () => {
  const url = `http://localhost:${PORT}`;
  console.log(`NEO Quiz Automation Tool running at ${url}`);

  // Auto-open browser
  const platform = process.platform;
  if (platform === 'darwin') {
    exec(`open ${url}`);
  } else if (platform === 'win32') {
    exec(`start ${url}`);
  } else {
    exec(`xdg-open ${url}`);
  }
});
