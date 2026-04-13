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

    const { items, template_url } = req.body;
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

        const results = await automation.processBatch(enrichedItems, {
          templateUrl: template_url || null,
        });

        // Update processing log
        const log = readProcessingLog();
        for (let ri = 0; ri < results.length; ri++) {
          const result = results[ri];
          const originalItem = enrichedItems[ri] || {};
          const formName = result.formName || result.form_name || (originalItem.quizData && originalItem.quizData.form_name);
          if (!formName) continue;
          log.entries[formName] = {
            status: result.success ? 'processed' : 'failed',
            form_url: result.newFormUrl || originalItem.form_url || null,
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

// GET /api/export
app.get('/api/export', async (req, res) => {
  try {
    const ExcelJS = require('exceljs');
    const manifest = readManifest();
    const log = readProcessingLog();

    const workbook = new ExcelJS.Workbook();
    workbook.creator = 'NEO Quiz Automation Tool';
    workbook.created = new Date();

    // ── Colors ──
    const headerFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1A1A2E' } };
    const headerFont = { bold: true, color: { argb: 'FFFFFFFF' }, size: 11 };
    const processedFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } };
    const failedFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFEE2E2' } };
    const pendingFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF1F5F9' } };
    const correctFill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFD1FAE5' } };
    const thinBorder = {
      top: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      bottom: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      left: { style: 'thin', color: { argb: 'FFE2E8F0' } },
      right: { style: 'thin', color: { argb: 'FFE2E8F0' } },
    };

    function styleHeader(sheet) {
      sheet.getRow(1).eachCell(cell => {
        cell.fill = headerFill;
        cell.font = headerFont;
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border = thinBorder;
      });
      sheet.getRow(1).height = 28;
    }

    // ═══════════════════════════════════════════════════════════════════════
    // Sheet 1: Summary
    // ═══════════════════════════════════════════════════════════════════════
    const summary = workbook.addWorksheet('Summary', {
      properties: { tabColor: { argb: 'FF2563EB' } },
    });
    summary.columns = [
      { header: '#', key: 'index', width: 5 },
      { header: 'Form Name', key: 'form_name', width: 48 },
      { header: 'Tribe', key: 'tribe', width: 10 },
      { header: 'Questions', key: 'question_count', width: 12 },
      { header: 'Status', key: 'status', width: 14 },
      { header: 'Form URL', key: 'form_url', width: 55 },
      { header: 'Processed Date', key: 'date', width: 22 },
      { header: 'Questions Done', key: 'questions_processed', width: 16 },
      { header: 'Topic Code', key: 'topic_code', width: 40 },
    ];
    styleHeader(summary);

    manifest.forms.forEach((form, i) => {
      const entry = log.entries[form.form_name] || {};
      const status = entry.status || 'not_started';
      const quizPath = path.join(QUIZ_JSON_DIR, form.filename);
      let topicCode = '';
      try {
        const qd = JSON.parse(fs.readFileSync(quizPath, 'utf-8'));
        topicCode = qd.topic_code || '';
      } catch (_) {}

      const row = summary.addRow({
        index: i + 1,
        form_name: form.form_name,
        tribe: form.tribe,
        question_count: form.question_count,
        status: status.replace('_', ' ').replace(/\b\w/g, c => c.toUpperCase()),
        form_url: entry.form_url || '',
        date: entry.date ? new Date(entry.date).toLocaleString() : '',
        questions_processed: entry.questions_processed || 0,
        topic_code: topicCode,
      });

      // Color the row based on status
      const fill = status === 'processed' ? processedFill : status === 'failed' ? failedFill : pendingFill;
      row.eachCell(cell => {
        cell.fill = fill;
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle', wrapText: true };
      });

      // Make URL clickable
      if (entry.form_url) {
        const urlCell = row.getCell('form_url');
        urlCell.value = { text: entry.form_url, hyperlink: entry.form_url };
        urlCell.font = { color: { argb: 'FF2563EB' }, underline: true };
      }
    });

    // Freeze header row
    summary.views = [{ state: 'frozen', ySplit: 1 }];

    // Auto-filter
    summary.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: manifest.forms.length + 1, column: 9 },
    };

    // ═══════════════════════════════════════════════════════════════════════
    // Sheet 2: All Questions (detailed)
    // ═══════════════════════════════════════════════════════════════════════
    const questions = workbook.addWorksheet('All Questions', {
      properties: { tabColor: { argb: 'FF16A34A' } },
    });
    questions.columns = [
      { header: 'Form Name', key: 'form_name', width: 40 },
      { header: 'Tribe', key: 'tribe', width: 8 },
      { header: 'Q#', key: 'q_number', width: 5 },
      { header: 'Question Text', key: 'text', width: 65 },
      { header: 'Type', key: 'type', width: 16 },
      { header: 'Option A', key: 'opt_a', width: 35 },
      { header: 'Option B', key: 'opt_b', width: 35 },
      { header: 'Option C', key: 'opt_c', width: 35 },
      { header: 'Option D', key: 'opt_d', width: 35 },
      { header: 'Correct', key: 'correct', width: 9 },
      { header: 'Needs Review', key: 'needs_review', width: 13 },
    ];
    styleHeader(questions);

    manifest.forms.forEach((form) => {
      const quizPath = path.join(QUIZ_JSON_DIR, form.filename);
      let quizData;
      try {
        quizData = JSON.parse(fs.readFileSync(quizPath, 'utf-8'));
      } catch (_) { return; }

      quizData.questions.forEach((q) => {
        const isTF = q.options.length === 2 &&
          q.options[0].toLowerCase() === 'true' &&
          q.options[1].toLowerCase() === 'false';
        const type = isTF ? 'True/False' :
          q.options.length === 1 ? 'Fill-in-the-blank' :
          `${q.options.length}-option MCQ`;

        const row = questions.addRow({
          form_name: form.form_name,
          tribe: form.tribe,
          q_number: q.number,
          text: q.text,
          type: type,
          opt_a: q.options[0] || '',
          opt_b: q.options[1] || '',
          opt_c: q.options[2] || '',
          opt_d: q.options[3] || '',
          correct: q.correct_answer_letter || '',
          needs_review: q.needs_review ? 'Yes' : '',
        });

        row.eachCell(cell => {
          cell.border = thinBorder;
          cell.alignment = { vertical: 'middle', wrapText: true };
        });

        // Highlight the correct answer cell
        const letters = ['opt_a', 'opt_b', 'opt_c', 'opt_d'];
        if (q.correct_answer_index >= 0 && q.correct_answer_index < letters.length) {
          const correctCell = row.getCell(letters[q.correct_answer_index]);
          correctCell.fill = correctFill;
          correctCell.font = { bold: true };
        }

        // Flag needs_review in orange
        if (q.needs_review) {
          row.getCell('needs_review').font = { color: { argb: 'FFD97706' }, bold: true };
        }
      });
    });

    questions.views = [{ state: 'frozen', ySplit: 1 }];
    questions.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: questions.rowCount, column: 11 },
    };

    // ═══════════════════════════════════════════════════════════════════════
    // Sheet 3: Statistics
    // ═══════════════════════════════════════════════════════════════════════
    const stats = workbook.addWorksheet('Statistics', {
      properties: { tabColor: { argb: 'FFD97706' } },
    });
    stats.columns = [
      { header: 'Metric', key: 'metric', width: 30 },
      { header: 'Value', key: 'value', width: 20 },
    ];
    styleHeader(stats);

    const totalForms = manifest.forms.length;
    const processedForms = manifest.forms.filter(f => (log.entries[f.form_name] || {}).status === 'processed').length;
    const failedForms = manifest.forms.filter(f => (log.entries[f.form_name] || {}).status === 'failed').length;
    const pendingForms = totalForms - processedForms - failedForms;
    let totalQuestions = 0;
    let totalProcessedQ = 0;
    manifest.forms.forEach(f => {
      totalQuestions += f.question_count;
      const entry = log.entries[f.form_name];
      if (entry) totalProcessedQ += entry.questions_processed || 0;
    });

    const tribes = [...new Set(manifest.forms.map(f => f.tribe))];
    const tribeStats = tribes.map(t => {
      const tribeForms = manifest.forms.filter(f => f.tribe === t);
      const tribeProcessed = tribeForms.filter(f => (log.entries[f.form_name] || {}).status === 'processed').length;
      return { tribe: t, total: tribeForms.length, processed: tribeProcessed };
    });

    const statsData = [
      { metric: 'Total Forms', value: totalForms },
      { metric: 'Processed', value: processedForms },
      { metric: 'Failed', value: failedForms },
      { metric: 'Pending', value: pendingForms },
      { metric: 'Completion Rate', value: `${totalForms > 0 ? Math.round((processedForms / totalForms) * 100) : 0}%` },
      { metric: '', value: '' },
      { metric: 'Total Questions (all forms)', value: totalQuestions },
      { metric: 'Questions Processed', value: totalProcessedQ },
      { metric: '', value: '' },
      { metric: 'Export Date', value: new Date().toLocaleString() },
      { metric: '', value: '' },
      { metric: '── By Tribe ──', value: '' },
    ];
    tribeStats.forEach(ts => {
      statsData.push({ metric: `${ts.tribe}`, value: `${ts.processed}/${ts.total} processed` });
    });

    statsData.forEach(d => {
      const row = stats.addRow(d);
      row.eachCell(cell => {
        cell.border = thinBorder;
        cell.alignment = { vertical: 'middle' };
      });
      if (d.metric.startsWith('──')) {
        row.eachCell(cell => { cell.font = { bold: true, color: { argb: 'FF64748B' } }; });
      }
    });

    // ── Send the file ──
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename="NEO_Quiz_Export_${new Date().toISOString().split('T')[0]}.xlsx"`);

    await workbook.xlsx.write(res);
    res.end();
  } catch (err) {
    console.error('Error in GET /api/export:', err);
    res.status(500).json({ error: 'Failed to generate export' });
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
