/* ─── NEO Quiz Automation — Frontend ──────────────────────────────────────── */

(function () {
  'use strict';

  // ─── State ───────────────────────────────────────────────────────────────
  let quizzes = [];
  let currentQuiz = null;      // { filename, quizData }
  let currentView = 'dashboard';
  let eventSource = null;      // SSE connection
  let processFormName = null;  // form being processed (single mode)
  let bulkItems = [];          // parsed bulk items

  // ─── DOM refs ────────────────────────────────────────────────────────────
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => document.querySelectorAll(sel);

  // ─── Init ────────────────────────────────────────────────────────────────
  async function init() {
    await fetchStatus();
    await loadQuizzes();
    bindEvents();
    connectSSE();
    startAutoRefresh();
  }

  // ─── Auto-refresh quiz data periodically ────────────────────────────────
  let autoRefreshTimer = null;
  function startAutoRefresh() {
    if (autoRefreshTimer) clearInterval(autoRefreshTimer);
    autoRefreshTimer = setInterval(async () => {
      try {
        await loadQuizzes();
        // If viewing a quiz detail, refresh that too
        if (currentQuiz && currentView === 'quiz') {
          const updated = quizzes.find(q => q.filename === currentQuiz.filename);
          if (updated && updated.status !== currentQuiz.meta.status) {
            currentQuiz.meta = updated;
            renderQuizDetail();
          }
        }
      } catch (_) { /* ignore fetch errors during polling */ }
    }, 10000);
  }

  // ─── API helpers ─────────────────────────────────────────────────────────
  async function api(path, options) {
    const res = await fetch('/api' + path, options);
    if (!res.ok) {
      const text = await res.text();
      throw new Error(text || res.statusText);
    }
    return res.json();
  }

  async function fetchStatus() {
    try {
      const status = await api('/status');
      $('#platform-badge').textContent = status.platform === 'darwin' ? 'macOS' : status.platform === 'win32' ? 'Windows' : 'Linux';
      updateRunningStatus(status.running);
    } catch (e) { console.error('Status fetch failed:', e); }
  }

  function updateRunningStatus(running) {
    const badge = $('#status-badge');
    if (running) {
      badge.textContent = 'Running';
      badge.className = 'badge badge-running';
    } else {
      badge.textContent = 'Idle';
      badge.className = 'badge badge-idle';
    }
  }

  // ─── Load quizzes ────────────────────────────────────────────────────────
  async function loadQuizzes() {
    quizzes = await api('/quizzes');
    renderQuizList();
    renderDashboard();
    populateTribeFilter();
  }

  function populateTribeFilter() {
    const tribes = [...new Set(quizzes.map(q => q.tribe))].sort();
    const sel = $('#filter-tribe');
    sel.innerHTML = '<option value="">All Tribes</option>';
    tribes.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t;
      opt.textContent = t;
      sel.appendChild(opt);
    });
  }

  // ─── Render quiz list ────────────────────────────────────────────────────
  function renderQuizList() {
    const search = ($('#search-input').value || '').toLowerCase();
    const tribe = $('#filter-tribe').value;
    const status = $('#filter-status').value;

    const filtered = quizzes.filter(q => {
      if (search && !q.form_name.toLowerCase().includes(search)) return false;
      if (tribe && q.tribe !== tribe) return false;
      if (status && q.status !== status) return false;
      return true;
    });

    const list = $('#quiz-list');
    list.innerHTML = '';
    filtered.forEach(q => {
      const li = document.createElement('li');
      li.className = currentQuiz && currentQuiz.filename === q.filename ? 'active' : '';
      li.innerHTML = `
        <span class="quiz-name" title="${esc(q.form_name)}">${esc(q.form_name)}</span>
        <span class="quiz-count">${q.question_count}q</span>
        <span class="status-dot ${q.status}"></span>
      `;
      li.addEventListener('click', () => selectQuiz(q));
      list.appendChild(li);
    });

    // Stats
    const total = quizzes.length;
    const processed = quizzes.filter(q => q.status === 'processed').length;
    const pending = quizzes.filter(q => q.status === 'not_started').length;
    $('#stat-total').textContent = `${total} quizzes`;
    $('#stat-processed').textContent = `${processed} processed`;
    $('#stat-pending').textContent = `${pending} pending`;
  }

  // ─── Render dashboard ────────────────────────────────────────────────────
  function renderDashboard() {
    const total = quizzes.length;
    const processed = quizzes.filter(q => q.status === 'processed').length;
    const failed = quizzes.filter(q => q.status === 'failed').length;
    const pending = total - processed - failed;
    $('#dash-total').textContent = total;
    $('#dash-processed').textContent = processed;
    $('#dash-pending').textContent = pending;
    $('#dash-failed').textContent = failed;
  }

  // ─── View switching ──────────────────────────────────────────────────────
  function showView(name) {
    currentView = name;
    $$('.view').forEach(v => v.classList.remove('active'));
    $(`#view-${name}`).classList.add('active');
    // Refresh data when navigating to dashboard or quiz list
    if (name === 'dashboard' || name === 'quiz') {
      loadQuizzes();
    }
  }

  // ─── Select a quiz ──────────────────────────────────────────────────────
  async function selectQuiz(quizMeta) {
    try {
      const quizData = await api(`/quiz/${quizMeta.filename}`);
      currentQuiz = { filename: quizMeta.filename, quizData, meta: quizMeta };
      renderQuizDetail();
      showView('quiz');
      renderQuizList(); // update active highlight
    } catch (e) {
      alert('Failed to load quiz: ' + e.message);
    }
  }

  // ─── Render quiz detail / editor ─────────────────────────────────────────
  function renderQuizDetail() {
    const { quizData, meta } = currentQuiz;

    $('#quiz-title').textContent = quizData.form_name;
    $('#quiz-meta').textContent = `${quizData.tribe} | ${quizData.question_count} questions | ${quizData.topic_code}`;

    // Status badge
    const statusBadge = $('#quiz-status-badge');
    statusBadge.textContent = meta.status.replace('_', ' ');
    statusBadge.className = `badge badge-${meta.status}`;

    // Form link
    const formLink = $('#quiz-form-link');
    if (meta.form_url) {
      formLink.href = meta.form_url;
      formLink.classList.remove('hidden');
    } else {
      formLink.classList.add('hidden');
    }

    // Reset button
    const resetBtn = $('#quiz-reset-btn');
    if (meta.status !== 'not_started') {
      resetBtn.classList.remove('hidden');
    } else {
      resetBtn.classList.add('hidden');
    }

    // Questions editor
    const editor = $('#questions-editor');
    editor.innerHTML = '';
    quizData.questions.forEach((q, idx) => {
      editor.appendChild(createQuestionCard(q, idx));
    });
  }

  function createQuestionCard(q, idx) {
    const card = document.createElement('div');
    card.className = `question-card${q.needs_review ? ' needs-review' : ''}`;
    const letters = 'ABCDEFGHIJ';
    const isTF = q.options.length === 2 &&
      q.options[0].toLowerCase() === 'true' &&
      q.options[1].toLowerCase() === 'false';
    const typeLabel = isTF ? 'True / False' :
      q.options.length === 1 ? 'Fill-in-the-blank' :
      `${q.options.length}-option`;

    let optionsHTML = q.options.map((opt, oi) => `
      <div class="option-row">
        <span class="option-letter">${letters[oi] || ''}</span>
        <div class="option-radio ${oi === q.correct_answer_index ? 'correct' : ''}"
             data-q="${idx}" data-o="${oi}" title="Click to mark as correct"></div>
        <input type="text" class="option-input" value="${esc(opt)}"
               data-q="${idx}" data-o="${oi}" />
        <button class="remove-option-btn" data-q="${idx}" data-o="${oi}" title="Remove option">&times;</button>
      </div>
    `).join('');

    card.innerHTML = `
      <div class="question-header">
        <span class="question-number">Q${q.number}</span>
        <span class="question-type">${typeLabel}${q.needs_review ? ' (needs review)' : ''}</span>
      </div>
      <input type="text" class="question-text-input" value="${esc(q.text)}" data-q="${idx}" />
      <div class="options-list" data-q="${idx}">
        ${optionsHTML}
      </div>
      <button class="add-option-btn" data-q="${idx}">+ Add option</button>
    `;
    return card;
  }

  // ─── Quiz editing events ─────────────────────────────────────────────────
  function handleEditorClick(e) {
    // Mark correct answer
    if (e.target.classList.contains('option-radio')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      currentQuiz.quizData.questions[qi].correct_answer_index = oi;
      const letters = 'ABCDEFGHIJ';
      currentQuiz.quizData.questions[qi].correct_answer_letter = letters[oi] || '';
      renderQuizDetail();
      return;
    }

    // Remove option
    if (e.target.classList.contains('remove-option-btn')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      const q = currentQuiz.quizData.questions[qi];
      if (q.options.length <= 1) return;
      q.options.splice(oi, 1);
      if (q.correct_answer_index >= oi && q.correct_answer_index > 0) {
        q.correct_answer_index--;
      }
      q.question_count = currentQuiz.quizData.questions.length;
      renderQuizDetail();
      return;
    }

    // Add option
    if (e.target.classList.contains('add-option-btn')) {
      const qi = parseInt(e.target.dataset.q);
      currentQuiz.quizData.questions[qi].options.push('New option');
      renderQuizDetail();
      return;
    }
  }

  function handleEditorInput(e) {
    // Question text
    if (e.target.classList.contains('question-text-input')) {
      const qi = parseInt(e.target.dataset.q);
      currentQuiz.quizData.questions[qi].text = e.target.value;
      return;
    }
    // Option text
    if (e.target.classList.contains('option-input')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      currentQuiz.quizData.questions[qi].options[oi] = e.target.value;
      return;
    }
  }

  // ─── Save quiz ───────────────────────────────────────────────────────────
  async function saveQuiz() {
    if (!currentQuiz) return;
    try {
      currentQuiz.quizData.question_count = currentQuiz.quizData.questions.length;
      await api(`/quiz/${currentQuiz.filename}`, {
        method: 'PUT',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(currentQuiz.quizData),
      });
      alert('Quiz saved successfully.');
    } catch (e) {
      alert('Save failed: ' + e.message);
    }
  }

  // ─── Approve & continue to process ──────────────────────────────────────
  function approveQuiz() {
    if (!currentQuiz) return;
    processFormName = currentQuiz.quizData.form_name;
    $('#process-title').textContent = currentQuiz.quizData.form_name;
    // Pre-fill URL if we have one
    if (currentQuiz.meta.form_url) {
      $('#form-url-input').value = currentQuiz.meta.form_url;
    } else {
      $('#form-url-input').value = '';
    }
    // Reset progress
    $('#progress-panel').classList.add('hidden');
    $('#start-process-btn').disabled = false;
    showView('process');
  }

  // ─── Start single form processing ───────────────────────────────────────
  async function startProcessing() {
    const url = $('#form-url-input').value.trim();
    if (!url) { alert('Please paste a form URL.'); return; }
    if (!processFormName) return;

    $('#start-process-btn').disabled = true;
    $('#progress-panel').classList.remove('hidden');
    resetProgressUI(processFormName, currentQuiz.quizData.questions.length);

    try {
      await api('/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: [{ form_name: processFormName, form_url: url }]
        }),
      });
    } catch (e) {
      alert('Failed to start: ' + e.message);
      $('#start-process-btn').disabled = false;
    }
  }

  // ─── Bulk processing ────────────────────────────────────────────────────
  function parseBulkInput() {
    const raw = $('#bulk-input').value.trim();
    if (!raw) { alert('Please paste form data.'); return; }

    const lines = raw.split('\n').filter(l => l.trim());
    bulkItems = [];
    lines.forEach(line => {
      // Split by tab, comma+space, or just comma
      const parts = line.split(/\t|,\s*/).map(s => s.trim());
      if (parts.length >= 2) {
        const name = parts[0];
        const url = parts[parts.length - 1]; // URL is usually last
        const quiz = quizzes.find(q => q.form_name === name);
        bulkItems.push({ form_name: name, form_url: url, found: !!quiz, question_count: quiz ? quiz.question_count : 0 });
      }
    });

    if (bulkItems.length === 0) {
      alert('Could not parse any form entries. Use format: Form Name, URL');
      return;
    }

    // Render preview table
    const preview = $('#bulk-preview');
    preview.classList.remove('hidden');
    preview.innerHTML = `
      <table>
        <thead><tr><th>Form Name</th><th>Questions</th><th>URL</th><th>Status</th></tr></thead>
        <tbody>
          ${bulkItems.map(b => `
            <tr>
              <td>${esc(b.form_name)}</td>
              <td>${b.question_count || '?'}</td>
              <td class="url-cell" title="${esc(b.form_url)}">${esc(b.form_url)}</td>
              <td>${b.found ? '<span class="badge badge-success">Found</span>' : '<span class="badge badge-error">Not found</span>'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    const startBtn = $('#bulk-start-btn');
    if (bulkItems.every(b => b.found)) {
      startBtn.classList.remove('hidden');
    } else {
      startBtn.classList.add('hidden');
      alert('Some form names were not found in the quiz data. Please check and try again.');
    }
  }

  async function startBulkProcessing() {
    if (bulkItems.length === 0) return;
    const validItems = bulkItems.filter(b => b.found);
    if (validItems.length === 0) return;

    // Switch to bulk progress view
    showView('bulk-progress');
    const list = $('#bulk-progress-list');
    list.innerHTML = validItems.map((b, i) => `
      <div class="bulk-progress-item" id="bulk-item-${i}">
        <span class="status-dot not_started"></span>
        <span class="form-name">${esc(b.form_name)}</span>
        <span class="question-progress">--</span>
        <span class="badge badge-not_started">Pending</span>
      </div>
    `).join('');
    $('#bulk-progress-bar').style.width = '0%';
    $('#bulk-progress-detail').textContent = 'Starting...';
    $('#bulk-progress-log').innerHTML = '';

    try {
      await api('/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          items: validItems.map(b => ({ form_name: b.form_name, form_url: b.form_url }))
        }),
      });
    } catch (e) {
      alert('Failed to start bulk processing: ' + e.message);
    }
  }

  // ─── Progress UI ─────────────────────────────────────────────────────────
  function resetProgressUI(formName, totalQuestions) {
    $('#progress-form-name').textContent = formName;
    $('#progress-status').textContent = 'Running';
    $('#progress-status').className = 'badge badge-running';
    $('#progress-bar').style.width = '0%';
    $('#progress-bar').className = 'progress-bar';
    $('#progress-detail').textContent = 'Initializing...';
    $('#progress-log').innerHTML = '';

    const steps = $('#progress-steps');
    steps.innerHTML = '';
    for (let i = 1; i <= totalQuestions; i++) {
      const step = document.createElement('span');
      step.className = 'progress-step pending';
      step.id = `step-q${i}`;
      step.textContent = `Q${i}`;
      steps.appendChild(step);
    }
  }

  // ─── SSE connection ──────────────────────────────────────────────────────
  function connectSSE() {
    if (eventSource) eventSource.close();
    eventSource = new EventSource('/api/progress');

    eventSource.addEventListener('message', (e) => {
      if (e.data === 'heartbeat') return;
      try {
        const data = JSON.parse(e.data);
        handleProgressEvent(data);
      } catch (err) { /* ignore parse errors */ }
    });

    eventSource.addEventListener('error', () => {
      // Auto-reconnect is built into EventSource
    });
  }

  function handleProgressEvent(data) {
    switch (data.type) {
      case 'batch_start':
        updateRunningStatus(true);
        break;

      case 'form_start':
        if (currentView === 'bulk-progress') {
          updateBulkItem(data.formIndex - 1, 'running', 'In Progress', '--');
          $('#bulk-progress-detail').textContent = `Processing ${data.formName} (${data.formIndex}/${data.totalForms})`;
          $('#bulk-progress-bar').style.width = `${((data.formIndex - 1) / data.totalForms) * 100}%`;
        } else {
          $('#progress-detail').textContent = `Opening form...`;
        }
        break;

      case 'title_replaced':
        if (currentView !== 'bulk-progress') {
          $('#progress-detail').textContent = 'Title replaced, scrolling to questions...';
        }
        break;

      case 'section_found':
        if (currentView !== 'bulk-progress') {
          $('#progress-detail').textContent = 'Found Self Assessment section, starting questions...';
        }
        break;

      case 'question_start': {
        const pct = ((data.question - 1) / data.totalQuestions) * 100;
        if (currentView === 'bulk-progress') {
          updateBulkItem(null, null, null, `Q${data.question}/${data.totalQuestions}`);
        } else {
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step active'; }
          $('#progress-bar').style.width = `${pct}%`;
          $('#progress-detail').textContent = `Q${data.question}/${data.totalQuestions}: ${(data.questionText || '').substring(0, 60)}...`;
        }
        break;
      }

      case 'question_done': {
        const pct = (data.question / data.totalQuestions) * 100;
        if (currentView === 'bulk-progress') {
          updateBulkItem(null, null, null, `Q${data.question}/${data.totalQuestions}`);
        } else {
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step done'; }
          $('#progress-bar').style.width = `${pct}%`;
        }
        break;
      }

      case 'question_error': {
        if (currentView !== 'bulk-progress') {
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step error'; }
        }
        break;
      }

      case 'cleanup_start':
        if (currentView !== 'bulk-progress') {
          $('#progress-detail').textContent = 'Cleaning up sample questions...';
        }
        break;

      case 'form_done':
        if (currentView === 'bulk-progress') {
          const idx = findBulkItemIndex(data.formName);
          if (idx !== null) updateBulkItem(idx, 'processed', 'Done', '');
        } else {
          $('#progress-bar').style.width = '100%';
          $('#progress-bar').classList.add('complete');
          $('#progress-status').textContent = 'Complete';
          $('#progress-status').className = 'badge badge-success';
          $('#progress-detail').textContent = 'Form processed successfully!';
        }
        // Refresh quiz list immediately so sidebar/dashboard reflect the new status
        loadQuizzes();
        break;

      case 'form_error':
        if (currentView === 'bulk-progress') {
          const idx = findBulkItemIndex(data.formName);
          if (idx !== null) updateBulkItem(idx, 'failed', 'Failed', data.error || '');
        } else {
          $('#progress-bar').classList.add('error');
          $('#progress-status').textContent = 'Failed';
          $('#progress-status').className = 'badge badge-error';
          $('#progress-detail').textContent = `Error: ${data.error || 'Unknown error'}`;
        }
        // Refresh quiz list immediately so sidebar/dashboard reflect the failure
        loadQuizzes();
        break;

      case 'batch_done':
        updateRunningStatus(false);
        if (currentView === 'bulk-progress') {
          const total = data.results ? data.results.length : 0;
          const success = data.results ? data.results.filter(r => r.success).length : 0;
          $('#bulk-progress-bar').style.width = '100%';
          $('#bulk-progress-bar').classList.add(success === total ? 'complete' : 'error');
          $('#bulk-progress-detail').textContent = `Complete: ${success}/${total} forms processed successfully`;
        }
        // Refresh quiz list to show updated statuses
        loadQuizzes();
        break;

      case 'auth_required':
        if (currentView !== 'bulk-progress') {
          $('#progress-detail').textContent = 'Sign-in required -- please sign in in the browser window...';
        }
        break;

      case 'log': {
        const logEl = currentView === 'bulk-progress' ? $('#bulk-progress-log') : $('#progress-log');
        if (logEl) {
          const line = document.createElement('div');
          line.className = 'log-line';
          if (data.message && data.message.includes('ERROR')) line.classList.add('log-error');
          if (data.message && data.message.includes('DONE')) line.classList.add('log-success');
          line.textContent = data.message || '';
          logEl.appendChild(line);
          logEl.scrollTop = logEl.scrollHeight;
        }
        break;
      }
    }
  }

  // Helper for bulk progress tracking
  let currentBulkIndex = 0;
  function findBulkItemIndex(formName) {
    const items = $$('.bulk-progress-item');
    for (let i = 0; i < items.length; i++) {
      const nameEl = items[i].querySelector('.form-name');
      if (nameEl && nameEl.textContent === formName) return i;
    }
    return null;
  }

  function updateBulkItem(index, dotClass, badgeText, progressText) {
    if (index !== null) currentBulkIndex = index;
    const item = $(`#bulk-item-${currentBulkIndex}`);
    if (!item) return;
    if (dotClass) {
      const dot = item.querySelector('.status-dot');
      dot.className = `status-dot ${dotClass === 'running' ? 'not_started' : dotClass}`;
      if (dotClass === 'running') dot.style.background = '#3b82f6';
    }
    if (badgeText) {
      const badge = item.querySelector('.badge');
      badge.textContent = badgeText;
      badge.className = `badge badge-${dotClass === 'running' ? 'running' : dotClass}`;
    }
    if (progressText !== undefined && progressText !== null) {
      item.querySelector('.question-progress').textContent = progressText;
    }
  }

  // ─── Reset quiz status ───────────────────────────────────────────────────
  async function resetQuizStatus() {
    if (!currentQuiz) return;
    if (!confirm(`Reset status for "${currentQuiz.quizData.form_name}"? This will mark it as not started.`)) return;
    try {
      await fetch(`/api/log/${encodeURIComponent(currentQuiz.quizData.form_name)}`, { method: 'DELETE' });
      await loadQuizzes();
      // Re-select to refresh detail view
      const updated = quizzes.find(q => q.filename === currentQuiz.filename);
      if (updated) selectQuiz(updated);
    } catch (e) {
      alert('Reset failed: ' + e.message);
    }
  }

  // ─── Event bindings ──────────────────────────────────────────────────────
  function bindEvents() {
    // Sidebar filters
    $('#search-input').addEventListener('input', renderQuizList);
    $('#filter-tribe').addEventListener('change', renderQuizList);
    $('#filter-status').addEventListener('change', renderQuizList);

    // Dashboard link (click header title)
    $('header h1').addEventListener('click', () => {
      currentQuiz = null;
      showView('dashboard');
      renderQuizList();
    });
    $('header h1').style.cursor = 'pointer';

    // Quiz detail actions
    $('#save-quiz-btn').addEventListener('click', saveQuiz);
    $('#approve-quiz-btn').addEventListener('click', approveQuiz);
    $('#quiz-reset-btn').addEventListener('click', resetQuizStatus);

    // Editor delegation
    $('#questions-editor').addEventListener('click', handleEditorClick);
    $('#questions-editor').addEventListener('input', handleEditorInput);

    // Setup Guide modal
    $('#setup-guide-btn').addEventListener('click', () => {
      $('#setup-modal').classList.remove('hidden');
    });
    $('#modal-close-btn').addEventListener('click', () => {
      $('#setup-modal').classList.add('hidden');
    });
    $('#setup-modal').addEventListener('click', (e) => {
      if (e.target === $('#setup-modal')) $('#setup-modal').classList.add('hidden');
    });
    $('#copy-prompt-btn').addEventListener('click', () => {
      const text = $('#setup-prompt').textContent;
      navigator.clipboard.writeText(text).then(() => {
        const btn = $('#copy-prompt-btn');
        btn.textContent = 'Copied!';
        btn.classList.add('copied');
        setTimeout(() => { btn.textContent = 'Copy to Clipboard'; btn.classList.remove('copied'); }, 2000);
      });
    });

    // Process
    $('#start-process-btn').addEventListener('click', startProcessing);

    // Bulk
    $('#bulk-parse-btn').addEventListener('click', parseBulkInput);
    $('#bulk-start-btn').addEventListener('click', startBulkProcessing);
  }

  // ─── Utility ─────────────────────────────────────────────────────────────
  function esc(str) {
    const div = document.createElement('div');
    div.textContent = str || '';
    return div.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // ─── Start ───────────────────────────────────────────────────────────────
  init();
})();
