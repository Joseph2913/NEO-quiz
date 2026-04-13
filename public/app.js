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
  let useTemplate = false;     // template duplication mode

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

  // ─── Bulk Tabs ──────────────────────────────────────────────────────────
  function switchBulkTab(tabName) {
    $$('.bulk-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tabName));
    $$('.bulk-tab-content').forEach(c => c.classList.remove('active'));
    $(`#bulk-tab-${tabName}`).classList.add('active');
    // Hide preview when switching tabs
    $('#bulk-preview').classList.add('hidden');
    $('#bulk-start-btn').classList.add('hidden');
    bulkItems = [];
  }

  // ─── Manual Picker ─────────────────────────────────────────────────────
  function getPickerQuizzes() {
    return quizzes.filter(q => q.status !== 'processed');
  }

  function createSearchDropdown(preselect) {
    const wrapper = document.createElement('div');
    wrapper.className = 'search-dropdown';
    wrapper.innerHTML = `
      <input type="text" class="search-dropdown-input" placeholder="Search or select a quiz..." autocomplete="off" />
      <span class="search-dropdown-arrow">&#9660;</span>
      <div class="search-dropdown-list"></div>
    `;

    const input = wrapper.querySelector('.search-dropdown-input');
    const list = wrapper.querySelector('.search-dropdown-list');
    let selectedValue = preselect || '';

    function renderList(filter) {
      const items = getPickerQuizzes();
      const query = (filter || '').toLowerCase();
      const filtered = query ? items.filter(q => q.form_name.toLowerCase().includes(query)) : items;

      if (filtered.length === 0) {
        list.innerHTML = '<div class="search-dropdown-empty">No quizzes found</div>';
        return;
      }
      list.innerHTML = filtered.map(q =>
        `<div class="search-dropdown-item${q.form_name === selectedValue ? ' selected' : ''}" data-value="${esc(q.form_name)}">
          <span>${esc(q.form_name)}</span>
          <span class="item-count">${q.question_count}q</span>
        </div>`
      ).join('');

      list.querySelectorAll('.search-dropdown-item').forEach(item => {
        item.addEventListener('mousedown', (e) => {
          e.preventDefault(); // prevent blur from firing first
          selectedValue = item.dataset.value;
          input.value = selectedValue;
          input.classList.add('has-selection');
          wrapper.classList.remove('open');
          updatePickerPreview();
        });
      });
    }

    // Open on focus
    input.addEventListener('focus', () => {
      wrapper.classList.add('open');
      if (selectedValue && input.value === selectedValue) {
        input.select(); // select all so typing replaces it
      }
      renderList(selectedValue ? '' : input.value);
    });

    // Filter as user types
    input.addEventListener('input', () => {
      selectedValue = '';
      input.classList.remove('has-selection');
      wrapper.classList.add('open');
      renderList(input.value);
      updatePickerPreview();
    });

    // Close on blur
    input.addEventListener('blur', () => {
      wrapper.classList.remove('open');
      // If user typed something that doesn't match, revert to selection
      if (selectedValue) {
        input.value = selectedValue;
        input.classList.add('has-selection');
      } else {
        // Check if typed text exactly matches a quiz
        const exact = getPickerQuizzes().find(q => q.form_name.toLowerCase() === input.value.toLowerCase());
        if (exact) {
          selectedValue = exact.form_name;
          input.value = exact.form_name;
          input.classList.add('has-selection');
          updatePickerPreview();
        }
      }
    });

    // Keyboard navigation
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        wrapper.classList.remove('open');
        input.blur();
      }
    });

    // Set initial value
    if (preselect) {
      input.value = preselect;
      input.classList.add('has-selection');
    }

    // Expose getter
    wrapper.getValue = () => selectedValue;

    return wrapper;
  }

  function addPickerRow(preselect, preurl) {
    const container = $('#manual-picker-rows');
    const row = document.createElement('div');
    row.className = 'picker-row';

    const dropdown = createSearchDropdown(preselect);
    row.appendChild(dropdown);

    const urlInput = document.createElement('input');
    urlInput.type = 'url';
    urlInput.className = 'picker-url';
    urlInput.placeholder = 'Paste the Microsoft Forms URL here...';
    if (preurl) { urlInput.value = preurl; urlInput.classList.add('has-value'); }
    if (useTemplate) { urlInput.style.display = 'none'; }
    urlInput.addEventListener('input', () => {
      urlInput.classList.toggle('has-value', urlInput.value.trim().length > 0);
      updatePickerPreview();
    });
    row.appendChild(urlInput);

    const removeBtn = document.createElement('button');
    removeBtn.className = 'picker-remove-btn';
    removeBtn.title = 'Remove';
    removeBtn.innerHTML = '&times;';
    removeBtn.addEventListener('click', () => {
      row.remove();
      updatePickerPreview();
    });
    row.appendChild(removeBtn);

    container.appendChild(row);
    return row;
  }

  function updatePickerPreview() {
    const rows = $$('.picker-row');
    bulkItems = [];
    rows.forEach(row => {
      const dropdown = row.querySelector('.search-dropdown');
      const name = dropdown ? dropdown.getValue() : '';
      const url = row.querySelector('.picker-url').value.trim();
      // In template mode, URL is not required per row
      if (name && (url || useTemplate)) {
        const quiz = quizzes.find(q => q.form_name === name);
        bulkItems.push({
          form_name: name,
          form_url: url || null,
          found: !!quiz,
          question_count: quiz ? quiz.question_count : 0,
        });
      }
    });
    renderBulkPreview();
  }

  // ─── CSV Download ──────────────────────────────────────────────────────
  function downloadCSVTemplate() {
    const pending = quizzes.filter(q => q.status !== 'processed');
    const header = 'Form Name,URL';
    const rows = pending.map(q => `"${q.form_name.replace(/"/g, '""')}",`);
    const csv = [header, ...rows].join('\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'neo_quiz_bulk_template.csv';
    link.click();
    URL.revokeObjectURL(link.href);
  }

  // ─── CSV Upload / Parse ────────────────────────────────────────────────
  function handleCSVUpload(file) {
    if (!file || !file.name.endsWith('.csv')) {
      alert('Please upload a .csv file.');
      return;
    }
    $('#csv-filename').textContent = `Loaded: ${file.name}`;
    $('#csv-filename').classList.remove('hidden');

    const reader = new FileReader();
    reader.onload = (e) => {
      const text = e.target.result;
      parseCSVText(text);
    };
    reader.readAsText(file);
  }

  function parseCSVText(text) {
    const lines = text.split('\n').filter(l => l.trim());
    bulkItems = [];

    // Skip header if present
    const startIdx = (lines[0] && lines[0].toLowerCase().includes('form name')) ? 1 : 0;

    for (let i = startIdx; i < lines.length; i++) {
      const line = lines[i].trim();
      if (!line) continue;

      // Parse CSV respecting quoted fields
      let name = '';
      let url = '';
      if (line.startsWith('"')) {
        const endQuote = line.indexOf('"', 1);
        if (endQuote > 0) {
          name = line.substring(1, endQuote).replace(/""/g, '"');
          url = line.substring(endQuote + 1).replace(/^[,\t\s]+/, '').trim();
        }
      } else {
        const parts = line.split(/[,\t]/).map(s => s.trim());
        name = parts[0] || '';
        url = parts[parts.length - 1] || '';
      }

      if (name && url) {
        const quiz = quizzes.find(q => q.form_name === name);
        bulkItems.push({ form_name: name, form_url: url, found: !!quiz, question_count: quiz ? quiz.question_count : 0 });
      } else if (name && !url) {
        // Row with name but no URL — skip silently (template row)
      }
    }

    if (bulkItems.length === 0) {
      alert('No valid entries found in the CSV. Make sure each row has a Form Name and a URL.');
      return;
    }

    renderBulkPreview();
  }

  // ─── Paste Text Parse ──────────────────────────────────────────────────
  function parseBulkInput() {
    const raw = $('#bulk-input').value.trim();
    if (!raw) { alert('Please paste form data.'); return; }

    const lines = raw.split('\n').filter(l => l.trim());
    bulkItems = [];
    lines.forEach(line => {
      const parts = line.split(/\t|,\s*/).map(s => s.trim());
      if (parts.length >= 2) {
        const name = parts[0];
        const url = parts[parts.length - 1];
        const quiz = quizzes.find(q => q.form_name === name);
        bulkItems.push({ form_name: name, form_url: url, found: !!quiz, question_count: quiz ? quiz.question_count : 0 });
      }
    });

    if (bulkItems.length === 0) {
      alert('Could not parse any form entries. Use format: Form Name, URL');
      return;
    }

    renderBulkPreview();
  }

  // ─── Shared Preview Render ─────────────────────────────────────────────
  function renderBulkPreview() {
    const preview = $('#bulk-preview');
    if (bulkItems.length === 0) {
      preview.classList.add('hidden');
      $('#bulk-start-btn').classList.add('hidden');
      return;
    }

    preview.classList.remove('hidden');
    const urlHeader = useTemplate ? 'Mode' : 'URL';
    preview.innerHTML = `
      <table>
        <thead><tr><th>Form Name</th><th>Questions</th><th>${urlHeader}</th><th>Status</th></tr></thead>
        <tbody>
          ${bulkItems.map(b => `
            <tr>
              <td>${esc(b.form_name)}</td>
              <td>${b.question_count || '?'}</td>
              <td class="url-cell">${
                useTemplate && !b.form_url
                  ? '<span style="color:#6366f1;font-size:12px;">From template</span>'
                  : `<span title="${esc(b.form_url || '')}">${esc(b.form_url || '')}</span>`
              }</td>
              <td>${b.found ? '<span class="badge badge-success">Found</span>' : '<span class="badge badge-error">Not found</span>'}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;

    const startBtn = $('#bulk-start-btn');
    const allFound = bulkItems.every(b => b.found);
    const anyValid = bulkItems.some(b => b.found);
    if (allFound) {
      startBtn.classList.remove('hidden');
    } else if (anyValid) {
      startBtn.classList.remove('hidden');
      // Show warning but still allow processing found items
    } else {
      startBtn.classList.add('hidden');
      alert('None of the form names matched the quiz data. Please check the names and try again.');
    }
  }

  // ─── Bulk Progress State ─────────────────────────────────────────────────
  let bulkFormStates = {}; // { formName: { index, totalQ, phase, questionsDone, status, error } }
  let bulkTotalForms = 0;

  async function startBulkProcessing() {
    if (bulkItems.length === 0) return;
    const validItems = bulkItems.filter(b => b.found);
    if (validItems.length === 0) return;

    // Initialize state
    bulkTotalForms = validItems.length;
    bulkFormStates = {};
    validItems.forEach((b, i) => {
      bulkFormStates[b.form_name] = {
        index: i,
        totalQ: b.question_count,
        phase: 'waiting',   // waiting | template | opening | title | questions | cleanup | done | failed
        questionsDone: 0,
        currentQuestion: 0,
        status: 'pending',  // pending | running | processed | failed
        error: null,
      };
    });

    // Switch to bulk progress view
    showView('bulk-progress');
    $('#bulk-complete-panel').classList.add('hidden');
    $('#bulk-progress-bar').style.width = '0%';
    $('#bulk-progress-bar').className = 'progress-bar';
    $('#bulk-progress-log').innerHTML = '';

    renderBulkCards();
    updateBulkSummary();

    // Validate template URL if in template mode
    const templateUrl = useTemplate ? ($('#template-url-input').value || '').trim() : null;
    if (useTemplate && !templateUrl) {
      alert('Please paste the template URL before starting.');
      return;
    }

    try {
      const payload = {
        items: validItems.map(b => ({ form_name: b.form_name, form_url: b.form_url })),
      };
      if (templateUrl) {
        payload.template_url = templateUrl;
      }

      await api('/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      });
    } catch (e) {
      alert('Failed to start bulk processing: ' + e.message);
    }
  }

  function renderBulkCards() {
    const list = $('#bulk-progress-list');
    list.innerHTML = '';

    Object.keys(bulkFormStates).forEach((formName) => {
      const s = bulkFormStates[formName];
      const cardClass = s.status === 'running' ? 'bp-active' :
                        s.status === 'processed' ? 'bp-done' :
                        s.status === 'failed' ? 'bp-failed' : '';

      const phases = useTemplate
        ? ['template', 'opening', 'title', 'questions', 'cleanup']
        : ['opening', 'title', 'questions', 'cleanup'];
      const phaseLabels = { template: 'Duplicate', opening: 'Opening', title: 'Title', questions: 'Questions', cleanup: 'Cleanup' };
      const phasesHTML = phases.map(p => {
        let cls = '';
        if (s.status === 'processed' || (s.status === 'running' && phases.indexOf(p) < phases.indexOf(s.phase))) cls = 'phase-done';
        else if (s.status === 'failed' && s.phase === p) cls = 'phase-error';
        else if (s.phase === p && s.status === 'running') cls = 'phase-active';
        return `<span class="bp-phase ${cls}">${phaseLabels[p]}</span>`;
      }).join('');

      // Question step pills
      let stepsHTML = '';
      if (s.totalQ > 0) {
        const pills = [];
        for (let i = 1; i <= s.totalQ; i++) {
          let stepCls = '';
          if (i < s.currentQuestion || (s.status === 'processed' && s.phase === 'done')) stepCls = 'step-done';
          else if (i === s.currentQuestion && s.phase === 'questions' && s.status === 'running') stepCls = 'step-active';
          pills.push(`<span class="bp-step ${stepCls}" id="bp-${s.index}-q${i}">Q${i}</span>`);
        }
        stepsHTML = `<div class="bp-steps">${pills.join('')}</div>`;
      }

      const progressPct = s.totalQ > 0 ? Math.round((s.questionsDone / s.totalQ) * 100) : 0;
      const barPct = s.status === 'processed' ? 100 : s.status === 'failed' ? progressPct : progressPct;

      let detailText = '';
      if (s.status === 'pending') detailText = 'Waiting...';
      else if (s.status === 'processed') detailText = s.newFormUrl ? `Complete! Form URL: ${s.newFormUrl}` : 'All questions filled, samples removed. Complete!';
      else if (s.status === 'failed') detailText = `Error: ${s.error || 'Unknown'}`;
      else if (s.phase === 'template') detailText = 'Duplicating form template...';
      else if (s.phase === 'opening') detailText = 'Opening form in browser...';
      else if (s.phase === 'title') detailText = 'Replacing form title...';
      else if (s.phase === 'questions') detailText = `Filling question ${s.currentQuestion} of ${s.totalQ}...`;
      else if (s.phase === 'cleanup') detailText = 'Deleting sample template questions...';

      const card = document.createElement('div');
      card.className = `bp-card ${cardClass}`;
      card.id = `bp-card-${s.index}`;
      card.innerHTML = `
        <div class="bp-card-header">
          <span class="bp-card-title">${esc(formName)}</span>
          <span class="bp-card-badge badge badge-${s.status === 'running' ? 'running' : s.status}">${
            s.status === 'pending' ? 'Pending' :
            s.status === 'running' ? 'In Progress' :
            s.status === 'processed' ? 'Complete' : 'Failed'
          }</span>
        </div>
        <div class="bp-phases">${phasesHTML}</div>
        <div class="bp-card-progress">
          <div class="bp-card-bar-container">
            <div class="bp-card-bar" style="width: ${barPct}%"></div>
          </div>
        </div>
        ${stepsHTML}
        <div class="bp-card-detail">${detailText}</div>
      `;
      list.appendChild(card);
    });
  }

  function updateBulkSummary() {
    const states = Object.values(bulkFormStates);
    const done = states.filter(s => s.status === 'processed').length;
    const running = states.filter(s => s.status === 'running').length;
    const pending = states.filter(s => s.status === 'pending').length;
    const failed = states.filter(s => s.status === 'failed').length;

    $('#bp-summary-done').textContent = `${done} done`;
    $('#bp-summary-running').textContent = `${running} running`;
    $('#bp-summary-pending').textContent = `${pending} pending`;
    $('#bp-summary-failed').textContent = `${failed} failed`;

    // Overall progress bar
    const total = states.length;
    const completedOrFailed = done + failed;
    $('#bulk-progress-bar').style.width = `${total > 0 ? (completedOrFailed / total) * 100 : 0}%`;
  }

  function showBulkComplete(results) {
    const total = results ? results.length : 0;
    const success = results ? results.filter(r => r.success).length : 0;
    const failed = total - success;

    const panel = $('#bulk-complete-panel');
    panel.classList.remove('hidden');

    if (failed === 0) {
      $('#bulk-complete-icon').textContent = String.fromCodePoint(0x2705);
      $('#bulk-complete-title').textContent = 'All Forms Processed Successfully!';
      $('#bulk-complete-message').textContent = `${success} of ${total} forms completed without errors.`;
      $('#bulk-progress-bar').classList.add('complete');
    } else if (success > 0) {
      $('#bulk-complete-icon').textContent = String.fromCodePoint(0x26A0, 0xFE0F);
      $('#bulk-complete-title').textContent = 'Bulk Processing Complete';
      $('#bulk-complete-message').textContent = `${success} of ${total} forms succeeded. ${failed} form(s) had errors.`;
      $('#bulk-progress-bar').classList.add('error');
    } else {
      $('#bulk-complete-icon').textContent = String.fromCodePoint(0x274C);
      $('#bulk-complete-title').textContent = 'Bulk Processing Failed';
      $('#bulk-complete-message').textContent = `All ${total} forms encountered errors.`;
      $('#bulk-progress-bar').classList.add('error');
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
    const isBulk = currentView === 'bulk-progress';
    const fn = data.formName;

    switch (data.type) {
      case 'batch_start':
        updateRunningStatus(true);
        break;

      case 'form_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].status = 'running';
          bulkFormStates[fn].phase = useTemplate ? 'template' : 'opening';
          renderBulkCards();
          updateBulkSummary();
        } else if (!isBulk) {
          $('#progress-detail').textContent = 'Opening form...';
        }
        break;

      case 'template_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'template';
          renderBulkCards();
        }
        break;

      case 'template_done':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'opening';
          if (data.newUrl) bulkFormStates[fn].newFormUrl = data.newUrl;
          renderBulkCards();
        }
        break;

      case 'title_replaced':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'title';
          renderBulkCards();
        } else if (!isBulk) {
          $('#progress-detail').textContent = 'Title replaced, scrolling to questions...';
        }
        break;

      case 'section_found':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'questions';
          bulkFormStates[fn].currentQuestion = 0;
          renderBulkCards();
        } else if (!isBulk) {
          $('#progress-detail').textContent = 'Found Self Assessment section, starting questions...';
        }
        break;

      case 'question_start': {
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'questions';
          bulkFormStates[fn].currentQuestion = data.question;
          bulkFormStates[fn].totalQ = data.totalQuestions || bulkFormStates[fn].totalQ;
          renderBulkCards();
        } else if (!isBulk) {
          const pct = ((data.question - 1) / data.totalQuestions) * 100;
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step active'; }
          $('#progress-bar').style.width = `${pct}%`;
          $('#progress-detail').textContent = `Q${data.question}/${data.totalQuestions}: ${(data.questionText || '').substring(0, 60)}...`;
        }
        break;
      }

      case 'question_done': {
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].questionsDone = data.question;
          bulkFormStates[fn].currentQuestion = data.question + 1;
          renderBulkCards();
        } else if (!isBulk) {
          const pct = (data.question / data.totalQuestions) * 100;
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step done'; }
          $('#progress-bar').style.width = `${pct}%`;
        }
        break;
      }

      case 'question_error': {
        if (isBulk && bulkFormStates[fn]) {
          // Mark the question as errored but keep going
          renderBulkCards();
        } else if (!isBulk) {
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step error'; }
        }
        break;
      }

      case 'cleanup_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'cleanup';
          renderBulkCards();
        } else if (!isBulk) {
          $('#progress-detail').textContent = 'Cleaning up sample questions...';
        }
        break;

      case 'form_done':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].status = 'processed';
          bulkFormStates[fn].phase = 'done';
          bulkFormStates[fn].questionsDone = bulkFormStates[fn].totalQ;
          if (data.newFormUrl) bulkFormStates[fn].newFormUrl = data.newFormUrl;
          renderBulkCards();
          updateBulkSummary();
        } else if (!isBulk) {
          $('#progress-bar').style.width = '100%';
          $('#progress-bar').classList.add('complete');
          $('#progress-status').textContent = 'Complete';
          $('#progress-status').className = 'badge badge-success';
          $('#progress-detail').textContent = 'Form processed successfully!';
        }
        loadQuizzes();
        break;

      case 'form_error':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].status = 'failed';
          bulkFormStates[fn].phase = bulkFormStates[fn].phase || 'opening';
          bulkFormStates[fn].error = data.error || 'Unknown error';
          renderBulkCards();
          updateBulkSummary();
        } else if (!isBulk) {
          $('#progress-bar').classList.add('error');
          $('#progress-status').textContent = 'Failed';
          $('#progress-status').className = 'badge badge-error';
          $('#progress-detail').textContent = `Error: ${data.error || 'Unknown error'}`;
        }
        loadQuizzes();
        break;

      case 'batch_done':
        updateRunningStatus(false);
        if (isBulk) {
          updateBulkSummary();
          showBulkComplete(data.results);
        }
        loadQuizzes();
        break;

      case 'auth_required':
        if (isBulk && bulkFormStates[fn]) {
          // Keep current phase, just note it
        } else if (!isBulk) {
          $('#progress-detail').textContent = 'Sign-in required -- please sign in in the browser window...';
        }
        break;

      case 'log': {
        const logEl = isBulk ? $('#bulk-progress-log') : $('#progress-log');
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

    // Export
    $('#export-btn').addEventListener('click', async () => {
      const btn = $('#export-btn');
      btn.disabled = true;
      btn.textContent = 'Exporting...';
      try {
        const res = await fetch('/api/export');
        if (!res.ok) throw new Error('Export failed');
        const blob = await res.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `NEO_Quiz_Export_${new Date().toISOString().split('T')[0]}.xlsx`;
        a.click();
        URL.revokeObjectURL(url);
      } catch (e) {
        alert('Export failed: ' + e.message);
      } finally {
        btn.disabled = false;
        btn.textContent = 'Export to Excel';
      }
    });

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

    // Template toggle
    $('#use-template-toggle').addEventListener('change', (e) => {
      useTemplate = e.target.checked;
      $('#template-url-input-group').classList.toggle('hidden', !useTemplate);
      // Hide/show URL columns in picker rows
      $$('.picker-url').forEach(el => {
        el.style.display = useTemplate ? 'none' : '';
      });
      // When in template mode, switch to manual picker (most natural)
      if (useTemplate) {
        switchBulkTab('manual');
      }
      updatePickerPreview();
    });

    // Bulk Tabs
    $$('.bulk-tab').forEach(tab => {
      tab.addEventListener('click', () => switchBulkTab(tab.dataset.tab));
    });

    // Manual Picker
    $('#add-picker-row-btn').addEventListener('click', () => addPickerRow());
    // Add one empty row by default
    addPickerRow();

    // CSV
    $('#csv-download-btn').addEventListener('click', downloadCSVTemplate);
    $('#csv-file-input').addEventListener('change', (e) => {
      if (e.target.files[0]) handleCSVUpload(e.target.files[0]);
    });
    // Drag & drop on CSV area
    const csvArea = $('#csv-upload-area');
    csvArea.addEventListener('dragover', (e) => { e.preventDefault(); csvArea.classList.add('dragover'); });
    csvArea.addEventListener('dragleave', () => csvArea.classList.remove('dragover'));
    csvArea.addEventListener('drop', (e) => {
      e.preventDefault();
      csvArea.classList.remove('dragover');
      if (e.dataTransfer.files[0]) handleCSVUpload(e.dataTransfer.files[0]);
    });

    // Paste
    $('#bulk-parse-btn').addEventListener('click', parseBulkInput);

    // Start (shared)
    $('#bulk-start-btn').addEventListener('click', startBulkProcessing);

    // Bulk complete -> back to dashboard
    $('#bulk-back-btn').addEventListener('click', () => {
      currentQuiz = null;
      showView('dashboard');
      renderQuizList();
    });
  }

  // ─── Utility ─────────────────────────────────────────────────────────────
  function esc(str) {
    const div = document.createElement('div');
    div.textContent = str || '';
    return div.innerHTML.replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // ─── Close dropdowns on outside click ────────────────────────────────────
  document.addEventListener('click', (e) => {
    if (!e.target.closest('.search-dropdown')) {
      $$('.search-dropdown.open').forEach(d => d.classList.remove('open'));
    }
  });

  // ─── Start ───────────────────────────────────────────────────────────────
  init();
})();
