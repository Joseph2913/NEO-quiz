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
  const NEO_TEMPLATE_URL = 'https://forms.cloud.microsoft/Pages/ShareFormPage.aspx?id=-PwcN9hMeUuH3N6aiZ96iJL6XI4jatJEuJk0OOXdqXtUNFU0NUhMU0IzTlhBRkFMSjI3MEc0TUJGOS4u&sharetoken=V5v5UgZt2EBKi9p5hl4O';

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
      const warningIcon = q.needs_review > 0 ? `<span class="quiz-warning-icon" title="${q.needs_review} question(s) need review">&#9888;</span>` : '';
      li.innerHTML = `
        <span class="quiz-name" title="${esc(q.form_name)}">${esc(q.form_name)}</span>
        ${warningIcon}
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

    // Error Check Summary
    const summaryPanel = $('#processing-summary');
    const summaryContent = $('#summary-content');
    const summaryToggle = $('#summary-toggle');
    if (meta.summary && meta.summary.length > 0) {
      summaryPanel.classList.remove('hidden');
      renderErrorCheckSummary(meta);
      summaryContent.innerHTML = renderProcessingTimeline(meta.summary);
      summaryToggle.classList.remove('open');
      summaryContent.classList.add('hidden');
    } else if (meta.status !== 'not_started') {
      summaryPanel.classList.remove('hidden');
      $('#error-check-header').innerHTML = '';
      $('#error-check-body').innerHTML = '<p style="color:#94a3b8;font-size:13px;padding:0 20px 12px;">No detailed report available. Run the automation again to generate one.</p>';
      summaryToggle.classList.remove('open');
      summaryContent.classList.add('hidden');
      summaryContent.innerHTML = '';
    } else {
      summaryPanel.classList.add('hidden');
    }

    // Warning banner for flagged questions
    const flaggedQuestions = detectQuizFlags(quizData);
    const warningBanner = $('#quiz-warning-banner');
    if (flaggedQuestions.length > 0) {
      const flagTypes = [...new Set(flaggedQuestions.flatMap(f => f.flags))];
      warningBanner.classList.remove('hidden');
      warningBanner.innerHTML = `
        <div class="warning-banner-icon">&#9888;</div>
        <div class="warning-banner-body">
          <div class="warning-banner-title">${flaggedQuestions.length} question${flaggedQuestions.length > 1 ? 's' : ''} need${flaggedQuestions.length === 1 ? 's' : ''} review</div>
          <div class="warning-banner-desc">${flagTypes.join(', ')}</div>
        </div>
        <button class="btn btn-sm btn-outline warning-jump-btn" onclick="document.querySelector('.question-card.needs-review')?.scrollIntoView({behavior:'smooth',block:'center'})">Jump to first</button>
      `;
    } else {
      warningBanner.classList.add('hidden');
      warningBanner.innerHTML = '';
    }

    // Questions editor
    const editor = $('#questions-editor');
    editor.innerHTML = '';
    quizData.questions.forEach((q, idx) => {
      editor.appendChild(createQuestionCard(q, idx));
    });
  }

  function renderErrorCheckSummary(meta) {
    const events = meta.summary || [];
    const questionsDone = events.filter(e => e.type === 'question_done').length;
    const questionErrors = events.filter(e => e.type === 'question_error');
    const titleError = events.find(e => e.type === 'title_error');
    const qcEvent = events.find(e => e.type === 'verification_done');
    const qcIssues = (qcEvent && qcEvent.issues) ? qcEvent.issues : [];
    const totalQ = meta.questions_total || 0;
    const allErrors = [...(titleError ? [{ step: 'Title', error: titleError.error || 'Title replacement failed' }] : [])];
    questionErrors.forEach(e => allErrors.push({ step: `Q${e.question}`, error: e.error || 'Question processing failed' }));
    qcIssues.forEach(issue => allErrors.push({ step: 'QC', error: issue }));
    const hasIssues = allErrors.length > 0;

    // Header
    let headerHtml = '<div class="error-check-title">';
    if (hasIssues) {
      headerHtml += `<span class="check-icon" style="color:#dc2626;">&#9888;</span>`;
      headerHtml += `<span>Error Check Summary</span>`;
    } else {
      headerHtml += `<span class="check-icon" style="color:#16a34a;">&#10003;</span>`;
      headerHtml += `<span>Error Check Summary</span>`;
    }
    headerHtml += '</div>';

    // Form link
    if (meta.form_url) {
      headerHtml += `<div class="error-check-link"><a href="${esc(meta.form_url)}" target="_blank" class="btn btn-primary btn-sm" style="text-decoration:none;">Open Form</a><span class="date-text">Processed ${meta.date ? new Date(meta.date).toLocaleString() : ''}</span></div>`;
    }

    // Stats
    headerHtml += '<div class="error-check-stats">';
    headerHtml += `<span style="color:#16a34a;">${questionsDone}/${totalQ} questions filled</span>`;
    if (hasIssues) {
      headerHtml += `<span style="color:#dc2626;">${allErrors.length} issue(s) found</span>`;
    } else {
      headerHtml += `<span style="color:#16a34a;">All checks passed</span>`;
    }
    headerHtml += '</div>';

    $('#error-check-header').innerHTML = headerHtml;

    // Body: error check table
    let bodyHtml = '<table class="error-check-table">';
    bodyHtml += '<thead><tr><th>Check</th><th>Status</th><th>Details</th></tr></thead><tbody>';

    // Title check
    const titleOk = !titleError && !qcIssues.some(i => i.toLowerCase().includes('title'));
    bodyHtml += `<tr><td>Form Title</td><td class="${titleOk ? 'check-pass' : 'check-fail'}">${titleOk ? 'Passed' : 'Failed'}</td><td>${titleOk ? `Set to "${esc(meta.form_name)}"` : esc(titleError ? titleError.error : 'Title verification failed')}</td></tr>`;

    // Per-question checks
    for (let i = 1; i <= totalQ; i++) {
      const qError = questionErrors.find(e => e.question === i);
      const qcFail = qcIssues.filter(issue => {
        const lower = issue.toLowerCase();
        return issue.startsWith(`Q${i}:`) || issue.startsWith(`Q${i} `) || lower.includes(`question ${i}`);
      });
      const hasQIssue = !!qError || qcFail.length > 0;

      let details = '';
      if (qError) details = esc(qError.error);
      if (qcFail.length > 0) {
        const qcDetails = qcFail.map(f => {
          // Strip the "Q3: " prefix for cleaner display
          const cleaned = f.replace(/^Q\d+:\s*/, '');
          return esc(cleaned);
        }).join('<br>');
        details += (details ? '<br>' : '') + qcDetails;
      }
      if (!hasQIssue) {
        const qStartEvt = events.find(e => e.type === 'question_start' && e.question === i);
        details = qStartEvt && qStartEvt.questionText ? esc(qStartEvt.questionText) : 'Filled successfully';
      }

      bodyHtml += `<tr><td>Question ${i}</td><td class="${hasQIssue ? 'check-fail' : 'check-pass'}">${hasQIssue ? 'Failed' : 'Passed'}</td><td>${details}</td></tr>`;
    }

    // Correct answers summary
    const answerIssues = qcIssues.filter(i => i.toLowerCase().includes('correct answer'));
    const answerOk = answerIssues.length === 0;
    const answersMarked = totalQ - answerIssues.length;
    bodyHtml += `<tr><td>Correct Answers</td><td class="${answerOk ? 'check-pass' : 'check-fail'}">${answerOk ? 'Passed' : 'Failed'}</td><td>${answerOk ? `All ${totalQ} answers marked correctly` : `${answersMarked}/${totalQ} marked — ${answerIssues.length} missing`}</td></tr>`;

    // Sample cleanup check
    const cleanupIssues = qcIssues.filter(i => i.toLowerCase().includes('sample question'));
    const cleanupOk = cleanupIssues.length === 0;
    bodyHtml += `<tr><td>Sample Cleanup</td><td class="${cleanupOk ? 'check-pass' : 'check-fail'}">${cleanupOk ? 'Passed' : 'Failed'}</td><td>${cleanupOk ? 'Sample questions removed' : cleanupIssues.map(i => esc(i)).join('<br>')}</td></tr>`;

    bodyHtml += '</tbody></table>';
    $('#error-check-body').innerHTML = bodyHtml;
  }

  function renderProcessingTimeline(events) {
    const eventLabels = {
      form_start: 'Processing started',
      template_start: 'Duplicating form template',
      template_done: 'Template duplicated',
      title_replaced: 'Form title replaced',
      section_found: 'Questions section found',
      question_start: 'Question started',
      question_done: 'Question completed',
      question_error: 'Question failed',
      cleanup_start: 'Cleaning up sample questions',
      form_done: 'Processing complete',
      form_error: 'Processing failed',
    };

    let html = '<div class="summary-timeline">';

    for (const evt of events) {
      const isError = evt.type === 'question_error' || evt.type === 'form_error';
      const isDone = evt.type === 'question_done' || evt.type === 'form_done' || evt.type === 'template_done' || evt.type === 'title_replaced' || evt.type === 'section_found';
      const dotClass = isError ? 'dot-error' : isDone ? 'dot-done' : 'dot-info';
      const dotIcon = isError ? '&#10007;' : isDone ? '&#10003;' : '&#8226;';

      let label = eventLabels[evt.type] || evt.type;
      let detail = '';

      if (evt.type === 'question_start' && evt.question) {
        label = `Question ${evt.question}${evt.totalQuestions ? ' of ' + evt.totalQuestions : ''} started`;
        if (evt.questionText) detail = evt.questionText;
      } else if (evt.type === 'question_done' && evt.question) {
        label = `Question ${evt.question} completed`;
      } else if (evt.type === 'question_error' && evt.question) {
        label = `Question ${evt.question} failed`;
      } else if (evt.type === 'template_done' && evt.newFormUrl) {
        detail = evt.newFormUrl;
      }

      const time = evt.timestamp ? new Date(evt.timestamp).toLocaleTimeString() : '';
      const errorMsg = isError && evt.error ? `<div class="summary-event-error">${esc(evt.error)}</div>` : '';
      const detailHtml = detail ? `<div class="summary-event-detail">${esc(detail)}</div>` : '';

      html += `
        <div class="summary-event">
          <div class="summary-dot ${dotClass}">${dotIcon}</div>
          <div class="summary-event-body">
            <div class="summary-event-label">${esc(label)}</div>
            ${detailHtml}
            ${errorMsg}
          </div>
          <div class="summary-event-time">${time}</div>
        </div>
      `;
    }

    html += '</div>';
    return html;
  }

  // ─── Flag detection ────────────────────────────────────────────────────
  function detectQuestionFlags(q) {
    const flags = [];
    if (q.needs_review && q.options.length <= 1) {
      flags.push('Open-ended — needs multiple choice options');
    } else if (q.needs_review && q.correct_answer_index === -1 && q.options.length > 1) {
      flags.push('No correct answer set');
    } else if (q.needs_review) {
      flags.push('Flagged for review');
    }
    if (!q.needs_review && (q.options.length === 3 || q.options.length === 5)) {
      flags.push(`Non-standard option count (${q.options.length} options)`);
    }
    return flags;
  }

  function detectQuizFlags(quizData) {
    const flagged = [];
    quizData.questions.forEach((q, idx) => {
      const flags = detectQuestionFlags(q);
      if (flags.length > 0) {
        flagged.push({ question: q, index: idx, flags });
      }
    });
    return flagged;
  }

  function createQuestionCard(q, idx, context) {
    const card = document.createElement('div');
    const flags = detectQuestionFlags(q);
    const hasFlags = flags.length > 0;
    card.className = `question-card${q.needs_review ? ' needs-review' : ''}${hasFlags ? ' has-flags' : ''}`;
    const letters = 'ABCDEFGHIJ';
    const isTF = q.options.length === 2 &&
      q.options[0].toLowerCase() === 'true' &&
      q.options[1].toLowerCase() === 'false';
    const typeLabel = isTF ? 'True / False' :
      q.options.length === 1 ? 'Fill-in-the-blank' :
      `${q.options.length}-option`;

    const flagBadgesHTML = flags.map(f => `<span class="flag-badge">${esc(f)}</span>`).join('');
    const ctx = context || 'main';

    let optionsHTML = q.options.map((opt, oi) => `
      <div class="option-row">
        <span class="option-letter">${letters[oi] || ''}</span>
        <div class="option-radio ${oi === q.correct_answer_index ? 'correct' : ''}"
             data-q="${idx}" data-o="${oi}" data-ctx="${ctx}" title="Click to mark as correct"></div>
        <input type="text" class="option-input" value="${esc(opt)}"
               data-q="${idx}" data-o="${oi}" data-ctx="${ctx}" />
        <button class="remove-option-btn" data-q="${idx}" data-o="${oi}" data-ctx="${ctx}" title="Remove option">&times;</button>
      </div>
    `).join('');

    card.innerHTML = `
      <div class="question-header">
        <span class="question-number">Q${q.number}</span>
        <span class="question-type">${typeLabel}</span>
        ${flagBadgesHTML}
      </div>
      <input type="text" class="question-text-input" value="${esc(q.text)}" data-q="${idx}" data-ctx="${ctx}" />
      <div class="options-list" data-q="${idx}">
        ${optionsHTML}
      </div>
      <button class="add-option-btn" data-q="${idx}" data-ctx="${ctx}">+ Add option</button>
    `;
    return card;
  }

  // ─── Quiz editing events ─────────────────────────────────────────────────
  function getQuizDataForContext(ctx) {
    if (!ctx || ctx === 'main') return currentQuiz ? currentQuiz.quizData : null;
    if (ctx.startsWith('bulk:')) {
      const filename = ctx.substring(5);
      return bulkWarningData[filename] ? bulkWarningData[filename].quizData : null;
    }
    return null;
  }

  function handleEditorClick(e) {
    const ctx = e.target.dataset.ctx || 'main';
    const quizData = getQuizDataForContext(ctx);
    if (!quizData) return;

    // Mark correct answer
    if (e.target.classList.contains('option-radio')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      quizData.questions[qi].correct_answer_index = oi;
      const letters = 'ABCDEFGHIJ';
      quizData.questions[qi].correct_answer_letter = letters[oi] || '';
      if (ctx === 'main') renderQuizDetail();
      else rerenderBulkWarningSection(ctx);
      return;
    }

    // Remove option
    if (e.target.classList.contains('remove-option-btn')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      const q = quizData.questions[qi];
      if (q.options.length <= 1) return;
      q.options.splice(oi, 1);
      if (q.correct_answer_index >= oi && q.correct_answer_index > 0) {
        q.correct_answer_index--;
      }
      if (ctx === 'main') renderQuizDetail();
      else rerenderBulkWarningSection(ctx);
      return;
    }

    // Add option
    if (e.target.classList.contains('add-option-btn')) {
      const qi = parseInt(e.target.dataset.q);
      quizData.questions[qi].options.push('New option');
      if (ctx === 'main') renderQuizDetail();
      else rerenderBulkWarningSection(ctx);
      return;
    }
  }

  function handleEditorInput(e) {
    const ctx = e.target.dataset.ctx || 'main';
    const quizData = getQuizDataForContext(ctx);
    if (!quizData) return;

    // Question text
    if (e.target.classList.contains('question-text-input')) {
      const qi = parseInt(e.target.dataset.q);
      quizData.questions[qi].text = e.target.value;
      return;
    }
    // Option text
    if (e.target.classList.contains('option-input')) {
      const qi = parseInt(e.target.dataset.q);
      const oi = parseInt(e.target.dataset.o);
      quizData.questions[qi].options[oi] = e.target.value;
      return;
    }
  }

  function rerenderBulkWarningSection(ctx) {
    const filename = ctx.startsWith('bulk:') ? ctx.substring(5) : ctx;
    const data = bulkWarningData[filename];
    if (!data) return;
    // Re-render the entire panel to reflect changes
    renderBulkWarningPanel();
    updateBulkStartButton();
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
  let singleUseTemplate = false; // template mode for single form processing

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
    // Reset template card
    singleUseTemplate = false;
    $('#single-template-card').classList.remove('selected');
    $('#single-url-group').style.display = '';
    // Reset progress
    $('#progress-panel').classList.add('hidden');
    $('#start-process-btn').disabled = false;
    showView('process');
  }

  // ─── Start single form processing ───────────────────────────────────────
  async function startProcessing() {
    const url = $('#form-url-input').value.trim();
    if (!singleUseTemplate && !url) { alert('Please paste a form URL or select the template.'); return; }
    if (!processFormName) return;

    $('#start-process-btn').disabled = true;
    $('#progress-panel').classList.remove('hidden');
    resetProgressUI(processFormName, currentQuiz.quizData.questions.length);

    try {
      const payload = {
        items: [{ form_name: processFormName, form_url: url || null }],
      };
      if (singleUseTemplate) {
        payload.template_url = NEO_TEMPLATE_URL;
      }
      await api('/process', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
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
    $('#bulk-action-buttons').classList.add('hidden');
    $('#bulk-warning-panel').classList.add('hidden');
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
      $('#bulk-action-buttons').classList.add('hidden');
      $('#bulk-warning-panel').classList.add('hidden');
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

    const actionButtons = $('#bulk-action-buttons');
    const startBtn = $('#bulk-start-btn');
    const allFound = bulkItems.every(b => b.found);
    const anyValid = bulkItems.some(b => b.found);

    // Check template URL is provided when in template mode
    const templateReady = !useTemplate || ($('#template-url-input').value || '').trim().length > 0;

    if (anyValid && templateReady) {
      actionButtons.classList.remove('hidden');
      startBtn.disabled = false;
    } else if (anyValid && !templateReady) {
      actionButtons.classList.remove('hidden');
      startBtn.disabled = true;
      startBtn.title = 'Select a template or paste a template URL first';
    } else {
      actionButtons.classList.add('hidden');
      if (bulkItems.length > 0) {
        alert('None of the form names matched the quiz data. Please check the names and try again.');
      }
    }

    // Check for warnings in selected quizzes
    renderBulkWarnings();
  }

  // ─── Bulk Warning Panel ─────────────────────────────────────────────────
  let bulkWarningData = {}; // { filename: { quizData, flagged: [...] } }
  let bulkIgnored = {};     // { "filename:qIndex": true }

  function isBulkWarningResolved() {
    // All flagged questions must be either ignored or no longer flagged
    for (const [filename, data] of Object.entries(bulkWarningData)) {
      for (const { index } of data.flagged) {
        const key = `${filename}:${index}`;
        if (!bulkIgnored[key]) {
          // Re-check if still flagged (user may have edited it)
          const q = data.quizData.questions[index];
          const flags = detectQuestionFlags(q);
          if (flags.length > 0) return false;
        }
      }
    }
    return true;
  }

  function updateBulkStartButton() {
    const startBtn = $('#bulk-start-btn');
    const hasWarnings = Object.keys(bulkWarningData).length > 0;
    if (hasWarnings && !isBulkWarningResolved()) {
      startBtn.disabled = true;
      startBtn.title = 'Review, edit, or ignore all flagged questions before processing';
      startBtn.textContent = 'Resolve Warnings to Start';
    } else {
      startBtn.disabled = false;
      startBtn.title = '';
      startBtn.textContent = 'Start Bulk Processing';
    }
  }

  function getUnresolvedCount() {
    let count = 0;
    for (const [filename, data] of Object.entries(bulkWarningData)) {
      for (const { index } of data.flagged) {
        const key = `${filename}:${index}`;
        if (!bulkIgnored[key]) {
          const q = data.quizData.questions[index];
          const flags = detectQuestionFlags(q);
          if (flags.length > 0) count++;
        }
      }
    }
    return count;
  }

  function ignoreBulkQuestion(filename, qIndex) {
    bulkIgnored[`${filename}:${qIndex}`] = true;
    renderBulkWarningPanel();
    updateBulkStartButton();
  }

  function ignoreAllForForm(filename) {
    const data = bulkWarningData[filename];
    if (!data) return;
    data.flagged.forEach(({ index }) => {
      bulkIgnored[`${filename}:${index}`] = true;
    });
    renderBulkWarningPanel();
    updateBulkStartButton();
  }

  function ignoreAllWarnings() {
    for (const [filename, data] of Object.entries(bulkWarningData)) {
      data.flagged.forEach(({ index }) => {
        bulkIgnored[`${filename}:${index}`] = true;
      });
    }
    renderBulkWarningPanel();
    updateBulkStartButton();
  }

  async function renderBulkWarnings() {
    const panel = $('#bulk-warning-panel');
    const saveBtn = $('#bulk-save-warnings-btn');
    bulkWarningData = {};
    bulkIgnored = {};

    // Load quiz data for each valid bulk item and detect flags
    const validItems = bulkItems.filter(b => b.found);
    let totalFlags = 0;

    for (const item of validItems) {
      const quizMeta = quizzes.find(q => q.form_name === item.form_name);
      if (!quizMeta) continue;
      try {
        const quizData = await api(`/quiz/${quizMeta.filename}`);
        const flagged = detectQuizFlags(quizData);
        if (flagged.length > 0) {
          bulkWarningData[quizMeta.filename] = { quizData, flagged, meta: quizMeta };
          totalFlags += flagged.length;
        }
      } catch (_) { /* skip on error */ }
    }

    if (totalFlags === 0) {
      panel.classList.add('hidden');
      panel.innerHTML = '';
      saveBtn.classList.add('hidden');
      updateBulkStartButton();
      return;
    }

    saveBtn.classList.remove('hidden');
    renderBulkWarningPanel();
    updateBulkStartButton();
  }

  function renderBulkWarningPanel() {
    const panel = $('#bulk-warning-panel');
    const unresolvedCount = getUnresolvedCount();
    const totalFlags = Object.values(bulkWarningData).reduce((sum, d) => sum + d.flagged.length, 0);
    const quizCount = Object.keys(bulkWarningData).length;
    const allResolved = unresolvedCount === 0;

    if (allResolved) {
      panel.innerHTML = `
        <div class="bulk-warning-header bulk-warning-resolved">
          <div class="bulk-warning-icon">&#10003;</div>
          <div class="bulk-warning-title">All ${totalFlags} warning${totalFlags > 1 ? 's' : ''} resolved</div>
          <div class="bulk-warning-count resolved">Ready to process</div>
        </div>
      `;
      panel.classList.remove('hidden');
      return;
    }

    let html = `
      <div class="bulk-warning-header">
        <div class="bulk-warning-icon">&#9888;</div>
        <div class="bulk-warning-header-body">
          <div class="bulk-warning-title">${unresolvedCount} flagged question${unresolvedCount > 1 ? 's' : ''} need${unresolvedCount === 1 ? 's' : ''} your attention</div>
          <div class="bulk-warning-subtitle">Expand each form below to review and fix issues, or ignore them to proceed.</div>
        </div>
        <div class="bulk-warning-header-actions">
          <span class="bulk-warning-count">${unresolvedCount} of ${totalFlags} unresolved</span>
          <button class="btn btn-sm btn-outline bulk-ignore-all-btn">Ignore All</button>
        </div>
      </div>
      <div class="bulk-warning-sections">
    `;

    for (const [filename, data] of Object.entries(bulkWarningData)) {
      const { quizData, flagged, meta } = data;
      const unresolvedInForm = flagged.filter(({ index }) => {
        const key = `${filename}:${index}`;
        if (bulkIgnored[key]) return false;
        return detectQuestionFlags(data.quizData.questions[index]).length > 0;
      }).length;
      const formResolved = unresolvedInForm === 0;

      html += `
        <div class="bulk-warning-section ${formResolved ? 'section-resolved' : ''}" data-filename="${esc(filename)}">
          <div class="bulk-warning-section-header" data-filename="${esc(filename)}">
            <span class="bulk-warning-collapse-icon">&#9654;</span>
            <span class="bulk-warning-section-name">${esc(quizData.form_name)}</span>
            <span class="badge badge-pending">${meta.tribe}</span>
            ${formResolved
              ? '<span class="bulk-warning-section-status resolved">&#10003; Resolved</span>'
              : `<span class="bulk-warning-section-count">${unresolvedInForm} warning${unresolvedInForm > 1 ? 's' : ''}</span>
                 <button class="btn btn-xs btn-outline bulk-ignore-form-btn" data-filename="${esc(filename)}">Ignore for this form</button>`
            }
          </div>
          <div class="bulk-warning-section-body hidden" data-filename="${esc(filename)}">
            <div class="bulk-warning-questions" id="bulk-warn-q-${esc(filename)}"></div>
          </div>
        </div>
      `;
    }

    html += '</div>';
    panel.innerHTML = html;
    panel.classList.remove('hidden');

    // Render question cards into each section
    for (const [filename, data] of Object.entries(bulkWarningData)) {
      const container = $(`#bulk-warn-q-${CSS.escape(filename)}`);
      if (!container) continue;
      data.flagged.forEach(({ question, index }) => {
        const key = `${filename}:${index}`;
        const isIgnored = !!bulkIgnored[key];
        const stillFlagged = detectQuestionFlags(question).length > 0;

        const wrapper = document.createElement('div');
        wrapper.className = `bulk-warn-item ${isIgnored ? 'ignored' : ''} ${!stillFlagged ? 'fixed' : ''}`;

        if (isIgnored) {
          wrapper.innerHTML = `
            <div class="bulk-warn-ignored-bar">
              <span>&#10003; Q${question.number} — Ignored</span>
              <button class="btn btn-xs btn-outline bulk-unignore-btn" data-key="${esc(key)}">Undo</button>
            </div>
          `;
        } else if (!stillFlagged) {
          wrapper.innerHTML = `
            <div class="bulk-warn-fixed-bar">
              <span>&#10003; Q${question.number} — Fixed</span>
            </div>
          `;
        } else {
          const card = createQuestionCard(question, index, `bulk:${filename}`);
          const ignoreBtn = document.createElement('button');
          ignoreBtn.className = 'btn btn-sm btn-outline bulk-ignore-q-btn';
          ignoreBtn.textContent = 'Ignore this warning';
          ignoreBtn.dataset.filename = filename;
          ignoreBtn.dataset.qindex = index;
          wrapper.appendChild(card);
          wrapper.appendChild(ignoreBtn);
        }

        container.appendChild(wrapper);
      });
    }

    // Bind collapse/expand
    $$('.bulk-warning-section-header').forEach(header => {
      header.addEventListener('click', (e) => {
        // Don't toggle if clicking an action button
        if (e.target.closest('.bulk-ignore-form-btn')) return;
        const fn = header.dataset.filename;
        const body = $(`.bulk-warning-section-body[data-filename="${CSS.escape(fn)}"]`);
        const icon = header.querySelector('.bulk-warning-collapse-icon');
        if (body.classList.contains('hidden')) {
          body.classList.remove('hidden');
          icon.innerHTML = '&#9660;';
          header.closest('.bulk-warning-section').classList.add('expanded');
        } else {
          body.classList.add('hidden');
          icon.innerHTML = '&#9654;';
          header.closest('.bulk-warning-section').classList.remove('expanded');
        }
      });
    });

    // Bind ignore buttons
    $$('.bulk-ignore-all-btn').forEach(btn => btn.addEventListener('click', ignoreAllWarnings));
    $$('.bulk-ignore-form-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        ignoreAllForForm(btn.dataset.filename);
      });
    });
    $$('.bulk-ignore-q-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        ignoreBulkQuestion(btn.dataset.filename, parseInt(btn.dataset.qindex));
      });
    });
    $$('.bulk-unignore-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        delete bulkIgnored[btn.dataset.key];
        renderBulkWarningPanel();
        updateBulkStartButton();
      });
    });
  }

  async function saveBulkWarningEdits() {
    let saved = 0;
    for (const [filename, data] of Object.entries(bulkWarningData)) {
      try {
        data.quizData.question_count = data.quizData.questions.length;
        await api(`/quiz/${filename}`, {
          method: 'PUT',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(data.quizData),
        });
        saved++;
      } catch (_) { /* skip on error */ }
    }
    if (saved > 0) {
      alert(`Saved changes to ${saved} quiz${saved > 1 ? 'zes' : ''}.`);
      await loadQuizzes();
      renderBulkWarnings();
    }
  }

  // ─── Bulk Progress State ─────────────────────────────────────────────────
  let bulkFormStates = {}; // { formName: { index, totalQ, phase, questionsDone, status, error } }
  let bulkTotalForms = 0;

  async function startBulkProcessing() {
    if (bulkItems.length === 0) return;
    const validItems = bulkItems.filter(b => b.found);
    if (validItems.length === 0) return;

    // Auto-save any warning edits before starting
    if (Object.keys(bulkWarningData).length > 0) {
      for (const [filename, data] of Object.entries(bulkWarningData)) {
        try {
          data.quizData.question_count = data.quizData.questions.length;
          await api(`/quiz/${filename}`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(data.quizData),
          });
        } catch (_) { /* best effort */ }
      }
    }

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
        ? ['template', 'opening', 'title', 'questions', 'cleanup', 'verify']
        : ['opening', 'title', 'questions', 'cleanup', 'verify'];
      const phaseLabels = { template: 'Duplicate', opening: 'Opening', title: 'Title', questions: 'Questions', cleanup: 'Cleanup', verify: 'Verify' };
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
  let singleFormPhase = 'opening';
  let singleFormErrors = []; // track question errors for summary

  function resetProgressUI(formName, totalQuestions) {
    singleFormPhase = 'opening';
    singleFormErrors = [];

    $('#progress-form-name').textContent = formName;
    $('#progress-status').textContent = 'Running';
    $('#progress-status').className = 'badge badge-running';
    $('#progress-bar').style.width = '0%';
    $('#progress-bar').className = 'progress-bar';
    $('#progress-detail').textContent = 'Initializing...';
    $('#progress-log').innerHTML = '';
    $('#progress-errors').classList.add('hidden');
    $('#progress-errors').innerHTML = '';
    $('#progress-form-link').classList.add('hidden');

    // Render phase indicators
    renderSinglePhases();

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

  function renderSinglePhases() {
    const phases = singleUseTemplate
      ? ['template', 'opening', 'title', 'questions', 'cleanup', 'verify']
      : ['opening', 'title', 'questions', 'cleanup', 'verify'];
    const labels = { template: 'Duplicate', opening: 'Opening', title: 'Title', questions: 'Questions', cleanup: 'Cleanup', verify: 'Verify' };
    const phaseOrder = phases;

    const container = $('#progress-phases');
    container.innerHTML = phaseOrder.map(p => {
      let cls = '';
      const currentIdx = phaseOrder.indexOf(singleFormPhase);
      const thisIdx = phaseOrder.indexOf(p);
      if (thisIdx < currentIdx) cls = 'phase-done';
      else if (p === singleFormPhase) cls = 'phase-active';
      return `<span class="bp-phase ${cls}">${labels[p]}</span>`;
    }).join('');
  }

  function showProgressFormLink(url) {
    const container = $('#progress-form-link');
    const link = $('#progress-form-url');
    link.href = url;
    container.classList.remove('hidden');
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
          singleFormPhase = singleUseTemplate ? 'template' : 'opening';
          renderSinglePhases();
          $('#progress-detail').textContent = singleUseTemplate ? 'Duplicating form template...' : 'Opening form...';
        }
        break;

      case 'template_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'template';
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'template';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Duplicating form template...';
        }
        break;

      case 'template_done':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'opening';
          if (data.newUrl) bulkFormStates[fn].newFormUrl = data.newUrl;
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'opening';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Template duplicated, loading form editor...';
          if (data.newUrl) {
            showProgressFormLink(data.newUrl);
          }
        }
        break;

      case 'title_replaced':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'title';
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'title';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Title replaced, scrolling to questions...';
        }
        break;

      case 'section_found':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'questions';
          bulkFormStates[fn].currentQuestion = 0;
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'questions';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Found questions section, starting...';
        }
        break;

      case 'question_start': {
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'questions';
          bulkFormStates[fn].currentQuestion = data.question;
          bulkFormStates[fn].totalQ = data.totalQuestions || bulkFormStates[fn].totalQ;
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'questions';
          renderSinglePhases();
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
          renderBulkCards();
        } else if (!isBulk) {
          const step = $(`#step-q${data.question}`);
          if (step) { step.className = 'progress-step error'; }
          singleFormErrors.push({
            question: data.question,
            error: data.error || 'Unknown error',
          });
        }
        break;
      }

      case 'title_error':
        if (!isBulk) {
          singleFormErrors.push({ question: 0, error: data.error || 'Title replacement failed' });
        }
        break;

      case 'cleanup_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'cleanup';
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'cleanup';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Cleaning up sample questions...';
        }
        break;

      case 'verification_start':
        if (isBulk && bulkFormStates[fn]) {
          bulkFormStates[fn].phase = 'verify';
          renderBulkCards();
        } else if (!isBulk) {
          singleFormPhase = 'verify';
          renderSinglePhases();
          $('#progress-detail').textContent = 'Running quality check...';
        }
        break;

      case 'verification_done':
        if (!isBulk && data.results && data.results.issues && data.results.issues.length > 0) {
          data.results.issues.forEach(issue => {
            singleFormErrors.push({ question: 0, error: issue });
          });
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
          singleFormPhase = 'done';
          // Mark all phases as done
          $('#progress-phases').querySelectorAll('.bp-phase').forEach(p => {
            p.className = 'bp-phase phase-done';
          });

          $('#progress-bar').style.width = '100%';

          if (data.newFormUrl) {
            showProgressFormLink(data.newFormUrl);
          }

          if (singleFormErrors.length > 0) {
            // Partially complete
            $('#progress-bar').classList.add('error');
            $('#progress-status').textContent = 'Partially Complete';
            $('#progress-status').className = 'badge badge-pending';
            const totalQ = $$('.progress-step').length;
            const failedQ = singleFormErrors.length;
            $('#progress-detail').textContent = `${totalQ - failedQ} of ${totalQ} questions filled successfully. ${failedQ} question(s) had issues.`;

            // Show error summary
            const errPanel = $('#progress-errors');
            errPanel.classList.remove('hidden');
            errPanel.innerHTML = `
              <div class="error-title">Issues found during processing:</div>
              ${singleFormErrors.map(e =>
                `<div class="error-item">Q${e.question}: ${e.error}</div>`
              ).join('')}
            `;
          } else {
            $('#progress-bar').classList.add('complete');
            $('#progress-status').textContent = 'Complete';
            $('#progress-status').className = 'badge badge-success';
            $('#progress-detail').textContent = 'Form processed successfully!';
          }
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
          // Mark the current phase as error
          const phaseEls = $('#progress-phases').querySelectorAll('.bp-phase');
          phaseEls.forEach(p => {
            if (p.classList.contains('phase-active')) {
              p.classList.remove('phase-active');
              p.classList.add('phase-error');
            }
          });

          $('#progress-bar').classList.add('error');
          $('#progress-status').textContent = 'Failed';
          $('#progress-status').className = 'badge badge-error';
          $('#progress-detail').textContent = `Error during ${singleFormPhase}: ${data.error || 'Unknown error'}`;
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

    // Processing summary toggle
    $('#summary-toggle').addEventListener('click', () => {
      const toggle = $('#summary-toggle');
      const content = $('#summary-content');
      toggle.classList.toggle('open');
      content.classList.toggle('hidden');
    });

    // Editor delegation (main editor + bulk warning editor)
    $('#questions-editor').addEventListener('click', handleEditorClick);
    $('#questions-editor').addEventListener('input', handleEditorInput);

    // Bulk warning panel editor delegation
    $('#bulk-warning-panel').addEventListener('click', handleEditorClick);
    $('#bulk-warning-panel').addEventListener('input', handleEditorInput);

    // Bulk warning save button
    $('#bulk-save-warnings-btn').addEventListener('click', saveBulkWarningEdits);

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

    // Single form template card
    $('#single-template-card').addEventListener('click', () => {
      const card = $('#single-template-card');
      singleUseTemplate = !card.classList.contains('selected');
      card.classList.toggle('selected');
      // Hide/show the manual URL input
      $('#single-url-group').style.display = singleUseTemplate ? 'none' : '';
    });

    // Process
    $('#start-process-btn').addEventListener('click', startProcessing);

    // Template toggle
    $('#use-template-toggle').addEventListener('change', (e) => {
      useTemplate = e.target.checked;
      $('#template-url-input-group').classList.toggle('hidden', !useTemplate);
      if (!useTemplate) {
        // Deselect the card and clear URL when toggling off
        $('#template-card-neo').classList.remove('selected');
        $('#template-url-input').value = '';
      }
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

    // Template card click
    $('#template-card-neo').addEventListener('click', () => {
      const card = $('#template-card-neo');
      const input = $('#template-url-input');
      const isSelected = card.classList.contains('selected');
      if (isSelected) {
        card.classList.remove('selected');
        input.value = '';
      } else {
        card.classList.add('selected');
        input.value = NEO_TEMPLATE_URL;
      }
      // Re-validate the start button
      if (bulkItems.length > 0) renderBulkPreview();
    });

    // If user types a custom URL, deselect the card and revalidate
    $('#template-url-input').addEventListener('input', () => {
      const card = $('#template-card-neo');
      if ($('#template-url-input').value.trim() !== NEO_TEMPLATE_URL) {
        card.classList.remove('selected');
      }
      // Re-validate the start button
      if (bulkItems.length > 0) renderBulkPreview();
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

    // Bulk complete -> back to dashboard (reset everything)
    $('#bulk-back-btn').addEventListener('click', () => {
      currentQuiz = null;

      // Clear bulk state
      bulkItems = [];
      bulkFormStates = {};
      bulkTotalForms = 0;

      // Reset template section
      useTemplate = false;
      $('#use-template-toggle').checked = false;
      $('#template-url-input-group').classList.add('hidden');
      $('#template-url-input').value = '';
      $('#template-card-neo').classList.remove('selected');

      // Reset picker rows
      $('#manual-picker-rows').innerHTML = '';
      addPickerRow();

      // Reset CSV tab
      $('#csv-file-input').value = '';
      $('#csv-filename').classList.add('hidden');
      $('#csv-filename').textContent = '';

      // Reset paste tab
      $('#bulk-input').value = '';

      // Hide preview, warnings, and action buttons
      $('#bulk-preview').classList.add('hidden');
      $('#bulk-action-buttons').classList.add('hidden');
      $('#bulk-warning-panel').classList.add('hidden');

      // Reset bulk progress view
      $('#bulk-progress-list').innerHTML = '';
      $('#bulk-progress-bar').style.width = '0%';
      $('#bulk-progress-bar').className = 'progress-bar';
      $('#bulk-progress-log').innerHTML = '';
      $('#bulk-complete-panel').classList.add('hidden');

      // Switch back to manual picker tab
      switchBulkTab('manual');

      // Refresh data and show dashboard
      loadQuizzes();
      showView('dashboard');
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
