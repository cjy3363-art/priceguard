/*
 * app.js — 화면 전환, 파일 업로드, 화면 그리기, 다운로드
 * 정제 로직은 clean.js(PriceCleaner), 규칙은 rules/병의원수가표.json 에 있다.
 * 업로드한 파일은 이 브라우저 메모리에서만 다루며, 어떤 네트워크 요청으로도 보내지 않는다.
 * (네트워크 요청은 같은 폴더의 규칙 JSON 파일을 읽는 1회뿐이다.)
 */
(function () {
  'use strict';

  var RULES_URL = 'rules/병의원수가표.json';
  var ACCEPT_EXT = ['.xlsx', '.xls', '.xlsm'];
  var VIEWS = ['upload', 'dashboard', 'validation', 'rules', 'history', 'benchmark'];
  var DASH_PAGE = 8;
  var C = window.PriceCleaner;

  var state = {
    rules: null, rulesSource: '',
    fileName: '', fileSize: 0, fileBytes: null, table: null, result: null, ranAt: null,
    filter: 'all', ruleFilter: '', view: 'diff', search: '', selected: null,
    dashFilter: 'all', dashPage: 0, logFilter: 'all',
    answer: null, ansFilter: 'mismatch', ansSearch: '', ansHighlight: true
  };

  /* ---------- DOM 도우미 (값은 항상 textContent 로 넣는다) ---------- */
  function $(id) { return document.getElementById(id); }
  function $$(sel) { return Array.prototype.slice.call(document.querySelectorAll(sel)); }
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v === null || v === undefined || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k.indexOf('on') === 0) el.addEventListener(k.slice(2), v);
      else el.setAttribute(k, v === true ? '' : v);
    });
    for (var i = 2; i < arguments.length; i++) append(el, arguments[i]);
    return el;
  }
  function append(el, kid) {
    if (kid === null || kid === undefined || kid === false) return;
    if (Array.isArray(kid)) { kid.forEach(function (k) { append(el, k); }); return; }
    el.appendChild(kid.nodeType ? kid : document.createTextNode(String(kid)));
  }
  function clear(el) { while (el.firstChild) el.removeChild(el.firstChild); }
  function icon(name, cls) {
    var NS = 'http://www.w3.org/2000/svg';
    var svg = document.createElementNS(NS, 'svg');
    svg.setAttribute('class', 'ic' + (cls ? ' ' + cls : ''));
    svg.setAttribute('aria-hidden', 'true');
    var use = document.createElementNS(NS, 'use');
    use.setAttribute('href', '#i-' + name);
    svg.appendChild(use);
    return svg;
  }
  function fmtNum(n) { return typeof n === 'number' ? n.toLocaleString('ko-KR') : n; }
  function pct(a, b) { return b ? (a / b * 100).toFixed(1) + '%' : '-'; }
  function pad2(n) { return String(n).padStart(2, '0'); }
  function today() { var d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function stamp(d) { return d.getFullYear() + '. ' + pad2(d.getMonth() + 1) + '. ' + pad2(d.getDate()) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()); }
  function fileSize(n) { return n >= 1048576 ? (n / 1048576).toFixed(2) + ' MB' : Math.max(1, Math.round(n / 1024)) + ' KB'; }

  function showError(msg) { var a = $('global-alert'); a.textContent = msg; a.hidden = false; window.scrollTo(0, 0); }
  function clearError() { $('global-alert').hidden = true; }

  function actionText(action) {
    var s = (state.rules.actions || {})[action] || action;
    return String(s).split(':')[0].trim();
  }
  function actionBadge(action) {
    var cls = action === 'auto_fix' ? 'badge-auto' : action === 'review' ? 'badge-review' : 'badge-keep';
    return h('span', { class: 'badge badge-action ' + cls, text: actionText(action) });
  }
  function statusBadge(status) {
    var cls = { '정상': 'badge-ok', '검토 필요': 'badge-review', '충돌': 'badge-conflict', '삭제됨': 'badge-del' }[status] || 'badge-keep';
    return h('span', { class: 'badge ' + cls, text: status });
  }
  function ruleById(id) { return state.rules.rules.filter(function (r) { return r.id === id; })[0] || null; }
  function unitOf(rule) {
    var u = rule && rule.expected_impact && rule.expected_impact.unit;
    return u ? String(u).split(/\s|\(/)[0] : '건';
  }
  function cellText(cell) {
    if (C.isBlank(cell)) return '(빈칸)';
    var v = C.display(cell, { markExcelDate: true });
    return typeof v === 'number' ? fmtNum(v) : String(v);
  }
  function cellNode(cell) {
    if (C.isBlank(cell)) return h('span', { class: 'blank-cell', text: '(빈칸)' });
    return document.createTextNode(cellText(cell));
  }
  function code(row) { return C.display(row.raw[state.result.keyCol]); }
  function nameOf(row) { var c = row.removed ? row.raw['시술명'] : row.cur['시술명']; return c && !C.isBlank(c) ? String(c.v).trim() : '(시술명 빈칸)'; }
  /** 행의 문제를 한 줄로 */
  function problemText(row) {
    if (row.removed) return '원본 ' + row.keptRow + '행과 완전 중복 · 삭제';
    if (row.reviews.length) return row.reviews.map(function (r) { return r.reason; }).join(' · ');
    if (row.fixes.length) return row.fixes.map(function (f) { return f.col + ' ' + (ruleById(f.ruleId) || {}).name; }).join(' · ');
    return '문제 없음';
  }

  /* =========================================================
     화면 전환
     ========================================================= */
  function currentView() {
    var v = (location.hash.match(/^#\/(\w+)/) || [])[1];
    return VIEWS.indexOf(v) >= 0 ? v : (state.result ? 'dashboard' : 'upload');
  }
  function go(view) {
    if (location.hash === '#/' + view) showView(view);
    else location.hash = '#/' + view;
  }
  function showView(view) {
    $$('.view').forEach(function (s) { s.hidden = s.getAttribute('data-view') !== view; });
    $$('.nav-link').forEach(function (a) {
      if (a.getAttribute('data-view') === view) a.setAttribute('aria-current', 'page');
      else a.removeAttribute('aria-current');
    });
    var sec = $('view-' + view);
    $('crumb-current').textContent = sec ? sec.getAttribute('data-title') : '';
    document.title = 'PriceGuard | 병의원 시술 가격표 검증';
    closeSidebar();
    window.scrollTo(0, 0);
  }
  window.addEventListener('hashchange', function () { showView(currentView()); });

  document.addEventListener('click', function (e) {
    var g = e.target.closest('[data-goto]');
    if (g) {
      e.preventDefault();
      var f = g.getAttribute('data-filter');
      if (f && state.result) { state.filter = f; state.ruleFilter = ''; renderValidation(); }
      go(g.getAttribute('data-goto'));
      return;
    }
    var d = e.target.closest('[data-download]');
    if (d && !d.disabled) download(d.getAttribute('data-download'));
  });

  /* 모바일 사이드바 */
  function openSidebar() { $('sidebar').classList.add('is-open'); $('sb-backdrop').hidden = false; $('menu-btn').setAttribute('aria-expanded', 'true'); }
  function closeSidebar() { $('sidebar').classList.remove('is-open'); $('sb-backdrop').hidden = true; $('menu-btn').setAttribute('aria-expanded', 'false'); }
  $('menu-btn').addEventListener('click', function () { $('sidebar').classList.contains('is-open') ? closeSidebar() : openSidebar(); });
  $('sb-backdrop').addEventListener('click', closeSidebar);
  document.addEventListener('keydown', function (e) { if (e.key === 'Escape') closeSidebar(); });

  function setHasResult(has) {
    $$('[data-needs-result]').forEach(function (e) { e.hidden = has; });
    $$('[data-has-result]').forEach(function (e) { e.hidden = !has; });
    $$('[data-download]').forEach(function (b) { b.disabled = !has; });
  }

  /* =========================================================
     규칙 불러오기
     ========================================================= */
  function validateRules(json) {
    if (!json || !Array.isArray(json.rules) || !Array.isArray(json.execution_order) || !json.dataset || !Array.isArray(json.dataset.columns)) {
      throw new Error('규칙 파일 형식이 올바르지 않습니다. (rules, execution_order, dataset 항목이 필요합니다)');
    }
  }

  function setRules(json, source) {
    validateRules(json);
    state.rules = json;
    state.rulesSource = source;
    var st = $('rules-status');
    clear(st);
    append(st, ['규칙 파일 ', h('b', { text: source }), ' · 규칙 ' + json.rules.length + '개',
      json.approval && json.approval.approved_on ? ' · 승인 ' + json.approval.approved_on : '']);
    $('sb-version').textContent = '규칙 v' + (json.schema_version || '-') + (json.approval && json.approval.approved_on ? ' · 승인 ' + json.approval.approved_on : '');
    updateRulesBadge();
    var warns = C.checkOrder(json);
    if (warns.length) showError('규칙 파일 확인 필요: ' + warns.join(' '));
    $('rules-picker').hidden = true;
    setUploadEnabled(true);
    renderRules();
    if (state.table) runClean();
  }

  function updateRulesBadge() {
    var t = $('rules-badge-text');
    if (!state.rules) { t.textContent = '규칙 파일 없음'; return; }
    t.textContent = state.result
      ? '정제 규칙 ' + state.result.summary.rulesExecuted + '개 적용됨'
      : '정제 규칙 ' + state.rules.rules.length + '개 준비됨';
  }

  function loadRules() {
    setUploadEnabled(false);
    fetch(RULES_URL, { cache: 'no-store', credentials: 'same-origin' })
      .then(function (res) { if (!res.ok) throw new Error('HTTP ' + res.status); return res.json(); })
      .then(function (json) { setRules(json, RULES_URL); })
      .catch(function () {
        $('rules-status').textContent = '규칙 파일을 자동으로 불러오지 못했습니다.';
        $('rules-badge-text').textContent = '규칙 파일 필요';
        $('sb-version').textContent = '규칙 파일 필요';
        $('rules-picker').hidden = false;
      });
  }

  $('rules-input').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f) return;
    f.text().then(function (txt) {
      try { clearError(); setRules(JSON.parse(txt), f.name + ' (직접 선택)'); }
      catch (err) { showError('규칙 파일을 읽지 못했습니다: ' + err.message); }
    });
  });

  /* =========================================================
     업로드
     ========================================================= */
  function setUploadEnabled(on) {
    $('dropzone').classList.toggle('is-disabled', !on);
    $('file-input').disabled = !on;
  }

  var dz = $('dropzone');
  ['dragenter', 'dragover'].forEach(function (t) {
    dz.addEventListener(t, function (e) { e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; dz.classList.add('is-over'); });
  });
  ['dragleave', 'dragend'].forEach(function (t) {
    dz.addEventListener(t, function () { dz.classList.remove('is-over'); });
  });
  dz.addEventListener('drop', function (e) {
    e.preventDefault();
    dz.classList.remove('is-over');
    var files = e.dataTransfer.files;
    if (files && files.length) handleFile(files[0]);
  });
  dz.addEventListener('click', function (e) {
    if (e.target.closest('label')) return;
    $('file-input').click();
  });
  dz.addEventListener('keydown', function (e) {
    if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('file-input').click(); }
  });
  // 페이지 다른 곳에 떨어뜨렸을 때 브라우저가 파일을 여는 것을 막는다
  window.addEventListener('dragover', function (e) { e.preventDefault(); });
  window.addEventListener('drop', function (e) { e.preventDefault(); });

  $('file-input').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (f) handleFile(f);
  });
  $('btn-other-file').addEventListener('click', function () { $('file-input').click(); });

  function handleFile(file) {
    clearError();
    if (!state.rules) { showError('규칙 파일이 아직 준비되지 않았습니다. 규칙 파일을 먼저 불러와 주세요.'); return; }
    var name = file.name || '';
    var ext = (name.match(/\.[^.]+$/) || [''])[0].toLowerCase();
    if (ACCEPT_EXT.indexOf(ext) < 0) {
      showError('"' + name + '"은(는) 지원하지 않는 파일 형식입니다. 엑셀 파일(' + ACCEPT_EXT.join(', ') + ')만 올릴 수 있습니다.');
      return;
    }
    file.arrayBuffer().then(function (buf) {
      var bytes = new Uint8Array(buf);
      var table;
      try { table = C.readWorkbook(XLSX, bytes, state.rules); }
      catch (err) { showError('파일을 읽지 못했습니다: ' + err.message); return; }
      state.fileName = name;
      state.fileSize = file.size || bytes.length;
      state.fileBytes = bytes;
      state.table = table;
      state.selected = null;
      state.filter = 'all'; state.ruleFilter = ''; state.search = '';
      state.dashFilter = 'all'; state.dashPage = 0; state.logFilter = 'all';
      $('search').value = '';
      resetAnswer();
      runClean();
      if (state.result) go('dashboard');
    }).catch(function (err) { showError('파일을 읽지 못했습니다: ' + err.message); });
  }

  function runClean() {
    try { state.result = C.clean(state.table, state.rules); }
    catch (err) { showError('정제 규칙을 실행하지 못했습니다: ' + err.message); return; }
    state.ranAt = new Date();
    if (state.result.warnings.length) showError('규칙 확인 필요: ' + state.result.warnings.join(' '));
    $('rerun-result').hidden = true;
    setHasResult(true);
    updateRulesBadge();
    renderUploadInfo();
    renderSidebarFile();
    renderDashboard();
    renderRules();
    renderValidation();
    renderHistory();
    renderBenchmarkHead();
  }

  /* =========================================================
     파일 정보 (업로드 화면, 사이드바)
     ========================================================= */
  function renderUploadInfo() {
    var m = state.table.meta;
    $('upload-feedback').hidden = false;
    $('meta-name').textContent = state.fileName;
    $('meta-size').textContent = '(' + fileSize(state.fileSize) + ')';
    $('meta-sheets').textContent = m.sheetCount + '개';
    $('meta-rows').textContent = fmtNum(m.dataRowCount) + '행';
    $('meta-cols').textContent = m.columnCount + '개';
    $('meta-rules').textContent = '정제 규칙 ' + state.result.summary.rulesExecuted + '개 적용 완료';
    var notes = [];
    var id = m.rowIdentity;
    if (id && id.source === 'inherited') {
      notes.push('정리본 파일입니다. "' + id.column + '" 컬럼의 원본 행 번호 ' + id.inherited + '개를 이어받아, 변경 기록과 검토 사유에 처음 원본의 행 번호를 씁니다.');
      if (id.missing) notes.push('원본 행 번호가 비어 있는 ' + id.missing + '행은 이 파일의 행 번호를 씁니다.');
      if (id.duplicated.length) notes.push('원본 행 번호가 겹치는 행이 있습니다: ' + id.duplicated.join(', ') + '행.');
    }
    if (m.sheetCount > 1) notes.push('시트가 여러 개입니다. 첫 번째 시트("' + m.sheetName + '")만 검사합니다.');
    if (m.extraColumns.length) notes.push('규칙에 없는 컬럼은 검사하지 않습니다: ' + m.extraColumns.join(', '));
    if (m.skippedEmptyRows) notes.push('완전히 빈 행 ' + m.skippedEmptyRows + '개는 데이터로 세지 않았습니다.');
    $('meta-note').textContent = notes.join(' ');
    $('meta-note').hidden = !notes.length;
    $('step-1').className = 'step is-done';
    $('step-2').className = 'step is-done';
    $('step-3').className = 'step is-active';
  }

  function renderSidebarFile() {
    $('sb-file-name').textContent = state.fileName;
    $('sb-file-name').title = state.fileName;
    $('sb-file-rows').textContent = '(' + fmtNum(state.result.summary.rowsBefore) + '행)';
    $('sb-dot').className = 'dot dot-ok';
    $('sb-status-text').textContent = '정제 완료';
    $('sb-status-text').parentNode.classList.add('is-ok');
  }

  /* =========================================================
     대시보드
     ========================================================= */
  function kpi(opts) {
    return h('div', { class: 'kpi' + (opts.tone ? ' tone-' + opts.tone : '') },
      h('div', { class: 'kpi-top' }, h('span', { class: 'kpi-label', text: opts.label }),
        opts.icon ? icon(opts.icon) : h('span', { class: 'kpi-meta', text: opts.meta || '' })),
      h('div', { class: 'kpi-value' }, h('span', { class: 'kpi-num', text: opts.value }), h('span', { class: 'kpi-unit', text: opts.unit || '' })),
      opts.bar !== undefined
        ? h('div', { class: 'kpi-barrow' }, (function () { var b = h('div'); b.style.width = Math.max(0, Math.min(100, opts.bar)) + '%'; return h('div', { class: 'kpi-bar' }, b); })(), h('span', { text: opts.foot || '' }))
        : h('div', { class: 'kpi-foot', text: opts.foot || '' }));
  }

  function thresholdText() {
    var rule = state.rules.rules.filter(function (r) { return r.condition && r.condition.type === 'price_ratio_outlier'; })[0];
    if (!rule) return null;
    var words = { ratio_lt: '미만', ratio_lte: '이하', ratio_gt: '초과', ratio_gte: '이상' };
    var f = rule.condition.flag_if || {};
    var parts = Object.keys(words).filter(function (k) { return f[k] !== undefined; }).map(function (k) { return f[k] + ' ' + words[k]; });
    return {
      rule: rule,
      text: '가격 이상치 기준(' + rule.id + ')은 임시 기준입니다. 현재 기준: ' + rule.condition.ratio + ' 비율이 ' + parts.join(' 또는 ') + '. ' +
        '이 값은 목업 데이터 검증용으로 정한 것이며, 실제 병의원 데이터에 적용하려면 별도의 검증과 기준 조정이 필요합니다.'
    };
  }

  function renderDashboard() {
    var r = state.result, s = r.summary;
    $('dash-file-name').textContent = state.fileName;
    $('dash-rows').textContent = fmtNum(s.rowsBefore) + '행 데이터 → 정리 후 ' + fmtNum(s.rowsAfter) + '행';
    $('dash-rules').textContent = '정제 규칙 ' + s.rulesExecuted + ' / ' + s.rulesTotal + '개 실행 · 해당 건 있는 규칙 ' + s.rulesWithHits + '개';
    $('dash-time').textContent = '검증 시각 ' + stamp(state.ranAt);

    var k = $('kpis');
    clear(k);
    append(k, [
      kpi({ label: '전체 데이터', icon: 'table', value: fmtNum(s.rowsBefore), unit: '행', foot: '원본 행 수 · 정리 후 ' + fmtNum(s.rowsAfter) + '행' }),
      kpi({ label: '자동 수정', meta: pct(s.autoFixRows, s.rowsBefore), tone: 'ok', value: fmtNum(s.autoFixCells), unit: '칸', foot: s.autoFixRows + '개 행 · 형식 통일' }),
      kpi({ label: '검토 필요', meta: pct(s.reviewRows, s.rowsBefore), tone: 'error', value: fmtNum(s.reviewRows), unit: '행', foot: '충돌 ' + s.conflictRows + '행 포함 · 사유 ' + s.reviewEntries + '건' }),
      kpi({ label: '중복 제거', meta: pct(s.removedRows, s.rowsBefore), value: fmtNum(s.removedRows), unit: '행', foot: '완전히 같은 행만 제거' })
    ]);

    var t = thresholdText();
    $('threshold-notice').hidden = !t;
    if (t) $('threshold-notice').textContent = t.text;

    var fixes = [], reviews = [];
    r.ruleStats.forEach(function (st) {
      var rule = ruleById(st.id);
      if (st.action === 'auto_fix') fixes.push({ id: st.id, name: st.name, value: st.entries, unit: unitOf(rule) });
      else if (st.action === 'review') reviews.push({ id: st.id, name: st.name, value: st.rowCount, unit: '행' });
    });
    if (r.fallbackStats.entries) reviews.push({ id: '형식', name: '알 수 없는 형식', value: r.fallbackStats.rowCount, unit: '행' });
    renderBars($('chart-fix'), fixes, false);
    renderBars($('chart-review'), reviews, true);
    $('err-total').textContent = '사유 ' + s.reviewEntries + '건 · 수정 ' + s.autoFixCells + '칸';
    $('err-foot').textContent = s.autoFixCells + '칸 자동 수정 · ' + s.removedRows + '행 중복 제거 · ' + s.reviewRows + '행 검토 표시';

    renderRecommendations();
    renderDashTable();
  }

  function renderBars(el, items, isReview) {
    clear(el);
    var max = Math.max.apply(null, items.map(function (i) { return i.value; }).concat([1]));
    items.forEach(function (it) {
      var fill = h('div', { class: 'bar-fill' + (isReview ? ' is-review' : '') });
      fill.style.width = it.value ? (it.value / max * 100) + '%' : '0';
      el.appendChild(h('div', { class: 'bar-item' + (it.value ? '' : ' is-zero'), title: it.id + ' ' + it.name + ': ' + it.value + it.unit },
        h('div', { class: 'bar-head' },
          h('span', { class: 'bar-name' }, h('span', { class: 'rid', text: it.id }), it.name),
          h('span', { class: 'bar-val', text: fmtNum(it.value) + it.unit })),
        h('div', { class: 'bar-track' }, fill)));
    });
  }

  function renderRecommendations() {
    var r = state.result;
    var list = r.active.filter(function (row) { return row.reviews.length; });
    var sorted = list.slice().sort(function (a, b) {
      var ca = a.status === '충돌' ? 1 : 0, cb = b.status === '충돌' ? 1 : 0;
      return (cb - ca) || (b.reviews.length - a.reviews.length) || (a.excelRow - b.excelRow);
    });
    $('rec-sub').textContent = '검토 사유가 많거나 충돌인 행부터 표시';
    $('rec-tag').textContent = list.length + '행 대기';
    var box = $('review-list');
    clear(box);
    sorted.slice(0, 5).forEach(function (row) {
      var first = row.reviews[0];
      box.appendChild(h('button', { class: 'rec-item', type: 'button', onclick: function () { openRow(row); } },
        h('div', { class: 'rec-top' },
          h('span', { class: 'rec-code', text: '#' + row.excelRow + ' ' + code(row) }),
          h('span', { class: 'rec-kind' + (row.status === '충돌' ? ' is-error' : ''), text: row.status === '충돌' ? '시술코드 충돌' : '검토 사유 ' + row.reviews.length + '건' })),
        h('div', { class: 'rec-name', text: nameOf(row) }),
        h('div', { class: 'rec-desc', text: row.reviews.map(function (x) { return x.reason; }).join(' · ') || first.reason })));
    });
    if (!sorted.length) box.appendChild(h('p', { class: 'caption muted', text: '검토가 필요한 행이 없습니다.' }));
    $('btn-review-all').textContent = '검토 항목 확인 (' + list.length + '행)';
  }

  var DASH_FILTERS = [
    { key: 'all', label: '전체 처리', test: function (r) { return r.removed || r.fixes.length || r.reviews.length; } },
    { key: 'fixed', label: '자동 수정', test: function (r) { return r.fixes.length > 0; } },
    { key: 'review', label: '검토 필요', test: function (r) { return !r.removed && r.reviews.length > 0; } },
    { key: 'removed', label: '중복 제거', test: function (r) { return r.removed; } }
  ];

  function renderSegmented(el, items, current, onPick) {
    clear(el);
    items.forEach(function (it) {
      el.appendChild(h('button', { class: 'seg', type: 'button', 'aria-pressed': current === it.key ? 'true' : 'false', onclick: function () { onPick(it.key); } },
        it.dot ? h('span', { class: 'dot ' + it.dot }) : null, it.label, it.n !== undefined ? h('span', { class: 'n', text: '(' + it.n + ')' }) : null));
    });
  }

  function renderDashTable() {
    var r = state.result;
    var f = DASH_FILTERS.filter(function (x) { return x.key === state.dashFilter; })[0];
    renderSegmented($('dash-seg'), DASH_FILTERS.map(function (x) { return { key: x.key, label: x.label, n: r.rows.filter(x.test).length }; }), state.dashFilter,
      function (k) { state.dashFilter = k; state.dashPage = 0; renderDashTable(); });
    var rows = r.rows.filter(f.test);
    var pages = Math.max(1, Math.ceil(rows.length / DASH_PAGE));
    state.dashPage = Math.min(state.dashPage, pages - 1);
    var slice = rows.slice(state.dashPage * DASH_PAGE, (state.dashPage + 1) * DASH_PAGE);

    var t = $('dash-table');
    clear(t.tHead);
    t.tHead.appendChild(h('tr', null, ['원본 행', '시술코드', '시술명', '정상가격', '이벤트가격', '발견된 문제', '처리 상태', '상세'].map(function (x, i) {
      return h('th', { class: i === 0 || i >= 6 ? 'c' : i === 3 || i === 4 ? 'num' : null, text: x });
    })));
    var tb = t.tBodies[0];
    clear(tb);
    slice.forEach(function (row) {
      var src = row.removed ? row.raw : row.cur;
      var bad = !row.removed && row.reviews.length;
      tb.appendChild(h('tr', null,
        h('td', { class: 'row-no', text: '#' + row.excelRow }),
        h('td', { class: 'mono', text: code(row) }),
        h('td', { text: nameOf(row) }),
        h('td', { class: 'num' }, cellNode(src['정상가격'])),
        h('td', { class: 'num' }, cellNode(src['이벤트가격'])),
        h('td', { class: 'wrap' }, h('span', { class: 'problem' + (bad ? ' is-error' : ''), text: problemText(row) })),
        h('td', { class: 'c' }, row.removed ? statusBadge('삭제됨') : row.reviews.length ? statusBadge(row.status) : h('span', { class: 'badge badge-auto', text: '수정됨' })),
        h('td', { class: 'c' }, h('button', { class: 'link-btn', type: 'button', onclick: function () { openRow(row); }, text: '상세' }))));
    });
    $('dash-page-info').textContent = '처리된 ' + rows.length + '행 중 ' + (rows.length ? (state.dashPage * DASH_PAGE + 1) + '–' + (state.dashPage * DASH_PAGE + slice.length) : 0) + '행 표시';
    $('dash-page').textContent = (state.dashPage + 1) + ' / ' + pages + ' 페이지';
    $('dash-prev').disabled = state.dashPage === 0;
    $('dash-next').disabled = state.dashPage >= pages - 1;
  }
  $('dash-prev').addEventListener('click', function () { state.dashPage--; renderDashTable(); });
  $('dash-next').addEventListener('click', function () { state.dashPage++; renderDashTable(); });

  /** 다른 화면에서 행을 열면 검증 화면으로 이동해 선택 */
  function openRow(row) {
    state.filter = 'all'; state.ruleFilter = ''; state.search = ''; $('search').value = '';
    state.selected = row;
    renderValidation();
    go('validation');
    var tr = $('cmp-table').querySelector('tr[data-row="' + row.excelRow + '"]');
    if (tr) tr.scrollIntoView({ block: 'center' });
  }

  $('btn-rerun').addEventListener('click', function () {
    var box = $('rerun-result');
    try {
      var t2 = C.readWorkbook(XLSX, state.fileBytes, state.rules);
      var r2 = C.clean(t2, state.rules);
      var same = C.signature(r2) === C.signature(state.result);
      box.className = 'alert ' + (same ? 'alert-ok' : 'alert-error');
      box.textContent = same
        ? '같은 파일을 다시 읽어 한 번 더 정제했습니다. 정리본, 검토 사유, 변경 기록이 모두 이전 결과와 같습니다.'
        : '두 번 실행한 결과가 다릅니다. 이 결과를 개발 담당자에게 알려 주세요.';
    } catch (err) {
      box.className = 'alert alert-error';
      box.textContent = '다시 실행하지 못했습니다: ' + err.message;
    }
    box.hidden = false;
  });

  /* =========================================================
     정제 규칙
     ========================================================= */
  function describeBasis(b) {
    if (!b) return null;
    if (typeof b === 'string') return b;
    return Object.keys(b).map(function (k) { return b[k]; }).join(' ');
  }

  function renderRules() {
    var list = $('rules-list');
    clear(list);
    if (!state.rules) return;
    var statsById = {};
    if (state.result) state.result.ruleStats.forEach(function (s) { statsById[s.id] = s; });
    var temp = thresholdText();

    state.rules.execution_order.forEach(function (step) {
      if (!step.rules.length) return;
      var group = h('div', { class: 'rule-group' },
        h('div', { class: 'rule-group-head' }, h('span', { class: 'step-num', text: step.step }), step.step + '단계 · ' + step.description));
      step.rules.forEach(function (id) {
        var rule = ruleById(id);
        if (!rule) return;
        var st = statsById[id];
        var unit = rule.action === 'review' ? '행' : unitOf(rule);
        var countNode = st
          ? h('span', { class: 'rule-count' }, '적용 ', h('b', { text: fmtNum(rule.action === 'review' ? st.rowCount : st.entries) }), ' ' + unit)
          : h('span', { class: 'rule-count', text: '파일 대기 중' });
        var exp = rule.expected_impact || {};
        var body = h('div', { class: 'rule-body' },
          h('dl', null,
            h('dt', { text: '처리 방식' }), h('dd', { text: (state.rules.actions || {})[rule.action] || rule.action }),
            h('dt', { text: '적용 컬럼' }), h('dd', { text: rule.columns.join(', ') }),
            h('dt', { text: '처리 이유' }), h('dd', { text: rule.why }),
            describeBasis(rule.criterion_basis) ? [h('dt', { text: '판단 기준' }), h('dd', { text: describeBasis(rule.criterion_basis) })] : null,
            h('dt', { text: '진단 예상' }), h('dd', { text: (exp.count !== undefined ? exp.count + ' ' + (exp.unit || '') : '-') + (exp.note ? ' — ' + exp.note : '') }),
            st ? [h('dt', { text: '실제 적용' }), h('dd', { text: st.entries + '건 · ' + st.rowCount + '개 행' })] : null),
          temp && temp.rule.id === rule.id ? h('div', { class: 'alert alert-warn', text: temp.text }) : null);
        group.appendChild(h('details', { class: 'rule' },
          h('summary', null,
            h('span', { class: 'badge badge-id', text: rule.id }),
            h('span', { class: 'rule-name' }, rule.name, h('span', { class: 'rule-cols', text: rule.columns.join(', ') })),
            actionBadge(rule.action),
            countNode),
          body));
      });
      list.appendChild(group);
    });
  }

  /* =========================================================
     데이터 검증 (원본·정리본 비교)
     ========================================================= */
  var FILTERS = [
    { key: 'all', label: '전체', test: function () { return true; } },
    { key: 'review', label: '검토 필요', dot: 'dot-error', test: function (r) { return !r.removed && r.reviews.length > 0; } },
    { key: 'conflict', label: '충돌', test: function (r) { return r.status === '충돌'; } },
    { key: 'changed', label: '자동 수정', dot: 'dot-ok', test: function (r) { return r.fixes.length > 0; } },
    { key: 'removed', label: '삭제된 행', test: function (r) { return r.removed; } },
    { key: 'ok', label: '정상', test: function (r) { return r.status === '정상'; } }
  ];

  function rowMatchesRule(row, id) {
    if (!id) return true;
    if (row.reviews.some(function (x) { return x.ruleId === id; })) return true;
    if (row.fixes.some(function (x) { return x.ruleId === id; })) return true;
    var rule = ruleById(id);
    return !!(row.removed && rule && rule.operation === 'remove_row');
  }

  function filteredRows() {
    var f = FILTERS.filter(function (x) { return x.key === state.filter; })[0] || FILTERS[0];
    var q = state.search.trim().toLowerCase();
    var key = state.result.keyCol;
    return state.result.rows.filter(function (row) {
      if (!f.test(row)) return false;
      if (!rowMatchesRule(row, state.ruleFilter)) return false;
      if (q) {
        var hay = [row.raw[key], row.cur[key], row.raw['시술명'], row.cur['시술명']].map(function (c) { return c ? String(c.v || '') : ''; }).join(' ').toLowerCase();
        if (hay.indexOf(q) < 0) return false;
      }
      return true;
    });
  }

  function renderValidation() {
    if (!state.result) return;
    var r = state.result, s = r.summary;

    var badges = $('val-badges');
    clear(badges);
    append(badges, [h('span', { class: 'badge badge-review', text: '검토 필요 ' + s.reviewRows + '행' }),
      h('span', { class: 'caption muted', text: '자동 수정 ' + s.autoFixCells + '칸 · 중복 제거 ' + s.removedRows + '행' })]);

    var okRows = r.rows.filter(function (x) { return x.status === '정상'; }).length;
    var k = $('val-kpis');
    clear(k);
    append(k, [
      kpi({ label: '총 검증 행', meta: '', value: fmtNum(s.rowsBefore), unit: '행' }),
      kpi({ label: '검토 대기', meta: pct(s.reviewRows, s.rowsBefore), tone: 'error', value: fmtNum(s.reviewRows), unit: '행' }),
      kpi({ label: '자동 수정', meta: pct(s.autoFixRows, s.rowsBefore), tone: 'ok', value: fmtNum(s.autoFixCells), unit: '칸' }),
      kpi({ label: '정상 통과', meta: pct(okRows, s.rowsBefore), value: fmtNum(okRows), unit: '행' })
    ]);

    renderSegmented($('filter-chips'), FILTERS.map(function (f) { return { key: f.key, label: f.label, dot: f.dot, n: r.rows.filter(f.test).length }; }), state.filter,
      function (key) { state.filter = key; renderValidation(); });

    var sel = $('filter-rule');
    clear(sel);
    sel.appendChild(h('option', { value: '', text: '오류 유형: 전체' }));
    r.ruleStats.forEach(function (st) {
      if (!st.entries) return;
      sel.appendChild(h('option', { value: st.id, text: st.id + ' ' + st.name + ' (' + st.rowCount + '행)' }));
    });
    if (r.fallbackStats.entries) sel.appendChild(h('option', { value: 'format_fallback', text: '알 수 없는 형식 (' + r.fallbackStats.rowCount + '행)' }));
    sel.value = state.ruleFilter;
    if (sel.value !== state.ruleFilter) { state.ruleFilter = ''; sel.value = ''; }
    $('view-mode').value = state.view;

    renderTable();
  }

  function renderTable() {
    var r = state.result;
    var cols = r.colNames;
    var thead = $('cmp-table').tHead;
    clear(thead);
    thead.appendChild(h('tr', null, h('th', { class: 'c', text: '행' }), h('th', { text: '처리 상태' }),
      cols.map(function (c) { return h('th', { class: c === '정상가격' || c === '이벤트가격' ? 'num' : null, text: c }); }),
      h('th', { text: '발견된 문제' })));

    var tbody = $('cmp-table').tBodies[0];
    clear(tbody);
    var rows = filteredRows();
    var frag = document.createDocumentFragment();
    rows.forEach(function (row) {
      var changed = {}, flagged = {};
      row.fixes.forEach(function (f) { changed[f.col] = f; });
      row.reviews.forEach(function (rv) { rv.cols.forEach(function (c) { flagged[c] = true; }); });
      var cls = [];
      if (row.removed) cls.push('is-removed');
      else if (row.status === '충돌') cls.push('is-conflict');
      else if (row.reviews.length) cls.push('is-review');
      if (state.selected === row) cls.push('is-selected');
      var tr = h('tr', { class: cls.join(' '), tabindex: '0', 'data-row': row.excelRow, 'aria-label': '원본 엑셀 ' + row.excelRow + '행 상세 보기' },
        h('td', { class: 'row-no', text: row.excelRow }),
        h('td', null, statusBadge(row.status)));
      cols.forEach(function (c) {
        var td = h('td', { class: 'val' });
        var src = row.removed || state.view === 'orig' ? row.raw[c] : row.cur[c];
        if (src.t === 'number') td.classList.add('num');
        td.appendChild(cellNode(src));
        if (!row.removed && changed[c]) {
          td.classList.add('chg');
          td.title = '원본: ' + cellText(changed[c].before) + ' → 정리본: ' + cellText(changed[c].after);
          if (state.view === 'diff') td.appendChild(h('span', { class: 'old' }, cellNode(changed[c].before)));
        }
        if (!row.removed && flagged[c]) td.classList.add('flag');
        tr.appendChild(td);
      });
      tr.appendChild(h('td', { class: 'problem-cell' }, h('span', { class: 'problem' + (!row.removed && row.reviews.length ? ' is-error' : ''), text: problemText(row) })));
      tr.addEventListener('click', function () { selectRow(row); });
      tr.addEventListener('keydown', function (e) { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); selectRow(row); } });
      frag.appendChild(tr);
    });
    tbody.appendChild(frag);
    $('cmp-count').textContent = '검색 결과 ' + rows.length + '행 / 전체 ' + r.rows.length + '행';
    $('cmp-foot-info').textContent = '표시 중 ' + rows.length + '행 · 전체 ' + r.rows.length + '행 (삭제된 중복 행 포함)';
    updateErrNav(rows);
    renderDetail();
  }

  function reviewRowsIn(rows) { return rows.filter(function (r) { return !r.removed && r.reviews.length; }); }
  function updateErrNav(rows) {
    var list = reviewRowsIn(rows || filteredRows());
    var idx = list.indexOf(state.selected);
    $('err-pos').textContent = list.length ? (idx >= 0 ? idx + 1 : '-') + ' / ' + list.length : '0 / 0';
    $('btn-prev-err').disabled = !list.length || idx === 0;
    $('btn-next-err').disabled = !list.length || idx === list.length - 1;
  }
  function stepErr(dir) {
    var list = reviewRowsIn(filteredRows());
    if (!list.length) return;
    var idx = list.indexOf(state.selected);
    var next = idx < 0 ? (dir > 0 ? 0 : list.length - 1) : Math.max(0, Math.min(list.length - 1, idx + dir));
    selectRow(list[next]);
    var tr = $('cmp-table').querySelector('tr[data-row="' + list[next].excelRow + '"]');
    if (tr) tr.scrollIntoView({ block: 'nearest' });
  }
  $('btn-prev-err').addEventListener('click', function () { stepErr(-1); });
  $('btn-next-err').addEventListener('click', function () { stepErr(1); });

  function selectRow(row) {
    state.selected = row;
    $$('#cmp-table tbody tr.is-selected').forEach(function (tr) { tr.classList.remove('is-selected'); });
    var tr = $('cmp-table').querySelector('tr[data-row="' + row.excelRow + '"]');
    if (tr) tr.classList.add('is-selected');
    updateErrNav();
    renderDetail();
    if (window.matchMedia('(max-width: 1280px)').matches) $('detail').scrollIntoView({ behavior: 'smooth', block: 'start' });
  }

  function renderDetail() {
    var box = $('detail');
    clear(box);
    var row = state.selected;
    if (!row) {
      box.appendChild(h('div', { class: 'insp-empty' }, icon('rule'), h('b', { text: '검토 상세' }), h('span', { class: 'caption', text: '왼쪽 표에서 행을 선택하면 검토 사유와 변경 내용을 보여줍니다.' })));
      return;
    }
    var r = state.result;
    box.appendChild(h('div', { class: 'insp-head' },
      h('div', { class: 'insp-title' }, h('h2', { text: '검토 상세' }), h('span', { class: 'caption muted', text: '원본 ' + row.excelRow + '행' })),
      statusBadge(row.status)));

    box.appendChild(h('dl', { class: 'insp-info' },
      h('div', null, h('dt', { text: '시술 정보' }), h('dd', { text: nameOf(row) })),
      h('div', null, h('dt', { text: '시술 코드' }), h('dd', { class: 'mono', text: code(row) })),
      h('div', null, h('dt', { text: '원본 엑셀 행 · 데이터순번' }), h('dd', { text: row.excelRow + '행 · ' + row.seq })),
      row.fileRow !== row.excelRow ? h('div', null, h('dt', { text: '이번에 올린 파일의 행' }), h('dd', { text: row.fileRow + '행' })) : null,
      h('div', null, h('dt', { text: '적용 규칙' }), h('dd', { text: row.removed ? (r.ruleStats.filter(function (s) { return ruleById(s.id).operation === 'remove_row'; })[0] || {}).id || '-' :
        (row.reviews.map(function (x) { return x.ruleId; }).concat(row.fixes.map(function (x) { return x.ruleId; })).filter(function (v, i, a) { return a.indexOf(v) === i; }).join(', ') || '없음') }))));

    if (row.removed) {
      box.appendChild(h('div', { class: 'alert alert-info', text: '완전 중복 행이라 삭제되었습니다. 같은 내용의 원본 엑셀 ' + row.keptRow + '행을 남겼습니다.' }));
    }

    if (row.reviews.length) {
      var sec = h('div', { class: 'insp-section' }, h('h4', { text: '검토 사유 (' + row.reviews.length + '건)' }));
      row.reviews.forEach(function (rv) {
        sec.appendChild(h('div', { class: 'reason' + (rv.flagType === 'conflict' ? ' is-conflict' : '') },
          h('div', { class: 'reason-top' }, h('span', { class: 'badge badge-id', text: rv.ruleId }), rv.reason),
          h('div', { class: 'reason-why', text: '이유: ' + rv.why })));
      });
      box.appendChild(sec);
    } else if (!row.removed) {
      box.appendChild(h('div', { class: 'insp-section' }, h('h4', { text: '검토 사유' }), h('p', { class: 'caption muted', text: '검토가 필요한 항목이 없습니다.' })));
    }

    if (row.fixes.length) {
      var fx = h('div', { class: 'insp-section' }, h('h4', { text: '자동 수정 (' + row.fixes.length + '칸)' }));
      row.fixes.forEach(function (f) {
        var rule = ruleById(f.ruleId);
        fx.appendChild(h('div', { class: 'reason is-fix' },
          h('div', { class: 'reason-top' }, h('span', { class: 'badge badge-id', text: f.ruleId }), f.col + ': ' + cellText(f.before) + ' → ' + cellText(f.after)),
          rule ? h('div', { class: 'reason-why', text: rule.name }) : null));
      });
      box.appendChild(fx);
    }

    var changed = {};
    row.fixes.forEach(function (f) { changed[f.col] = true; });
    var tb = h('tbody');
    r.colNames.forEach(function (c) {
      tb.appendChild(h('tr', { class: changed[c] ? 'changed' : null },
        h('th', { text: c }),
        h('td', null, cellNode(row.raw[c])),
        h('td', null, row.removed ? '-' : cellNode(row.cur[c]))));
    });
    box.appendChild(h('div', { class: 'insp-section' }, h('h4', { text: '값 비교' }),
      h('table', { class: 'kv' }, h('thead', null, h('tr', null, h('th', { text: '컬럼' }), h('th', { text: '원본' }), h('th', { text: row.removed ? '-' : '정리본' }))), tb)));
  }

  $('filter-rule').addEventListener('change', function (e) { state.ruleFilter = e.target.value; renderTable(); });
  $('view-mode').addEventListener('change', function (e) { state.view = e.target.value; renderTable(); });
  $('search').addEventListener('input', function (e) { state.search = e.target.value; renderTable(); });
  $('btn-reset-filter').addEventListener('click', function () {
    state.filter = 'all'; state.ruleFilter = ''; state.search = ''; state.view = 'diff';
    $('search').value = '';
    renderValidation();
  });

  /* =========================================================
     변경 기록 · 다운로드
     ========================================================= */
  function baseName() { return (state.fileName || '가격표').replace(/\.[^.]+$/, ''); }

  function download(kind) {
    if (!state.result) return;
    try {
      var wbs = C.buildWorkbooks(XLSX, state.result, state.rules);
      var names = { cleaned: '_정리본_', log: '_변경기록_', removed: '_제거된중복행_' };
      XLSX.writeFile(wbs[kind], baseName() + names[kind] + today() + '.xlsx', { compression: true });
    } catch (err) { showError('파일을 만들지 못했습니다: ' + err.message); }
  }

  var LOG_FILTERS = [
    { key: 'all', label: '전체', test: function () { return true; } },
    { key: 'fix', label: '자동 수정', test: function (e) { return /^자동 수정$/.test(e['처리방식']); } },
    { key: 'review', label: '검토 표시', test: function (e) { return e['처리방식'] === '검토 표시'; } },
    { key: 'remove', label: '행 삭제', test: function (e) { return /행 삭제/.test(e['처리방식']); } }
  ];

  function renderHistory() {
    var r = state.result;
    $('dl-removed').disabled = !r.removed.length;
    $('dl-removed-desc').textContent = r.removed.length
      ? '완전 중복으로 제거된 ' + r.removed.length + '행입니다. 각 행이 어느 행과 같아서 지워졌는지 함께 기록됩니다.'
      : '제거된 중복 행이 없습니다.';
    renderLogTable();

    var pv = $('removed-preview');
    clear(pv);
    $('removed-card').hidden = !r.removed.length;
    if (!r.removed.length) return;
    var cols = r.colNames;
    var tb = h('tbody');
    r.removed.forEach(function (row) {
      tb.appendChild(h('tr', null, h('td', { class: 'row-no', text: '#' + row.excelRow }), h('td', { class: 'row-no', text: '#' + row.keptRow }),
        cols.map(function (c) { return h('td', { class: row.raw[c].t === 'number' ? 'num' : null }, cellNode(row.raw[c])); })));
    });
    pv.appendChild(h('table', { class: 'tbl' },
      h('thead', null, h('tr', null, h('th', { class: 'c', text: '원본 행' }), h('th', { class: 'c', text: '유지한 행' }), cols.map(function (c) { return h('th', { text: c }); }))), tb));
  }

  function renderLogTable() {
    var r = state.result;
    var fields = state.rules.change_log && state.rules.change_log.fields || [];
    renderSegmented($('log-seg'), LOG_FILTERS.map(function (f) { return { key: f.key, label: f.label, n: r.log.filter(f.test).length }; }), state.logFilter,
      function (k) { state.logFilter = k; renderLogTable(); });
    var f = LOG_FILTERS.filter(function (x) { return x.key === state.logFilter; })[0];
    var rows = r.log.filter(f.test);
    $('log-sub').textContent = '전체 ' + r.log.length + '건 · 원본_엑셀행번호는 머리글을 포함한 실제 엑셀 행 번호, 데이터순번은 머리글을 뺀 순번입니다.';
    var t = $('log-table');
    clear(t.tHead);
    t.tHead.appendChild(h('tr', null, fields.map(function (x) { return h('th', { text: x }); })));
    var tb = t.tBodies[0];
    clear(tb);
    var frag = document.createDocumentFragment();
    rows.forEach(function (e) {
      frag.appendChild(h('tr', null, fields.map(function (x) {
        var v = e[x];
        var cls = x === '원본_엑셀행번호' || x === '데이터순번' ? 'row-no' : x === '규칙ID' || x === '시술코드' ? 'mono' : x === '처리이유' || x === '변경전값' ? 'wrap' : null;
        return h('td', { class: cls, text: v === null || v === undefined ? '' : typeof v === 'number' ? fmtNum(v) : String(v) });
      })));
    });
    tb.appendChild(frag);
  }

  /* =========================================================
     정답 대조
     ========================================================= */
  function renderBenchmarkHead() {
    $('bm-target-name').textContent = baseName() + '_정리본';
    $('bm-target-rows').textContent = '정제 완료 (' + fmtNum(state.result.summary.rowsAfter) + '행)';
  }

  function resetAnswer() {
    state.answer = null;
    clear($('ans-result'));
    $('ans-file').textContent = '선택된 정답 파일 없음';
    $('ans-file-sub').textContent = '정답 파일은 이 브라우저에서만 읽고 저장·전송하지 않습니다.';
    $('ans-pick-label').textContent = '파일 선택';
  }

  function decodeText(buf) {
    try { return new TextDecoder('utf-8', { fatal: true }).decode(buf); }
    catch (e) { return new TextDecoder('euc-kr').decode(buf); }
  }

  $('answer-input').addEventListener('change', function (e) {
    var f = e.target.files && e.target.files[0];
    e.target.value = '';
    if (!f || !state.result) return;
    if (!/\.csv$/i.test(f.name)) { showError('정답 파일은 CSV(.csv)만 선택할 수 있습니다.'); return; }
    clearError();
    f.arrayBuffer().then(function (buf) {
      var res;
      try { res = C.compareWithAnswer(XLSX, decodeText(new Uint8Array(buf)), state.result); }
      catch (err) { showError('정답과 비교하지 못했습니다: ' + err.message); return; }
      state.answer = res;
      state.ansFilter = res.mismatched ? 'mismatch' : 'all';
      state.ansSearch = '';
      $('ans-file').textContent = f.name;
      $('ans-file-sub').textContent = '정답 ' + fmtNum(res.answerRows) + '행 · 이 브라우저에서만 읽음';
      $('ans-pick-label').textContent = '변경';
      renderAnswer();
    });
  });

  var ANS_FILTERS = [
    { key: 'all', label: '1:1 대응 전체', test: function () { return true; } },
    { key: 'mismatch', label: '불일치', dot: 'dot-error', test: function (p) { return p.cols.length > 0; } },
    { key: 'flagged', label: '검토 표시됨', test: function (p) { return p.cols.length > 0 && p.flagged; } },
    { key: 'missed', label: '표시 못함', test: function (p) { return p.cols.length > 0 && !p.flagged; } },
    { key: 'match', label: '완전 일치', test: function (p) { return !p.cols.length; } },
    { key: 'held', label: '대응 보류', dot: 'dot-idle', held: true }
  ];

  function basisRow(term, desc, value) {
    return h('div', { class: 'basis-row' }, h('dt', { text: term }), h('dd', null, desc, value !== undefined ? h('b', { class: 'basis-val', text: value }) : null));
  }

  function renderAnswer() {
    var res = state.answer;
    var box = $('ans-result');
    clear(box);
    if (!res) return;
    var keyName = res.keyMode === 'row' ? '원본 행 번호(' + res.keyColumnName + ')' : res.keyColumnName;

    box.appendChild(h('div', { class: 'alert alert-info', text: res.keyMode === 'row'
      ? '정답 파일의 "' + res.keyColumnName + '" 컬럼(원본 행 번호)으로 행을 맞췄습니다. 같은 번호가 한쪽에 두 번 이상 있으면 행을 고르지 않고 "대응 보류"로 둡니다.'
      : '정답 파일에 원본 행 번호 컬럼이 없어 "' + res.keyColumnName + '"로 행을 맞췄습니다. 같은 시술코드가 한쪽에 두 번 이상 있으면 어느 행이 정답과 대응하는지 정하지 않고 "대응 보류"로 둡니다.' }));
    if (res.missingColumns.length) box.appendChild(h('div', { class: 'alert alert-warn', text: '정답 파일에 없어 비교하지 못한 컬럼: ' + res.missingColumns.join(', ') }));

    var rate = res.pairedRows ? res.matched / res.pairedRows * 100 : 0;
    box.appendChild(h('div', { class: 'kpi-grid' },
      kpi({ label: '일치율', meta: '1:1 대응 행 기준', value: rate.toFixed(1), unit: '%', bar: rate, foot: res.matched + ' / ' + res.pairedRows }),
      kpi({ label: '완전 일치', meta: '', tone: 'ok', value: fmtNum(res.matched), unit: '행', foot: '비교 컬럼 ' + res.comparedColumns.length + '개가 모두 같음' }),
      kpi({ label: '불일치', meta: '다른 칸 ' + res.diffCells, tone: 'error', value: fmtNum(res.mismatched), unit: '행', foot: '검토 표시됨 ' + res.mismatchFlagged + ' · 표시 못함 ' + res.mismatchUnflagged }),
      kpi({ label: '대응 보류', meta: '일치율에서 제외', value: fmtNum(res.heldCodes), unit: '개 코드', foot: '정리본 ' + res.heldCleanedRows + '행 · 정답 ' + res.heldAnswerRows + '행' })));

    var pct2 = function (a, b) { return b ? (a / b * 100).toFixed(1) + '%' : '-'; };
    box.appendChild(h('div', { class: 'card basis' },
      h('div', { class: 'card-title-row' }, h('div', { class: 'card-title' }, icon('info', 'ic-primary'), '계산 기준')),
      h('dl', null,
        basisRow('1:1 대응', keyName + '이(가) 정리본과 정답에 각각 한 번씩만 있는 행 → ', res.pairedRows + '행'),
        basisRow('일치', '1:1 대응 행 중 비교 컬럼 ' + res.comparedColumns.length + '개(' + res.comparedColumns.join(', ') + ')가 모두 같은 행 → ', res.matched + '행'),
        basisRow('불일치', '1:1 대응 행 중 한 칸 이상 다른 행 → ', res.mismatched + '행 (' + res.diffCells + '칸)'),
        basisRow('대응 보류', '같은 키가 어느 한쪽에 두 번 이상 있어 대응할 정답 행을 정할 수 없는 경우. 일치·불일치·일치율에 넣지 않고, 자동으로 지우거나 고치지 않습니다 → ',
          res.heldCodes + '개 코드 (정리본 ' + res.heldCleanedRows + '행 · 정답 ' + res.heldAnswerRows + '행)'),
        basisRow('일치율', '일치 ÷ 1:1 대응 행 = ' + res.matched + ' ÷ ' + res.pairedRows + ' = ', pct2(res.matched, res.pairedRows)),
        basisRow('참고', '정답 전체 기준: 일치 ÷ 정답 행 = ' + res.matched + ' ÷ ' + res.answerRows + ' = ', pct2(res.matched, res.answerRows)),
        basisRow('행 수 확인', '정답 ' + res.answerRows + '행 = 1:1 ' + res.pairedRows + ' + 보류 ' + res.heldAnswerRows + ' + 정답에만 ' + res.onlyAnswer.length +
          ' · 정리본 ' + res.cleanedRows + '행 = 1:1 ' + res.pairedRows + ' + 보류 ' + res.heldCleanedRows + ' + 정리본에만 ' + res.onlyCleaned.length))));

    var tools = h('div', { class: 'card toolbar' });
    var seg = h('div', { class: 'segmented', role: 'group', 'aria-label': '대조 필터' });
    renderSegmented(seg, ANS_FILTERS.map(function (f) { return { key: f.key, label: f.label, dot: f.dot, n: f.held ? res.heldCodes : res.pairs.filter(f.test).length }; }), state.ansFilter,
      function (k) { state.ansFilter = k; renderAnswer(); });
    var search = h('input', { type: 'search', placeholder: '시술코드, 시술명 또는 컬럼 검색', 'aria-label': '대조 결과 검색', value: state.ansSearch });
    search.addEventListener('input', function (e) { state.ansSearch = e.target.value; renderAnswerTable(); });
    var tg = h('input', { type: 'checkbox', id: 'hl-toggle' });
    tg.checked = state.ansHighlight;
    tg.addEventListener('change', function (e) { state.ansHighlight = e.target.checked; $('ans-table').classList.toggle('no-hl', !state.ansHighlight); });
    append(tools, [h('div', { class: 'tb-group' }, seg),
      h('div', { class: 'tb-group tb-end' }, h('div', { class: 'search' }, icon('search'), search),
        h('label', { class: 'toggle', for: 'hl-toggle' }, tg, h('span', { class: 'track' }), '차이 강조 표시'))]);
    box.appendChild(tools);

    var card = h('div', { class: 'card card-flush' },
      h('div', { class: 'table-wrap table-tall' }, h('table', { class: 'tbl' + (state.ansHighlight ? '' : ' no-hl'), id: 'ans-table' }, h('thead'), h('tbody'))),
      h('div', { class: 'tbl-foot' }, h('span', { id: 'ans-foot' }), h('span', { text: '정답 파일의 값은 비교에만 쓰며 정리본을 바꾸지 않습니다.' })));
    box.appendChild(card);
    renderAnswerTable();

    if (res.onlyCleaned.length || res.onlyAnswer.length) {
      var ul = h('ul', { class: 'side-list' });
      res.onlyCleaned.forEach(function (x) {
        ul.appendChild(h('li', { text: '정리본에만 있음: 원본 엑셀 ' + x.row.excelRow + '행 (' + C.display(x.row.cur[state.result.keyCol]) + ') · ' + (x.flagged ? '검토 표시됨' : '검토 표시 없음') }));
      });
      res.onlyAnswer.forEach(function (x) { ul.appendChild(h('li', { text: '정답에만 있음: ' + x.key + ' (정답 CSV ' + x.line + '행)' })); });
      box.appendChild(h('div', { class: 'card' }, h('div', { class: 'card-title-row' }, h('div', { class: 'card-title', text: '한쪽에만 있는 행' })), ul));
    }
  }

  function cmpCell(pair, col, numeric) {
    var cleaned = C.isBlank(pair.row.cur[col]) ? '(빈칸)' : cellText(pair.row.cur[col]);
    var diff = pair.cols.filter(function (d) { return d.col === col; })[0];
    if (!diff) return h('div', { class: 'cmp-cell' }, h('span', { text: cleaned }));
    var ans = diff.answer === '' ? '(빈칸)' : numeric && isFinite(Number(diff.answer)) ? fmtNum(Number(diff.answer)) : diff.answer;
    return h('div', { class: 'cmp-cell' },
      h('span', null, h('span', { class: 'lbl', text: '정리' }), h('span', { class: 'bad', text: cleaned })),
      h('span', null, h('span', { class: 'lbl', text: '정답' }), h('span', { class: 'ans', text: ans })));
  }

  function ansText(col, v) {
    if (v === '') return '(빈칸)';
    return (col === '정상가격' || col === '이벤트가격') && isFinite(Number(v)) ? fmtNum(Number(v)) : v;
  }

  /** 대응 보류: 같은 키의 정리본 행과 정답 행을 모두 보여주고, 어느 쪽도 고르지 않는다 */
  function renderHeldTable(t) {
    var res = state.answer;
    var q = state.ansSearch.trim().toLowerCase();
    var groups = res.held.filter(function (g) {
      if (!q) return true;
      var hay = [g.key].concat(g.cleaned.map(function (c) { return nameOf(c.row); })).concat(g.answers.map(function (a) { return a.values['시술명'] || ''; })).join(' ').toLowerCase();
      return hay.indexOf(q) >= 0;
    });
    var show = ['시술명', '시술부위', '시술횟수', '정상가격', '이벤트가격'];
    clear(t.tHead);
    t.tHead.appendChild(h('tr', null, h('th', { text: '구분' }), h('th', { text: '행' }),
      show.map(function (c) { return h('th', { class: c === '정상가격' || c === '이벤트가격' ? 'num' : null, text: c }); }),
      h('th', { text: '이벤트 기간' }), h('th', { text: '정리기 상태 · 검토 사유' }), h('th', { text: '정답 행과 다른 칸 (참고)' })));
    var tb = t.tBodies[0];
    clear(tb);
    var frag = document.createDocumentFragment();
    var span = show.length + 5;
    groups.forEach(function (g) {
      frag.appendChild(h('tr', { class: 'held-head' }, h('td', { colspan: String(span) },
        h('span', { class: 'badge badge-held', text: '대응 보류' }),
        h('span', { class: 'mono held-code', text: g.key }),
        h('span', { class: 'caption muted', text: '정리본 ' + g.cleaned.length + '행 · 정답 ' + g.answers.length + '행 — 어느 행이 정답과 대응하는지 정하지 않았습니다. 자동으로 지우거나 고치지 않습니다.' }))));
      g.cleaned.forEach(function (c) {
        var r = c.row;
        frag.appendChild(h('tr', { class: 'held-row' },
          h('td', null, h('span', { class: 'badge badge-id', text: '정리본' })),
          h('td', { class: 'row-no', text: '원본 ' + r.excelRow + '행' }),
          show.map(function (col) { return h('td', { class: col === '정상가격' || col === '이벤트가격' ? 'num' : null }, cellNode(r.cur[col])); }),
          h('td', { text: cellText(r.cur['이벤트시작일']) + ' ~ ' + cellText(r.cur['이벤트종료일']) }),
          h('td', { class: 'wrap' }, statusBadge(r.status), ' ', h('span', { class: 'problem', text: r.reasonText || '' })),
          h('td', { class: 'wrap' }, c.cols === null ? h('span', { class: 'caption muted', text: '정답 행이 여러 개라 계산하지 않음' })
            : c.cols.length ? h('span', { class: 'problem', text: c.cols.length + '칸: ' + c.cols.map(function (d) { return d.col; }).join(', ') })
              : h('span', { class: 'caption muted', text: '0칸' }))));
      });
      g.answers.forEach(function (a) {
        frag.appendChild(h('tr', { class: 'held-row is-answer' },
          h('td', null, h('span', { class: 'badge badge-auto', text: '정답' })),
          h('td', { class: 'row-no', text: 'CSV ' + a.line + '행' }),
          show.map(function (col) { return h('td', { class: col === '정상가격' || col === '이벤트가격' ? 'num' : null, text: ansText(col, a.values[col] || '') }); }),
          h('td', { text: (a.values['이벤트시작일'] || '(빈칸)') + ' ~ ' + (a.values['이벤트종료일'] || '(빈칸)') }),
          h('td', { class: 'caption muted', text: '-' }),
          h('td', { class: 'caption muted', text: '-' })));
      });
    });
    tb.appendChild(frag);
    $('ans-foot').textContent = '대응 보류 ' + groups.length + '개 코드 표시 · 정리본 ' + res.heldCleanedRows + '행 · 정답 ' + res.heldAnswerRows + '행 (일치율 계산에서 제외)';
  }

  function renderAnswerTable() {
    var res = state.answer;
    var t = $('ans-table');
    if (!t) return;
    if (state.ansFilter === 'held') { renderHeldTable(t); return; }
    var f = ANS_FILTERS.filter(function (x) { return x.key === state.ansFilter; })[0];
    var q = state.ansSearch.trim().toLowerCase();
    var rows = res.pairs.filter(function (p) {
      if (!f.test(p)) return false;
      if (!q) return true;
      var hay = [code(p.row), nameOf(p.row)].concat(p.cols.map(function (d) { return d.col; })).join(' ').toLowerCase();
      return hay.indexOf(q) >= 0;
    });
    clear(t.tHead);
    t.tHead.appendChild(h('tr', null, h('th', { class: 'c', text: '행' }), h('th', { text: '시술코드' }), h('th', { text: '시술명 (정리본 vs 정답)' }),
      h('th', { class: 'num', text: '정상가격 대조' }), h('th', { class: 'num', text: '이벤트가격 대조' }), h('th', { class: 'c', text: '판정 결과' }),
      h('th', { text: '그 밖의 차이' }), h('th', { class: 'c', text: '검토 표시' }), h('th', { text: '정리기의 검토 사유' })));
    var tb = t.tBodies[0];
    clear(tb);
    var frag = document.createDocumentFragment();
    var main = ['시술명', '정상가격', '이벤트가격'];
    rows.forEach(function (p) {
      var bad = p.cols.length > 0;
      var others = p.cols.filter(function (d) { return main.indexOf(d.col) < 0; });
      frag.appendChild(h('tr', { class: bad ? 'is-mismatch' : null },
        h('td', { class: 'row-no', text: p.row.excelRow }),
        h('td', { class: 'mono', text: code(p.row) }),
        h('td', null, cmpCell(p, '시술명', false)),
        h('td', { class: 'num' }, cmpCell(p, '정상가격', true)),
        h('td', { class: 'num' }, cmpCell(p, '이벤트가격', true)),
        h('td', { class: 'c' }, h('span', { class: 'verdict ' + (bad ? 'is-bad' : 'is-ok'), text: bad ? '불일치 ' + p.cols.length + '칸' : '일치' })),
        h('td', { class: 'wrap' }, others.length ? others.map(function (d, i) {
          return h('div', { class: 'cmp-cell' }, h('span', null, h('span', { class: 'lbl', text: d.col }), h('span', { class: 'bad', text: d.cleaned === '' ? '(빈칸)' : d.cleaned }), ' → ', h('span', { class: 'ans', text: d.answer === '' ? '(빈칸)' : d.answer })));
        }) : h('span', { class: 'caption muted', text: '-' })),
        h('td', { class: 'c' }, bad ? h('span', { class: 'badge ' + (p.flagged ? 'badge-review' : 'badge-conflict'), text: p.flagged ? '검토 표시됨' : '표시 못함' }) : h('span', { class: 'caption muted', text: '-' })),
        h('td', { class: 'wrap' }, h('span', { class: 'problem', text: p.row.reasonText || '-' }))));
    });
    tb.appendChild(frag);
    $('ans-foot').textContent = '표시 중 ' + rows.length + '행 · 1:1 대응 ' + res.pairs.length + '행 · 대응 보류 ' + res.heldCodes + '개 코드는 "대응 보류" 탭에서 확인';
  }

  /* =========================================================
     시작
     ========================================================= */
  if (!window.XLSX || !C) {
    showError('엑셀 처리 파일(lib/xlsx.full.min.js 또는 clean.js)을 불러오지 못했습니다. 폴더 구조를 확인해 주세요.');
    return;
  }
  setHasResult(false);
  renderDetail();
  showView(currentView());
  loadRules();
})();
