/* APQP Smart Review — 화면 (대시보드 · 서류 업로드 · 정합성 검토 · 공정 대조표 · 승인 검토 보고서 · 검토 이력 · 설정)
   계산은 js/logic.js, 저장은 js/store.js, AI 연결은 js/ai-endpoint.js · js/ai-panel.js 에 있습니다.
   화면 틀·부품은 같은 수강생의 과제 01(data10-07)과 같습니다 — 2단계에서 「협력사 품질 포털」의 메뉴로 합칩니다. */
(function () {
  'use strict';
  var L = window.APLogic;
  var S = window.APStore;
  var Sample = window.APSample;
  var XLSX = window.XLSX;

  var db = S.loadDb();
  var ui = {
    imp: null,
    newForm: false,
    find: { verdict: 'issues', rule: '', doc: '', q: '' },
    rep: { aiText: '', parsed: null, showPrompt: false },
    hist: {}
  };

  // ── 작은 도구 ─────────────────────────────────────────────
  function h(tag, attrs) {
    var el = document.createElement(tag);
    if (attrs) Object.keys(attrs).forEach(function (k) {
      var v = attrs[k];
      if (v == null || v === false) return;
      if (k === 'class') el.className = v;
      else if (k === 'text') el.textContent = v;
      else if (k === 'value') el.value = v;
      else if (k.slice(0, 2) === 'on' && typeof v === 'function') el.addEventListener(k.slice(2), v);
      else if (v === true) el.setAttribute(k, '');
      else el.setAttribute(k, v);
    });
    for (var i = 2; i < arguments.length; i++) add(el, arguments[i]);
    return el;
  }
  function add(el, c) {
    if (c == null || c === false) return;
    if (Array.isArray(c)) { c.forEach(function (x) { add(el, x); }); return; }
    el.appendChild(typeof c === 'string' || typeof c === 'number' ? document.createTextNode(String(c)) : c);
  }
  function toast(msg, isError) {
    var t = document.getElementById('toast');
    t.textContent = msg;
    t.className = 'toast' + (isError ? ' error' : '');
    t.hidden = false;
    clearTimeout(toast._t);
    toast._t = setTimeout(function () { t.hidden = true; }, 3800);
  }
  function save() {
    if (!S.saveDb(db)) {
      document.getElementById('storeBanner').hidden = false;
      if (S.lastError() === 'quota') toast('브라우저 저장 공간이 찼습니다. 설정에서 JSON 백업을 받아 두세요.', true);
    }
  }
  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function today() { var d = new Date(); return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate()); }
  function nowIso() { var d = new Date(); return today() + 'T' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) + ':' + pad2(d.getSeconds()); }
  function field(label, input, hint, cls) {
    return h('label', { class: 'field' + (cls ? ' ' + cls : '') }, h('span', null, label), input, hint ? h('small', { class: 'hint' }, hint) : null);
  }
  function select(options, value, attrs) {
    var s = h('select', attrs || {});
    options.forEach(function (o) {
      var v = Array.isArray(o) ? o[0] : o, t = Array.isArray(o) ? o[1] : o;
      s.appendChild(h('option', { value: v, selected: String(v) === String(value) }, t));
    });
    return s;
  }
  var onDialogClose = null;
  function openDialog(title, content, actions) {
    var d = document.getElementById('dialog');
    document.getElementById('dialogTitle').textContent = title;
    var c = document.getElementById('dialogContent'); c.textContent = ''; add(c, content);
    var a = document.getElementById('dialogActions'); a.textContent = ''; add(a, actions);
    if (!d.open) d.showModal();
  }
  function closeDialog() { var d = document.getElementById('dialog'); if (d.open) d.close(); }
  document.getElementById('dialog').addEventListener('close', function () { var f = onDialogClose; onDialogClose = null; if (f) f(); });
  function confirmBox(title, text, okLabel, onOk, danger) {
    openDialog(title, h('p', { class: 'pre' }, text), [
      h('button', { type: 'button', class: 'btn', onclick: closeDialog }, '그만두기'),
      h('button', { type: 'button', class: 'btn ' + (danger ? 'btn-danger' : 'btn-primary'), onclick: function () { closeDialog(); onOk(); } }, okLabel)
    ]);
  }
  function download(name, blob) {
    var a = h('a', { href: URL.createObjectURL(blob), download: name });
    document.body.appendChild(a); a.click();
    setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
  }
  function writeXlsx(name, sheets, widths) {
    var wb = XLSX.utils.book_new();
    Object.keys(sheets).forEach(function (n) {
      var ws = XLSX.utils.aoa_to_sheet(sheets[n]);
      var w = (widths && widths[n]) || [];
      var cols = 0; sheets[n].forEach(function (r) { cols = Math.max(cols, r.length); });
      ws['!cols'] = Array.from({ length: cols }, function (_, i) { return { wch: w[i] || 14 }; });
      XLSX.utils.book_append_sheet(wb, ws, n.replace(/[\\\/?*\[\]:]/g, '-').slice(0, 31));
    });
    var out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    download(name, new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }));
  }
  function copyText(s) {
    function fallback() {
      var ta = h('textarea', { style: 'position:fixed;left:-9999px;top:0' });
      ta.value = s; document.body.appendChild(ta); ta.select();
      var ok = false;
      try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
      ta.remove();
      toast(ok ? '복사했습니다' : '복사하지 못했습니다. 글상자를 눌러 직접 복사하세요', !ok);
    }
    if (navigator.clipboard && window.isSecureContext) navigator.clipboard.writeText(s).then(function () { toast('복사했습니다'); }, fallback);
    else fallback();
  }
  function tag() { return db._sample ? '_예시데이터' : ''; }
  function safeName(s) { return String(s || '').replace(/[\\\/?*\[\]:"<>|\s]+/g, ''); }
  function pageHead(title, sub, right) {
    return h('div', { class: 'page-head' }, h('div', { class: 'ph-text' }, h('h1', null, title), sub ? h('p', { class: 'note' }, sub) : null), right || null);
  }

  // ── 검토 건 ───────────────────────────────────────────────
  function byId(id) { return db.reviews.filter(function (r) { return r.id === id; })[0] || null; }
  function cur() { return byId(db.current) || db.reviews[db.reviews.length - 1] || null; }
  function setCur(id) { if (db.current !== id) { db.current = id; ui.rep = { aiText: '', parsed: null, showPrompt: false }; save(); } }
  function resultOf(rv) { return L.runReview(rv.docs, db.settings); }
  function prevOf(rv) { return L.previousOf(db.reviews, rv); }
  function compareOf(rv, res) { var p = prevOf(rv); return p ? { prev: p, cmp: L.compareResults(resultOf(p), res || resultOf(rv)) } : null; }
  function docCount(rv) { return L.DOC_IDS.filter(function (id) { return rv.docs[id]; }).length; }
  function sortedReviews() {
    return db.reviews.slice().sort(function (a, b) { return a.supplier < b.supplier ? -1 : a.supplier > b.supplier ? 1 : a.partNo < b.partNo ? -1 : a.partNo > b.partNo ? 1 : a.round - b.round; });
  }
  function latestPerPart() {
    var m = {};
    db.reviews.forEach(function (r) { var k = L.reviewKey(r); if (!m[k] || r.round > m[k].round) m[k] = r; });
    return Object.keys(m).map(function (k) { return m[k]; });
  }
  /* 화면 위 「검토 건」 고르기 — base 는 그 화면 주소(#/review 등) */
  function picker(base) {
    var s = select(sortedReviews().map(function (r) { return [r.id, L.reviewTitle(r)]; }), (cur() || {}).id, { 'aria-label': '검토 건' });
    s.addEventListener('change', function () { setCur(s.value); location.hash = base + '/' + encodeURIComponent(s.value); });
    return h('div', { class: 'picker' }, field('검토 건', s));
  }
  function decisionPill(code, label) { return h('span', { class: 'dec dec-' + code }, label || L.DECISION[code]); }
  function verdictBadge(v) { return h('span', { class: 'vd vd-' + v }, L.VERDICT[v]); }
  function noReview(main, title) {
    main.appendChild(pageHead(title));
    main.appendChild(h('section', { class: 'card' },
      h('p', null, '아직 검토 건이 없습니다. 「서류 업로드」에서 검토 건(협력사·품번)을 만들고 APQP 서류를 올리거나, 예시 데이터로 흐름을 먼저 볼 수 있습니다.'),
      h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/upload' }, '서류 업로드로'), sampleButton(false))));
  }
  function noDocs(main) {
    main.appendChild(h('section', { class: 'card' }, h('p', null, '이 검토 건에는 아직 서류가 없습니다.'), h('a', { class: 'btn btn-primary', href: '#/upload' }, '서류 업로드로')));
  }

  // ── 아이콘·차트 부품 ──────────────────────────────────────
  var SVGNS = 'http://www.w3.org/2000/svg';
  function sv(tagName, attrs) {
    var el = document.createElementNS(SVGNS, tagName);
    Object.keys(attrs || {}).forEach(function (k) { el.setAttribute(k, attrs[k]); });
    for (var i = 2; i < arguments.length; i++) { var c = arguments[i]; if (c == null) continue; el.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); }
    return el;
  }
  var ICON = {
    doc: 'M6 3h9l4 4v14H6zM9 12h7M9 16h5M14 3v5h5',
    docs: 'M8 3h8l4 4v12H8zM15 3v5h5M4 7v14h11',
    alert: 'M12 4 2.8 19.5h18.4zM12 10v4.5M12 17.2v.3',
    check: 'M4 4h16v16H4zM8 12.5l3 3 5-6',
    miss: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M8.5 8.5l7 7M15.5 8.5l-7 7',
    ask: 'M12 3.5a8.5 8.5 0 1 0 0 17a8.5 8.5 0 1 0 0-17M9.5 9.5a2.5 2.5 0 1 1 3.5 2.3c-.7.3-1 .9-1 1.7M12 16.5v.3',
    gauge: 'M4 16a8 8 0 1 1 16 0M12 16l4-5M7 16h0M17 16h0',
    stamp: 'M9 4h6v5l3 4H6l3-4zM5 17h14v3H5z',
    chev: 'm9 6 6 6-6 6',
    chip: 'M7 7h10v10H7zM10 10h4v4h-4zM9.5 3v4M14.5 3v4M9.5 17v4M14.5 17v4M3 9.5h4M3 14.5h4M17 9.5h4M17 14.5h4',
    down: 'M12 4v11M7.5 10.5 12 15l4.5-4.5M4 19h16',
    redo: 'M4 12a8 8 0 0 1 14-5.3L20 9M20 4v5h-5M20 12a8 8 0 0 1-14 5.3L4 15M4 20v-5h5'
  };
  function icon(name, cls) { return sv('svg', { viewBox: '0 0 24 24', class: 'ico' + (cls ? ' ' + cls : ''), 'aria-hidden': 'true' }, sv('path', { d: ICON[name] })); }
  function kpi(title, iconName, value, unit, note, tone) {
    return h('div', { class: 'dk' + (tone ? ' tone-' + tone : '') },
      h('span', { class: 'dk-ico' }, icon(iconName)),
      h('div', { class: 'dk-body' }, h('div', { class: 'dk-t' }, title), h('div', { class: 'dk-v' }, value, unit ? h('small', null, unit) : null), note ? h('div', { class: 'dk-note' }, note) : null));
  }
  function panel(title, more, body, cls) {
    return h('section', { class: 'dp' + (cls ? ' ' + cls : '') },
      h('div', { class: 'dp-head' }, h('h2', null, title), more ? h('a', { class: 'dp-more', href: more[0] }, more[1] || '더보기', icon('chev')) : null),
      body);
  }
  function hbars(items, opt) {
    var max = Math.max.apply(null, items.map(function (x) { return x.value; }).concat([1]));
    return h('ul', { class: 'hbars' + (opt && opt.cls ? ' ' + opt.cls : ''), 'aria-label': opt && opt.label },
      items.map(function (x) {
        return h('li', { class: 'hb' + (x.cls ? ' ' + x.cls : '') }, h('span', { class: 'hb-k', title: x.key }, x.key),
          h('span', { class: 'hb-track' }, h('span', { class: 'hb-fill', style: 'width:' + (x.value ? Math.max(2, x.value / max * 100) : 0) + '%' })),
          h('span', { class: 'hb-v' }, x.value + (x.unit || '건')));
      }));
  }

  // ── 라우팅 ────────────────────────────────────────────────
  function routeParts() { return location.hash.replace(/^#\/?/, '').split('/'); }
  function route() { return routeParts()[0] || 'home'; }
  function routeArg() { var p = routeParts()[1]; try { return p ? decodeURIComponent(p) : ''; } catch (e) { return ''; } }
  function render() {
    var r = route();
    var main = document.getElementById('main');
    main.textContent = '';
    document.querySelectorAll('.sidebar a[data-route]').forEach(function (a) {
      if (a.getAttribute('data-route') === r) a.setAttribute('aria-current', 'page'); else a.removeAttribute('aria-current');
    });
    var arg = routeArg();
    if (arg && byId(arg)) setCur(arg);
    document.getElementById('sampleBanner').hidden = !db._sample;
    if (!S.available()) document.getElementById('storeBanner').hidden = false;
    document.body.setAttribute('data-route', r);
    if (r === 'upload') renderUpload(main);
    else if (r === 'review') renderReview(main);
    else if (r === 'matrix') renderMatrix(main);
    else if (r === 'report') renderReport(main);
    else if (r === 'history') renderHistory(main);
    else if (r === 'settings') renderSettings(main);
    else renderHome(main);
  }
  window.addEventListener('hashchange', function () { setMenu(false); render(); window.scrollTo(0, 0); });

  var menuBtn = document.getElementById('menuBtn'), sidebar = document.getElementById('sidebar'), backdrop = document.getElementById('sideBackdrop');
  function setMenu(open) {
    document.body.classList.toggle('menu-open', open);
    menuBtn.setAttribute('aria-expanded', open ? 'true' : 'false');
    menuBtn.setAttribute('aria-label', open ? '메뉴 닫기' : '메뉴 열기');
    backdrop.hidden = !open;
    if (open) { var first = sidebar.querySelector('a'); if (first) first.focus(); }
  }
  menuBtn.addEventListener('click', function () { setMenu(!document.body.classList.contains('menu-open')); });
  backdrop.addEventListener('click', function () { setMenu(false); menuBtn.focus(); });
  document.addEventListener('keydown', function (e) {
    if (e.key === 'Escape' && document.body.classList.contains('menu-open')) { setMenu(false); menuBtn.focus(); }
  });
  sidebar.addEventListener('click', function (e) { if (e.target.closest('a')) setMenu(false); });
  var WEEK = ['일', '월', '화', '수', '목', '금', '토'];
  function nowText() {
    var d = new Date();
    document.getElementById('nowText').textContent = d.getFullYear() + '. ' + pad2(d.getMonth() + 1) + '. ' + pad2(d.getDate()) + ' (' + WEEK[d.getDay()] + ') ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes());
  }
  nowText(); setInterval(nowText, 30000);

  // ── 예시 데이터 ───────────────────────────────────────────
  function sampleButton(primary) {
    return h('button', { type: 'button', class: 'btn' + (primary ? ' btn-primary' : ''), onclick: loadSample }, '예시 데이터로 시작');
  }
  /* 예시 통합 문서를 실제 업로드와 같은 길(병합 칸 채우기 → 서류 종류 추정 → 열 찾기)로 읽습니다 */
  function sampleReview(round) {
    var rv = { id: 'r' + (++db.seq), supplier: Sample.META.supplier, partNo: Sample.META.partNo, partName: Sample.META.partName, stage: Sample.META.stage, round: round,
      submittedAt: Sample.DATES[round], createdAt: nowIso(), docs: {}, opinion: {}, aiText: '' };
    Sample.workbook(round).forEach(function (sh) {
      var aoa = L.fillMerges(JSON.parse(JSON.stringify(sh.aoa)), sh.merges), g = L.guessDocType(aoa, sh.name, Sample.fileName(round));
      if (!g.type) return;
      var p = L.parseDoc(aoa, g.type);
      rv.docs[g.type] = { fileName: Sample.fileName(round), sheet: sh.name, at: nowIso(), rows: p.rows };
    });
    return L.cleanReview(rv);
  }
  function loadSample() {
    function go() {
      var keep = { settings: db.settings, mask: db.mask, mapping: db.mapping };
      db = L.emptyDb(); Object.assign(db, keep);
      var r1 = sampleReview(1), r2 = sampleReview(2);
      db.reviews = [r1, r2]; db.current = r1.id; db._sample = true;
      db.aliases = L.updateAliases(db.aliases, db.reviews);
      save();
      toast('예시 검토 건 2개(1차 제출 · 2차 재제출)를 넣었습니다. 1차 제출부터 보여 드립니다.');
      location.hash = '#/review/' + r1.id;
    }
    if (db.reviews.length && !db._sample) confirmBox('예시 데이터로 바꾸기', '지금 있는 검토 건 ' + db.reviews.length + '개를 지우고 예시 데이터를 넣습니다.\n필요하면 먼저 설정에서 JSON 백업을 받아 두세요.', '바꾸기', go, true);
    else go();
  }

  // ══ 대시보드 ═════════════════════════════════════════════
  function renderHome(main) {
    main.appendChild(h('section', { class: 'hero' },
      h('div', { class: 'hero-text' },
        h('h1', null, 'APQP 서류 검토 현황'),
        h('p', null, '협력사가 낸 초도품 APQP 서류(Flow Chart · 관리계획서 · PFMEA · 검사기준서 · 외주품 리스트)를 서로 대조해 누락·불일치를 찾고, 승인 검토 보고서까지 만듭니다.')),
      db.reviews.length ? h('span', { class: 'hero-chip' }, icon('docs'), '검토 건 ' + db.reviews.length + '개') : null));
    if (!db.reviews.length) {
      main.appendChild(h('section', { class: 'dp dash-empty' },
        h('h2', null, '아직 검토 건이 없습니다'),
        h('p', null, '품번 하나의 APQP 서류를 Excel 로 올리면 이 화면에 검토 현황이 쌓입니다. 실제 서류가 아직 없으면 예시 데이터(가상, 1차 제출 + 2차 재제출)로 흐름을 볼 수 있습니다.'),
        h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/upload' }, '서류 업로드'), sampleButton(false))));
      main.appendChild(flowGuide());
      return;
    }
    var latest = latestPerPart(), dec = { approve: 0, conditional: 0, reject: 0 }, open = 0, sumScore = 0, byRule = {};
    var results = {};
    latest.forEach(function (rv) {
      var res = resultOf(rv); results[rv.id] = res;
      dec[res.summary.decision.code]++; open += res.summary.issues + res.summary.counts.check; sumScore += res.summary.score;
      res.findings.forEach(function (f) { if (f.verdict !== 'ok') byRule[f.rule] = (byRule[f.rule] || 0) + 1; });
    });
    main.appendChild(h('section', { class: 'dkpis' },
      kpi('검토 품번', 'docs', latest.length, '개', '검토 건(제출 차수 포함) ' + db.reviews.length + '개'),
      kpi('열린 지적', 'alert', open, '건', '품번별 최신 제출 기준 · 확인 필요 포함', 'warn'),
      kpi('반려 제안', 'miss', dec.reject, '건', '조건부 승인 ' + dec.conditional + ' · 승인 ' + dec.approve, 'bad'),
      kpi('평균 적합률', 'gauge', Math.round(sumScore / latest.length), '%', '판정 가운데 「적합」의 비율', 'ok')));

    var tb = h('tbody');
    sortedReviews().slice().sort(function (a, b) { return (b.submittedAt || '') < (a.submittedAt || '') ? -1 : 1; }).slice(0, 8).forEach(function (rv) {
      var res = results[rv.id] || resultOf(rv), sm = res.summary;
      function go() { setCur(rv.id); location.hash = '#/review/' + encodeURIComponent(rv.id); }
      tb.appendChild(h('tr', { class: 'clickable', tabindex: '0', onclick: go, onkeydown: function (e) { if (e.key === 'Enter') go(); } },
        h('td', { class: 'nw' }, rv.supplier), h('td', { class: 'nw' }, rv.partNo), h('td', { class: 'num' }, rv.round + '차'), h('td', { class: 'nw' }, rv.submittedAt || ''),
        h('td', { class: 'num' }, sm.score + '%'), h('td', { class: 'num' }, sm.issues + sm.counts.check + '건'),
        h('td', null, decisionPill(sm.decision.code)), h('td', null, rv.opinion.decision ? decisionPill(rv.opinion.decision) : h('span', { class: 'note' }, '미정'))));
    });
    var recent = panel('최근 검토 건', ['#/history', '검토 이력'], h('div', { class: 'table-wrap flat' }, h('table', { class: 'list dash-table' },
      h('thead', null, h('tr', null, ['협력사', '품번', '차수', '제출일', '적합률', '지적', '의견 제안', '검토자 결정'].map(function (t) { return h('th', null, t); }))), tb)), 'a-recent');
    var ruleItems = L.RULES.map(function (r) { return { key: r.name, value: byRule[r.id] || 0 }; }).filter(function (x) { return x.value; }).sort(function (a, b) { return b.value - a.value; });
    var rules = panel('규칙별 열린 지적', ['#/settings', '규칙 설정'], ruleItems.length ? hbars(ruleItems, { label: '규칙별 열린 지적', cls: 'wide' }) : h('p', { class: 'note' }, '품번별 최신 제출에 열린 지적이 없습니다.'));
    main.appendChild(h('div', { class: 'dash-low' }, recent, rules));
    main.appendChild(flowGuide());
  }
  function flowGuide() {
    var steps = [
      ['#/upload', '서류 업로드', '검토 건(협력사·품번)을 만들고 APQP 서류 Excel 을 올립니다'],
      ['#/review', '정합성 검토', '규칙으로 서류끼리 대조해 누락·불일치·확인 필요를 찾습니다'],
      ['#/matrix', '공정 대조표', '공정 × 서류 표로 어느 서류에 무엇이 빠졌는지 봅니다'],
      ['#/report', '승인 검토 보고서', '승인 의견 제안을 검토하고 보고서를 Excel·인쇄·Word 로 냅니다'],
      ['#/history', '검토 이력', '재제출 때 앞 결과와 비교해 무엇이 고쳐졌는지 봅니다']
    ];
    return panel('일하는 순서', null, h('ol', { class: 'flow' }, steps.map(function (s, i) {
      return h('li', null, h('a', { href: s[0] }, h('span', { class: 'flow-n' }, String(i + 1)), h('span', { class: 'flow-t' }, h('strong', null, s[1]), h('span', null, s[2])), icon('chev', 'q-chev')));
    })), 'flow-panel');
  }

  // ══ 서류 업로드 ═══════════════════════════════════════════
  function decodeCsv(buf) {
    var u8 = new Uint8Array(buf), t;
    try { t = new TextDecoder('utf-8', { fatal: true }).decode(u8); }
    catch (e) { try { t = new TextDecoder('euc-kr').decode(u8); } catch (e2) { t = new TextDecoder('utf-8').decode(u8); } }
    return t.replace(/^﻿/, '');
  }
  function sheetState(name, aoa, fileName) {
    var g = L.guessDocType(aoa, name, fileName);
    var st = { name: name, aoa: aoa, type: g.type, guessed: g.type, hr: null, hn: null, ov: null };
    applySavedMapping(st);
    return st;
  }
  function sigOf(st) {
    if (!st.type) return '';
    var d = L.detectColumns(st.aoa, st.type);
    return st.type + '|' + d.headers.map(L.normHead).join('|');
  }
  function applySavedMapping(st) {
    var sig = sigOf(st);
    if (sig && db.mapping && db.mapping[sig]) st.ov = Object.assign({}, db.mapping[sig]);
  }
  function readOne(f) {
    if (/\.pdf$/i.test(f.name)) return readPdf(f);
    return f.arrayBuffer().then(function (buf) {
      var wb = /\.csv$/i.test(f.name) ? XLSX.read(decodeCsv(buf), { type: 'string' }) : XLSX.read(buf, { type: 'array' });
      var sheets = wb.SheetNames.map(function (n) {
        var ws = wb.Sheets[n];
        var aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
        return sheetState(n, L.fillMerges(aoa, ws['!merges']), f.name);
      });
      return { name: f.name, sheets: sheets };
    }).catch(function (e) { return { name: f.name, error: '읽지 못했습니다 (' + (e && e.message || e) + ')' }; });
  }
  // ── PDF 읽기(pdf.js, vendor/pdfjs) — data09-08 과 같은 방식 ──
  // 처음 필요할 때만 불러옵니다. file:// 에서는 브라우저가 Worker 를 막으므로 worker 스크립트를 일반 스크립트로 먼저 넣어
  // pdf.js 가 화면 스레드에서 돌게 합니다. 파일은 바이트로만 넘겨 밖으로 요청이 나가지 않습니다.
  var pdfLoading = null;
  function withPdf(cb, fail) {
    if (window.pdfjsLib && window.pdfjsWorker) { cb(window.pdfjsLib); return; }
    if (!pdfLoading) {
      pdfLoading = [];
      var addScript = function (src, next) {
        var s = document.createElement('script'); s.src = src;
        s.onload = next; s.onerror = function () { var q = pdfLoading; pdfLoading = null; (q || []).forEach(function (x) { x.fail(); }); };
        document.head.appendChild(s);
      };
      addScript('vendor/pdfjs/pdf.worker.min.js', function () {
        addScript('vendor/pdfjs/pdf.min.js', function () {
          window.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdfjs/pdf.worker.min.js';
          var q = pdfLoading; pdfLoading = null; q.forEach(function (x) { x.ok(window.pdfjsLib); });
        });
      });
    }
    pdfLoading.push({ ok: cb, fail: fail });
  }
  function readPdf(f) {
    return f.arrayBuffer().then(function (buf) {
      return new Promise(function (resolve, reject) {
        withPdf(function (lib) {
          lib.getDocument({ data: new Uint8Array(buf), isEvalSupported: false, disableFontFace: true }).promise.then(function (doc) {
            var all = [], p = Promise.resolve();
            for (var i = 1; i <= doc.numPages; i++) (function (n) {
              p = p.then(function () { return doc.getPage(n); }).then(function (pg) {
                var vp = pg.getViewport({ scale: 1 });
                return pg.getTextContent().then(function (tc) { all = all.concat(L.pdfTextItems(tc, n, vp.height)); });
              });
            })(i);
            return p.then(function () { return { pages: doc.numPages, items: all }; });
          }).then(resolve, reject);
        }, function () { reject(new Error('PDF 읽기 도구(vendor/pdfjs)를 불러오지 못했습니다. 폴더째 받았는지 확인해 주세요')); });
      });
    }).then(function (r) {
      if (!r.items.length) return { name: f.name, error: '이 PDF 에는 글자 층이 없습니다(스캔·그림 PDF). 1단계에서는 글자가 들어 있는 PDF 만 읽습니다 — 스캔본은 원본 Excel 을 받아 올려 주세요.' };
      var aoa = L.pdfLinesToAoa(L.itemsToLines(r.items));
      return { name: f.name, pdf: true, sheets: [sheetState('PDF ' + r.pages + '쪽', aoa, f.name)] };
    }).catch(function (e) { return { name: f.name, error: 'PDF 를 읽지 못했습니다 (' + (e && e.message || e) + ')' }; });
  }
  function readFiles(fileList) {
    var files = Array.prototype.slice.call(fileList || []);
    if (!files.length) return;
    if (!cur()) { toast('먼저 검토 건(협력사·품번)을 만들어 주세요.', true); return; }
    toast(files.length + '개 파일을 읽는 중…');
    Promise.all(files.map(readOne)).then(function (list) { ui.imp = (ui.imp || []).concat(list); render(); });
  }
  function templateButton() {
    return h('button', { type: 'button', class: 'btn', onclick: function () {
      writeXlsx('APQP_서류_양식.xlsx', {
        'Flow Chart': [['공정번호', '공정명', '공정 구분', '외주 여부', '비고'], [10, '(예) 소재 입고검사', '검사', '', ''], [20, '(예) 절단', '가공', '', ''], [30, '(예) 도장', '외주', 'Y', '']],
        '관리계획서': [['공정번호', '공정명', '제품 특성', '공정 특성', '특별특성', '규격/공차', '평가·측정방법', '시료 크기', '시료 주기', '관리방법', '대응계획'], [10, '(예) 소재 입고검사', '(예) 소재 두께', '', '', '3.2±0.1 mm', '마이크로미터', '3개', '매 로트', '수입검사 성적서', '반품']],
        'PFMEA': [['공정번호', '공정명', '요구사항', '잠재적 고장형태', '잠재적 고장영향', 'S', '분류', '잠재적 원인', 'O', '현 공정관리', 'D', 'RPN'], [10, '(예) 소재 입고검사', '(예) 소재 두께 규격', '두께 미달', '강도 저하', 6, '', '공급사 관리 미흡', 3, '수입검사', 4, 72]],
        '검사기준서': [['검사 구분', '공정번호', '검사항목', '특별특성', '규격', '검사방법', '시료수', '검사주기'], ['수입검사', 10, '(예) 소재 두께', '', '3.2±0.1 mm', '마이크로미터', '3개', '매 로트']],
        '외주품 리스트': [['공정번호', '외주 공정', '외주처', '품번', '비고'], [30, '(예) 도장', '(예) 외주처', '(예) SMP-0000', '']]
      });
    } }, '빈 양식 받기');
  }
  function renderUpload(main) {
    main.appendChild(pageHead('서류 업로드', '검토 건(협력사·품번)을 정하고 APQP 서류를 올립니다. 파일은 서버로 보내지 않고 이 브라우저에서만 읽습니다.'));
    var rv = cur();
    // 1. 검토 건
    var c1 = h('section', { class: 'card' }, h('h2', null, '1. 검토 건'));
    if (rv && !ui.newForm) {
      c1.appendChild(picker('#/upload'));
      c1.appendChild(h('p', { class: 'note' }, (rv.stage || '초도품') + ' · ' + rv.round + '차 제출' + (rv.partName ? ' · ' + rv.partName : '') + (prevOf(rv) ? ' · 앞 제출(' + prevOf(rv).round + '차)과 비교됩니다' : '')));
      c1.appendChild(h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.newForm = { supplier: rv.supplier, partNo: rv.partNo, partName: rv.partName, stage: rv.stage, resubmit: true }; render(); } }, '재제출 등록 (' + L.nextRound(db.reviews, rv.supplier, rv.partNo) + '차)'),
        h('button', { type: 'button', class: 'btn', onclick: function () { ui.newForm = {}; render(); } }, '새 검토 건')));
    } else c1.appendChild(newReviewForm());
    main.appendChild(c1);
    if (!rv || ui.newForm) { main.appendChild(h('section', { class: 'card' }, h('h2', null, '2. 서류 올리기'), h('p', { class: 'note' }, '검토 건을 만든 뒤 서류를 올립니다.'), h('div', { class: 'btn-row' }, templateButton(), sampleButton(false)))); return; }

    // 2. 서류 올리기
    var input = h('input', { type: 'file', accept: '.xlsx,.xls,.csv,.pdf', multiple: true, onchange: function (e) { readFiles(e.target.files); e.target.value = ''; } });
    var drop = h('label', { class: 'drop' }, icon('down', 'drop-ico'), h('strong', null, '서류 파일 고르기'), h('span', { class: 'note' }, '.xlsx · .xls · .csv · .pdf(글자가 든 PDF) — 여러 개를 함께, 또는 시트 여러 개가 든 통합 문서 하나도 됩니다(여기로 끌어 놓아도 됨)'), input);
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('over'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('over'); });
    drop.addEventListener('drop', function (e) { e.preventDefault(); drop.classList.remove('over'); readFiles(e.dataTransfer.files); });
    var tb = h('tbody');
    L.DOC_IDS.forEach(function (id) {
      var d = rv.docs[id];
      tb.appendChild(h('tr', null, h('td', { class: 'nw' }, h('strong', null, L.docLabel(id))),
        h('td', null, d ? h('span', { class: 'badge ok' }, '넣음') : h('span', { class: 'badge muted' }, id === 'outs' ? '없음(외주 공정이 있으면 필수)' : '없음')),
        h('td', null, d ? d.fileName + (d.sheet ? ' · ' + d.sheet : '') : ''), h('td', { class: 'num' }, d ? d.rows.length + '행' : ''),
        h('td', null, d ? h('button', { type: 'button', class: 'btn btn-sm2', onclick: function () { confirmBox('서류 빼기', L.docLabel(id) + ' 을 이 검토 건에서 뺍니다.', '빼기', function () { delete rv.docs[id]; save(); render(); }, true); } }, '빼기') : null)));
    });
    main.appendChild(h('section', { class: 'card' },
      h('h2', null, '2. 서류 올리기 — ' + L.reviewTitle(rv)),
      h('p', null, '서류 5가지: ', h('strong', null, 'Flow Chart · 관리계획서 · PFMEA · 검사기준서 · 외주품 리스트'), '. 시트 이름·파일 이름·열 이름으로 어느 서류인지 추정하고, 제목 줄·두 줄 머리·병합 칸이 있어도 머리행을 찾습니다. 틀리면 아래에서 고릅니다.'),
      drop,
      h('div', { class: 'btn-row', style: 'margin-top:12px' }, templateButton(), sampleButton(false)),
      h('p', { class: 'note', style: 'margin-top:10px' }, '시험용 가상 파일: 리포의 samples/예시_APQP_SMP-2041_1차제출.xlsx · 2차재제출.xlsx · 검사기준서 PDF'),
      h('h3', { style: 'margin-top:14px' }, '이 검토 건의 서류'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['서류', '상태', '파일 · 시트', '행', ''].map(function (t) { return h('th', null, t); }))), tb)),
      docCount(rv) ? h('div', { class: 'btn-row', style: 'margin-top:12px' }, h('a', { class: 'btn btn-primary', href: '#/review/' + encodeURIComponent(rv.id) }, '정합성 검토로')) : null));
    if (ui.imp) main.appendChild(importCard(rv));
  }
  function newReviewForm() {
    var f = ui.newForm || {};
    var sup = h('input', { value: f.supplier || '', placeholder: '예: 협력사 B', autocomplete: 'off' });
    var pn = h('input', { value: f.partNo || '', placeholder: '예: SMP-2041', autocomplete: 'off' });
    var pname = h('input', { value: f.partName || '', placeholder: '예: 브래킷 ASSY', autocomplete: 'off' });
    var stage = select(['초도품', '변경점', '재승인', '기타'], f.stage || '초도품', { 'aria-label': '구분' });
    var date = h('input', { type: 'date', value: today() });
    var roundIn = h('input', { type: 'number', min: '1', inputmode: 'numeric', value: String(L.nextRound(db.reviews, f.supplier || '', f.partNo || '')) });
    function upd() { roundIn.value = String(L.nextRound(db.reviews, sup.value, pn.value)); }
    sup.addEventListener('input', upd); pn.addEventListener('input', upd);
    return h('div', null,
      f.resubmit ? h('p', { class: 'alert info' }, '같은 협력사·품번으로 재제출 검토 건을 만듭니다. 서류를 다시 올리면 앞 제출과 비교됩니다.') : null,
      h('div', { class: 'form-grid three' },
        field('협력사 *', sup), field('품번 *', pn), field('품명', pname), field('구분', stage), field('제출 차수', roundIn, '같은 협력사·품번이 있으면 다음 차수를 넣어 둡니다'), field('제출일', date)),
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          if (!sup.value.trim() || !pn.value.trim()) { toast('협력사와 품번을 적어 주세요.', true); return; }
          var round = Math.max(1, parseInt(roundIn.value, 10) || 1);
          if (db.reviews.some(function (r) { return L.reviewKey(r) === L.reviewKey({ supplier: sup.value, partNo: pn.value }) && r.round === round; })) { toast('같은 협력사·품번의 ' + round + '차 검토 건이 이미 있습니다.', true); return; }
          var rv = L.cleanReview({ id: 'r' + (++db.seq), supplier: sup.value, partNo: pn.value, partName: pname.value, stage: stage.value, round: round, submittedAt: date.value, createdAt: nowIso(), docs: {} });
          if (db._sample) { delete db._sample; }
          db.reviews.push(rv); db.current = rv.id; db.aliases = L.updateAliases(db.aliases, db.reviews);
          ui.newForm = false; ui.imp = null; save();
          toast('검토 건을 만들었습니다. 이제 서류를 올리세요.');
          location.hash = '#/upload/' + encodeURIComponent(rv.id);
          render();
        } }, '검토 건 만들기'),
        db.reviews.length ? h('button', { type: 'button', class: 'btn', onclick: function () { ui.newForm = false; render(); } }, '그만두기') : null));
  }
  function parsedOf(st) {
    if (!st.type) return null;
    return L.parseDoc(st.aoa, st.type, { headerRow: st.hr, headerRows: st.hn, overrides: st.ov });
  }
  var PREVIEW_COLS = {
    flow: [['procNo', '공정번호'], ['procName', '공정명'], ['kind', '구분'], ['outsourced', '외주']],
    cp: [['procNo', '공정번호'], ['procName', '공정명'], ['char', '특성'], ['special', '특별특성'], ['sample', '시료'], ['freq', '주기'], ['control', '관리방법']],
    pfmea: [['procNo', '공정번호'], ['procName', '공정명'], ['char', '요구사항'], ['mode', '고장형태'], ['S', 'S'], ['O', 'O'], ['D', 'D'], ['RPN', 'RPN']],
    insp: [['procNo', '공정번호'], ['item', '검사항목'], ['special', '특별특성'], ['sample', '시료수'], ['freq', '주기']],
    outs: [['procNo', '공정번호'], ['procName', '외주 공정'], ['vendor', '외주처'], ['partNo', '품번']]
  };
  function previewTable(type, rows) {
    var cols = PREVIEW_COLS[type];
    return h('div', { class: 'table-wrap' }, h('table', { class: 'list' },
      h('thead', null, h('tr', null, h('th', null, '행'), cols.map(function (c) { return h('th', null, c[1]); }))),
      h('tbody', null, rows.map(function (r) {
        return h('tr', null, h('td', { class: 'num' }, r._r), cols.map(function (c) { var v = r[c[0]]; return h('td', { class: c[0] === 'char' || c[0] === 'mode' ? 'wide' : 'nw' }, v === true ? '외주' : v === false || v == null ? '' : String(v)); }));
      }))));
  }
  function importCard(rv) {
    var card = h('section', { class: 'card import' }, h('h2', null, '3. 서류 종류 · 열 연결 확인'));
    var ok = 0, typeCount = {};
    ui.imp.forEach(function (file) {
      var box = h('div', { class: 'imp-file' }, h('h3', null, file.name + (file.pdf ? ' (PDF 시범 읽기)' : '')));
      card.appendChild(box);
      if (file.error) { box.appendChild(h('p', { class: 'alert error' }, file.error)); return; }
      if (file.pdf) box.appendChild(h('p', { class: 'alert info' }, 'PDF 는 글자 위치로 표를 다시 맞춘 것이라 칸이 어긋날 수 있습니다. 아래 미리보기를 꼭 확인해 주세요. 스캔본·DRM 이 걸린 PDF 는 읽지 못합니다.'));
      file.sheets.forEach(function (st) {
        var sb = h('div', { class: 'imp-sheet' });
        box.appendChild(sb);
        var ts = select([['', '사용 안 함']].concat(L.DOC_TYPES.map(function (d) { return [d.id, d.label]; })), st.type, { 'aria-label': st.name + ' 서류 종류' });
        ts.addEventListener('change', function () { st.type = ts.value; st.hr = null; st.hn = null; st.ov = null; applySavedMapping(st); render(); });
        sb.appendChild(h('div', { class: 'imp-sheet-head' }, h('strong', null, '시트 「' + st.name + '」'),
          field('서류 종류', ts, st.guessed ? '추정: ' + L.docLabel(st.guessed) : '어느 서류인지 추정하지 못했습니다', 'inline')));
        if (!st.type) return;
        typeCount[st.type] = (typeCount[st.type] || 0) + 1;
        var p = parsedOf(st), det = p.det;
        // 머리행 고르기
        var hrOpts = [['-1', '(머리행 없음)']];
        for (var r = 0; r < Math.min(20, st.aoa.length); r++) hrOpts.push([String(r), (r + 1) + '행: ' + (st.aoa[r] || []).map(L.str).filter(Boolean).join(' | ').slice(0, 60)]);
        var hrSel = select(hrOpts, String(det.headerRow), { 'aria-label': '머리행' });
        hrSel.addEventListener('change', function () { st.hr = +hrSel.value; st.hn = 1; st.ov = null; render(); });
        var two = h('input', { type: 'checkbox', checked: det.headerRows === 2, onchange: function (e) { st.hr = det.headerRow; st.hn = e.target.checked ? 2 : 1; st.ov = null; render(); } });
        sb.appendChild(h('div', { class: 'form-grid' }, field('머리행', hrSel, '자동으로 찾은 줄입니다. 틀리면 고르세요.'), h('label', { class: 'check' }, two, '두 줄 머리(위 줄 「특성」 아래 「제품·공정」처럼)')));
        if (det.headerRow < 0) { sb.appendChild(h('p', { class: 'alert warn' }, '머리행(공정번호·공정명 같은 열 이름 줄)을 찾지 못했습니다. 머리행을 골라 주세요.')); return; }
        var grid = h('div', { class: 'map-grid' });
        L.docType(st.type).fields.forEach(function (f) {
          var opts = [['-1', '(연결 안 함)']].concat(det.headers.map(function (x, i) { return [String(i), colName(i) + ': ' + (x || '(빈 칸)')]; }));
          var s = select(opts, det.map[f.id] == null ? '-1' : String(det.map[f.id]), { 'aria-label': f.label });
          s.addEventListener('change', function () { st.ov = Object.assign({}, st.ov || {}); st.ov[f.id] = +s.value; render(); });
          var need = L.docType(st.type).req.some(function (g) { return g.indexOf(f.id) >= 0 && !g.some(function (id) { return det.map[id] != null; }); });
          grid.appendChild(field(f.label, s, null, need ? 'miss' : ''));
        });
        sb.appendChild(grid);
        if (det.missing.length) {
          sb.appendChild(h('p', { class: 'alert warn' }, '꼭 있어야 하는 열을 연결해 주세요: ' + det.missing.map(function (g) { return g.split('|').map(function (id) { return L.docType(st.type).fields.filter(function (f) { return f.id === id; })[0].label; }).join(' 또는 '); }).join(', ')));
          return;
        }
        ok++;
        sb.appendChild(h('p', { class: 'note' }, '읽은 행 ' + p.rows.length + '개 (머리 ' + (det.headerRow + 1) + '행' + (det.headerRows === 2 ? '·' + (det.headerRow + 2) + '행 두 줄' : '') + ')' + (p.skipped ? ' · 건너뛴 행 ' + p.skipped + '개(필수 칸이 빔·합계 등)' : '') + (rv.docs[st.type] ? ' · 넣으면 지금 있는 ' + L.docLabel(st.type) + ' 을 바꿉니다' : '')));
        sb.appendChild(previewTable(st.type, p.rows.slice(0, 5)));
      });
    });
    var dups = Object.keys(typeCount).filter(function (k) { return typeCount[k] > 1; });
    if (dups.length) card.appendChild(h('p', { class: 'alert warn' }, '같은 서류로 고른 시트가 둘 이상입니다: ' + dups.map(L.docLabel).join(', ') + ' — 마지막 시트가 들어갑니다. 하나만 남기고 「사용 안 함」으로 바꿔 주세요.'));
    card.appendChild(h('div', { class: 'btn-row', style: 'margin-top:14px' },
      h('button', { type: 'button', class: 'btn btn-primary', disabled: !ok, onclick: function () { doImport(rv); } }, '검토 건에 넣기 (서류 ' + ok + '개)'),
      h('button', { type: 'button', class: 'btn', onclick: function () { ui.imp = null; render(); } }, '그만두기')));
    return card;
  }
  function colName(i) { var s = ''; i++; while (i > 0) { var m = (i - 1) % 26; s = String.fromCharCode(65 + m) + s; i = Math.floor((i - 1) / 26); } return s; }
  function doImport(rv) {
    var n = 0, names = [];
    db.mapping = db.mapping || {};
    ui.imp.forEach(function (file) {
      if (file.error) return;
      file.sheets.forEach(function (st) {
        if (!st.type) return;
        var p = parsedOf(st);
        if (!p || p.det.headerRow < 0 || p.det.missing.length) return;
        if (st.ov) db.mapping[st.type + '|' + p.det.headers.map(L.normHead).join('|')] = st.ov;
        rv.docs[st.type] = { fileName: file.name, sheet: st.name, at: nowIso(), rows: p.rows };
        n++; names.push(L.docLabel(st.type));
      });
    });
    var c = L.cleanReview(rv);
    rv.docs = c.docs;
    db.aliases = L.updateAliases(db.aliases, db.reviews);
    save();
    ui.imp = null;
    toast('서류 ' + n + '개를 넣었습니다: ' + names.join(', '));
    location.hash = '#/review/' + encodeURIComponent(rv.id);
  }

  // ══ 정합성 검토 ═══════════════════════════════════════════
  function evChip(e, rv) {
    return h('button', { type: 'button', class: 'ev', onclick: function () { showRow(rv, e); } }, L.evText(e));
  }
  /* 근거 보기 — 그 서류의 그 행(정규화된 값) */
  function showRow(rv, e) {
    var d = rv.docs[e.doc];
    var r = d && e.row ? d.rows.filter(function (x) { return x._r === e.row; })[0] : null;
    var labels = {}; (L.docType(e.doc) || { fields: [] }).fields.forEach(function (f) { labels[f.id] = f.label.replace(/\(선택\)/, ''); });
    labels.char = labels.char || '특성'; labels.RPN = 'RPN';
    var body = r ? h('dl', { class: 'kv' }, Object.keys(r).filter(function (k) { return k !== '_r' && k !== 'rpnCalc' && r[k] !== '' && r[k] != null && !(k === 'char' && e.doc === 'cp'); }).map(function (k) {
      var v = r[k]; if (v === true) v = '예'; if (v === false) v = '아니오';
      return [h('dt', null, labels[k] || k), h('dd', null, String(v) + (k === 'RPN' && r.rpnCalc ? ' (S×O×D 로 계산)' : ''))];
    })) : h('p', null, d ? (e.text || '') + (d.fileName ? ' — ' + d.fileName : '') : '이 서류는 제출되지 않았습니다.');
    openDialog(L.docLabel(e.doc) + (e.row ? ' ' + e.row + '행' : ''), [d ? h('p', { class: 'note' }, d.fileName + (d.sheet ? ' · 시트 「' + d.sheet + '」' : '') + (e.row ? ' · 시트의 ' + e.row + '번째 줄' : '')) : null, body],
      [h('button', { type: 'button', class: 'btn btn-primary', onclick: closeDialog }, '닫기')]);
  }
  function renderReview(main) {
    var rv = cur();
    if (!rv) { noReview(main, '정합성 검토'); return; }
    main.appendChild(pageHead('정합성 검토', 'APQP 서류끼리 규칙으로 대조한 결과입니다. 근거를 누르면 그 서류의 행을 볼 수 있습니다.', picker('#/review')));
    if (!docCount(rv)) { noDocs(main); return; }
    var res = resultOf(rv), sm = res.summary, c = sm.counts;
    main.appendChild(h('section', { class: 'dkpis six' },
      kpi('적합률', 'gauge', sm.score, '%', '판정 ' + sm.total + '건 가운데 적합', 'ok'),
      kpi('적합', 'check', c.ok, '건', null, 'ok'),
      kpi('불일치', 'alert', c.mismatch, '건', null, 'warn'),
      kpi('누락', 'miss', c.missing, '건', null, 'bad'),
      kpi('확인 필요', 'ask', c.check, '건', '사람이 보고 판단', 'hold'),
      h('div', { class: 'dk dk-dec' }, h('span', { class: 'dk-ico' }, icon('stamp')), h('div', { class: 'dk-body' }, h('div', { class: 'dk-t' }, '승인 의견 제안'), h('div', { class: 'dk-v sm' }, decisionPill(sm.decision.code)), h('div', { class: 'dk-note' }, '중대 ' + sm.major + ' · 경미 ' + sm.minor)))));
    var cp = compareOf(rv, res);
    if (cp) main.appendChild(h('p', { class: 'alert info' }, '앞 제출(' + cp.prev.round + '차) 대비: 고쳐진 지적 ' + cp.cmp.fixed.length + '건 · 남은 지적 ' + cp.cmp.remain.length + '건 · 새로 생긴 지적 ' + cp.cmp.added.length + '건 (적합률 ' + cp.cmp.prevScore + '% → ' + cp.cmp.curScore + '%) — ', h('a', { href: '#/history/' + encodeURIComponent(rv.id) }, '검토 이력에서 자세히')));
    var missingDocs = L.DOC_IDS.filter(function (id) { return !rv.docs[id]; });
    if (missingDocs.length) main.appendChild(h('p', { class: 'alert warn' }, '없는 서류: ' + missingDocs.map(L.docLabel).join(', ') + ' — 그 서류가 필요한 규칙은 건너뜁니다.'));

    // 규칙별
    var rt = h('tbody');
    L.RULES.forEach(function (r) {
      var b = sm.byRule[r.id], n = b.ok + b.mismatch + b.missing + b.check;
      rt.appendChild(h('tr', { class: b.on ? null : 'off' }, h('td', null, h('strong', null, r.name), h('div', { class: 'note' }, r.desc)), h('td', { class: 'nw' }, b.on ? h('span', { class: 'sev sev-' + b.sev }, L.SEV[b.sev]) : h('span', { class: 'badge muted' }, '끔')),
        h('td', { class: 'num' }, b.on ? n : '-'), h('td', { class: 'num' }, b.ok), h('td', { class: 'num' + (b.mismatch ? ' hot' : '') }, b.mismatch), h('td', { class: 'num' + (b.missing ? ' hot' : '') }, b.missing), h('td', { class: 'num' + (b.check ? ' warm' : '') }, b.check)));
    });
    main.appendChild(h('section', { class: 'card' }, h('div', { class: 'card-head' }, h('h2', null, '규칙별 결과'), h('a', { class: 'dp-more', href: '#/settings' }, '규칙 켜기·끄기·기준값', icon('chev'))),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list rules-table' }, h('thead', null, h('tr', null, ['규칙', '심각도', '판정', '적합', '불일치', '누락', '확인 필요'].map(function (t) { return h('th', null, t); }))), rt))));

    // 판정 목록
    var st = ui.find;
    var vSel = select([['issues', '지적만 (불일치·누락·확인 필요)'], ['', '전체'], ['ok', '적합'], ['mismatch', '불일치'], ['missing', '누락'], ['check', '확인 필요']], st.verdict, { 'aria-label': '판정' });
    var rSel = select([['', '모든 규칙']].concat(L.RULES.map(function (r) { return [r.id, r.name]; })), st.rule, { 'aria-label': '규칙' });
    var dSel = select([['', '모든 서류']].concat(L.DOC_TYPES.map(function (d) { return [d.id, d.label]; })), st.doc, { 'aria-label': '근거 서류' });
    var q = h('input', { type: 'search', value: st.q, placeholder: '공정번호·항목·내용', 'aria-label': '찾기' });
    vSel.addEventListener('change', function () { st.verdict = vSel.value; render(); });
    rSel.addEventListener('change', function () { st.rule = rSel.value; render(); });
    dSel.addEventListener('change', function () { st.doc = dSel.value; render(); });
    q.addEventListener('change', function () { st.q = q.value.trim(); render(); });
    var order = { missing: 0, mismatch: 1, check: 2, ok: 3 };
    var list = res.findings.filter(function (f) {
      if (st.verdict === 'issues' ? f.verdict === 'ok' : st.verdict && f.verdict !== st.verdict) return false;
      if (st.rule && f.rule !== st.rule) return false;
      if (st.doc && !f.ev.some(function (e) { return e.doc === st.doc; })) return false;
      if (st.q && (f.procNo + ' ' + f.item + ' ' + f.msg).toLowerCase().indexOf(st.q.toLowerCase()) < 0) return false;
      return true;
    }).sort(function (a, b) { return order[a.verdict] - order[b.verdict] || (a.sev === b.sev ? 0 : a.sev === 'major' ? -1 : 1) || L.procCmp(a.procNo, b.procNo); });
    var ft = h('tbody');
    list.slice(0, 300).forEach(function (f) {
      ft.appendChild(h('tr', { class: 'v-' + f.verdict },
        h('td', { class: 'nw' }, verdictBadge(f.verdict), f.verdict === 'mismatch' || f.verdict === 'missing' ? h('span', { class: 'sev sev-' + f.sev }, L.SEV[f.sev]) : null),
        h('td', { class: 'nw' }, L.ruleMeta(f.rule).name), h('td', { class: 'num' }, f.procNo || '-'),
        h('td', { class: 'wide' }, f.msg), h('td', null, h('span', { class: 'evs' }, f.ev.map(function (e) { return evChip(e, rv); })))));
    });
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '판정 목록'),
      h('div', { class: 'filters' }, field('판정', vSel), field('규칙', rSel), field('근거 서류', dSel), field('찾기', q)),
      h('p', { class: 'note', style: 'margin-top:10px' }, list.length + '건' + (list.length > 300 ? ' (앞 300건만 표시)' : '') + ' · 중대 = 승인 의견을 반려 쪽으로 미는 지적, 경미 = 보완 조건. 규칙마다 설정에서 바꿉니다.'),
      list.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list findings' }, h('thead', null, h('tr', null, ['판정', '규칙', '공정', '내용', '근거'].map(function (t) { return h('th', null, t); }))), ft))
        : h('p', { class: 'alert info' }, st.verdict === 'issues' ? '지적이 없습니다.' : '조건에 맞는 판정이 없습니다.'),
      h('div', { class: 'btn-row', style: 'margin-top:12px' }, h('a', { class: 'btn', href: '#/matrix/' + encodeURIComponent(rv.id) }, '공정 대조표'), h('a', { class: 'btn btn-primary', href: '#/report/' + encodeURIComponent(rv.id) }, '승인 검토 보고서로'))));
  }

  // ══ 공정 대조표 ═══════════════════════════════════════════
  var COV_TEXT = { ok: '있음', missing: '누락', extra: 'Flow 에 없음', na: '-', nodoc: '서류 없음' };
  function renderMatrix(main) {
    var rv = cur();
    if (!rv) { noReview(main, '공정 대조표'); return; }
    main.appendChild(pageHead('공정 대조표', '공정 × 서류 — 공정번호가 어느 서류에 있는지 한 표로 봅니다.', picker('#/matrix')));
    if (!docCount(rv)) { noDocs(main); return; }
    var res = resultOf(rv), cov = res.coverage, okSim = db.settings.rules.proc_name.okSim / 100;
    var tb = h('tbody');
    cov.rows.forEach(function (r) {
      tb.appendChild(h('tr', { class: r.excluded ? 'excl' : null },
        h('th', { scope: 'row', class: 'num' }, r.procNo),
        h('td', { class: 'cov-name' }, r.name, r.outsourced ? h('span', { class: 'flag part' }, '외주') : null, r.excluded ? h('span', { class: 'flag' }, '대조 제외') : null),
        L.DOC_IDS.map(function (id) {
          var c = r.cells[id];
          var diff = (id === 'cp' || id === 'pfmea' || id === 'outs') && r.names[id] && r.name && L.similarity(r.name, r.names[id]) < okSim;
          return h('td', { class: 'cov cov-' + c.status + (diff ? ' cov-diff' : ''), title: c.rows.length ? c.rows.map(function (x) { return x + '행'; }).join(', ') : null },
            COV_TEXT[c.status] + (c.count > 1 ? ' ' + c.count : ''), diff ? h('small', null, r.names[id]) : null);
        })));
    });
    main.appendChild(h('section', { class: 'card' },
      h('ul', { class: 'legend cov-legend' }, ['ok', 'missing', 'extra', 'na'].map(function (k) { return h('li', null, h('i', { class: 'cov-' + k }), COV_TEXT[k] === '-' ? '대상 아님' : COV_TEXT[k]); }), h('li', null, h('i', { class: 'cov-diff-i' }), '공정명이 다름(작은 글씨)')),
      h('p', { class: 'note' }, '「누락」은 Flow Chart 의 공정이 관리계획서·PFMEA 에 없거나, 외주 공정이 외주품 리스트에 없는 칸입니다. 숫자는 그 공정의 행 수입니다.' + (cov.inspHasNo ? '' : ' 검사기준서에 공정번호 열이 없어 검사기준서 칸은 대상 아님으로 둡니다(항목 이름으로 대조).') + ' 이동·보관처럼 대조에서 뺀 공정은 흐리게 보입니다(설정에서 바꿈).'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list cov-table' },
        h('thead', null, h('tr', null, h('th', null, '공정번호'), h('th', null, '공정명(Flow Chart)'), L.DOC_TYPES.map(function (d) { return h('th', null, d.short); }))), tb))));
  }

  // ══ 승인 검토 보고서 ══════════════════════════════════════
  function requestsFromFindings(res) {
    return res.findings.filter(function (f) { return f.verdict === 'missing' || f.verdict === 'mismatch'; }).map(function (f) {
      return '[' + L.ruleMeta(f.rule).name + '] ' + f.msg + ' — 보완 바랍니다.';
    }).join('\n');
  }
  function renderReport(main) {
    var rv = cur();
    if (!rv) { noReview(main, '승인 검토 보고서'); return; }
    main.appendChild(pageHead('승인 검토 보고서', '도구의 승인 의견은 규칙 판정에서 나온 제안입니다. 검토자가 결정하고 의견을 적어 보고서로 냅니다.', picker('#/report')));
    if (!docCount(rv)) { noDocs(main); return; }
    var res = resultOf(rv), sm = res.summary, op = rv.opinion, cp = compareOf(rv, res);

    // 승인 의견
    var decSel = select([['', '미정'], ['approve', '승인'], ['conditional', '조건부 승인'], ['reject', '반려']], op.decision, { 'aria-label': '검토자 결정' });
    var who = h('input', { value: op.reviewer || '', placeholder: '예: 부품품질팀 홍길동', autocomplete: 'off' });
    var cm = h('textarea', { rows: '5' }); cm.value = op.comment || '';
    var rq = h('textarea', { rows: '5', placeholder: '한 줄에 하나씩' }); rq.value = op.requests || '';
    function store() { op.decision = decSel.value; op.reviewer = who.value.trim(); op.comment = cm.value; op.requests = rq.value; op.at = nowIso(); save(); }
    main.appendChild(h('section', { class: 'card dec-card' },
      h('h2', null, '승인 의견'),
      h('div', { class: 'dec-row' }, h('span', { class: 'note' }, '도구의 제안'), decisionPill(sm.decision.code), h('span', { class: 'note' }, '적합률 ' + sm.score + '% · 중대 ' + sm.major + ' · 경미 ' + sm.minor + ' · 확인 필요 ' + sm.counts.check)),
      h('ul', { class: 'reasons' }, sm.decision.reasons.map(function (x) { return h('li', null, x); })),
      h('p', { class: 'note' }, '제안 기준(설정에서 바꿈): 중대 지적 ' + db.settings.decision.rejectMajor + '건 이상 또는 경미 지적 ' + db.settings.decision.rejectMinor + '건 이상이면 반려, 지적·확인 필요가 있으면 조건부 승인, 없으면 승인. 최종 결정은 검토자가 합니다.'),
      h('div', { class: 'form-grid' }, field('검토자 결정', decSel), field('검토자', who)),
      field('검토 의견', cm),
      field('보완 요청 (협력사에 보낼 것)', rq),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { store(); toast('승인 의견을 저장했습니다.'); render(); } }, '저장'),
        h('button', { type: 'button', class: 'btn', onclick: function () { rq.value = requestsFromFindings(res); store(); toast('누락·불일치 지적으로 보완 요청을 채웠습니다. 문장을 다듬어 주세요.'); } }, '지적 사항으로 보완 요청 채우기'))));

    main.appendChild(aiCard(rv, res, cp));

    // 내보내기
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '보고서 내보내기'),
      h('p', { class: 'note' }, 'Excel 은 요약 · 판정 목록 · 공정 대조표' + (cp ? ' · 재제출 비교' : '') + ' · 규칙 설정 · 서류 데이터 시트로 나옵니다. 인쇄·Word 는 아래 미리보기 그대로입니다. 지금 쓰시는 보고서 양식을 받으면 칸 배치를 맞춥니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { store(); downloadReport(rv); } }, icon('down'), 'Excel 보고서(.xlsx)'),
        h('button', { type: 'button', class: 'btn', onclick: function () { store(); render(); setTimeout(function () { window.print(); }, 50); } }, '인쇄 / PDF'),
        h('button', { type: 'button', class: 'btn', onclick: function () { store(); render(); downloadWord(rv); } }, 'Word(.doc) 저장'))));
    main.appendChild(reportPreview(rv, res, cp));
  }
  function reportFileBase(rv) { return 'APQP_승인검토_' + safeName(rv.partNo) + '_' + rv.round + '차' + tag() + '_' + today(); }
  function downloadReport(rv) {
    var res = resultOf(rv), cp = compareOf(rv, res);
    var sheets = L.reportSheets({ review: rv, result: res, prev: cp && cp.prev, compare: cp && cp.cmp, settings: db.settings, today: today() });
    writeXlsx(reportFileBase(rv) + '.xlsx', sheets, { '요약': [22, 34, 30, 10, 10, 10, 10], '판정 목록': [5, 9, 7, 26, 9, 18, 70, 28], '공정 대조표': [9, 18, 14, 6, 12, 12, 12, 12, 12, 30], '재제출 비교': [10, 24, 9, 18, 60, 60], '규칙 설정': [28, 6, 7, 60, 50], '서류 데이터': [12, 5, 8, 14, 22, 8, 10, 8, 22, 4, 4, 4, 6, 10] });
    toast(reportFileBase(rv) + '.xlsx 을 내려받았습니다.');
  }
  function downloadWord(rv) {
    var pv = document.querySelector('.report-preview');
    if (!pv) return;
    var css = 'body{font-family:"맑은 고딕",sans-serif;font-size:10.5pt;word-break:keep-all}h2{font-size:15pt;color:#13274b}h3{font-size:12pt;color:#13274b;margin-top:14pt}table{border-collapse:collapse;width:100%;margin:6pt 0}th,td{border:1px solid #999;padding:3pt 5pt;vertical-align:top;font-size:9.5pt}th{background:#eef1f5}.note{color:#555}';
    var html = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word"><head><meta charset="utf-8"><title>' + reportFileBase(rv) + '</title><style>' + css + '</style></head><body>' + pv.innerHTML + '</body></html>';
    download(reportFileBase(rv) + '.doc', new Blob(['﻿' + html], { type: 'application/msword' }));
    toast(reportFileBase(rv) + '.doc 을 내려받았습니다(Word 로 열립니다).');
  }
  function kvTable(rows) { return h('div', { class: 'table-wrap' }, h('table', { class: 'list rp-kv' }, h('tbody', null, rows.map(function (r) { return h('tr', null, h('th', null, r[0]), h('td', null, r[1])); })))); }
  function reportPreview(rv, res, cp) {
    var sm = res.summary, op = rv.opinion;
    var issues = res.findings.filter(function (f) { return f.verdict !== 'ok'; });
    var order = { missing: 0, mismatch: 1, check: 2 };
    issues.sort(function (a, b) { return order[a.verdict] - order[b.verdict] || L.procCmp(a.procNo, b.procNo); });
    var pv = h('section', { class: 'card report-preview' },
      h('h2', null, '초도품 APQP 승인 검토 보고서'),
      h('p', { class: 'note' }, '작성일 ' + today() + ' · APQP Smart Review(1단계)'),
      kvTable([['협력사', rv.supplier], ['품번 · 품명', rv.partNo + (rv.partName ? ' · ' + rv.partName : '')], ['구분 · 제출 차수', (rv.stage || '초도품') + ' · ' + rv.round + '차'], ['제출일', rv.submittedAt || '']]),
      h('h3', null, '1. 제출 서류'),
      h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['서류', '파일 · 시트', '행'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, L.DOC_IDS.map(function (id) { var d = rv.docs[id]; return h('tr', null, h('td', null, L.docLabel(id)), h('td', null, d ? d.fileName + (d.sheet ? ' · ' + d.sheet : '') : '(제출 안 됨)'), h('td', { class: 'num' }, d ? d.rows.length : 0)); })))),
      h('h3', null, '2. 정합성 판정 요약'),
      kvTable([['판정', '전체 ' + sm.total + '건 — 적합 ' + sm.counts.ok + ' · 불일치 ' + sm.counts.mismatch + ' · 누락 ' + sm.counts.missing + ' · 확인 필요 ' + sm.counts.check], ['적합률', sm.score + '%'], ['지적', '중대 ' + sm.major + '건 · 경미 ' + sm.minor + '건'],
        ['승인 의견 제안(도구)', sm.decision.label + ' — ' + sm.decision.reasons.join(' / ')]].concat(cp ? [['앞 제출 비교', cp.prev.round + '차 → ' + rv.round + '차: 고쳐짐 ' + cp.cmp.fixed.length + ' · 남음 ' + cp.cmp.remain.length + ' · 새로 생김 ' + cp.cmp.added.length + ' (적합률 ' + cp.cmp.prevScore + '% → ' + cp.cmp.curScore + '%)']] : [])),
      h('h3', null, '3. 검토자 결정'),
      kvTable([['결정', op.decision ? L.DECISION[op.decision] : '(미정)'], ['검토자', op.reviewer || ''], ['검토 의견', h('div', { class: 'pre' }, op.comment || '')], ['보완 요청', h('div', { class: 'pre' }, op.requests || '')]]),
      h('h3', null, '4. 지적 사항 (' + issues.length + '건)'),
      issues.length ? h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['No', '판정', '규칙', '공정', '내용', '근거'].map(function (t) { return h('th', null, t); }))),
        h('tbody', null, issues.map(function (f, i) {
          return h('tr', null, h('td', { class: 'num' }, i + 1), h('td', { class: 'nw' }, L.VERDICT[f.verdict] + (f.verdict === 'check' ? '' : ' · ' + L.SEV[f.sev])), h('td', null, L.ruleMeta(f.rule).name), h('td', { class: 'num' }, f.procNo), h('td', { class: 'wide' }, f.msg), h('td', null, f.ev.map(L.evText).join(', ')));
        })))) : h('p', null, '지적 사항이 없습니다.'),
      h('p', { class: 'note' }, '※ 승인 의견은 규칙 판정에서 나온 제안이며, 최종 승인은 검토자가 결정합니다. 서류 원본은 이 보고서에 담지 않았습니다.'));
    return pv;
  }
  // AI 로 보고서 문장 — 사실만 담은 프롬프트(협력사·품번은 별칭) → 복사 → 답 붙여 넣기
  function aiCard(rv, res, cp) {
    db.aliases = L.updateAliases(db.aliases, db.reviews);
    var info = L.buildAiPrompt({ review: rv, result: res, compare: cp && cp.cmp, aliases: db.mask ? db.aliases : null });
    var mask = h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); render(); } });
    var box = h('textarea', { class: 'prompt-box', readonly: true, 'aria-label': 'AI 프롬프트', rows: '14' }); box.value = info.prompt;
    var ans = h('textarea', { class: 'prompt-box answer', 'aria-label': 'AI 답 붙여 넣기', rows: '10', placeholder: 'AI 답을 여기에 붙여 넣으세요 — 「검토 의견:」 「보완 요청:」 「확인 필요:」 줄을 읽습니다.' });
    ans.value = ui.rep.aiText;
    ans.addEventListener('input', function () { ui.rep.aiText = ans.value; });
    function readAnswer() {
      var r = L.parseAiReport(ans.value, db.aliases);
      ui.rep.parsed = r; ui.rep.aiText = ans.value; render();
      if (!r.opinion && !r.requests.length) toast('「검토 의견:」 줄을 찾지 못했습니다. 답 형식을 확인해 주세요.', true);
    }
    var send = window.AIPanel.sendButton('설정한 AI 서버로 보내기', function () { return info.prompt; }, function (text) { ans.value = text; ui.rep.aiText = text; readAnswer(); }, { toast: toast, settingsHref: '#/settings' });
    var p = ui.rep.parsed, preview = null;
    if (p) {
      preview = h('div', { class: 'ai-preview' },
        h('h3', null, '읽은 결과'),
        p.opinion ? h('p', null, h('strong', null, '검토 의견: '), p.opinion) : null,
        p.requests.length ? h('div', null, h('strong', null, '보완 요청 ' + p.requests.length + '건'), h('ul', null, p.requests.map(function (x) { return h('li', null, x); }))) : null,
        p.checks.length ? h('div', null, h('strong', null, '확인 필요'), h('ul', null, p.checks.map(function (x) { return h('li', null, x); }))) : null,
        h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var op = rv.opinion;
          op.comment = p.opinion + (p.checks.length ? '\n\n확인 필요:\n' + p.checks.map(function (x) { return '- ' + x; }).join('\n') : '');
          if (p.requests.length) op.requests = p.requests.join('\n');
          op.at = nowIso(); rv.aiText = ui.rep.aiText; ui.rep.parsed = null; ui.rep.aiText = '';
          save(); render(); toast('AI 문장을 검토 의견·보완 요청에 넣었습니다. 사실과 맞는지 검토해 주세요.');
        } }, '검토 의견·보완 요청에 반영')));
    }
    return h('section', { class: 'card ai-card', id: 'aiCard' },
      h('div', { class: 'card-head' }, h('h2', null, icon('chip', 'ai-chip'), 'AI 로 보고서 문장 다듬기')),
      h('ol', { class: 'steps' },
        h('li', null, '아래 프롬프트를 복사해 회사가 허용한 AI(사내 AI 또는 ChatGPT 등)에 붙여 넣습니다. 프롬프트에는 도구가 대조한 사실(판정·근거 행)만 들어 있고 서류 원본은 들어 있지 않습니다.'),
        h('li', null, 'AI 답 전체를 복사해 「AI 답」 칸에 붙여 넣고 「답 읽기」를 누릅니다.'),
        h('li', null, '읽은 결과를 확인하고 「검토 의견·보완 요청에 반영」 → 위 승인 의견 칸에서 다듬습니다.')),
      h('label', { class: 'check' }, mask, '협력사·외주처 이름과 품번을 별칭(업체1, 품목1)으로 가리기 — 권장. 대응표는 이 브라우저에만 있고, AI 답의 별칭은 원래 이름으로 되돌립니다.'),
      info.leaks.length ? h('p', { class: 'alert warn' }, '가리지 못한 이름이 프롬프트에 남아 있습니다: ' + info.leaks.join(', ') + ' — 보내기 전에 직접 지워 주세요.') : null,
      h('details', { class: 'prompt-details', open: ui.rep.showPrompt, ontoggle: function (e) { ui.rep.showPrompt = e.target.open; } }, h('summary', null, '프롬프트 보기 (' + info.prompt.length + '자)'), box),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn btn-primary', onclick: function () { copyText(info.prompt); } }, '프롬프트 복사'), send),
      field('AI 답', ans),
      h('div', { class: 'btn-row' }, h('button', { type: 'button', class: 'btn', onclick: readAnswer }, '답 읽기'), ui.rep.parsed ? h('button', { type: 'button', class: 'btn', onclick: function () { ui.rep.parsed = null; ui.rep.aiText = ''; render(); } }, '답 지우기') : null),
      preview);
  }

  // ══ 검토 이력 ═════════════════════════════════════════════
  function renderHistory(main) {
    main.appendChild(pageHead('검토 이력', '품번별 제출 차수와 결과입니다. 재제출이면 바로 앞 제출과 비교해 무엇이 고쳐졌는지 보여 줍니다.'));
    if (!db.reviews.length) {
      main.appendChild(h('section', { class: 'card' }, h('p', null, '아직 검토 건이 없습니다.'), h('div', { class: 'btn-row' }, h('a', { class: 'btn btn-primary', href: '#/upload' }, '서류 업로드로'), sampleButton(false))));
      return;
    }
    var groups = {};
    sortedReviews().forEach(function (r) { (groups[L.reviewKey(r)] = groups[L.reviewKey(r)] || []).push(r); });
    var focus = cur();
    Object.keys(groups).forEach(function (k) {
      var list = groups[k], first = list[0];
      var tb = h('tbody');
      var results = {};
      list.forEach(function (rv) {
        var res = resultOf(rv), sm = res.summary; results[rv.id] = res;
        tb.appendChild(h('tr', { class: focus && focus.id === rv.id ? 'sel' : null },
          h('td', { class: 'num' }, rv.round + '차'), h('td', { class: 'nw' }, rv.submittedAt || ''), h('td', { class: 'num' }, docCount(rv) + '/5'),
          h('td', { class: 'num' }, sm.score + '%'), h('td', { class: 'nw' }, '중대 ' + sm.major + ' · 경미 ' + sm.minor + ' · 확인 ' + sm.counts.check),
          h('td', null, decisionPill(sm.decision.code)), h('td', null, rv.opinion.decision ? decisionPill(rv.opinion.decision) : h('span', { class: 'note' }, '미정')),
          h('td', { class: 'nw' }, h('div', { class: 'btn-row tight' },
            h('a', { class: 'btn btn-sm2', href: '#/review/' + encodeURIComponent(rv.id) }, '열기'),
            h('a', { class: 'btn btn-sm2', href: '#/history/' + encodeURIComponent(rv.id) }, '비교'),
            h('button', { type: 'button', class: 'btn btn-sm2 btn-danger', onclick: function () { confirmBox('검토 건 지우기', L.reviewTitle(rv) + ' 을 지웁니다. 되돌릴 수 없습니다.', '지우기', function () { db.reviews = db.reviews.filter(function (x) { return x.id !== rv.id; }); if (db.current === rv.id) db.current = (db.reviews[db.reviews.length - 1] || {}).id || ''; save(); render(); }, true); } }, '지우기')))));
      });
      // 비교: 이 묶음 안에서 고른 차수(없으면 가장 최신) vs 그 앞
      var target = focus && L.reviewKey(focus) === k && prevOf(focus) ? focus : list[list.length - 1];
      var prev = prevOf(target), cmpBox = null;
      if (prev) {
        var cmp = L.compareResults(results[prev.id] || resultOf(prev), results[target.id] || resultOf(target));
        function items(arr, fn) { return arr.length ? h('ul', { class: 'cmp-list' }, arr.map(function (x) { return h('li', null, fn(x)); })) : h('p', { class: 'note' }, '없음'); }
        cmpBox = h('div', { class: 'cmp' },
          h('h3', null, prev.round + '차 → ' + target.round + '차 비교 · 적합률 ' + cmp.prevScore + '% → ' + cmp.curScore + '%'),
          h('div', { class: 'cmp-grid' },
            h('div', { class: 'cmp-col cmp-fixed' }, h('h4', null, '고쳐진 지적 ' + cmp.fixed.length + '건'), items(cmp.fixed, function (f) { return [verdictBadge(f.verdict), ' ', L.ruleMeta(f.rule).name + ' — ' + f.msg]; })),
            h('div', { class: 'cmp-col cmp-remain' }, h('h4', null, '남은 지적 ' + cmp.remain.length + '건'), items(cmp.remain, function (x) { return [verdictBadge(x.cur.verdict), ' ', L.ruleMeta(x.cur.rule).name + ' — ' + x.cur.msg]; })),
            h('div', { class: 'cmp-col cmp-added' }, h('h4', null, '새로 생긴 지적 ' + cmp.added.length + '건'), items(cmp.added, function (f) { return [verdictBadge(f.verdict), ' ', L.ruleMeta(f.rule).name + ' — ' + f.msg]; }))));
      } else if (list.length > 1) cmpBox = h('p', { class: 'note' }, target.round + '차는 첫 제출이라 비교할 앞 제출이 없습니다. 다른 차수의 「비교」를 누르세요.');
      main.appendChild(h('section', { class: 'card' },
        h('div', { class: 'card-head' }, h('h2', null, first.supplier + ' · ' + first.partNo + (first.partName ? ' · ' + first.partName : '')),
          h('a', { class: 'btn btn-sm2', href: '#/upload/' + encodeURIComponent(list[list.length - 1].id), onclick: function () { setCur(list[list.length - 1].id); ui.newForm = { supplier: first.supplier, partNo: first.partNo, partName: first.partName, stage: first.stage, resubmit: true }; } }, '재제출 등록')),
        h('div', { class: 'table-wrap' }, h('table', { class: 'list' }, h('thead', null, h('tr', null, ['차수', '제출일', '서류', '적합률', '지적', '의견 제안', '검토자 결정', ''].map(function (t) { return h('th', null, t); }))), tb)),
        cmpBox));
    });
  }

  // ══ 설정 ═══════════════════════════════════════════════════
  function renderSettings(main) {
    main.appendChild(pageHead('설정', '정합성 규칙(켜기·끄기·심각도·기준값), 이름 대조 기준, 승인 의견 제안 기준, AI 연결, 데이터 백업'));
    main.appendChild(rulesCard());
    main.appendChild(h('section', { class: 'card' }, h('h2', null, 'AI 프롬프트 가리기'),
      h('label', { class: 'check' }, h('input', { type: 'checkbox', checked: db.mask, onchange: function (e) { db.mask = e.target.checked; save(); toast(db.mask ? '협력사·품번을 별칭으로 가립니다.' : '프롬프트에 실제 이름이 들어갑니다.', !db.mask); } }), '협력사·외주처 이름과 품번을 별칭으로 가리기(기본·권장)'),
      h('p', { class: 'note' }, '대응표: 협력사·외주처 ' + Object.keys(db.aliases.suppliers).length + '곳 · 품번 ' + Object.keys(db.aliases.parts).length + '개. 같은 이름은 늘 같은 별칭이 됩니다. 대응표는 이 브라우저와 JSON 백업에만 있습니다.')));
    main.appendChild(window.AIPanel.settingsCard({ toast: toast }));
    main.appendChild(h('section', { class: 'card' }, h('h2', null, '문서보안(DRM)·공유'),
      h('ul', null,
        h('li', null, '이 도구는 파일을 브라우저 안에서만 읽습니다. DRM 이 걸린 Excel·PDF 는 브라우저가 내용을 읽을 수 없어, 사내 규정에 맞게 해제한 사본이나 협력사 제출 원본을 써야 합니다(기획서 10장 ④ 확인 중).'),
        h('li', null, '데이터는 각 브라우저 안에만 저장됩니다. 팀이 같은 검토 건을 보려면 지금은 JSON 백업 파일을 주고받아야 합니다.'),
        h('li', null, '팀 공유 DB·로그인과, 같은 수강생의 과제 01(협력사 공정불량 Master Plan)과 하나로 묶는 「협력사 품질 포털」은 2단계에서 다룹니다.'))));
    main.appendChild(dataCard());
  }
  function rulesCard() {
    var st = db.settings, ins = [];
    function paramInput(r, p, val) {
      var el;
      if (p.type === 'num') el = h('input', { type: 'number', min: String(p.min), max: String(p.max), value: String(val), inputmode: 'numeric' });
      else if (p.type === 'bool') { el = h('input', { type: 'checkbox', checked: !!val }); ins.push({ r: r.id, p: p, el: el }); return h('label', { class: 'check' }, el, p.label); }
      else if (p.type === 'select') el = select(p.options, val);
      else el = h('input', { value: val, autocomplete: 'off' });
      ins.push({ r: r.id, p: p, el: el });
      return field(p.label, el);
    }
    var list = h('div', { class: 'dict' });
    L.RULES.forEach(function (r) {
      var x = st.rules[r.id];
      var on = h('input', { type: 'checkbox', checked: x.on }), sev = select([['major', '중대 — 반려 쪽'], ['minor', '경미 — 보완 조건']], x.sev, { 'aria-label': r.name + ' 심각도' });
      ins.push({ r: r.id, p: { id: 'on', type: 'bool' }, el: on }); ins.push({ r: r.id, p: { id: 'sev', type: 'sev' }, el: sev });
      list.appendChild(h('details', { class: 'dict-item' + (x.on ? '' : ' off') },
        h('summary', null, h('strong', null, r.name), x.on ? h('span', { class: 'sev sev-' + x.sev }, L.SEV[x.sev]) : h('span', { class: 'badge muted' }, '끔')),
        h('p', { class: 'note' }, r.desc),
        h('div', { class: 'form-grid' }, h('label', { class: 'check' }, on, '이 규칙 사용'), field('심각도(누락·불일치일 때)', sev)),
        r.params.length ? h('div', { class: 'form-grid' }, r.params.map(function (p) { return paramInput(r, p, x[p.id]); })) : null));
    });
    var mIns = {}, dIns = {};
    function numField(p, val, store) { store[p.id] = h('input', { type: 'number', min: String(p.min), max: String(p.max), value: String(val), inputmode: 'numeric' }); return field(p.label, store[p.id]); }
    return h('section', { class: 'card' }, h('h2', null, '정합성 규칙'),
      h('p', { class: 'note' }, '시작 규칙 8개입니다. 「정합하다」고 보는 실제 대조 규칙을 받으면(기획서 10장 ②) 여기 기준값을 바꾸거나 규칙을 더합니다. 기준값은 1단계 가정입니다.'),
      list,
      h('h3', { style: 'margin-top:16px' }, '항목(특성) 이름 대조'),
      h('p', { class: 'note' }, 'PFMEA 요구사항·관리계획서 특성·검사기준서 검사항목처럼 서류마다 조금씩 다르게 적은 이름을 「닮음」으로 짝짓습니다. 띄어쓰기·부호·괄호 속은 무시합니다.'),
      h('div', { class: 'form-grid' }, L.MATCH_PARAMS.map(function (p) { return numField(p, st.match[p.id], mIns); })),
      h('h3', { style: 'margin-top:16px' }, '승인 의견 제안 기준'),
      h('div', { class: 'form-grid' }, L.DECISION_PARAMS.map(function (p) { return numField(p, st.decision[p.id], dIns); })),
      h('div', { class: 'btn-row', style: 'margin-top:12px' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          var o = JSON.parse(JSON.stringify(db.settings));
          ins.forEach(function (x) {
            var t = o.rules[x.r];
            if (x.p.type === 'bool') t[x.p.id] = x.el.checked;
            else if (x.p.type === 'num') t[x.p.id] = x.el.value;
            else t[x.p.id] = x.el.value;
          });
          Object.keys(mIns).forEach(function (k) { o.match[k] = mIns[k].value; });
          Object.keys(dIns).forEach(function (k) { o.decision[k] = dIns[k].value; });
          db.settings = L.cleanSettings(o); save(); toast('규칙을 저장했습니다. 판정을 다시 계산했습니다.'); render();
        } }, '규칙 저장'),
        h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          confirmBox('시작값으로 되돌리기', '규칙 켜기·끄기·심각도·기준값을 처음 값으로 되돌립니다.', '되돌리기', function () { db.settings = L.defaultSettings(); save(); render(); }, true);
        } }, '시작값으로')));
  }
  function dataCard() {
    var input = h('input', { type: 'file', accept: '.json,application/json', onchange: function (e) {
      var f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      f.text().then(function (t) {
        var o = JSON.parse(t), d = L.restoreDb(o && o.db ? o.db : o);
        confirmBox('백업 되살리기', f.name + ' — 검토 건 ' + d.reviews.length + '개.\n지금 데이터를 이 백업으로 바꿉니다.', '되살리기', function () { db = d; save(); toast('되살렸습니다.'); render(); }, true);
      }).catch(function () { toast('JSON 백업 파일을 읽지 못했습니다.', true); });
    } });
    return h('section', { class: 'card' }, h('h2', null, '데이터 백업'),
      h('p', { class: 'note' }, '검토 건·서류(정규화한 행)·승인 의견·규칙·별칭 대응표를 JSON 한 파일로 받습니다. AI 키는 넣지 않습니다.'),
      h('div', { class: 'btn-row' },
        h('button', { type: 'button', class: 'btn btn-primary', onclick: function () {
          download('data10-02_백업' + tag() + '_' + today() + '.json', new Blob([JSON.stringify({ app: 'data10-02', at: nowIso(), db: db }, null, 1)], { type: 'application/json' }));
        } }, 'JSON 백업 받기'),
        h('label', { class: 'btn file-btn' }, '백업 되살리기', input),
        db.reviews.length ? h('button', { type: 'button', class: 'btn btn-danger', onclick: function () {
          confirmBox('전부 지우기', '검토 건 ' + db.reviews.length + '개를 모두 지웁니다. 규칙·AI 설정은 남습니다.', '지우기', function () {
            var keep = { settings: db.settings, mask: db.mask, mapping: db.mapping };
            db = L.emptyDb(); Object.assign(db, keep); save(); toast('지웠습니다.'); render();
          }, true);
        } }, '데이터 전부 지우기') : null));
  }

  render();
})();
