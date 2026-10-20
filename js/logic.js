/* APQP Smart Review — 순수 로직 (화면·저장소와 무관, node 에서도 불러 씀)

   1. 공통 도구·정규화(공정번호·이름 닮음·주기·시료수·특별특성 기호)
   2. 서류 종류와 열 찾기(머리행 · 두 줄 머리 · 병합 칸 채우기 · 서류 종류 추정)
   3. 서류 읽기(정규화 행)
   4. PDF 글자 → 표(글자 층이 있는 PDF)
   5. 정합성 규칙과 검토 실행(판정 · 공정 대조표 · 점수 · 승인 의견 제안)
   6. 재제출 비교
   7. AI 프롬프트·답 읽기(협력사·품번 별칭)
   8. 보고서(Excel 시트)
   9. 저장 형식
*/
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.APLogic = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ── 1. 공통 도구·정규화 ───────────────────────────────────── */
  function str(v) { return v == null ? '' : String(v).replace(/\s+/g, ' ').trim(); }
  function uniq(a) { var s = {}, o = []; a.forEach(function (x) { if (!s[x]) { s[x] = 1; o.push(x); } }); return o; }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }
  function round(n, d) { var p = Math.pow(10, d || 0); return Math.round(n * p) / p; }
  function pad(n) { return (n < 10 ? '0' : '') + n; }

  /* 엑셀 날짜 일련번호 · Date · 「2026-10-05」「2026.10.5」「2026/10/05」「2026년 10월 5일」 → 'YYYY-MM-DD' */
  function parseDate(v) {
    if (v == null || v === '') return '';
    if (v instanceof Date) return isNaN(v) ? '' : v.getFullYear() + '-' + pad(v.getMonth() + 1) + '-' + pad(v.getDate());
    if (typeof v === 'number' && v > 20000 && v < 80000) { var t = new Date(Date.UTC(1899, 11, 30) + Math.floor(v) * 86400000); return t.getUTCFullYear() + '-' + pad(t.getUTCMonth() + 1) + '-' + pad(t.getUTCDate()); }
    var m = /^(\d{4})\s*[-./년]\s*(\d{1,2})\s*[-./월]\s*(\d{1,2})/.exec(str(v));
    if (!m) return '';
    var y = +m[1], mo = +m[2], d = +m[3], c = new Date(Date.UTC(y, mo - 1, d));
    return c.getUTCMonth() === mo - 1 && y > 1990 ? y + '-' + pad(mo) + '-' + pad(d) : '';
  }

  /* 공정번호 — 「OP10」「OP-10」「P10」「공정 10」「010」 → 「10」, 「10-1」「10.1」 → 「10-1」. 비면 '' */
  function normProcNo(v) {
    if (typeof v === 'number' && isFinite(v)) return String(v).replace(/\.0+$/, '').replace('.', '-');
    var s = str(v).toUpperCase().replace(/\s+/g, '');
    if (!s) return '';
    s = s.replace(/^(OP|OPR|PROC|P|NO\.?|공정번호|공정|#)[-_.:]?(?=\d)/, '');
    s = s.replace(/[._/]/g, '-');
    return s.split('-').map(function (p) { return /^\d+$/.test(p) ? String(+p) : p; }).join('-');
  }
  /* 공정번호 자연 순서(10 < 10-1 < 20 < 100) */
  function procCmp(a, b) {
    var x = String(a).split(/(\d+)/), y = String(b).split(/(\d+)/);
    for (var i = 0; i < Math.max(x.length, y.length); i++) {
      var p = x[i] == null ? '' : x[i], q = y[i] == null ? '' : y[i];
      if (p === q) continue;
      var np = /^\d+$/.test(p), nq = /^\d+$/.test(q);
      if (np && nq) return +p - +q;
      return p < q ? -1 : 1;
    }
    return 0;
  }
  /* 이름 비교용 — 띄어쓰기·문장부호·괄호 기호를 빼고 소문자로 */
  function normName(s) { return str(s).toLowerCase().replace(/[\s·・.,:;'"`~!@#$%^&*_\-+=|\\/<>?()[\]{}「」『』【】]/g, ''); }
  function stripParen(s) { return str(s).replace(/\([^)]*\)|\[[^\]]*\]|（[^）]*）/g, ''); }
  function bigrams(s) { var o = []; for (var i = 0; i < s.length - 1; i++) o.push(s.slice(i, i + 2)); return o; }
  /* 두 이름이 얼마나 닮았는지(0~1).
     같으면 1, 괄호 속(「용접(CO2)」)을 빼고 같으면 1, 한쪽이 다른 쪽을 품으면 0.5 + 0.5 × (짧은/긴),
     그 밖에는 두 글자 묶음(bigram) 겹침 비율(Dice). */
  function similarity(a, b) {
    var x = normName(a), y = normName(b);
    if (!x || !y) return 0;
    if (x === y) return 1;
    var xs = normName(stripParen(a)), ys = normName(stripParen(b));
    if (xs && ys && xs === ys) return 1;
    var s = x.length <= y.length ? x : y, l = x.length <= y.length ? y : x;
    var contain = s.length >= 2 && l.indexOf(s) >= 0 ? 0.5 + 0.5 * (s.length / l.length) : 0;
    var A = bigrams(x), B = bigrams(y);
    if (!A.length || !B.length) return contain;
    var used = {}, hit = 0;
    A.forEach(function (g) { for (var j = 0; j < B.length; j++) if (!used[j] && B[j] === g) { used[j] = 1; hit++; break; } });
    return round(Math.max(contain, 2 * hit / (A.length + B.length)), 3);
  }

  /* 주기 — 같은 뜻의 여러 표기를 한 가지로(「매 로트」「1회/LOT」「로트당」 → lot:1) */
  function normFreq(v) {
    var s = str(v).toLowerCase().replace(/\s+/g, '');
    if (!s) return '';
    if (/전수|100%|全數|all/.test(s)) return 'all';
    if (/초[·.,]?중[·.,]?종|초중종|초물/.test(s)) return 'fml';
    if (/작업시작|시업|start/.test(s)) return 'start';
    if (/변경|교체시|셋업|setup/.test(s)) return 'change';
    var n = 1, m = /(\d+)\s*(회|번|times?|x)/.exec(s) || /^(\d+)\//.exec(s);
    if (m) n = +m[1];
    var unit = /lot|로트|롯트/.test(s) ? 'lot' : /교대|shift|조/.test(s) ? 'shift' : /매일|일|day|daily/.test(s) ? 'day' : /주|week/.test(s) ? 'week' : /월|month/.test(s) ? 'month' : /년|year/.test(s) ? 'year' : /시간|hr|hour/.test(s) ? 'hour' : '';
    if (unit) {
      var per = /(\d+)\s*(일|주|개월|월|시간|교대|로트|lot)/.exec(s.replace(/(\d+)\s*(회|번)/, ''));
      return unit + ':' + n + (per && +per[1] > 1 ? '/' + per[1] : '');
    }
    return s;
  }
  /* 시료수 — 「5개」「5 EA」「n=5」「5pcs」 → 「5」, 「전수」「100%」 → all */
  function normSample(v) {
    var s = str(v).toLowerCase().replace(/\s+/g, '');
    if (!s) return '';
    if (/전수|100%|全數|all/.test(s)) return 'all';
    var m = /^(?:n=?)?(\d+)(개|ea|pcs?|p|매|점|본)?(\/.*)?$/.exec(s);
    if (m) return m[1] + (m[3] ? m[3] : '');
    return s;
  }
  /* 특별특성 기호 대응표 — 「◆=CC, ◇=SC」 글을 읽습니다 */
  var DEFAULT_SYMBOLS = '◆=CC, ◇=SC, ★=CC, ☆=SC, Ⓒ=CC, Ⓢ=SC, CC=CC, SC=SC';
  function parseSymbolMap(text) {
    var m = {};
    String(text || '').split(/[,\n]/).forEach(function (pair) {
      var p = pair.split('=');
      if (p.length === 2 && str(p[0]) && str(p[1])) m[str(p[0]).toUpperCase()] = str(p[1]).toUpperCase();
    });
    return m;
  }
  function normSymbol(v, map) {
    var s = str(v).toUpperCase().replace(/\s+/g, '');
    if (!s || s === '-' || s === '—' || s === '/') return '';
    if (map && map[s]) return map[s];
    return s;
  }
  function splitList(text) { return uniq(String(text || '').split(/[,\n]/).map(str).filter(Boolean)); }
  function truthy(v) { var s = str(v).toUpperCase(); return !!s && /^(Y|YES|O|○|●|◎|V|✓|✔|예|해당|외주|1|TRUE)$/.test(s); }

  /* ── 2. 서류 종류와 열 찾기 ────────────────────────────────────
     aliases: 머리글을 정규화(띄어쓰기·부호 제거, 소문자)한 글과 비교합니다.
     같으면 강하게(100+길이), 머리글이 별칭을 품으면 약하게(길이) 맞춥니다. 「=」로 시작하는 별칭은 같을 때만. */
  var PROC_NO = ['공정번호', '공정no', '공정코드', '공정순서', '공정순번', 'opno', 'op번호', 'operationno', 'processno', 'processnumber', 'processstep', '=op', '=no', '=공정no'];
  var PROC_NAME = ['공정명', '공정이름', '공정내용', '작업명', '작업내용', 'processname', 'operationname', 'operationdescription', 'processdescription', '공정'];
  var DOC_TYPES = [
    { id: 'flow', label: 'Flow Chart', short: 'Flow', hint: /flow|흐름|공정\s*도|프로세스\s*맵/i,
      fields: [
        { id: 'procNo', label: '공정번호', aliases: PROC_NO },
        { id: 'procName', label: '공정명', aliases: PROC_NAME },
        { id: 'kind', label: '공정 구분(선택)', aliases: ['공정구분', '구분', '공정유형', '공정종류', '작업구분', '기호', 'type', 'symbol'] },
        { id: 'outsrc', label: '외주 여부(선택)', aliases: ['외주여부', '외주', '사내외주', '작업처', '사내사외', 'outsourcing', 'outsourced', '외주구분'] },
        { id: 'note', label: '비고(선택)', aliases: ['비고', '설명', 'remark', 'note'] }
      ], req: [['procNo'], ['procName']] },
    { id: 'cp', label: '관리계획서', short: '관리계획서', hint: /관리\s*계획|control\s*plan|(^|[^a-z])q?cp([^a-z]|$)|qc\s*공정/i,
      fields: [
        { id: 'procNo', label: '공정번호', aliases: PROC_NO },
        { id: 'procName', label: '공정명', aliases: PROC_NAME },
        { id: 'charNo', label: '특성번호(선택)', aliases: ['특성번호', '특성no', 'characteristicno', '=char.no'] },
        { id: 'charProd', label: '제품 특성', aliases: ['제품특성', '특성제품', '제품', 'productcharacteristic', 'product', '관리항목', '관리특성'] },
        { id: 'charProc', label: '공정 특성', aliases: ['공정특성', '특성공정', '공정파라미터', '공정조건', 'processcharacteristic', 'processparameter'] },
        { id: 'special', label: '특별특성', aliases: ['특별특성', '특별특성기호', '특별특성분류', '특성분류', '분류', '=class', 'sccc', '=기호', 'specialcharacteristic', 'specialchar'] },
        { id: 'spec', label: '규격·공차(선택)', aliases: ['규격공차', '제품공정규격', '규격', '공차', 'specification', 'spec', 'tolerance', '기준'] },
        { id: 'measure', label: '평가·측정방법', aliases: ['평가측정방법', '측정방법', '평가방법', '검사방법', '측정기', '측정기구', '게이지', '검사구', 'evaluationmeasurementtechnique', 'measurement', 'measuringmethod'] },
        { id: 'sample', label: '시료 크기', aliases: ['시료크기', '시료수', '샘플수', '샘플크기', 'samplesize', '시료', '크기', 'n수', 'sample', '=n'] },
        { id: 'freq', label: '주기', aliases: ['시료주기', '주기', '빈도', '검사주기', '측정주기', '검사시기', 'frequency', 'freq', 'samplefreq'] },
        { id: 'control', label: '관리방법', aliases: ['관리방법', '관리방식', '관리수단', 'controlmethod', '관리도구'] },
        { id: 'reaction', label: '대응계획(선택)', aliases: ['대응계획', '조치방법', '이상시조치', 'reactionplan'] }
      ], req: [['procNo'], ['charProd', 'charProc']] },
    { id: 'pfmea', label: 'PFMEA', short: 'PFMEA', hint: /fmea/i,
      fields: [
        { id: 'procNo', label: '공정번호', aliases: PROC_NO },
        { id: 'procName', label: '공정명', aliases: PROC_NAME.concat(['공정기능', '공정단계']) },
        { id: 'char', label: '요구사항·특성', aliases: ['요구사항', '공정요구사항', '기능요구사항', '특성', '관리항목', 'requirement', 'requirements', 'item'] },
        { id: 'mode', label: '잠재 고장형태', aliases: ['잠재적고장형태', '잠재고장형태', '고장형태', '고장모드', 'failuremode', 'potentialfailuremode'] },
        { id: 'effect', label: '고장영향(선택)', aliases: ['잠재적고장영향', '고장영향', '영향', 'effect', 'failureeffect'] },
        { id: 'cause', label: '원인(선택)', aliases: ['잠재적원인', '고장원인', '원인', '메커니즘', 'cause', 'failurecause'] },
        { id: 'S', label: '심각도 S', aliases: ['심각도', '엄격도', 'severity', '=s', '=sev'] },
        { id: 'O', label: '발생도 O(선택)', aliases: ['발생도', 'occurrence', '=o', '=occ'] },
        { id: 'D', label: '검출도 D(선택)', aliases: ['검출도', 'detection', '=d', '=det'] },
        { id: 'RPN', label: 'RPN(선택)', aliases: ['rpn', '위험우선순위', '위험우선수'] },
        { id: 'special', label: '분류·특별특성(선택)', aliases: ['분류', '특별특성', '=class', 'classification'] },
        { id: 'prevent', label: '현 관리(선택)', aliases: ['현공정관리', '현관리', '예방관리', '검출관리', 'currentcontrol', 'currentprocesscontrol'] }
      ], req: [['procNo'], ['mode', 'char'], ['S']] },
    { id: 'insp', label: '검사기준서', short: '검사기준서', hint: /검사\s*기준|검사\s*표준|검사\s*규격|검사\s*지도|inspection/i,
      fields: [
        { id: 'procNo', label: '공정번호(선택)', aliases: PROC_NO },
        { id: 'procName', label: '검사 구분(선택)', aliases: PROC_NAME.concat(['검사구분', '검사단계']) },
        { id: 'item', label: '검사항목', aliases: ['검사항목', '측정항목', '검사내용', '관리항목', '항목', '특성', 'inspectionitem', 'item', 'characteristic'] },
        { id: 'special', label: '특별특성', aliases: ['특별특성', '특별특성기호', '특성분류', '분류', '=class', '=기호', 'specialchar'] },
        { id: 'spec', label: '규격·판정기준(선택)', aliases: ['규격', '판정기준', '검사기준', '기준', '공차', 'spec', 'specification', 'criteria'] },
        { id: 'method', label: '검사방법(선택)', aliases: ['검사방법', '측정방법', '측정기', '검사구', '측정기구', 'method', 'gauge'] },
        { id: 'sample', label: '시료수', aliases: ['시료수', '시료크기', '샘플수', '샘플링', '시료', 'n수', 'samplesize', 'sample', '=n'] },
        { id: 'freq', label: '주기', aliases: ['검사주기', '주기', '빈도', '측정주기', '검사시기', 'frequency', 'freq'] }
      ], req: [['item']] },
    { id: 'outs', label: '외주품 리스트', short: '외주 리스트', hint: /외주/i,
      fields: [
        { id: 'procNo', label: '공정번호(선택)', aliases: PROC_NO },
        { id: 'procName', label: '외주 공정', aliases: ['외주공정', '외주공정명', '외주내용', '외주작업'].concat(PROC_NAME) },
        { id: 'vendor', label: '외주처(선택)', aliases: ['외주처', '외주업체', '외주사', '업체명', '업체', '협력사', 'vendor', 'subcontractor'] },
        { id: 'partNo', label: '품번(선택)', aliases: ['품번', '부품번호', 'partno', '=p/n', '=pn'] },
        { id: 'note', label: '비고(선택)', aliases: ['비고', 'remark', 'note'] }
      ], req: [['procNo', 'procName']] }
  ];
  var DOC_IDS = DOC_TYPES.map(function (d) { return d.id; });
  function docType(id) { return DOC_TYPES.filter(function (d) { return d.id === id; })[0] || null; }
  function docLabel(id) { var d = docType(id); return d ? d.label : id; }

  function normHead(v) { return str(v).replace(/[\s·・.:_\-/()[\]（）&,]/g, '').toLowerCase(); }
  /* 머리글 칸들 → { 항목id: 열번호 } (한 열은 한 항목에만) */
  function matchColumns(labels, fields) {
    var pairs = [];
    labels.forEach(function (lab, c) {
      var n = normHead(lab);
      if (!n) return;
      fields.forEach(function (f) {
        var best = 0;
        f.aliases.forEach(function (a) {
          var exact = a.charAt(0) === '=', b = exact ? a.slice(1) : a;
          if (n === b) best = Math.max(best, 100 + b.length);
          else if (!exact && b.length >= 2 && n.indexOf(b) >= 0) best = Math.max(best, b.length);
        });
        if (best) pairs.push({ c: c, f: f.id, s: best });
      });
    });
    pairs.sort(function (a, b) { return b.s - a.s || a.c - b.c; });
    var map = {}, usedC = {};
    pairs.forEach(function (p) { if (map[p.f] == null && !usedC[p.c]) { map[p.f] = p.c; usedC[p.c] = 1; } });
    return map;
  }
  function missingOf(type, map) {
    return docType(type).req.filter(function (g) { return !g.some(function (id) { return map[id] != null; }); })
      .map(function (g) { return g.join('|'); });
  }
  /* 두 줄 머리 잇기. 병합 정보가 없는 파일(CSV 등)에서는 묶음 머리(「특성」)가 첫 칸에만 있으므로,
     위 칸이 비고 아래 칸이 있으면 바로 왼쪽 묶음 머리를 이어 씁니다(위·아래가 같은 칸은 묶음이 아님). */
  function combineRows(up, low) {
    var n = Math.max((up || []).length, (low || []).length), out = [], group = '';
    for (var c = 0; c < n; c++) {
      var u = str(up && up[c]), l = str(low && low[c]);
      if (u) group = u !== l && l ? u : '';
      else if (l && group) u = group;
      else group = '';
      out.push(!l ? u : !u ? l : u === l ? u : u + ' ' + l);
    }
    return out;
  }
  /* 위 20행 안에서 머리행을 찾습니다. 「특성 > 제품 · 공정」처럼 두 줄로 된 머리도 이어 붙여 봅니다
     (두 줄 머리는 한 줄보다 더 많은 항목을 맞힐 때만 고릅니다 — 첫 자료 행을 머리로 먹지 않게). */
  function detectColumns(aoa, type) {
    var dt = docType(type), best = { headerRow: -1, headerRows: 1, map: {}, score: 0 };
    if (!dt) return null;
    var lim = Math.min(20, aoa.length), singles = [];
    for (var r = 0; r < Math.min(21, aoa.length); r++) singles.push(matchColumns((aoa[r] || []).map(str), dt.fields));
    for (r = 0; r < lim; r++) {
      var one = singles[r], s1 = Object.keys(one).length;
      var s1b = s1 + (missingOf(type, one).length ? 0 : 0.5);
      if (s1b > best.score) best = { headerRow: r, headerRows: 1, map: one, score: s1b };
      if (r + 1 < aoa.length) {
        var two = matchColumns(combineRows(aoa[r], aoa[r + 1]), dt.fields), s2 = Object.keys(two).length;
        var s2b = s2 + (missingOf(type, two).length ? 0 : 0.5);
        var sNext = singles[r + 1] ? Object.keys(singles[r + 1]).length : 0;
        if (s2 > s1 && s2 > sNext && s2b > best.score) best = { headerRow: r, headerRows: 2, map: two, score: s2b };
      }
    }
    if (best.score < 2) best = { headerRow: -1, headerRows: 1, map: {}, score: 0 };
    return finishDet(aoa, type, best);
  }
  function headerLabels(aoa, headerRow, headerRows) {
    if (headerRow < 0) return [];
    return headerRows === 2 ? combineRows(aoa[headerRow], aoa[headerRow + 1]) : (aoa[headerRow] || []).map(str);
  }
  function finishDet(aoa, type, d) {
    var headers = headerLabels(aoa, d.headerRow, d.headerRows);
    return { type: type, headerRow: d.headerRow, headerRows: d.headerRows, map: d.map, headers: headers, missing: missingOf(type, d.map), score: d.score };
  }
  /* 병합 칸 채우기 — SheetJS 는 병합 칸의 첫 칸에만 값을 둡니다. merges = ws['!merges'] */
  function fillMerges(aoa, merges) {
    (merges || []).forEach(function (m) {
      var v = aoa[m.s.r] ? aoa[m.s.r][m.s.c] : null;
      if (v == null || v === '') return;
      for (var r = m.s.r; r <= m.e.r; r++) {
        if (!aoa[r]) aoa[r] = [];
        for (var c = m.s.c; c <= m.e.c; c++) if (r !== m.s.r || c !== m.s.c) aoa[r][c] = v;
      }
    });
    return aoa;
  }
  /* 시트 하나가 어느 서류인지 — 시트 이름·파일 이름 낱말(강함) + 맞는 열 개수 */
  function guessDocType(aoa, sheetName, fileName) {
    var scores = {}, best = '', bestScore = 0;
    DOC_TYPES.forEach(function (dt) {
      var d = detectColumns(aoa, dt.id), s = d.headerRow < 0 ? 0 : Object.keys(d.map).length + (d.missing.length ? 0 : 3);
      if (dt.hint.test(sheetName || '')) s += 20;
      else if (dt.hint.test(fileName || '')) s += 8;
      scores[dt.id] = s;
    });
    // 외주품 리스트·PFMEA 는 이름 낱말이 겹치기 쉬워(「외주 공정 관리계획서」) 이름이 둘 다 맞으면 열 점수로 가립니다
    DOC_IDS.forEach(function (id) { if (scores[id] > bestScore) { bestScore = scores[id]; best = id; } });
    return { type: bestScore >= 4 ? best : '', scores: scores };
  }

  /* ── 3. 서류 읽기 ──────────────────────────────────────────────
     opts = { headerRow, headerRows, overrides: {항목id: 열번호(-1 = 끊기)} } — 사용자가 고친 경우.
     결과 행: { _r(시트 행 번호, 1부터), procNo, procName, … }. 공정번호·공정명은 빈 칸이면 위 행 값을 이어 씁니다(병합 안 된 「같음」 칸). */
  var FILL_DOWN = { cp: true, pfmea: true, insp: true, outs: false, flow: false };
  function parseDoc(aoa, type, opts) {
    opts = opts || {};
    var det = detectColumns(aoa, type);
    if (opts.headerRow != null && opts.headerRow >= 0) {
      var hr = opts.headerRow, hn = opts.headerRows === 2 ? 2 : 1;
      if (hr !== det.headerRow || hn !== det.headerRows) det = finishDet(aoa, type, { headerRow: hr, headerRows: hn, map: matchColumns(headerLabels(aoa, hr, hn), docType(type).fields), score: 0 });
    }
    if (opts.overrides) {
      Object.keys(opts.overrides).forEach(function (k) { var v = opts.overrides[k]; if (v == null || v < 0) delete det.map[k]; else det.map[k] = +v; });
      det.missing = missingOf(type, det.map);
    }
    var rows = [], skipped = 0;
    if (det.headerRow < 0 || det.missing.length) return { det: det, rows: rows, skipped: 0 };
    var M = det.map, ids = Object.keys(M), last = { procNo: '', procName: '' };
    for (var r = det.headerRow + det.headerRows; r < aoa.length; r++) {
      var row = aoa[r] || [], o = { _r: r + 1 };
      ids.forEach(function (id) { o[id] = str(row[M[id]]); });
      var content = ids.filter(function (id) { return id !== 'procNo' && id !== 'procName' && o[id]; }).length;
      if (!content && !o.procNo && !o.procName) { if (row.some(function (v) { return str(v); })) skipped++; continue; }
      if (/^(합계|소계|total|계|비고|작성|검토|승인)$/i.test(o.procNo || o.procName || '')) { skipped++; continue; }
      if (FILL_DOWN[type] && content) {
        if (!o.procNo && M.procNo != null) { o.procNo = last.procNo; if (!o.procName) o.procName = last.procName; }
        else if (o.procNo && !o.procName && normProcNo(o.procNo) === last.procNo) o.procName = last.procName;
      }
      o.procNo = normProcNo(o.procNo);
      if (o.procNo && o.procNo !== last.procNo) { last.procNo = o.procNo; last.procName = o.procName; }
      else if (o.procName) last.procName = o.procName;
      if (type === 'flow' && !o.procNo && !o.procName) { skipped++; continue; }
      if (type === 'cp') { o.char = [o.charProd, o.charProc].filter(Boolean).join(' / '); if (!o.char) { skipped++; continue; } }
      if (type === 'pfmea') {
        ['S', 'O', 'D', 'RPN'].forEach(function (k) { var n = parseFloat(o[k]); o[k] = isFinite(n) ? n : null; });
        if (o.RPN == null && o.S != null && o.O != null && o.D != null) { o.RPN = o.S * o.O * o.D; o.rpnCalc = true; }
        if (!o.mode && !o.char) { skipped++; continue; }
      }
      if (type === 'insp' && !o.item) { skipped++; continue; }
      if (type === 'flow') { o.outsourced = /외주/.test(o.kind || '') || truthy(o.outsrc) || /외주/.test(o.procName || ''); }
      rows.push(o);
    }
    return { det: det, rows: rows, skipped: skipped };
  }

  /* ── 4. PDF 글자 → 표 ─────────────────────────────────────────
     itemsToLines·pdfTextItems 는 1차수 data09-08(CMM PDF 성적서 읽기)의 것을 옮겨 왔습니다.
     items: [{ str, x, y, w, page }] (y 는 위에서 아래로). 같은 높이(±2pt)를 한 줄로, 가까운 조각(gap pt 미만)은 한 구절로. */
  function itemsToLines(items, gap) {
    gap = gap == null ? 6 : gap;
    var list = (items || []).filter(function (t) { return String(t.str || '').trim() !== ''; })
      .map(function (t) { return { str: String(t.str), x: Number(t.x), y: Number(t.y), w: Number(t.w) || 0, page: Number(t.page) || 1 }; });
    list.sort(function (a, b) { return a.page - b.page || a.y - b.y || a.x - b.x; });
    var lines = [];
    list.forEach(function (t) {
      var ln = lines[lines.length - 1];
      if (!ln || ln.page !== t.page || Math.abs(ln.y - t.y) > 2) { ln = { page: t.page, y: t.y, items: [] }; lines.push(ln); }
      ln.items.push(t);
    });
    lines.forEach(function (ln) {
      ln.items.sort(function (a, b) { return a.x - b.x; });
      var ph = [];
      ln.items.forEach(function (t) {
        var s = t.str.replace(/\s+/g, ' ').trim(), p = ph[ph.length - 1], d = p ? t.x - p.right : Infinity;
        if (p && d < gap) { p.text += (d > 0.8 ? ' ' : '') + s; p.right = Math.max(p.right, t.x + t.w); }
        else ph.push({ text: s, x: t.x, right: t.x + t.w });
      });
      ln.phrases = ph;
      ln.text = ph.map(function (p) { return p.text; }).join('  ');
    });
    return lines;
  }
  function pdfTextItems(textContent, page, pageHeight) {
    return (textContent.items || []).filter(function (it) { return it.str && it.str.trim(); }).map(function (it) {
      return { str: it.str, x: it.transform[4], y: pageHeight - it.transform[5], w: it.width, page: page };
    });
  }
  /* 줄 → 표(배열의 배열). 머리글 낱말이 가장 많이 맞는 줄을 머리행으로 보고, 그 구절들의 가운데를 열 기준으로 삼아
     아래 줄의 구절을 가장 가까운 열 칸에 넣습니다. 쪽마다 되풀이되는 같은 머리행은 한 번만 둡니다.
     머리행 위의 줄(제목·품번 등)은 한 칸짜리 행으로 남깁니다. */
  function headerHits(texts) {
    var n = 0;
    texts.forEach(function (t) {
      var h = normHead(t);
      if (!h) return;
      if (DOC_TYPES.some(function (dt) { return dt.fields.some(function (f) { return f.aliases.some(function (a) { var e = a.charAt(0) === '=', b = e ? a.slice(1) : a; return h === b || (!e && b.length >= 3 && h.indexOf(b) >= 0); }); }); })) n++;
    });
    return n;
  }
  function pdfLinesToAoa(lines) {
    var hi = -1, hs = 1;
    lines.forEach(function (ln, i) { var s = headerHits(ln.phrases.map(function (p) { return p.text; })); if (s > hs && ln.phrases.length >= 2) { hs = s; hi = i; } });
    if (hi < 0) return lines.map(function (ln) { return ln.phrases.map(function (p) { return p.text; }); });
    var head = lines[hi].phrases, centers = head.map(function (p) { return (p.x + p.right) / 2; }), sig = lines[hi].text;
    var aoa = [];
    lines.forEach(function (ln, i) {
      if (i < hi) { aoa.push([ln.text]); return; }
      if (i > hi && ln.text === sig) return;
      var row = head.map(function () { return null; });
      ln.phrases.forEach(function (p) {
        var c = (p.x + p.right) / 2, k = 0, best = Infinity;
        centers.forEach(function (x, j) {
          // 칸 왼쪽 끝이 머리글 범위 안이면 그 칸으로(왼쪽 정렬 글), 아니면 가운데가 가장 가까운 칸
          var d = p.x >= head[j].x - 3 && p.x <= head[j].right ? -1 : Math.abs(c - x);
          if (d < best) { best = d; k = j; }
        });
        row[k] = row[k] == null ? p.text : row[k] + ' ' + p.text;
      });
      aoa.push(row);
    });
    return aoa;
  }

  /* ── 5. 정합성 규칙과 검토 실행 ─────────────────────────────── */
  var VERDICT = { ok: '적합', mismatch: '불일치', missing: '누락', check: '확인 필요' };
  var SEV = { major: '중대', minor: '경미' };
  var RULES = [
    { id: 'docs', name: '필수 서류 제출', desc: 'Flow Chart·관리계획서·PFMEA·검사기준서가 모두 있는지. 외주품 리스트는 Flow Chart 에 외주 공정이 있을 때만 필수.', sev: 'major', params: [] },
    { id: 'proc_presence', name: '공정번호 존재 일치', desc: 'Flow Chart 의 공정번호가 관리계획서·PFMEA 에 모두 있는지. 반대로 관리계획서·PFMEA 에만 있는 공정번호는 확인 필요.', sev: 'major',
      params: [{ id: 'exclude', label: '대조에서 뺄 공정(공정 구분·공정명 낱말, 쉼표)', type: 'text', def: '이동, 운반, 보관, 대기, 저장' }] },
    { id: 'proc_name', name: '공정명 일치', desc: '같은 공정번호의 공정명이 Flow Chart 와 관리계획서·PFMEA·외주품 리스트에서 같은지(띄어쓰기·부호·괄호 속은 무시하고 닮음으로 비교).', sev: 'minor',
      params: [{ id: 'okSim', label: '적합으로 볼 닮음(%) 이상', type: 'num', def: 80, min: 50, max: 100 }, { id: 'checkSim', label: '확인 필요로 볼 닮음(%) 이상 — 그 아래는 불일치', type: 'num', def: 50, min: 0, max: 99 }] },
    { id: 'pfmea_high', name: 'PFMEA 고위험 항목 → 관리계획서', desc: 'PFMEA 에서 RPN 이 기준 이상이거나 심각도(S)가 기준 이상인 항목이, 관리계획서의 같은 공정에 관리방법과 함께 있는지.', sev: 'major',
      params: [{ id: 'rpnMin', label: 'RPN 기준(이상)', type: 'num', def: 100, min: 1, max: 1000 }, { id: 'sMin', label: '심각도 S 기준(이상)', type: 'num', def: 9, min: 1, max: 10 }] },
    { id: 'special', name: '특별특성 기호 일치', desc: '관리계획서의 특별특성 기호가 검사기준서의 같은 항목에 같은 기호로 있는지(기호 대응표로 ◆·CC 같은 다른 표기를 맞춤).', sev: 'major',
      params: [{ id: 'symbols', label: '기호 대응표(「기호=분류」, 쉼표)', type: 'text', def: DEFAULT_SYMBOLS }] },
    { id: 'insp_item', name: '관리계획서 검사 항목 → 검사기준서', desc: '관리계획서의 검사 항목이 검사기준서에 있는지.', sev: 'major',
      params: [{ id: 'scope', label: '대상', type: 'select', def: 'insp', options: [['insp', '검사 공정의 항목만(공정명에 아래 낱말)'], ['all', '관리계획서의 모든 항목']] },
        { id: 'keywords', label: '검사 공정 낱말(쉼표)', type: 'text', def: '검사, 측정, 시험' }] },
    { id: 'insp_cond', name: '검사 주기·시료수 일치', desc: '관리계획서와 검사기준서에 함께 있는 항목의 주기·시료수가 같은지(「매 로트」=「1회/LOT」, 「5개」=「5 EA」처럼 같은 뜻은 같게 봄).', sev: 'minor',
      params: [{ id: 'freq', label: '주기 비교', type: 'bool', def: true }, { id: 'sample', label: '시료수 비교', type: 'bool', def: true }] },
    { id: 'outsourced', name: '외주 공정 → 외주품 리스트', desc: 'Flow Chart 에서 외주로 표시된 공정(공정 구분·외주 여부·공정명의 「외주」)이 외주품 리스트에 있는지.', sev: 'major', params: [] }
  ];
  var MATCH_PARAMS = [
    { id: 'charOk', label: '항목(특성) 이름 — 같다고 볼 닮음(%) 이상', def: 60, min: 30, max: 100 },
    { id: 'charCheck', label: '항목(특성) 이름 — 확인 필요로 볼 닮음(%) 이상', def: 40, min: 0, max: 99 }
  ];
  var DECISION_PARAMS = [
    { id: 'rejectMajor', label: '반려 제안 — 중대 지적(누락·불일치) 건수 이상', def: 1, min: 1, max: 99 },
    { id: 'rejectMinor', label: '반려 제안 — 경미 지적 건수 이상', def: 6, min: 1, max: 999 }
  ];
  var DECISION = { approve: '승인', conditional: '조건부 승인', reject: '반려' };
  function ruleMeta(id) { return RULES.filter(function (r) { return r.id === id; })[0]; }
  function defaultSettings() {
    var s = { rules: {}, match: {}, decision: {} };
    RULES.forEach(function (r) { var o = { on: true, sev: r.sev }; r.params.forEach(function (p) { o[p.id] = p.def; }); s.rules[r.id] = o; });
    MATCH_PARAMS.forEach(function (p) { s.match[p.id] = p.def; });
    DECISION_PARAMS.forEach(function (p) { s.decision[p.id] = p.def; });
    return s;
  }
  function numIn(v, p) { var n = Number(v); return isFinite(n) && n >= p.min && n <= p.max ? Math.round(n) : p.def; }
  function cleanSettings(o) {
    var s = defaultSettings();
    if (!o || typeof o !== 'object') return s;
    RULES.forEach(function (r) {
      var x = o.rules && o.rules[r.id];
      if (!x || typeof x !== 'object') return;
      var t = s.rules[r.id];
      t.on = x.on !== false;
      if (SEV[x.sev]) t.sev = x.sev;
      r.params.forEach(function (p) {
        if (x[p.id] == null) return;
        if (p.type === 'num') t[p.id] = numIn(x[p.id], p);
        else if (p.type === 'bool') t[p.id] = x[p.id] !== false;
        else if (p.type === 'select') t[p.id] = p.options.some(function (q) { return q[0] === x[p.id]; }) ? x[p.id] : p.def;
        else t[p.id] = String(x[p.id]);
      });
    });
    if (s.rules.proc_name.checkSim >= s.rules.proc_name.okSim) s.rules.proc_name.checkSim = s.rules.proc_name.okSim - 1;
    MATCH_PARAMS.forEach(function (p) { if (o.match && o.match[p.id] != null) s.match[p.id] = numIn(o.match[p.id], p); });
    if (s.match.charCheck >= s.match.charOk) s.match.charCheck = s.match.charOk - 1;
    DECISION_PARAMS.forEach(function (p) { if (o.decision && o.decision[p.id] != null) s.decision[p.id] = numIn(o.decision[p.id], p); });
    return s;
  }

  function rowsOf(docs, id) { return docs && docs[id] && Array.isArray(docs[id].rows) ? docs[id].rows : null; }
  function procSet(rows) { var m = {}; (rows || []).forEach(function (r) { if (r.procNo) (m[r.procNo] = m[r.procNo] || []).push(r); }); return m; }
  function hasWord(text, words) { var t = normName(text); return words.some(function (w) { var n = normName(w); return n && t.indexOf(n) >= 0; }); }
  function ev(doc, r, text) { return { doc: doc, row: r && r._r, text: text == null ? '' : text }; }
  /* 후보 행들 가운데 이름이 가장 닮은 것 — keys: 비교할 칸들(가장 높은 닮음을 씀) */
  function bestMatch(name, cands, keys) {
    var best = null;
    (cands || []).forEach(function (c) {
      keys.forEach(function (k) {
        var s = similarity(name, c[k]);
        if (s > 0 && (!best || s > best.sim)) best = { row: c, sim: s, key: k };
      });
    });
    return best;
  }
  function procLabel(no, name) { return no ? (name ? no + ' ' + name : no) : (name || ''); }

  /* docs = { flow:{rows}, cp:{rows}, pfmea:{rows}, insp:{rows}, outs:{rows} } (없는 서류는 빠짐), settings = cleanSettings 결과.
     결과 = { findings[], coverage, summary } — findings 의 verdict 는 ok·mismatch·missing·check */
  function runReview(docs, settings) {
    var S = cleanSettings(settings), R = S.rules, M = S.match;
    var F = [], seq = 0;
    function add(rule, verdict, procNo, item, msg, evidence) {
      F.push({ id: 'f' + (++seq), rule: rule, verdict: verdict, sev: R[rule].sev, procNo: procNo || '', item: item || '', msg: msg, ev: evidence || [] });
    }
    var flow = rowsOf(docs, 'flow'), cp = rowsOf(docs, 'cp'), pf = rowsOf(docs, 'pfmea'), insp = rowsOf(docs, 'insp'), outs = rowsOf(docs, 'outs');
    var fP = procSet(flow), cP = procSet(cp), pP = procSet(pf), iP = procSet(insp), oP = procSet(outs);
    var exclude = splitList(R.proc_presence.exclude);
    var flowProcs = Object.keys(fP).sort(procCmp);
    function excluded(no) { var r = fP[no][0]; return exclude.length && (hasWord(r.kind || '', exclude) || hasWord(r.procName || '', exclude)); }
    var outsrcFlow = (flow || []).filter(function (r) { return r.outsourced; });
    var charOk = M.charOk / 100, charCheck = M.charCheck / 100;

    // 필수 서류
    if (R.docs.on) {
      ['flow', 'cp', 'pfmea', 'insp'].forEach(function (id) {
        if (rowsOf(docs, id)) add('docs', 'ok', '', docLabel(id), docLabel(id) + ' 제출됨 (' + docs[id].rows.length + '행)', [ev(id, null, docs[id].fileName || '')]);
        else add('docs', 'missing', '', docLabel(id), docLabel(id) + ' 이 제출되지 않았습니다.', []);
      });
      if (outs) add('docs', 'ok', '', docLabel('outs'), docLabel('outs') + ' 제출됨 (' + outs.length + '행)', [ev('outs', null, docs.outs.fileName || '')]);
      else if (!flow) add('docs', 'check', '', docLabel('outs'), 'Flow Chart 가 없어 외주품 리스트가 필요한지 알 수 없습니다.', []);
      else if (outsrcFlow.length) add('docs', 'missing', '', docLabel('outs'), 'Flow Chart 에 외주 공정 ' + outsrcFlow.length + '개가 있는데 외주품 리스트가 제출되지 않았습니다.', outsrcFlow.map(function (r) { return ev('flow', r, procLabel(r.procNo, r.procName)); }));
    }

    // 공정번호 존재 일치
    if (R.proc_presence.on && flow) {
      var targets = [['cp', cP, cp], ['pfmea', pP, pf]].filter(function (t) { return t[2]; });
      flowProcs.forEach(function (no) {
        var fr = fP[no][0];
        if (excluded(no)) return;
        targets.forEach(function (t) {
          if (t[1][no]) add('proc_presence', 'ok', no, docLabel(t[0]), procLabel(no, fr.procName) + ' — ' + docLabel(t[0]) + ' 에 있음', [ev('flow', fr, fr.procName), ev(t[0], t[1][no][0], t[1][no][0].procName || '')]);
          else add('proc_presence', 'missing', no, docLabel(t[0]), 'Flow Chart 의 공정 ' + procLabel(no, fr.procName) + ' 이 ' + docLabel(t[0]) + ' 에 없습니다.', [ev('flow', fr, fr.procName)]);
        });
      });
      targets.forEach(function (t) {
        Object.keys(t[1]).sort(procCmp).forEach(function (no) {
          if (fP[no]) return;
          var r = t[1][no][0];
          add('proc_presence', 'check', no, docLabel(t[0]), docLabel(t[0]) + ' 의 공정 ' + procLabel(no, r.procName) + ' 이 Flow Chart 에 없습니다.', [ev(t[0], r, r.procName)]);
        });
      });
    }

    // 공정명 일치 — 기준: Flow Chart(없으면 관리계획서)
    if (R.proc_name.on) {
      var okS = R.proc_name.okSim / 100, ckS = R.proc_name.checkSim / 100;
      var baseId = flow ? 'flow' : cp ? 'cp' : null, baseP = flow ? fP : cP;
      if (baseId) {
        [['cp', cP], ['pfmea', pP], ['outs', oP]].forEach(function (t) {
          if (t[0] === baseId || !rowsOf(docs, t[0])) return;
          Object.keys(t[1]).sort(procCmp).forEach(function (no) {
            if (!baseP[no]) return;
            var bn = baseP[no][0].procName, names = uniq(t[1][no].map(function (r) { return r.procName; }).filter(Boolean));
            if (!bn || !names.length) return;
            names.forEach(function (nm) {
              var r = t[1][no].filter(function (x) { return x.procName === nm; })[0], s = similarity(bn, nm);
              var e = [ev(baseId, baseP[no][0], bn), ev(t[0], r, nm)];
              var pct = Math.round(s * 100) + '%';
              if (s >= okS) add('proc_name', 'ok', no, docLabel(t[0]), '공정 ' + no + ' 공정명 일치: 「' + bn + '」 / 「' + nm + '」 (닮음 ' + pct + ')', e);
              else if (s >= ckS) add('proc_name', 'check', no, docLabel(t[0]), '공정 ' + no + ' 공정명이 조금 다릅니다: ' + docLabel(baseId) + ' 「' + bn + '」 · ' + docLabel(t[0]) + ' 「' + nm + '」 (닮음 ' + pct + ')', e);
              else add('proc_name', 'mismatch', no, docLabel(t[0]), '공정 ' + no + ' 공정명이 다릅니다: ' + docLabel(baseId) + ' 「' + bn + '」 · ' + docLabel(t[0]) + ' 「' + nm + '」 (닮음 ' + pct + ')', e);
            });
          });
        });
      }
    }

    // PFMEA 고위험 → 관리계획서
    if (R.pfmea_high.on && pf && cp) {
      pf.forEach(function (r) {
        var why = [];
        if (r.RPN != null && r.RPN >= R.pfmea_high.rpnMin) why.push('RPN ' + r.RPN + (r.rpnCalc ? '(S×O×D)' : ''));
        if (r.S != null && r.S >= R.pfmea_high.sMin) why.push('S ' + r.S);
        if (!why.length) return;
        var name = r.char || r.mode, label = [r.char, r.mode].filter(Boolean).join(' — ');
        var cands = r.procNo && cP[r.procNo] ? cP[r.procNo] : [];
        var e0 = ev('pfmea', r, label + ' · ' + why.join(', '));
        if (!cands.length) { add('pfmea_high', 'missing', r.procNo, name, 'PFMEA 고위험 항목(' + why.join(', ') + ') 「' + label + '」 — 관리계획서에 공정 ' + (r.procNo || '(번호 없음)') + ' 이 없습니다.', [e0]); return; }
        var b = bestMatch(r.char || '', cands, ['charProd', 'charProc']), b2 = bestMatch(r.mode || '', cands, ['charProd', 'charProc']), b3 = bestMatch(r.cause || '', cands, ['charProd', 'charProc']);
        [b2, b3].forEach(function (x) { if (x && (!b || x.sim > b.sim)) b = x; });
        if (!b || b.sim < charCheck) {
          add('pfmea_high', 'missing', r.procNo, name, 'PFMEA 고위험 항목(' + why.join(', ') + ') 「' + label + '」 이 관리계획서 공정 ' + r.procNo + ' 에 없습니다. 같은 공정의 관리항목: ' + cands.map(function (c) { return c.char; }).join(', '), [e0].concat(cands.slice(0, 3).map(function (c) { return ev('cp', c, c.char); })));
        } else if (b.sim < charOk) {
          add('pfmea_high', 'check', r.procNo, name, 'PFMEA 고위험 항목(' + why.join(', ') + ') 「' + label + '」 과 관리계획서 「' + b.row.char + '」 이 같은 항목인지 확인해 주세요 (닮음 ' + Math.round(b.sim * 100) + '%).', [e0, ev('cp', b.row, b.row.char)]);
        } else if (!b.row.control && !b.row.measure) {
          add('pfmea_high', 'mismatch', r.procNo, name, 'PFMEA 고위험 항목(' + why.join(', ') + ') 「' + label + '」 은 관리계획서에 있지만 관리방법이 비어 있습니다.', [e0, ev('cp', b.row, b.row.char)]);
        } else {
          add('pfmea_high', 'ok', r.procNo, name, 'PFMEA 고위험 항목(' + why.join(', ') + ') 「' + label + '」 → 관리계획서 「' + b.row.char + '」 · 관리방법 ' + (b.row.control || b.row.measure), [e0, ev('cp', b.row, b.row.char + ' · ' + (b.row.control || b.row.measure))]);
        }
      });
    }

    // 관리계획서 항목 ↔ 검사기준서 항목 짝 찾기(같은 공정번호 먼저, 없으면 전체에서)
    function inspMatch(c) {
      if (!insp) return null;
      var same = c.procNo && iP[c.procNo] ? iP[c.procNo] : [];
      var b = bestMatch(c.charProd || '', same, ['item']), b2 = bestMatch(c.charProc || '', same, ['item']);
      if (b2 && (!b || b2.sim > b.sim)) b = b2;
      if (!b || b.sim < charOk) {
        var a = bestMatch(c.charProd || '', insp, ['item']), a2 = bestMatch(c.charProc || '', insp, ['item']);
        if (a2 && (!a || a2.sim > a.sim)) a = a2;
        if (a && (!b || a.sim > b.sim)) b = a;
      }
      return b;
    }

    // 특별특성 기호
    if (R.special.on && cp && insp) {
      var sm = parseSymbolMap(R.special.symbols), seenInsp = {};
      cp.forEach(function (c) {
        var cs = normSymbol(c.special, sm), b = inspMatch(c);
        if (b && b.sim >= charOk) seenInsp[b.row._r] = c;
        if (!cs) {
          if (b && b.sim >= charOk && normSymbol(b.row.special, sm)) add('special', 'mismatch', c.procNo, c.char, '검사기준서 「' + b.row.item + '」 에는 특별특성 「' + b.row.special + '」 이 있는데 관리계획서 「' + c.char + '」 에는 기호가 없습니다.', [ev('cp', c, c.char), ev('insp', b.row, b.row.item + ' · ' + b.row.special)]);
          return;
        }
        if (!b || b.sim < charCheck) { add('special', 'missing', c.procNo, c.char, '관리계획서의 특별특성 항목 「' + c.char + '」(' + c.special + ') 이 검사기준서에 없습니다.', [ev('cp', c, c.char + ' · ' + c.special)]); return; }
        var e = [ev('cp', c, c.char + ' · ' + c.special), ev('insp', b.row, b.row.item + ' · ' + (b.row.special || '(기호 없음)'))];
        if (b.sim < charOk) { add('special', 'check', c.procNo, c.char, '관리계획서 「' + c.char + '」(' + c.special + ') 과 검사기준서 「' + b.row.item + '」 이 같은 항목인지 확인해 주세요 (닮음 ' + Math.round(b.sim * 100) + '%).', e); return; }
        var is = normSymbol(b.row.special, sm);
        if (is === cs) add('special', 'ok', c.procNo, c.char, '특별특성 일치: 「' + c.char + '」 ' + c.special + (b.row.special !== c.special ? ' = ' + b.row.special : ''), e);
        else add('special', 'mismatch', c.procNo, c.char, '특별특성 기호가 다릅니다: 관리계획서 「' + c.char + '」 ' + c.special + (sm[str(c.special).toUpperCase()] ? '(' + cs + ')' : '') + ' · 검사기준서 「' + b.row.item + '」 ' + (b.row.special ? b.row.special + (is ? '(' + is + ')' : '') : '(기호 없음)'), e);
      });
      insp.forEach(function (r) {
        if (!normSymbol(r.special, sm) || seenInsp[r._r]) return;
        add('special', 'check', r.procNo, r.item, '검사기준서의 특별특성 항목 「' + r.item + '」(' + r.special + ') 과 짝이 되는 관리계획서 항목을 찾지 못했습니다.', [ev('insp', r, r.item + ' · ' + r.special)]);
      });
    }

    // 관리계획서 검사 항목 → 검사기준서 (+ 주기·시료수)
    if ((R.insp_item.on || R.insp_cond.on) && cp && insp) {
      var kws = splitList(R.insp_item.keywords);
      var scope = cp.filter(function (c) { return R.insp_item.scope === 'all' || hasWord(c.procName || '', kws); });
      scope.forEach(function (c) {
        var b = inspMatch(c), e0 = ev('cp', c, c.char);
        if (!b || b.sim < charCheck) { if (R.insp_item.on) add('insp_item', 'missing', c.procNo, c.char, '관리계획서 ' + procLabel(c.procNo, c.procName) + ' 의 검사 항목 「' + c.char + '」 이 검사기준서에 없습니다.', [e0]); return; }
        var e = [e0, ev('insp', b.row, b.row.item)];
        if (b.sim < charOk) { if (R.insp_item.on) add('insp_item', 'check', c.procNo, c.char, '관리계획서 「' + c.char + '」 과 검사기준서 「' + b.row.item + '」 이 같은 항목인지 확인해 주세요 (닮음 ' + Math.round(b.sim * 100) + '%).', e); return; }
        if (R.insp_item.on) add('insp_item', 'ok', c.procNo, c.char, '검사 항목 「' + c.char + '」 → 검사기준서 「' + b.row.item + '」', e);
        if (!R.insp_cond.on) return;
        var diffs = [], checks = [];
        [['freq', '주기', normFreq], ['sample', '시료수', normSample]].forEach(function (k) {
          if (!R.insp_cond[k[0]]) return;
          var a = c[k[0]], z = b.row[k[0]];
          if (!a && !z) return;
          if (!a || !z) { checks.push(k[1] + '(' + (a ? '검사기준서' : '관리계획서') + '에 없음: ' + (a || z) + ')'); return; }
          if (k[2](a) !== k[2](z)) diffs.push(k[1] + ' 관리계획서 「' + a + '」 · 검사기준서 「' + z + '」');
        });
        var ec = [ev('cp', c, c.char + ' · 주기 ' + (c.freq || '-') + ' · 시료 ' + (c.sample || '-')), ev('insp', b.row, b.row.item + ' · 주기 ' + (b.row.freq || '-') + ' · 시료 ' + (b.row.sample || '-'))];
        if (diffs.length) add('insp_cond', 'mismatch', c.procNo, c.char, '「' + c.char + '」 ' + diffs.join(' / '), ec);
        else if (checks.length) add('insp_cond', 'check', c.procNo, c.char, '「' + c.char + '」 ' + checks.join(' / '), ec);
        else add('insp_cond', 'ok', c.procNo, c.char, '「' + c.char + '」 주기·시료수 일치 (' + [c.freq, c.sample].filter(Boolean).join(' · ') + ')', ec);
      });
    }

    // 외주 공정 → 외주품 리스트
    if (R.outsourced.on && flow && outs) {
      var okN = R.proc_name.okSim / 100, used = {};
      outsrcFlow.forEach(function (r) {
        var hit = r.procNo && oP[r.procNo] ? { row: oP[r.procNo][0], by: '공정번호' } : null;
        if (!hit) { var b = bestMatch(stripParen(r.procName).replace(/외주/g, ''), outs, ['procName']); if (b && b.sim >= okN) hit = { row: b.row, by: '공정명' }; }
        if (hit) { used[hit.row._r] = 1; add('outsourced', 'ok', r.procNo, r.procName, '외주 공정 ' + procLabel(r.procNo, r.procName) + ' → 외주품 리스트(' + hit.by + ')' + (hit.row.vendor ? ' · 외주처 ' + hit.row.vendor : ''), [ev('flow', r, r.procName), ev('outs', hit.row, procLabel(hit.row.procNo, hit.row.procName))]); }
        else add('outsourced', 'missing', r.procNo, r.procName, 'Flow Chart 의 외주 공정 ' + procLabel(r.procNo, r.procName) + ' 이 외주품 리스트에 없습니다.', [ev('flow', r, procLabel(r.procNo, r.procName) + (r.kind ? ' · ' + r.kind : ''))]);
      });
      outs.forEach(function (o) {
        if (used[o._r]) return;
        var inFlow = o.procNo && fP[o.procNo];
        add('outsourced', 'check', o.procNo, o.procName, '외주품 리스트의 ' + procLabel(o.procNo, o.procName) + (inFlow ? ' 이 Flow Chart 에서는 외주로 표시되어 있지 않습니다.' : ' 과 맞는 Flow Chart 공정을 찾지 못했습니다.'), [ev('outs', o, procLabel(o.procNo, o.procName))].concat(inFlow ? [ev('flow', fP[o.procNo][0], fP[o.procNo][0].procName)] : []));
      });
    }

    var coverage = buildCoverage(docs, fP, cP, pP, iP, oP, flowProcs, excluded);
    return { findings: F, coverage: coverage, summary: summarize(F, S) };
  }

  /* 공정 대조표 — 공정 × 서류. 칸 상태: ok(있음) · missing(있어야 하는데 없음) · extra(Flow 에 없는 공정) · na(대상 아님) · nodoc(서류 없음) */
  function buildCoverage(docs, fP, cP, pP, iP, oP, flowProcs, excluded) {
    var all = uniq(flowProcs.concat(Object.keys(cP), Object.keys(pP), Object.keys(iP), Object.keys(oP))).sort(procCmp);
    var maps = { flow: fP, cp: cP, pfmea: pP, insp: iP, outs: oP };
    var inspHasNo = Object.keys(iP).length > 0;
    var rows = all.map(function (no) {
      var fr = fP[no] && fP[no][0], cells = {}, names = {};
      DOC_IDS.forEach(function (id) {
        var list = maps[id][no] || [];
        names[id] = list[0] ? list[0].procName : '';
        var expected = false;
        if (fr && !excluded(no)) expected = id === 'cp' || id === 'pfmea' || (id === 'outs' && fr.outsourced);
        var st;
        if (!rowsOf(docs, id)) st = 'nodoc';
        else if (list.length) st = id !== 'flow' && !fr ? 'extra' : 'ok';
        else if (id === 'insp' && !inspHasNo) st = 'na';
        else st = expected ? 'missing' : 'na';
        cells[id] = { status: st, count: list.length, rows: list.map(function (r) { return r._r; }) };
      });
      return { procNo: no, name: fr ? fr.procName : (names.cp || names.pfmea || names.insp || names.outs || ''), kind: fr ? fr.kind || '' : '', outsourced: !!(fr && fr.outsourced), excluded: !!(fr && excluded(no)), cells: cells, names: names };
    });
    return { rows: rows, inspHasNo: inspHasNo };
  }
  function summarize(F, S) {
    var c = { ok: 0, mismatch: 0, missing: 0, check: 0 }, byRule = {}, major = 0, minor = 0;
    RULES.forEach(function (r) { byRule[r.id] = { ok: 0, mismatch: 0, missing: 0, check: 0, on: S.rules[r.id].on, sev: S.rules[r.id].sev }; });
    F.forEach(function (f) {
      c[f.verdict]++; byRule[f.rule][f.verdict]++;
      if (f.verdict === 'mismatch' || f.verdict === 'missing') { if (f.sev === 'major') major++; else minor++; }
    });
    var total = F.length, issues = c.mismatch + c.missing;
    var score = total ? Math.round(c.ok / total * 100) : 0;
    var D = S.decision, reasons = [], code;
    if (major >= D.rejectMajor) { code = 'reject'; reasons.push('중대 지적 ' + major + '건 (반려 기준 ' + D.rejectMajor + '건 이상)'); }
    if (minor >= D.rejectMinor) { code = 'reject'; reasons.push('경미 지적 ' + minor + '건 (반려 기준 ' + D.rejectMinor + '건 이상)'); }
    if (!code && (minor || major || c.check)) { code = 'conditional'; if (major) reasons.push('중대 지적 ' + major + '건'); if (minor) reasons.push('경미 지적 ' + minor + '건 — 보완 조건'); if (c.check) reasons.push('확인 필요 ' + c.check + '건 — 확인 후 승인'); }
    if (!code) { code = total ? 'approve' : 'conditional'; reasons.push(total ? '규칙에 걸린 지적 없음' : '판정할 서류가 없습니다'); }
    return { counts: c, total: total, issues: issues, major: major, minor: minor, score: score, byRule: byRule, decision: { code: code, label: DECISION[code], reasons: reasons } };
  }

  /* ── 6. 재제출 비교 ─────────────────────────────────────────── */
  function findingKey(f) { return f.rule + '|' + f.procNo + '|' + normName(f.item); }
  function compareResults(prev, cur) {
    function open(res) { var m = {}; (res ? res.findings : []).forEach(function (f) { if (f.verdict !== 'ok') { var k = findingKey(f); if (!m[k]) m[k] = f; } }); return m; }
    var a = open(prev), b = open(cur), fixed = [], remain = [], added = [];
    Object.keys(a).forEach(function (k) { if (b[k]) remain.push({ prev: a[k], cur: b[k] }); else fixed.push(a[k]); });
    Object.keys(b).forEach(function (k) { if (!a[k]) added.push(b[k]); });
    function byProc(x, y) { var f = x.prev || x, g = y.prev || y; return procCmp(f.procNo, g.procNo) || (f.rule < g.rule ? -1 : f.rule > g.rule ? 1 : 0); }
    fixed.sort(byProc); remain.sort(byProc); added.sort(byProc);
    return { fixed: fixed, remain: remain, added: added, prevScore: prev ? prev.summary.score : null, curScore: cur ? cur.summary.score : null };
  }
  function reviewKey(rv) { return normName(rv.supplier) + '|' + normName(rv.partNo); }
  /* 같은 협력사·품번의 바로 앞 제출(차수가 더 작은 것 중 가장 큰 것) */
  function previousOf(reviews, rv) {
    var k = reviewKey(rv), best = null;
    (reviews || []).forEach(function (x) { if (x.id !== rv.id && reviewKey(x) === k && x.round < rv.round && (!best || x.round > best.round)) best = x; });
    return best;
  }
  function nextRound(reviews, supplier, partNo) {
    var k = normName(supplier) + '|' + normName(partNo), n = 0;
    (reviews || []).forEach(function (x) { if (reviewKey(x) === k) n = Math.max(n, x.round || 0); });
    return n + 1;
  }

  /* ── 7. AI 프롬프트·답 읽기 ─────────────────────────────────
     협력사(외주처 포함) 이름은 「업체1」, 품번은 「품목1」 같은 별칭으로 바꿉니다(기본). 대응표는 이 브라우저에만 둡니다. */
  var ALIAS = { suppliers: '업체', parts: '품목' };
  function assignAliases(map, names, prefix) {
    var out = {}, used = 0, re = new RegExp('^' + prefix + '(\\d+)$');
    Object.keys(map || {}).forEach(function (k) { var m = re.exec(map[k]); if (m) { out[k] = map[k]; used = Math.max(used, +m[1]); } });
    uniq((names || []).map(str).filter(Boolean)).forEach(function (k) { if (!out[k]) out[k] = prefix + (++used); });
    return out;
  }
  function namesOf(reviews) {
    var sup = [], parts = [];
    (reviews || []).forEach(function (rv) {
      sup.push(rv.supplier); parts.push(rv.partNo);
      var o = rv.docs && rv.docs.outs;
      if (o && o.rows) o.rows.forEach(function (r) { if (r.vendor) sup.push(r.vendor); if (r.partNo) parts.push(r.partNo); });
    });
    return { suppliers: sup.map(str).filter(Boolean).sort(), parts: parts.map(str).filter(Boolean).sort() };
  }
  function updateAliases(aliases, reviews) {
    aliases = aliases || {};
    var n = namesOf(reviews);
    return { suppliers: assignAliases(aliases.suppliers, n.suppliers, ALIAS.suppliers), parts: assignAliases(aliases.parts, n.parts, ALIAS.parts) };
  }
  function escRe(s) { return s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&'); }
  function nameRegex(names) {
    var list = names.filter(function (n) { return n.length >= 2; }).sort(function (a, b) { return b.length - a.length; }).map(escRe);
    return list.length ? new RegExp('(^|[^0-9A-Za-z])(' + list.join('|') + ')(?![0-9A-Za-z])', 'g') : null;
  }
  function maskText(text, aliases) {
    var s = text == null ? '' : String(text), all = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { all[n] = m[n]; }); });
    var re = nameRegex(Object.keys(all));
    return re ? s.replace(re, function (x, pre, n) { return pre + all[n]; }) : s;
  }
  function unmaskText(text, aliases) {
    var back = {};
    ['suppliers', 'parts'].forEach(function (k) { var m = (aliases && aliases[k]) || {}; Object.keys(m).forEach(function (n) { back[m[n]] = n; }); });
    return String(text == null ? '' : text).replace(new RegExp('(' + ALIAS.suppliers + '|' + ALIAS.parts + ')(\\d+)(?![0-9개건종가])', 'g'), function (all) { return back[all] || all; });
  }
  function findLeaks(text, aliases) {
    var s = String(text || ''), names = [];
    ['suppliers', 'parts'].forEach(function (k) { names = names.concat(Object.keys((aliases && aliases[k]) || {})); });
    return names.filter(function (n) { var re = nameRegex([n]); return re && re.test(s); });
  }
  function evText(e) { return docLabel(e.doc) + (e.row ? ' ' + e.row + '행' : ''); }
  /* ctx = { review, result, compare(없어도 됨), aliases(null 이면 가리지 않음), maxItems } */
  function buildAiPrompt(ctx) {
    var al = ctx.aliases || null, rv = ctx.review, res = ctx.result, sm = res.summary, L = [];
    function M(s) { return al ? maskText(s, al) : String(s == null ? '' : s); }
    L.push('너는 부품품질팀의 초도품 APQP 승인 검토를 돕는 보조야. 아래 [사실]은 협력사가 낸 APQP 서류(Flow Chart·관리계획서·PFMEA·검사기준서·외주품 리스트)를 도구가 규칙으로 대조한 결과야.');
    L.push('[사실]에 있는 내용과 숫자만 근거로 삼아줘. 숫자를 새로 만들거나 바꾸지 말고, 서류에 없는 원인은 반드시 「추정」이라고 표시해줘. 최종 승인 판단은 검토자가 하니 단정하지 말아줘.');
    if (al) L.push('협력사 이름과 품번은 보안상 「' + ALIAS.suppliers + '1」「' + ALIAS.parts + '1」 같은 별칭으로 바꿔 두었어. 답에서도 별칭을 띄어 쓰지 말고 그대로 써줘.');
    L.push('');
    L.push('[사실]');
    L.push('- 검토 건: ' + M(rv.supplier) + ' · 품번 ' + M(rv.partNo) + (rv.partName ? ' (' + rv.partName + ')' : '') + ' · ' + (rv.stage || '초도품') + ' ' + rv.round + '차 제출' + (rv.submittedAt ? ' (' + rv.submittedAt + ')' : ''));
    L.push('- 제출 서류: ' + DOC_IDS.map(function (id) { var d = rv.docs && rv.docs[id]; return docLabel(id) + (d ? ' ' + d.rows.length + '행' : ' 없음'); }).join(', '));
    L.push('- 판정 ' + sm.total + '건: 적합 ' + sm.counts.ok + ' · 불일치 ' + sm.counts.mismatch + ' · 누락 ' + sm.counts.missing + ' · 확인 필요 ' + sm.counts.check + ' (정합률 ' + sm.score + '%)');
    L.push('- 중대 지적 ' + sm.major + '건 · 경미 지적 ' + sm.minor + '건 → 도구의 승인 의견 제안: ' + sm.decision.label + ' (' + sm.decision.reasons.join('; ') + ')');
    var issues = res.findings.filter(function (f) { return f.verdict !== 'ok'; }), max = ctx.maxItems || 40;
    L.push('- 지적 사항 ' + issues.length + '건' + (issues.length > max ? ' (앞 ' + max + '건만)' : '') + ':');
    issues.slice(0, max).forEach(function (f, i) {
      L.push('  ' + (i + 1) + '. [' + VERDICT[f.verdict] + (f.verdict === 'check' ? '' : '·' + SEV[f.sev]) + '] ' + ruleMeta(f.rule).name + ' — ' + M(f.msg) + ' (근거: ' + f.ev.map(evText).join(', ') + ')');
    });
    if (ctx.compare) {
      var cm = ctx.compare;
      L.push('- 앞 제출과 비교: 고쳐진 지적 ' + cm.fixed.length + '건 · 남은 지적 ' + cm.remain.length + '건 · 새로 생긴 지적 ' + cm.added.length + '건 (정합률 ' + cm.prevScore + '% → ' + cm.curScore + '%)');
      cm.fixed.slice(0, 10).forEach(function (f) { L.push('  · 고쳐짐: ' + M(f.msg)); });
    }
    L.push('');
    L.push('[요청]');
    L.push('APQP 승인 검토 보고서에 넣을 문장을 아래 형식으로만 써줘.');
    L.push('검토 의견: (3~5문장. 전체 정합 상태, 중대 지적의 요지, 승인 의견 제안의 근거)');
    L.push('보완 요청:');
    L.push('- (협력사에 보낼 보완 요청을 지적 사항마다 한 줄씩, 어느 서류의 무엇을 고칠지 구체적으로)');
    L.push('확인 필요:');
    L.push('- (검토자가 현장·협력사에 확인할 것을 한 줄씩)');
    var text = L.join('\n');
    return { prompt: text, leaks: al ? findLeaks(text, al) : [] };
  }
  /* AI 답 → { opinion, requests[], checks[] } (별칭은 원래 이름으로) */
  function parseAiReport(text, aliases) {
    var t = unmaskText(text, aliases), out = { opinion: '', requests: [], checks: [] }, mode = '';
    t.split(/\r?\n/).forEach(function (line) {
      var s = line.trim();
      if (!s) return;
      var m;
      if ((m = /^[#\s*]*검토\s*의견\s*\**\s*[:：]\s*\**\s*(.*)$/.exec(s))) { mode = 'op'; if (m[1].trim()) out.opinion += (out.opinion ? ' ' : '') + m[1].trim(); return; }
      if ((m = /^[#\s*]*보완\s*요청\s*\**\s*[:：]\s*\**\s*(.*)$/.exec(s))) { mode = 'rq'; if (m[1].trim()) out.requests.push(m[1].trim()); return; }
      if ((m = /^[#\s*]*확인\s*필요\s*\**\s*[:：]\s*\**\s*(.*)$/.exec(s))) { mode = 'ck'; if (m[1].trim()) out.checks.push(m[1].trim()); return; }
      var item = s.replace(/^([-*·•]|\d+[.)])\s*/, '').trim();
      if (!item) return;
      if (mode === 'op') out.opinion += (out.opinion ? ' ' : '') + s;
      else if (mode === 'rq') out.requests.push(item);
      else if (mode === 'ck') out.checks.push(item);
      else out.opinion += (out.opinion ? ' ' : '') + s;
    });
    return out;
  }

  /* ── 8. 보고서 (Excel 시트 = 배열의 배열) ──────────────────────
     ctx = { review, result, prev(앞 제출, 없어도 됨), compare, settings, today } */
  function reportSheets(ctx) {
    var rv = ctx.review, res = ctx.result, sm = res.summary, S = cleanSettings(ctx.settings), op = rv.opinion || {};
    var sum = [], list = [], cov = [], rules = [], data = [];
    sum.push(['초도품 APQP 승인 검토 보고서']);
    sum.push(['작성일', ctx.today]);
    sum.push([]);
    sum.push(['협력사', rv.supplier]);
    sum.push(['품번', rv.partNo]);
    sum.push(['품명', rv.partName || '']);
    sum.push(['구분 · 제출 차수', (rv.stage || '초도품') + ' · ' + rv.round + '차']);
    sum.push(['제출일', rv.submittedAt || '']);
    sum.push([]);
    sum.push(['제출 서류', '파일', '시트', '행 수']);
    DOC_IDS.forEach(function (id) { var d = rv.docs && rv.docs[id]; sum.push([docLabel(id), d ? d.fileName : '(제출 안 됨)', d ? d.sheet : '', d ? d.rows.length : 0]); });
    sum.push([]);
    sum.push(['판정 요약', '건수']);
    sum.push(['전체 판정', sm.total]);
    ['ok', 'mismatch', 'missing', 'check'].forEach(function (k) { sum.push([VERDICT[k], sm.counts[k]]); });
    sum.push(['정합률(%)', sm.score]);
    sum.push(['중대 지적', sm.major]);
    sum.push(['경미 지적', sm.minor]);
    sum.push([]);
    sum.push(['규칙별', '사용', '심각도', '적합', '불일치', '누락', '확인 필요']);
    RULES.forEach(function (r) { var b = sm.byRule[r.id]; sum.push([r.name, b.on ? '사용' : '끔', SEV[b.sev], b.ok, b.mismatch, b.missing, b.check]); });
    sum.push([]);
    sum.push(['승인 의견 제안(도구)', sm.decision.label, sm.decision.reasons.join(' / ')]);
    sum.push(['검토자 결정', op.decision ? DECISION[op.decision] : '(미정)']);
    sum.push(['검토자', op.reviewer || '']);
    sum.push(['검토 의견', op.comment || '']);
    if (op.requests) sum.push(['보완 요청', op.requests]);
    if (ctx.compare) {
      sum.push([]);
      sum.push(['앞 제출 비교', (ctx.prev ? ctx.prev.round + '차' : '') + ' → ' + rv.round + '차']);
      sum.push(['정합률(%)', ctx.compare.prevScore + ' → ' + ctx.compare.curScore]);
      sum.push(['고쳐진 지적', ctx.compare.fixed.length]);
      sum.push(['남은 지적', ctx.compare.remain.length]);
      sum.push(['새로 생긴 지적', ctx.compare.added.length]);
    }
    sum.push([]);
    sum.push(['※ 승인 의견은 규칙 판정에서 나온 제안입니다. 최종 승인은 검토자가 결정합니다.']);

    list.push(['No', '판정', '심각도', '규칙', '공정번호', '항목', '내용', '근거']);
    var order = { missing: 0, mismatch: 1, check: 2, ok: 3 };
    res.findings.slice().sort(function (a, b) { return order[a.verdict] - order[b.verdict] || (a.sev === b.sev ? 0 : a.sev === 'major' ? -1 : 1) || procCmp(a.procNo, b.procNo); }).forEach(function (f, i) {
      list.push([i + 1, VERDICT[f.verdict], f.verdict === 'ok' || f.verdict === 'check' ? '' : SEV[f.sev], ruleMeta(f.rule).name, f.procNo, f.item, f.msg, f.ev.map(evText).join(', ')]);
    });

    var mark = { ok: '○', missing: '누락', extra: 'Flow 에 없음', na: '-', nodoc: '(서류 없음)' };
    cov.push(['공정번호', '공정명(Flow Chart)', '공정 구분', '외주'].concat(DOC_IDS.map(docLabel)).concat(['공정명이 다른 서류']));
    res.coverage.rows.forEach(function (r) {
      var diff = ['cp', 'pfmea', 'outs'].filter(function (id) { return r.names[id] && r.name && similarity(r.name, r.names[id]) < S.rules.proc_name.okSim / 100; }).map(function (id) { return docLabel(id) + ': ' + r.names[id]; });
      cov.push([r.procNo, r.name, r.kind + (r.excluded ? ' (대조 제외)' : ''), r.outsourced ? '외주' : ''].concat(DOC_IDS.map(function (id) { var c = r.cells[id]; return c.status === 'ok' || c.status === 'extra' ? mark[c.status] + (c.count > 1 ? ' ' + c.count + '행' : '') : mark[c.status]; })).concat([diff.join(' / ')]));
    });
    cov.push([]);
    cov.push(['표시: ○ 있음 · 누락 = 있어야 하는데 없음 · Flow 에 없음 = Flow Chart 에 없는 공정 · - 대상 아님']);

    rules.push(['규칙', '사용', '심각도', '설명', '기준값']);
    RULES.forEach(function (r) {
      var x = S.rules[r.id];
      rules.push([r.name, x.on ? '사용' : '끔', SEV[x.sev], r.desc, r.params.map(function (p) { var v = x[p.id]; return p.label + ': ' + (p.type === 'bool' ? (v ? '예' : '아니오') : p.type === 'select' ? p.options.filter(function (o) { return o[0] === v; })[0][1] : v); }).join(' / ')]);
    });
    rules.push(['항목 이름 대조', '', '', '같다고 볼 닮음 ' + S.match.charOk + '% · 확인 필요 ' + S.match.charCheck + '%', '']);
    rules.push(['승인 의견 제안', '', '', '중대 지적 ' + S.decision.rejectMajor + '건 이상 또는 경미 지적 ' + S.decision.rejectMinor + '건 이상 → 반려 / 지적·확인 필요가 있으면 조건부 승인 / 없으면 승인', '']);

    data.push(['서류', '행', '공정번호', '공정명', '항목·특성', '특별특성', '주기', '시료수', '관리·검사방법', 'S', 'O', 'D', 'RPN', '외주']);
    DOC_IDS.forEach(function (id) {
      var d = rv.docs && rv.docs[id];
      if (!d) return;
      d.rows.forEach(function (r) {
        data.push([docLabel(id), r._r, r.procNo, r.procName, r.char || r.item || [r.mode].filter(Boolean).join('') || r.kind || r.vendor || '', r.special || '', r.freq || '', r.sample || '',
          r.control || r.measure || r.method || '', r.S == null ? '' : r.S, r.O == null ? '' : r.O, r.D == null ? '' : r.D, r.RPN == null ? '' : r.RPN, id === 'flow' ? (r.outsourced ? '외주' : '') : id === 'outs' ? (r.vendor || '') : '']);
      });
    });
    var out = { '요약': sum, '판정 목록': list, '공정 대조표': cov };
    if (ctx.compare) {
      var cmp = [['구분', '규칙', '공정번호', '항목', '앞 제출 내용', '이번 제출 내용']];
      ctx.compare.fixed.forEach(function (f) { cmp.push(['고쳐짐', ruleMeta(f.rule).name, f.procNo, f.item, VERDICT[f.verdict] + ': ' + f.msg, '']); });
      ctx.compare.remain.forEach(function (x) { cmp.push(['남음', ruleMeta(x.cur.rule).name, x.cur.procNo, x.cur.item, VERDICT[x.prev.verdict] + ': ' + x.prev.msg, VERDICT[x.cur.verdict] + ': ' + x.cur.msg]); });
      ctx.compare.added.forEach(function (f) { cmp.push(['새로 생김', ruleMeta(f.rule).name, f.procNo, f.item, '', VERDICT[f.verdict] + ': ' + f.msg]); });
      out['재제출 비교'] = cmp;
    }
    out['규칙 설정'] = rules;
    out['서류 데이터'] = data;
    return out;
  }

  /* ── 9. 저장 형식 ─────────────────────────────────────────── */
  var SCHEMA_VERSION = 1;
  var DOC_ROW_KEYS = { flow: ['procNo', 'procName', 'kind', 'outsrc', 'note'], cp: ['procNo', 'procName', 'charNo', 'charProd', 'charProc', 'special', 'spec', 'measure', 'sample', 'freq', 'control', 'reaction'],
    pfmea: ['procNo', 'procName', 'char', 'mode', 'effect', 'cause', 'special', 'prevent'], insp: ['procNo', 'procName', 'item', 'special', 'spec', 'method', 'sample', 'freq'], outs: ['procNo', 'procName', 'vendor', 'partNo', 'note'] };
  function cleanDocRow(id, r) {
    var o = { _r: Math.max(0, parseInt(r && r._r, 10) || 0) };
    DOC_ROW_KEYS[id].forEach(function (k) { if (r[k] != null && r[k] !== '') o[k] = str(r[k]); });
    if (o.procNo) o.procNo = normProcNo(o.procNo);
    if (id === 'cp') o.char = [o.charProd, o.charProc].filter(Boolean).join(' / ');
    if (id === 'pfmea') { ['S', 'O', 'D', 'RPN'].forEach(function (k) { var n = parseFloat(r[k]); o[k] = isFinite(n) ? n : null; }); if (r.rpnCalc) o.rpnCalc = true; }
    if (id === 'flow') o.outsourced = !!r.outsourced;
    return o;
  }
  function cleanReview(r) {
    if (!r || typeof r !== 'object' || !str(r.id)) return null;
    var o = { id: str(r.id), supplier: str(r.supplier), partNo: str(r.partNo), partName: str(r.partName), stage: str(r.stage) || '초도품', round: Math.max(1, parseInt(r.round, 10) || 1),
      submittedAt: parseDate(r.submittedAt), createdAt: str(r.createdAt), docs: {}, opinion: {}, aiText: r.aiText == null ? '' : String(r.aiText) };
    if (r.docs && typeof r.docs === 'object') DOC_IDS.forEach(function (id) {
      var d = r.docs[id];
      if (!d || !Array.isArray(d.rows)) return;
      o.docs[id] = { fileName: str(d.fileName), sheet: str(d.sheet), at: str(d.at), rows: d.rows.filter(function (x) { return x && typeof x === 'object'; }).map(function (x) { return cleanDocRow(id, x); }) };
    });
    var op = r.opinion || {};
    o.opinion = { decision: DECISION[op.decision] ? op.decision : '', reviewer: str(op.reviewer), comment: op.comment == null ? '' : String(op.comment), requests: op.requests == null ? '' : String(op.requests), at: str(op.at) };
    return o;
  }
  function emptyDb() { return { schemaVersion: SCHEMA_VERSION, reviews: [], seq: 0, current: '', settings: defaultSettings(), mask: true, aliases: { suppliers: {}, parts: {} }, mapping: {} }; }
  function restoreDb(o) {
    var db = emptyDb();
    if (!o || typeof o !== 'object') return db;
    if (Array.isArray(o.reviews)) db.reviews = o.reviews.map(cleanReview).filter(Boolean);
    if (typeof o.seq === 'number') db.seq = o.seq;
    db.current = db.reviews.some(function (r) { return r.id === o.current; }) ? o.current : (db.reviews[db.reviews.length - 1] || {}).id || '';
    db.settings = cleanSettings(o.settings);
    db.mask = o.mask !== false;
    if (o.aliases && typeof o.aliases === 'object') db.aliases = { suppliers: assignAliases(o.aliases.suppliers, [], ALIAS.suppliers), parts: assignAliases(o.aliases.parts, [], ALIAS.parts) };
    if (o.mapping && typeof o.mapping === 'object') db.mapping = o.mapping;
    if (o._sample) db._sample = true;
    return db;
  }
  function reviewTitle(rv) { return rv ? rv.supplier + ' · ' + rv.partNo + ' · ' + rv.round + '차' + (rv.submittedAt ? ' (' + rv.submittedAt + ')' : '') : ''; }

  return {
    // 정규화
    str: str, parseDate: parseDate, normProcNo: normProcNo, procCmp: procCmp, normName: normName, similarity: similarity, normFreq: normFreq, normSample: normSample,
    parseSymbolMap: parseSymbolMap, normSymbol: normSymbol, DEFAULT_SYMBOLS: DEFAULT_SYMBOLS,
    // 서류
    DOC_TYPES: DOC_TYPES, DOC_IDS: DOC_IDS, docType: docType, docLabel: docLabel, normHead: normHead, matchColumns: matchColumns, detectColumns: detectColumns,
    headerLabels: headerLabels, fillMerges: fillMerges, guessDocType: guessDocType, parseDoc: parseDoc,
    // PDF
    itemsToLines: itemsToLines, pdfTextItems: pdfTextItems, pdfLinesToAoa: pdfLinesToAoa,
    // 규칙·검토
    VERDICT: VERDICT, SEV: SEV, RULES: RULES, MATCH_PARAMS: MATCH_PARAMS, DECISION_PARAMS: DECISION_PARAMS, DECISION: DECISION, ruleMeta: ruleMeta,
    defaultSettings: defaultSettings, cleanSettings: cleanSettings, runReview: runReview, evText: evText,
    // 비교
    findingKey: findingKey, compareResults: compareResults, previousOf: previousOf, nextRound: nextRound, reviewKey: reviewKey,
    // AI
    ALIAS: ALIAS, assignAliases: assignAliases, updateAliases: updateAliases, maskText: maskText, unmaskText: unmaskText, findLeaks: findLeaks,
    buildAiPrompt: buildAiPrompt, parseAiReport: parseAiReport,
    // 보고서·저장
    reportSheets: reportSheets, emptyDb: emptyDb, restoreDb: restoreDb, cleanReview: cleanReview, reviewTitle: reviewTitle, SCHEMA_VERSION: SCHEMA_VERSION
  };
});
