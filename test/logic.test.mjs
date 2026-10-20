// 실행: node test/logic.test.mjs   (의존성 없음 — vendor 의 SheetJS 만 씀)
// 실제 APQP 서류는 아직 받지 않았습니다. 여기서는 가상 표와 예시 데이터(협력사 B · SMP-2041)로 검사합니다.
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import { createRequire } from 'node:module';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');
const E = require('../js/ai-endpoint.js');
const XLSX = require('../vendor/xlsx.full.min.js');

let passed = 0;
function test(name, fn) {
  try { fn(); passed++; console.log('  ok  ' + name); }
  catch (e) { console.error('  FAIL ' + name + '\n       ' + e.message); process.exitCode = 1; }
}
const clone = o => JSON.parse(JSON.stringify(o));
/* 예시 통합 문서 → { 서류id: { rows } } (앱의 예시 불러오기와 같은 길) */
function sampleDocs(round) {
  const docs = {};
  Sample.workbook(round).forEach(sh => {
    const aoa = L.fillMerges(clone(sh.aoa), sh.merges);
    const g = L.guessDocType(aoa, sh.name, Sample.fileName(round));
    if (g.type) docs[g.type] = { rows: L.parseDoc(aoa, g.type).rows, fileName: Sample.fileName(round), sheet: sh.name };
  });
  return docs;
}
const issues = res => res.findings.filter(f => f.verdict !== 'ok');
const sig = f => [f.rule, f.verdict, f.procNo].join(' ');

console.log('정규화');
test('공정번호 — OP10·P-10·공정 10·010·10.1 · 숫자 10', () => {
  assert.equal(L.normProcNo('OP10'), '10');
  assert.equal(L.normProcNo('op-10'), '10');
  assert.equal(L.normProcNo('P10'), '10');
  assert.equal(L.normProcNo('공정 10'), '10');
  assert.equal(L.normProcNo('010'), '10');
  assert.equal(L.normProcNo('10.1'), '10-1');
  assert.equal(L.normProcNo(10), '10');
  assert.equal(L.normProcNo(''), '');
  assert.deepEqual(['100', '20', '10-1', '10', '5'].sort(L.procCmp), ['5', '10', '10-1', '20', '100']);
});
test('이름 닮음 — 같음 1 · 괄호 속 무시 1 · 품는 이름 0.5~1 · 다른 이름 낮음', () => {
  assert.equal(L.similarity('출하검사', '출하 검사'), 1);
  assert.equal(L.similarity('용접(CO2)', '용접'), 1);
  assert.equal(L.similarity('절단', '레이저 절단'), 0.7);
  assert.ok(L.similarity('볼트 체결 토크', '부품 누락 여부') < 0.4);
  assert.equal(L.similarity('', '용접'), 0);
});
test('주기 — 「매 로트」=「매 LOT」=「1회/로트」 · 「1회/일」≠「매 로트」 · 「2회/일」≠「1회/일」', () => {
  assert.equal(L.normFreq('매 로트'), L.normFreq('매 LOT'));
  assert.equal(L.normFreq('1회/로트'), L.normFreq('로트당'));
  assert.equal(L.normFreq('1회/교대'), L.normFreq('교대당 1회'));
  assert.equal(L.normFreq('매일'), L.normFreq('1회/일'));
  assert.notEqual(L.normFreq('1회/일'), L.normFreq('매 로트'));
  assert.notEqual(L.normFreq('2회/일'), L.normFreq('1회/일'));
  assert.equal(L.normFreq('전수'), L.normFreq('100%'));
  assert.equal(L.normFreq('초·중·종물'), L.normFreq('초중종'));
});
test('시료수 — 「5개」=「5 EA」=「n=5」 · 「전수」=「100%」 · 「3개」≠「5개」', () => {
  assert.equal(L.normSample('5개'), '5');
  assert.equal(L.normSample('5 EA'), '5');
  assert.equal(L.normSample('n=5'), '5');
  assert.equal(L.normSample('전수'), L.normSample('100%'));
  assert.notEqual(L.normSample('3개'), L.normSample('5개'));
});
test('특별특성 기호 대응표 — ◆=CC · ◇=SC · 모르는 기호는 그대로 · 「-」는 없음', () => {
  const m = L.parseSymbolMap(L.DEFAULT_SYMBOLS);
  assert.equal(L.normSymbol('◆', m), 'CC');
  assert.equal(L.normSymbol('cc', m), 'CC');
  assert.equal(L.normSymbol('◇', m), 'SC');
  assert.equal(L.normSymbol('HS', m), 'HS');
  assert.equal(L.normSymbol('-', m), '');
});

console.log('서류 읽기');
test('관리계획서 두 줄 머리(특성 > 번호·제품·공정 / 방법 > …) + 병합 칸', () => {
  const sh = Sample.cp(1), aoa = L.fillMerges(clone(sh.aoa), sh.merges);
  const d = L.detectColumns(aoa, 'cp');
  assert.equal(d.headerRow, 4); assert.equal(d.headerRows, 2);
  assert.deepEqual(d.missing, []);
  assert.equal(d.map.charProd, 4); assert.equal(d.map.charProc, 5); assert.equal(d.map.sample, 9); assert.equal(d.map.freq, 10); assert.equal(d.map.control, 11);
  const p = L.parseDoc(aoa, 'cp');
  assert.equal(p.rows.length, 14);
  const r = p.rows.find(x => x.char === '절단 출력·속도');
  assert.equal(r.procNo, '20'); assert.equal(r.procName, '레이저 절단'); assert.equal(r._r, 10);
});
test('병합 정보 없는 두 줄 머리(CSV) — 묶음 머리를 오른쪽으로 이어 쓰고 공정번호·공정명은 위 행 값을 이어 씀', () => {
  const t = fs.readFileSync(path.join(root, 'samples/예시_관리계획서_병합없음.csv'), 'utf8').replace(/^﻿/, '');
  const wb = XLSX.read(t, { type: 'string' });
  const aoa = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], { header: 1, raw: true, defval: null });
  assert.equal(L.guessDocType(aoa, '', '예시_관리계획서_병합없음.csv').type, 'cp');
  const p = L.parseDoc(aoa, 'cp');
  assert.equal(p.det.headers[5], '특성 공정');
  assert.equal(p.rows.length, 14);
  assert.equal(p.rows[1].procNo, '10'); assert.equal(p.rows[1].procName, '소재 입고검사');
  assert.equal(p.rows[3].charProc, '절단 출력·속도');
});
test('제목 줄 아래 머리행 · 「No」 일련번호 열보다 「공정번호」가 먼저 · 첫 자료 행을 두 줄 머리로 먹지 않음', () => {
  const sh = Sample.outs(2), p = L.parseDoc(sh.aoa, 'outs');
  assert.equal(p.det.headerRow, 3); assert.equal(p.det.headerRows, 1);
  assert.equal(p.det.map.procNo, 1);
  assert.equal(p.rows.length, 2);
  assert.equal(p.rows[1].vendor, '외주처 나');
});
test('서류 종류 추정 — 시트 이름으로, 이름이 없어도 열로', () => {
  Sample.workbook(1).forEach(sh => {
    const aoa = L.fillMerges(clone(sh.aoa), sh.merges);
    assert.equal(L.guessDocType(aoa, sh.name, '').type, Sample.SHEET_TYPE[sh.name] || '', sh.name);
    if (Sample.SHEET_TYPE[sh.name]) assert.equal(L.guessDocType(aoa, 'Sheet1', 'a.xlsx').type, Sample.SHEET_TYPE[sh.name], sh.name + ' (열로만)');
  });
});
test('PFMEA — RPN 이 비면 S×O×D · 합계 줄·빈 줄 건너뜀 · Flow Chart 외주 표시 3가지', () => {
  const aoa = [['공정번호', '공정명', '고장형태', 'S', 'O', 'D', 'RPN'], ['OP10', '절단', '길이 불량', 5, 4, 6, null], [], ['합계', null, null, null, null, null, null], [20, '용접', '균열', '9', 2, 2, 36]];
  const p = L.parseDoc(aoa, 'pfmea');
  assert.equal(p.rows.length, 2);
  assert.equal(p.rows[0].RPN, 120); assert.equal(p.rows[0].rpnCalc, true); assert.equal(p.rows[0].procNo, '10');
  assert.equal(p.rows[1].S, 9);
  const f = L.parseDoc([['공정번호', '공정명', '공정 구분', '외주 여부'], [10, '절단', '가공', ''], [20, '도장', '가공', 'Y'], [30, '열처리', '외주', ''], [40, '도금(외주)', '', '']], 'flow');
  assert.deepEqual(f.rows.map(r => r.outsourced), [false, true, true, true]);
});
test('필수 열이 없으면 읽지 않고 · 사용자가 열을 고르면 다시 읽음 · 머리행을 직접 고름', () => {
  const aoa = [['제목'], ['번호', '이름'], [10, '절단'], [20, '용접']];
  const d = L.detectColumns(aoa, 'flow');
  assert.equal(d.headerRow, -1);
  assert.equal(L.parseDoc(aoa, 'flow').rows.length, 0);
  const p = L.parseDoc(aoa, 'flow', { headerRow: 1, overrides: { procNo: 0, procName: 1 } });
  assert.deepEqual(p.det.missing, []);
  assert.deepEqual(p.rows.map(r => r.procNo + r.procName), ['10절단', '20용접']);
});
test('PDF 글자 조각 → 표 — 머리행 기준 칸 맞춤 · 다음 쪽 같은 머리행은 한 번만', () => {
  const it = (str, x, y, page) => ({ str, x, y, w: str.length * 6, page: page || 1 });
  const items = [it('검사기준서', 40, 30), it('검사항목', 40, 80), it('특별특성', 160, 80), it('시료수', 260, 80), it('검사주기', 340, 80),
    it('용접', 40, 100), it('각장', 66, 100), it('◇', 170, 100), it('5 EA', 262, 100), it('1회/교대', 342, 100),
    it('검사항목', 40, 80, 2), it('특별특성', 160, 80, 2), it('시료수', 260, 80, 2), it('검사주기', 340, 80, 2),
    it('외관', 40, 100, 2), it('5 EA', 262, 100, 2), it('매 LOT', 342, 100, 2)];
  const aoa = L.pdfLinesToAoa(L.itemsToLines(items));
  assert.deepEqual(aoa, [['검사기준서'], ['검사항목', '특별특성', '시료수', '검사주기'], ['용접 각장', '◇', '5 EA', '1회/교대'], ['외관', null, '5 EA', '매 LOT']]);
  const p = L.parseDoc(aoa, 'insp');
  assert.equal(p.rows.length, 2); assert.equal(p.rows[0].special, '◇');
});

console.log('정합성 규칙 — 예시 1차 제출(결함 5가지를 심음)');
const S0 = L.defaultSettings();
const R1 = L.runReview(sampleDocs(1), S0), R2 = L.runReview(sampleDocs(2), S0);
test('심어 둔 결함 5가지 + 공정명 확인 필요 1건이 정확히 잡힘(그 밖의 지적 없음)', () => {
  assert.deepEqual(issues(R1).map(sig).sort(), [
    'insp_cond mismatch 80', 'outsourced missing 60', 'pfmea_high missing 70', 'proc_name check 20', 'proc_presence missing 30', 'special mismatch 50'
  ].sort());
});
test('지적의 근거 — 어느 서류 몇 행', () => {
  const f = R1.findings.find(x => x.rule === 'pfmea_high' && x.verdict === 'missing');
  assert.equal(f.ev[0].doc, 'pfmea'); assert.ok(f.ev[0].row > 4);
  const sp = R1.findings.find(x => x.rule === 'special' && x.verdict !== 'ok');
  assert.deepEqual(sp.ev.map(e => e.doc), ['cp', 'insp']);
  assert.equal(sp.ev[1].row, 8);
  assert.match(sp.msg, /◇\(SC\).*◆\(CC\)/);
});
test('S≥9 항목(용접 강도, RPN 81)은 RPN 기준 아래여도 대조하고, 관리계획서에 관리방법과 함께 있어 적합', () => {
  const f = R1.findings.find(x => x.rule === 'pfmea_high' && x.procNo === '40');
  assert.equal(f.verdict, 'ok'); assert.match(f.msg, /S 9/);
});
test('검사 주기·시료수 — 「매 로트」/「매 LOT」, 「5개」/「5 EA」는 같게 보고 「1회/일」만 불일치', () => {
  const c = R1.findings.filter(x => x.rule === 'insp_cond');
  assert.equal(c.filter(x => x.verdict === 'ok').length, 5);
  assert.match(c.find(x => x.verdict === 'mismatch').msg, /주기 관리계획서 「매 로트」 · 검사기준서 「1회\/일」/);
});
test('이동·보관 공정(15·55·90)은 공정번호 대조에서 빠짐 · 외주 공정 45 는 리스트에 있어 적합', () => {
  assert.ok(!R1.findings.some(x => x.rule === 'proc_presence' && ['15', '55', '90'].includes(x.procNo)));
  assert.equal(R1.findings.find(x => x.rule === 'outsourced' && x.procNo === '45').verdict, 'ok');
});
test('점수·승인 의견 제안 — 1차 중대 4건 → 반려 · 2차 확인 필요 1건 → 조건부 승인', () => {
  const s = R1.summary;
  assert.equal(s.major, 4); assert.equal(s.minor, 1); assert.equal(s.counts.check, 1);
  assert.equal(s.decision.code, 'reject');
  assert.equal(s.score, Math.round(s.counts.ok / s.total * 100));
  assert.deepEqual(issues(R2).map(sig), ['proc_name check 20']);
  assert.equal(R2.summary.decision.code, 'conditional');
});
test('공정 대조표 — 30 PFMEA 누락 · 60 외주 리스트 누락 · 15 대조 제외 · 검사기준서는 공정번호가 있는 곳만', () => {
  const row = no => R1.coverage.rows.find(r => r.procNo === no);
  assert.equal(row('30').cells.pfmea.status, 'missing');
  assert.equal(row('30').cells.cp.status, 'ok');
  assert.equal(row('60').cells.outs.status, 'missing');
  assert.equal(row('45').cells.outs.status, 'ok');
  assert.equal(row('15').excluded, true); assert.equal(row('15').cells.pfmea.status, 'na');
  assert.equal(row('50').cells.insp.count, 3);
  assert.equal(R1.coverage.rows.length, 12);
});
test('재제출 비교 — 고쳐짐 5 · 남음 1(공정명) · 새로 생김 0', () => {
  const c = L.compareResults(R1, R2);
  assert.equal(c.fixed.length, 5); assert.equal(c.remain.length, 1); assert.equal(c.added.length, 0);
  assert.equal(c.remain[0].cur.rule, 'proc_name');
  assert.equal(c.prevScore, R1.summary.score); assert.equal(c.curScore, R2.summary.score);
  const back = L.compareResults(R2, R1);
  assert.equal(back.added.length, 5);
});

test('재제출 비교 키에 공정번호 포함 — 같은 규칙·같은 서류라도 다른 공정이면 고쳐짐 1 + 새로 생김 1', () => {
  const mk = no => ({ findings: [{ rule: 'proc_presence', verdict: 'missing', procNo: no, item: 'PFMEA', msg: no }], summary: { score: 0 } });
  const c = L.compareResults(mk('30'), mk('35'));
  assert.equal(c.fixed.length, 1); assert.equal(c.added.length, 1); assert.equal(c.remain.length, 0);
});
test('PFMEA 고위험 항목이 관리계획서에 있어도 관리방법·측정방법이 비면 불일치', () => {
  const docs = sampleDocs(2), c = docs.cp.rows.find(r => r.char === '용접 강도(균열)');
  c.control = ''; c.measure = '';
  const f = L.runReview(docs, S0).findings.find(x => x.rule === 'pfmea_high' && x.procNo === '40');
  assert.equal(f.verdict, 'mismatch'); assert.match(f.msg, /관리방법이 비어/);
});
test('외주 공정 — 이름이 달라도 공정번호가 같으면 외주품 리스트에 있는 것으로 봄', () => {
  const docs = sampleDocs(2);
  docs.outs.rows.find(r => r.procNo === '60').procName = '분체 코팅 외주';
  assert.equal(L.runReview(docs, S0).findings.find(f => f.rule === 'outsourced' && f.procNo === '60').verdict, 'ok');
});

console.log('규칙 설정');
test('규칙 끄기 · 심각도 바꾸기 → 판정과 승인 의견이 따라 바뀜', () => {
  const s = L.cleanSettings(clone(S0));
  s.rules.proc_presence.on = false; s.rules.special.on = false; s.rules.pfmea_high.on = false; s.rules.outsourced.sev = 'minor';
  const r = L.runReview(sampleDocs(1), s);
  assert.ok(!r.findings.some(f => f.rule === 'proc_presence' || f.rule === 'special' || f.rule === 'pfmea_high'));
  assert.equal(r.summary.major, 0); assert.equal(r.summary.minor, 2);
  assert.equal(r.summary.decision.code, 'conditional');
  s.decision.rejectMinor = 2;
  assert.equal(L.runReview(sampleDocs(1), s).summary.decision.code, 'reject');
});
test('PFMEA 기준값 — RPN 150 이면 공정 70(RPN 140)은 대상 아님 · S 10 이면 용접 강도(S 9)도 대상 아님', () => {
  const s = clone(S0); s.rules.pfmea_high.rpnMin = 150; s.rules.pfmea_high.sMin = 10;
  const r = L.runReview(sampleDocs(1), L.cleanSettings(s));
  assert.equal(r.findings.filter(f => f.rule === 'pfmea_high').length, 0);
});
test('검사 항목 대상 「모든 항목」 — 검사 공정이 아닌 항목(경도·도막 두께 등)도 검사기준서에서 찾음', () => {
  const s = clone(S0); s.rules.insp_item.scope = 'all';
  const r = L.runReview(sampleDocs(2), L.cleanSettings(s));
  const miss = r.findings.filter(f => f.rule === 'insp_item' && f.verdict === 'missing').map(f => f.item);
  assert.ok(miss.includes('경도') && miss.includes('도막 두께'));
});
test('특별특성 — 관리계획서에 기호가 없는데 검사기준서에 있으면 불일치 · 기호 대응표를 바꾸면 ◆=◇ 도 같게 볼 수 있음', () => {
  const docs = sampleDocs(2);
  docs.cp.rows.find(r => r.char === '용접 각장').special = '';
  assert.ok(L.runReview(docs, S0).findings.some(f => f.rule === 'special' && f.verdict === 'mismatch' && f.item === '용접 각장'));
  const s = clone(S0); s.rules.special.symbols = '◆=특별, ◇=특별';
  const r = L.runReview(sampleDocs(1), L.cleanSettings(s));
  assert.equal(r.findings.find(f => f.rule === 'special' && f.item === '용접 각장').verdict, 'ok');
});
test('필수 서류 — 없는 서류는 누락 · 외주품 리스트는 외주 공정이 있을 때만 필수', () => {
  const docs = sampleDocs(1); delete docs.pfmea; delete docs.outs;
  const r = L.runReview(docs, S0);
  const d = r.findings.filter(f => f.rule === 'docs' && f.verdict === 'missing').map(f => f.item).sort();
  assert.deepEqual(d, ['PFMEA', '외주품 리스트'].sort());
  docs.flow.rows.forEach(x => { x.outsourced = false; });
  assert.deepEqual(L.runReview(docs, S0).findings.filter(f => f.rule === 'docs' && f.verdict === 'missing').map(f => f.item), ['PFMEA']);
});
test('관리계획서에만 있는 공정 · 외주 리스트에만 있는 공정 → 확인 필요', () => {
  const docs = sampleDocs(2);
  docs.cp.rows.push({ _r: 99, procNo: '75', procName: '세척', char: '청정도', charProd: '청정도' });
  docs.outs.rows.push({ _r: 9, procNo: '', procName: '도금', vendor: '외주처 다' });
  const r = L.runReview(docs, S0);
  assert.ok(r.findings.some(f => f.rule === 'proc_presence' && f.verdict === 'check' && f.procNo === '75'));
  assert.ok(r.findings.some(f => f.rule === 'outsourced' && f.verdict === 'check' && f.item === '도금'));
  assert.equal(r.coverage.rows.find(x => x.procNo === '75').cells.cp.status, 'extra');
});
test('설정 정리 — 범위 밖 값·뒤집힌 기준은 바로잡음', () => {
  const s = L.cleanSettings({ rules: { proc_name: { okSim: 70, checkSim: 90 }, pfmea_high: { rpnMin: -5, sMin: 99 }, docs: { sev: 'nope', on: false } }, match: { charOk: 50, charCheck: 70 }, decision: { rejectMajor: 0 } });
  assert.equal(s.rules.proc_name.checkSim, 69);
  assert.equal(s.rules.pfmea_high.rpnMin, 100); assert.equal(s.rules.pfmea_high.sMin, 9);
  assert.equal(s.rules.docs.sev, 'major'); assert.equal(s.rules.docs.on, false);
  assert.equal(s.match.charCheck, 49);
  assert.equal(s.decision.rejectMajor, 1);
});

console.log('AI 프롬프트');
const reviews = [L.cleanReview({ id: 'r1', supplier: '협력사 B', partNo: 'SMP-2041', round: 1, docs: sampleDocs(1) }), L.cleanReview({ id: 'r2', supplier: '협력사 B', partNo: 'SMP-2041', round: 2, docs: sampleDocs(2) })];
const AL = L.updateAliases({}, reviews);
test('협력사·외주처·품번을 별칭으로 — 프롬프트에 실제 이름이 남지 않음', () => {
  assert.deepEqual(Object.keys(AL.suppliers).sort(), ['외주처 가', '외주처 나', '협력사 B']);
  const info = L.buildAiPrompt({ review: reviews[0], result: R1, compare: null, aliases: AL });
  assert.ok(!/협력사 B|SMP-2041|외주처 [가나]/.test(info.prompt), '이름이 남음');
  assert.deepEqual(info.leaks, []);
  assert.match(info.prompt, /반려/);
  assert.match(info.prompt, /검토 의견:/);
  const plain = L.buildAiPrompt({ review: reviews[0], result: R1, aliases: null });
  assert.match(plain.prompt, /협력사 B/);
});
test('AI 답 읽기 — 검토 의견·보완 요청·확인 필요, 별칭은 되돌리고 「품목 3개」는 그대로', () => {
  const a = AL.suppliers['협력사 B'], p = AL.parts['SMP-2041'];
  const ans = ['**검토 의견:** ' + a + ' 의 ' + p + ' 서류는 중대 지적 4건이 있어', '반려를 제안합니다.', '', '보완 요청:', '- PFMEA 에 공정 30 벤딩을 더해 주세요.', '2) 품목 3개의 외주 리스트를 고쳐 주세요.', '확인 필요:', '- 공정 20 공정명'].join('\n');
  const r = L.parseAiReport(ans, AL);
  assert.equal(r.opinion, '협력사 B 의 SMP-2041 서류는 중대 지적 4건이 있어 반려를 제안합니다.');
  assert.deepEqual(r.requests, ['PFMEA 에 공정 30 벤딩을 더해 주세요.', '품목 3개의 외주 리스트를 고쳐 주세요.']);
  assert.deepEqual(r.checks, ['공정 20 공정명']);
});
test('AI 연결 — 주소 정리·키 없을 때 머리글 없음(data10-07 과 같은 파일)', () => {
  assert.equal(E.normalizeBaseUrl(' https://llm.example.com/v1/chat/completions/ '), 'https://llm.example.com/v1');
  assert.equal(E.STORE_KEY, 'data10-02.ai');
  assert.equal(E.buildChatRequest({ baseUrl: 'https://x/v1', model: 'm' }, []).init.headers.Authorization, undefined);
});

console.log('보고서·저장');
test('보고서 시트 — 요약·판정 목록·공정 대조표·재제출 비교·규칙 설정·서류 데이터', () => {
  const rv = clone(reviews[1]); rv.opinion = { decision: 'conditional', reviewer: '검토자', comment: '의견', requests: '요청' };
  const sh = L.reportSheets({ review: rv, result: R2, prev: reviews[0], compare: L.compareResults(R1, R2), settings: S0, today: '2026-10-20' });
  assert.deepEqual(Object.keys(sh), ['요약', '판정 목록', '공정 대조표', '재제출 비교', '규칙 설정', '서류 데이터']);
  assert.equal(sh['판정 목록'].length, R2.findings.length + 1);
  assert.equal(sh['재제출 비교'].filter(r => r[0] === '고쳐짐').length, 5);
  assert.ok(sh['요약'].some(r => r[0] === '검토자 결정' && r[1] === '조건부 승인'));
  assert.ok(sh['요약'].some(r => r[0] === '승인 의견 제안(도구)' && r[1] === '조건부 승인'));
  const total = L.DOC_IDS.reduce((n, id) => n + rv.docs[id].rows.length, 0);
  assert.equal(sh['서류 데이터'].length, total + 1);
  const noCmp = L.reportSheets({ review: reviews[0], result: R1, settings: S0, today: '2026-10-20' });
  assert.ok(!noCmp['재제출 비교']);
});
test('앞 제출 찾기 · 다음 차수 — 띄어쓰기·대소문자 다른 이름도 같은 품번', () => {
  const list = [{ id: 'a', supplier: '협력사 B', partNo: 'SMP-2041', round: 1 }, { id: 'b', supplier: '협력사B', partNo: 'smp-2041', round: 3 }, { id: 'c', supplier: '협력사 B', partNo: 'SMP-2041', round: 2 }, { id: 'd', supplier: '협력사 C', partNo: 'SMP-2041', round: 1 }];
  assert.equal(L.previousOf(list, list[1]).id, 'c');
  assert.equal(L.previousOf(list, list[0]), null);
  assert.equal(L.nextRound(list, '협력사 B', 'SMP-2041'), 4);
  assert.equal(L.nextRound(list, '협력사 Z', 'X'), 1);
});
test('저장 형식 되살리기 — 이상한 값은 버리고 고침', () => {
  const db = L.restoreDb({
    reviews: [{ id: 'r1', supplier: ' 협력사 B ', partNo: 'SMP-2041', round: 0, submittedAt: '2026.10.6', docs: { flow: { rows: [{ _r: '5', procNo: 'OP10', procName: '절단', outsourced: 1 }, null] }, xx: { rows: [] }, cp: { rows: 'x' } }, opinion: { decision: '아무거나' } }, null, { supplier: '아이디 없음' }],
    current: 'zz', settings: { rules: { docs: { on: false } } }, mask: false, aliases: { suppliers: { 'A': '업체3', 'B': '이상한값' } }
  });
  assert.equal(db.reviews.length, 1);
  const r = db.reviews[0];
  assert.equal(r.supplier, '협력사 B'); assert.equal(r.round, 1); assert.equal(r.submittedAt, '2026-10-06');
  assert.deepEqual(Object.keys(r.docs), ['flow']);
  assert.deepEqual(r.docs.flow.rows, [{ _r: 5, procNo: '10', procName: '절단', outsourced: true }]);
  assert.equal(r.opinion.decision, '');
  assert.equal(db.current, 'r1');
  assert.equal(db.settings.rules.docs.on, false);
  assert.equal(db.mask, false);
  assert.deepEqual(db.aliases.suppliers, { A: '업체3' });
});
test('예시 파일(samples/*.xlsx)이 있고 다시 읽으면 예시 원본과 같음', () => {
  for (const round of [1, 2]) {
    const book = XLSX.read(fs.readFileSync(path.join(root, 'samples', Sample.fileName(round))), { type: 'buffer' });
    const docs = {};
    book.SheetNames.forEach(n => {
      const ws = book.Sheets[n];
      const aoa = L.fillMerges(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }), ws['!merges']);
      const g = L.guessDocType(aoa, n, Sample.fileName(round));
      if (g.type) docs[g.type] = L.parseDoc(aoa, g.type).rows;
    });
    const want = sampleDocs(round);
    L.DOC_IDS.forEach(id => assert.deepEqual(docs[id], want[id].rows, round + '차 ' + id));
  }
  assert.ok(fs.existsSync(path.join(root, 'samples/예시_APQP_SMP-2041_검사기준서_2차.pdf')));
});

console.log('\n' + passed + '개 통과' + (process.exitCode ? ' — 실패 있음' : ''));
