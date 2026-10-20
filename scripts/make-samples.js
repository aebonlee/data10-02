// 가상 샘플 파일 만들기: node scripts/make-samples.js
// js/sample-data.js 의 예시(협력사 B · 품번 SMP-2041, 전부 지어낸 값)를 Excel 통합 문서 두 개(1차 제출 · 2차 재제출)로 씁니다.
// 방금 쓴 파일을 앱과 같은 방식(병합 칸 채우기 → 서류 종류 추정 → 머리행·열 찾기)으로 다시 읽어, 원본 표를 바로 읽은 결과와 같은지 확인합니다.
// 검사기준서 PDF 는 python3 scripts/make-sample-pdf.py 로 따로 만듭니다(reportlab 필요).
const fs = require('fs');
const path = require('path');
const XLSX = require('../vendor/xlsx.full.min.js');
const L = require('../js/logic.js');
const Sample = require('../js/sample-data.js');

const out = path.join(__dirname, '..', 'samples');
fs.mkdirSync(out, { recursive: true });

function direct(round) {
  const docs = {};
  Sample.workbook(round).forEach(sh => {
    const type = Sample.SHEET_TYPE[sh.name];
    if (!type) return;
    docs[type] = L.parseDoc(L.fillMerges(JSON.parse(JSON.stringify(sh.aoa)), sh.merges), type).rows;
  });
  return docs;
}

for (const round of [1, 2]) {
  const wb = XLSX.utils.book_new();
  Sample.workbook(round).forEach(sh => {
    const ws = XLSX.utils.aoa_to_sheet(sh.aoa);
    ws['!merges'] = sh.merges;
    ws['!cols'] = sh.cols.map(w => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, sh.name);
  });
  const file = path.join(out, Sample.fileName(round));
  fs.writeFileSync(file, XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' }));

  // 다시 읽어 확인
  const book = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
  const want = direct(round), got = {};
  book.SheetNames.forEach(n => {
    const ws = book.Sheets[n];
    const aoa = L.fillMerges(XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null }), ws['!merges']);
    const g = L.guessDocType(aoa, n, path.basename(file));
    if (!g.type) return;
    if (g.type !== Sample.SHEET_TYPE[n]) throw new Error(n + ' 시트를 ' + g.type + ' 로 추정함');
    got[g.type] = L.parseDoc(aoa, g.type).rows;
  });
  for (const id of L.DOC_IDS) {
    if (JSON.stringify(got[id]) !== JSON.stringify(want[id])) throw new Error(path.basename(file) + ' · ' + id + ' 다시 읽기 불일치');
  }
  const res = L.runReview(Object.fromEntries(L.DOC_IDS.map(id => [id, { rows: got[id] }])), L.defaultSettings());
  const c = res.summary.counts;
  console.log(path.basename(file), '— 시트 5개 다시 읽기 같음 ·', L.DOC_IDS.map(id => id + ' ' + got[id].length + '행').join(', '),
    '· 판정 적합 ' + c.ok + ' 불일치 ' + c.mismatch + ' 누락 ' + c.missing + ' 확인 ' + c.check + ' → ' + res.summary.decision.label);
}

// CSV 한 개 — 병합 칸이 없는 관리계획서(공정번호·공정명을 첫 행에만 적은 모양)도 읽히는지 보는 용도
const cp = Sample.cp(1).aoa.slice(4);
const csv = cp.map(row => row.map(v => { const s = v == null ? '' : String(v); return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s; }).join(',')).join('\r\n');
fs.writeFileSync(path.join(out, '예시_관리계획서_병합없음.csv'), '﻿' + csv);
console.log('예시_관리계획서_병합없음.csv — 두 줄 머리, 병합 없음');
