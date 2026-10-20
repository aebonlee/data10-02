// 헤드리스 브라우저로 주요 흐름 점검 (선택 — playwright-core 가 있는 PC 에서만)
//   PLAYWRIGHT_CORE=/경로/node_modules/playwright-core node scripts/check-browser.mjs
// file:// 로 index.html 을 열어 예시 데이터 → 정합성 검토(근거 보기) → 공정 대조표 → 승인 의견·AI 답 반영 → 보고서 Excel·Word
// → 검토 이력 비교 → 실제 업로드 길(새 검토 건 · 1차 xlsx · 재제출 · PDF) → 설정(규칙 끄기) 순서로 누르고, JS 오류와 폭별 가로 넘침을 잽니다.
import { createRequire } from 'node:module';
import path from 'node:path';
import fs from 'node:fs';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const { chromium } = require(process.env.PLAYWRIGHT_CORE || 'playwright-core');
const XLSX = require(path.join(root, 'vendor/xlsx.full.min.js'));
const url = 'file://' + path.join(root, 'index.html');

const errors = [];
const problems = [];
const browser = await chromium.launch();
const ctx = await browser.newContext({ acceptDownloads: true, viewport: { width: 1280, height: 900 } });
const page = await ctx.newPage();
page.on('pageerror', e => errors.push('pageerror: ' + e.message));
page.on('console', m => { if (m.type() === 'error') errors.push('console: ' + m.text()); });

function ok(cond, msg) { if (!cond) problems.push(msg); else console.log('  ok  ' + msg); }
async function go(hash) { await page.evaluate(h => { location.hash = h; }, hash); await page.waitForTimeout(150); }
async function clickText(sel, text) { await page.locator(sel, { hasText: text }).first().click(); await page.waitForTimeout(150); }
async function text(sel) { return (await page.locator(sel).first().textContent()) || ''; }

await page.goto(url);
// 점검기 자체 확인용: BREAK=1 이면 일부러 넓은 상자를 넣어 넘침을 잡는지 봅니다
if (process.env.BREAK) await page.addStyleTag({ content: '.page-head { min-width: 900px; }' });
ok((await text('main')).includes('아직 검토 건이 없습니다'), '빈 대시보드');
await clickText('button', '예시 데이터로 시작');
ok(await page.locator('#sampleBanner').isVisible(), '예시 데이터 띠 보임');
ok((await text('main h1')).includes('정합성 검토'), '정합성 검토 화면');
ok((await text('.dk-dec')).includes('반려'), '1차 제출 → 반려 제안');
const nIssues = await page.locator('table.findings tbody tr').count();
ok(nIssues === 6, '지적 6건 (' + nIssues + ')');
ok(await page.locator('table.findings .vd-missing').count() === 3, '누락 3건');

// 근거 보기
await page.locator('table.findings .ev', { hasText: 'PFMEA' }).first().click();
await page.waitForTimeout(150);
ok(await page.locator('#dialog').isVisible() && (await text('#dialog')).includes('행'), '근거 행 보기 대화상자');
await clickText('#dialog button', '닫기');

// 필터
await page.locator('select[aria-label="판정"]').selectOption('');
await page.waitForTimeout(150);
ok(await page.locator('table.findings tbody tr').count() > 50, '전체 판정 보기');
await page.locator('select[aria-label="판정"]').selectOption('issues');
await page.waitForTimeout(100);

// 공정 대조표
await go('#/matrix');
ok(await page.locator('td.cov-missing').count() === 2, '대조표 누락 칸 2개(30 PFMEA · 60 외주)');
ok(await page.locator('td.cov-diff').count() === 1, '공정명 다른 칸 1개(20 레이저 절단)');

// 승인 검토 보고서
await go('#/report');
await page.locator('select[aria-label="검토자 결정"]').selectOption('reject');
await page.locator('.dec-card input').first().fill('부품품질팀 검토자');
await clickText('button', '지적 사항으로 보완 요청 채우기');
ok((await page.locator('.dec-card textarea').nth(1).inputValue()).includes('PFMEA'), '보완 요청 자동 채움');
await clickText('.dec-card button', '저장');
const prompt = await page.locator('textarea.prompt-box[readonly]').inputValue();
ok(!/협력사 B|SMP-2041|외주처 [가나]/.test(prompt), '프롬프트에 협력사·품번·외주처 없음(별칭)');
const alias = (/검토 건: (업체\d+)/.exec(prompt) || [])[1];
await page.locator('textarea.prompt-box.answer').fill(['검토 의견: ' + alias + ' 서류에 중대 지적 4건이 있어 반려를 제안합니다.', '보완 요청:', '- PFMEA 공정 30 추가', '- 외주품 리스트에 도장 추가', '확인 필요:', '- 공정 20 공정명'].join('\n'));
await clickText('button', '답 읽기');
ok((await text('.ai-preview')).includes('보완 요청 2건'), 'AI 답 읽음');
await clickText('button', '검토 의견·보완 요청에 반영');
ok((await page.locator('.dec-card textarea').first().inputValue()).includes('협력사 B 서류에'), 'AI 의견 별칭 되돌림');
ok((await text('.report-preview')).includes('반려') && (await text('.report-preview')).includes('부품품질팀 검토자'), '보고서 미리보기에 결정·검토자');
const [dl] = await Promise.all([page.waitForEvent('download'), clickText('button', 'Excel 보고서')]);
const file = path.join(root, 'samples', '.check-report.xlsx');
await dl.saveAs(file);
const wb = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
fs.unlinkSync(file);
ok(wb.SheetNames.join(',') === '요약,판정 목록,공정 대조표,규칙 설정,서류 데이터', '1차 보고서 시트: ' + wb.SheetNames.join(','));
const sum = XLSX.utils.sheet_to_json(wb.Sheets['요약'], { header: 1 });
ok(sum.some(r => r[0] === '검토자 결정' && r[1] === '반려'), '보고서 요약에 검토자 결정');
const [dw] = await Promise.all([page.waitForEvent('download'), clickText('button', 'Word')]);
ok(/\.doc$/.test(dw.suggestedFilename()), 'Word(.doc) 내려받기 ' + dw.suggestedFilename());

// 2차 재제출 · 검토 이력
await go('#/history');
ok((await text('main')).includes('1차 → 2차'), '이력 — 1차 → 2차 비교');
ok((await text('.cmp-fixed h4')).includes('5건') && (await text('.cmp-remain h4')).includes('1건') && (await text('.cmp-added h4')).includes('0건'), '고쳐짐 5 · 남음 1 · 새로 생김 0');
await page.locator('table.list tbody tr', { hasText: '2차' }).first().locator('a', { hasText: '열기' }).click();
await page.waitForTimeout(150);
ok((await text('.dk-dec')).includes('조건부 승인'), '2차 → 조건부 승인 제안');
ok((await text('main')).includes('앞 제출(1차) 대비: 고쳐진 지적 5건'), '검토 화면의 앞 제출 비교');
await go('#/report');
const [dl2] = await Promise.all([page.waitForEvent('download'), clickText('button', 'Excel 보고서')]);
await dl2.saveAs(file);
const wb2 = XLSX.read(fs.readFileSync(file), { type: 'buffer' });
fs.unlinkSync(file);
ok(wb2.SheetNames.includes('재제출 비교'), '2차 보고서에 재제출 비교 시트');

// 대시보드
await go('#/home');
ok((await text('main')).includes('최근 검토 건') && await page.locator('.dkpis .dk').count() === 4, '대시보드');

// 실제 업로드 길 — 새 검토 건 → 1차 xlsx → 열 연결 확인 → 넣기
await go('#/upload');
await clickText('button', '새 검토 건');
const inputs = page.locator('.card .form-grid input');
await inputs.nth(0).fill('협력사 Q');
await inputs.nth(1).fill('SMP-9001');
await clickText('button', '검토 건 만들기');
ok((await text('main')).includes('협력사 Q · SMP-9001 · 1차'), '새 검토 건 만들기');
await page.locator('.drop input[type=file]').setInputFiles(path.join(root, 'samples', '예시_APQP_SMP-2041_1차제출.xlsx'));
await page.waitForSelector('.import');
const types = await page.locator('.imp-sheet select[aria-label$="서류 종류"]').evaluateAll(s => s.map(x => x.value));
ok(types.join(',') === 'flow,cp,pfmea,insp,outs,', '시트 6개 서류 종류 추정: ' + types.join(','));
ok((await text('.import')).includes('읽은 행 14개 (머리 5행·6행 두 줄)'), '관리계획서 두 줄 머리 찾음');
await clickText('button', '검토 건에 넣기');
ok((await text('.dk-dec')).includes('반려') && await page.locator('table.findings tbody tr').count() === 6, '올린 1차 파일 → 지적 6건 · 반려');

// 재제출 + PDF(검사기준서)
await go('#/upload');
await clickText('button', '재제출 등록');
await clickText('button', '검토 건 만들기');
ok((await text('main')).includes('SMP-9001 · 2차'), '재제출 2차 검토 건');
await page.locator('.drop input[type=file]').setInputFiles(path.join(root, 'samples', '예시_APQP_SMP-2041_2차재제출.xlsx'));
await page.waitForSelector('.import');
await page.locator('.imp-sheet select[aria-label="검사기준서 서류 종류"]').selectOption('');
await page.locator('.drop input[type=file]').setInputFiles(path.join(root, 'samples', '예시_APQP_SMP-2041_검사기준서_2차.pdf'));
await page.waitForFunction(() => document.querySelectorAll('.imp-file').length === 2, null, { timeout: 15000 });
ok((await text('.imp-file:nth-of-type(2)')).includes('PDF 시범 읽기'), 'PDF 읽음');
const pdfType = await page.locator('.imp-file').nth(1).locator('select[aria-label$="서류 종류"]').inputValue();
ok(pdfType === 'insp', 'PDF → 검사기준서 추정 (' + pdfType + ')');
ok((await page.locator('.imp-file').nth(1).textContent()).includes('읽은 행 7개'), 'PDF 검사기준서 7행');
await clickText('button', '검토 건에 넣기');
ok((await text('.dk-dec')).includes('조건부 승인'), 'xlsx + PDF 재제출 → 조건부 승인');
ok((await text('main')).includes('고쳐진 지적 5건'), '업로드한 재제출도 앞 제출과 비교');
await go('#/upload');
ok((await text('main')).includes('예시_APQP_SMP-2041_검사기준서_2차.pdf'), '검사기준서는 PDF 에서 들어감');

// 설정 — 규칙 끄기
await go('#/settings');
await page.locator('.dict-item summary', { hasText: '공정명 일치' }).click();
await page.locator('.dict-item', { hasText: '공정명 일치' }).locator('label.check', { hasText: '이 규칙 사용' }).locator('input').uncheck();
await clickText('button', '규칙 저장');
await go('#/review');
ok((await text('.dk-dec')).includes('승인') && !(await text('.dk-dec')).includes('조건부'), '공정명 규칙 끄면 2차 → 승인 제안');

// 폭별 가로 넘침 — 안쪽 가로 스크롤 상자(overflow-x auto) 안은 의도된 것이라 뺍니다
for (const w of [360, 390, 768, 1024, 1280, 1440]) {
  await page.setViewportSize({ width: w, height: 900 });
  for (const r of ['#/home', '#/upload', '#/review', '#/matrix', '#/report', '#/history', '#/settings']) {
    await go(r);
    const bad = await page.evaluate(() => {
      const cw = document.documentElement.clientWidth, out = [];
      for (const el of document.querySelectorAll('main *, .topbar *')) {
        const b = el.getBoundingClientRect();
        if (!b.width || !b.height) continue;
        let p = el.parentElement, inner = false;
        while (p) { const ox = getComputedStyle(p).overflowX; if (ox === 'auto' || ox === 'scroll' || ox === 'hidden') { inner = true; break; } p = p.parentElement; }
        if (!inner && b.right > cw + 1) out.push(el.tagName + '.' + el.className + ' ' + Math.round(b.right) + '>' + cw);
      }
      return { out: out.slice(0, 3), page: document.documentElement.scrollWidth > cw + 1 };
    });
    if (bad.page || bad.out.length) problems.push(w + 'px ' + r + ' 넘침: ' + bad.out.join(' | '));
  }
}
console.log('  폭 6가지 × 화면 7개 넘침 점검 끝');
if (process.env.SHOTS) {
  for (const [w, r] of [[390, '#/review'], [1280, '#/review'], [1280, '#/matrix'], [390, '#/matrix'], [1280, '#/history'], [1280, '#/report'], [390, '#/upload']]) {
    await page.setViewportSize({ width: w, height: 1000 });
    await go(r);
    await page.screenshot({ path: path.join(process.env.SHOTS, r.slice(2) + '_' + w + '.png'), fullPage: true });
  }
}

await browser.close();
if (errors.length) problems.push(...errors);
if (problems.length) { console.error('\n문제 ' + problems.length + '건\n' + problems.join('\n')); process.exitCode = 1; }
else console.log('\n브라우저 점검 통과');
