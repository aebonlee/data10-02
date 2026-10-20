# 예시 검사기준서 PDF 만들기 (가상 데이터 — 실제 부품·회사 아님)
#   python3 scripts/make-sample-pdf.py      (reportlab 필요: pip install reportlab, node 필요)
# js/sample-data.js 의 2차 재제출 검사기준서 표를 글자가 들어 있는 PDF(글자 층 있음)로 씁니다.
# 「서류 업로드」에서 PDF 를 고르면 글자 위치로 표를 다시 맞춰 읽는지 시험하는 용도입니다(1차수 data09-08 의 PDF 샘플 방식).
import json
import os
import subprocess
import sys

from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
FONT_CANDIDATES = [
    os.path.expanduser('~/Library/Fonts/NanumGothic-Regular.ttf'),
    '/Library/Fonts/NanumGothic-Regular.ttf',
    '/usr/share/fonts/truetype/nanum/NanumGothic.ttf',
    'C:/Windows/Fonts/NanumGothic.ttf',
]
font = next((f for f in FONT_CANDIDATES if os.path.exists(f)), None)
if not font:
    sys.exit('한글 TTF 글꼴(NanumGothic, OFL)을 찾지 못했습니다.')
pdfmetrics.registerFont(TTFont('KR', font))

js = "const S=require('./js/sample-data.js');console.log(JSON.stringify({meta:S.META,date:S.DATES[2],aoa:S.insp(2).aoa}))"
data = json.loads(subprocess.check_output(['node', '-e', js], cwd=ROOT))
aoa = data['aoa']
head_i = next(i for i, r in enumerate(aoa) if r and r[0] == '검사 구분')
head, body = aoa[head_i], aoa[head_i + 1:]

W, H = landscape(A4)
X = [40, 120, 180, 330, 400, 540, 650, 720]  # 칸 왼쪽 끝(pt)
out = os.path.join(ROOT, 'samples', '예시_APQP_SMP-2041_검사기준서_2차.pdf')
c = canvas.Canvas(out, pagesize=(W, H))
c.setTitle('검사기준서 (가상 예시)')
y = H - 50
c.setFont('KR', 15)
c.drawString(40, y, '검사기준서 — 가상 예시(실제 회사 자료 아님)')
y -= 22
c.setFont('KR', 9.5)
c.drawString(40, y, '품번: %s   품명: %s   협력사: %s   제출: 2차 (%s)' % (data['meta']['partNo'], data['meta']['partName'], data['meta']['supplier'], data['date']))
y -= 30
c.setFont('KR', 10)
c.line(36, y + 14, W - 36, y + 14)
for x, t in zip(X, head):
    c.drawString(x, y, str(t))
c.line(36, y - 6, W - 36, y - 6)
y -= 24
for r in body:
    for x, t in zip(X, r):
        if t is not None and t != '':
            c.drawString(x, y, str(t))
    c.setStrokeColorRGB(.8, .8, .8)
    c.line(36, y - 7, W - 36, y - 7)
    y -= 22
c.setFont('KR', 8.5)
c.drawString(40, 40, '이 파일은 APQP Smart Review 시험용 가상 데이터입니다.')
c.save()
print(out)
