"""Build 'B. Top lawyers' sheet: individual rankings (Chambers Bands 1-3; Legal 500 Hall of Fame, Leading partners,
Next Generation Partners) in Corporate/M&A, Private equity, Banking & finance and Dispute resolution, Sweden."""
import json, sys
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
sys.argv, path = [sys.argv[0], '/dev/null'], sys.argv[1]
exec(open('research/scripts/build_rankings.py').read().split('AREAS = [')[0])  # CANON / canon()

data = json.load(open('research/data/top_lawyers_chambers2026_legal500_sweden.json'))
KEEP = {'Chambers': ['Band 1', 'Band 2', 'Band 3'], 'Legal 500': ['Hall of Fame', 'Leading partners', 'Next Generation Partners']}
AREAS = ['Corporate / M&A', 'Private equity', 'Banking & finance', 'Dispute resolution']
HL = {'MAQS': 'FFE699', 'Magnusson': 'F8CBAD', 'Delphi': 'F8CBAD', 'Lindahl': 'F8CBAD'}
rows = [d for d in data if d[2] in KEEP[d[0]]]
rank_order = {k: i for src in KEEP.values() for i, k in enumerate(src)}

wb = load_workbook(path)
if 'B. Top lawyers' in wb.sheetnames: del wb['B. Top lawyers']
ws = wb.create_sheet('B. Top lawyers')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864'); SUB = PatternFill('solid', fgColor='D9E1F2')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
ws['A1'] = 'B3. Top-ranked lawyers, Sweden: Corporate / M&A, Private equity, Banking & finance, Dispute resolution'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Top three ranking levels per directory. Chambers Europe 2026: Band 1, Band 2, Band 3 (individual bands; Senior Statespeople, Eminent Practitioners, '
            'Up and Coming and Bands 4-5 excluded). Legal 500 EMEA (current edition): Hall of Fame, Leading partners, Next Generation Partners (Leading associates excluded; '
            'Legal 500 does not split "Leading partners" into tiers). Legal 500 has no separate Sweden private equity table (PE sits within Corporate / M&A), so Private equity '
            'is Chambers only. Firm = as listed by the directory (reflects recent moves, e.g. partners now at BAHR). Extracted 13 Oct 2026. Highlight: MAQS (yellow); Magnusson, Delphi, Lindahl (orange).')
ws['A2'].alignment = W; ws.merge_cells('A2:H2'); ws.row_dimensions[2].height = 75

# --- Summary: count of top-ranked lawyers per firm per area
r = 4
ws.cell(r, 1, 'Summary: number of top-ranked lawyers per firm (unique names per area, either directory)').font = Font(bold=True, size=11)
r += 1
cnt = defaultdict(lambda: defaultdict(set))
for src, area, rank, name, firm in rows:
    cnt[canon(firm)][area].add(name)
firms = sorted(cnt, key=lambda f: (-sum(len(v) for v in cnt[f].values()), f))
hdrs = ['Firm'] + AREAS + ['Total']
for c, h in enumerate(hdrs, 1):
    x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = C; x.border = B
top = r + 1
for f in firms:
    r += 1
    ws.cell(r, 1, f).border = B
    for c, a in enumerate(AREAS, 2):
        x = ws.cell(r, c, len(cnt[f][a]) or None); x.alignment = C; x.border = B
    x = ws.cell(r, len(hdrs), f'=SUM(B{r}:E{r})'); x.alignment = C; x.border = B; x.font = Font(bold=True)
    if f in HL:
        for c in range(1, len(hdrs) + 1): ws.cell(r, c).fill = PatternFill('solid', fgColor=HL[f])
r += 2

# --- Detail per practice area
for area in AREAS:
    ws.cell(r, 1, area).font = Font(bold=True, size=12, color='1F3864'); r += 1
    for c, h in enumerate(['Directory', 'Ranking', 'Lawyer', 'Firm', 'Firm (as listed)'], 1):
        x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = C; x.border = B
    sel = sorted([d for d in rows if d[1] == area], key=lambda d: (d[0] != 'Chambers', rank_order[d[2]], canon(d[4]), d[3]))
    for src, a, rank, name, firm in sel:
        r += 1
        f = canon(firm)
        for c, v in enumerate([src, rank, name, f, firm], 1):
            x = ws.cell(r, c, v); x.border = B; x.alignment = W
            if f in HL: x.fill = PatternFill('solid', fgColor=HL[f])
    r += 2
ws.cell(r, 1, 'Highlighted firms below the top three levels (not shown above):').font = Font(bold=True)
r += 1
low = [d for d in data if canon(d[4]) in HL and d[2] not in KEEP[d[0]]]
if not low:
    ws.cell(r, 1, 'None ranked.'); r += 1
for src, area, rank, name, firm in sorted(low, key=lambda d: (canon(d[4]), d[1], d[2])):
    for c, v in enumerate([src, rank, name, canon(firm), area], 1):
        x = ws.cell(r, c, v); x.border = B; x.fill = PatternFill('solid', fgColor=HL[canon(firm)])
    r += 1
ws.cell(r, 1, 'MAQS has no individually ranked lawyers in these four practice areas in either directory.').font = Font(italic=True, size=9)
r += 2
ws.cell(r, 1, 'Sources: https://chambers.com/legal-guide/europe-7 (Sweden ranking tables); https://www.legal500.com/c/sweden/<practice-area>/lawyers').font = Font(italic=True, size=9)
for c, w in zip(range(1, 9), (26, 18, 18, 18, 22, 10, 10, 10)):
    ws.column_dimensions[get_column_letter(c)].width = w
order = ['A. Market size', 'A. Nordic market', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'B. Top lawyers', 'B. Partner pipeline', 'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts', 'G. MAQS key data']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('rows', len(rows)); print([(f, sum(len(v) for v in cnt[f].values())) for f in firms][:15])
print({f: {a: len(cnt[f][a]) for a in AREAS} for f in ('MAQS', 'Magnusson', 'Delphi', 'Lindahl') if f in cnt})
