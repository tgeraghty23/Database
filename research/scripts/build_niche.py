"""Build 'A. Niche firms' sheet: Swedish ranked independents incl. PE-owned firms and Legal 500 firms to watch."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
sys.argv, path = [sys.argv[0], '/dev/null'], sys.argv[1]
exec(open('research/scripts/build_rankings.py').read().split('def short')[0])  # loads ch, l5, CANON, canon, AREAS

INTL = {'Baker McKenzie', 'DLA Piper', 'White & Case', 'Linklaters', 'Bird & Bird', 'Eversheds Sutherland', 'CMS Wistrand',
        'Kilpatrick Townsend & Stockton Advokat KB', 'BAHR', 'Schjødt', 'Roschier', 'Snellman', 'Magnusson',
        'Barker Brettell LLP', 'Potter Clarkson LLP', 'Ernst & Young AB'}
LARGE = {'Vinge', 'Mannheimer Swartling', 'Setterwalls', 'Cederquist', 'Gernandt & Danielsson', 'Lindahl', 'Delphi', 'Cirio', 'MAQS'}
BORDER = {'Hammarskiöld': 'Mid-size full-service', 'Foyen': 'Mid-size full-service', 'Fylgia': 'Mid-size full-service',
          'Kanter': 'Mid-size corporate', 'Wigge & Partners': 'Mid-size corporate'}
AGRD = {'Allié', 'Born', 'Morris Law', 'Next Law', 'Synch', 'TM & Partners'}
# Partners / lawyers (firm-supplied to Legal 500 or Chambers, 2026); source tag
HC = {'Kahn Pedersen': (6, 19, 'L500/Chambers'), 'Sandart & Partners': (10, 12, 'L500'), 'REAL Advokatbyrå': (5, 20, 'L500/Chambers'),
      'Kompass': (4, 15, 'L500'), 'Hellström': (22, 49, 'L500 (22 partners, 10 senior assoc., 17 assoc.)'), 'Wigge & Partners': (11, 40, 'Firm website, Oct-26'),
      'Kanter': (13, 39, 'Chambers (IFLR1000: 10 partners / Realtid: 21 lawyers; conflicting)'), 'Kastell': (9, 18, 'Firm website, Oct-26'), 'Fylgia': (21, 35, 'Firm website, Oct-26'),
      'Carler': (7, None, 'Chambers'), 'Ström': (5, None, 'Chambers'), 'Westerberg & Partners': (12, 24, 'Firm website, Oct-26'),
      'Hammarskiöld': (None, 47, 'Chambers'), 'Foyen': (23, 73, 'Firm website, Oct-26'), 'TM & Partners': (18, 63, 'Firm website, Oct-26'), 'Morris Law': (None, 56, 'L500'),
      'AG Advokat': (None, 41, 'L500'), 'Gulliksson': (None, 41, 'L500'), 'TIME DANOWSKY Advokatbyrå AB': (None, 22, 'L500'), 'Norburg & Scherp': (None, 18, 'L500'),
      'Born': (9, 26, 'Firm website, Oct-26'), 'Harvest': (5, 27, 'Firm website, Oct-26'), 'Synch': (9, 38, 'Firm website, Oct-26'),
      'A1 Advokater': (None, None, 'Chambers: 17 staff')}

firms = {}
for area, ck, lk in AREAS:
    for d, src, key in (('CH', ch, ck), ('L5', l5, lk)):
        if not key: continue
        for label, n in src[key]:
            f = canon(n); e = firms.setdefault(f, {'CH': [], 'L5': [], 'FTW': []})
            if label == 'Firms to watch': e['FTW'].append(area)
            else: e[d].append((area, int(label[-1])))

rows = []
for f, e in firms.items():
    if f in INTL or f in LARGE: continue
    if f in AGRD: cat = 'Ranked independent: PE-owned (AGRD / Axcel)'
    elif f in BORDER: cat = 'Borderline: ' + BORDER[f]
    elif not e['CH'] and not e['L5']: cat = 'Legal 500 firm to watch only'
    else: cat = 'Ranked independent'
    best = sorted(e['CH'] + e['L5'], key=lambda x: x[1])
    top = '; '.join(sorted({a for a, t in best if t == best[0][1]})) if best else ''
    hc = HC.get(f, (None, None, ''))
    rows.append([f, cat, len(e['CH']), len(e['L5']), len(e['FTW']), (min(t for _, t in best) if best else None), top,
                 '; '.join(e['FTW']), hc[0], hc[1], hc[2]])
order = {'Ranked independent': 0, 'Ranked independent: PE-owned (AGRD / Axcel)': 1, 'Legal 500 firm to watch only': 3}
rows.sort(key=lambda r: (order.get(r[1], 2), -(r[2] + r[3]), r[0]))

wb = load_workbook(path)
if 'A. Niche firms' in wb.sheetnames: del wb['A. Niche firms']
ws = wb.create_sheet('A. Niche firms', wb.sheetnames.index('A. Growth drivers') + 1)
ws['A1'] = 'A5. Niche / independent Swedish law firms ranked by Chambers or Legal 500'; ws['A1'].font = Font(bold=True, size=14)
n_core = sum(1 for r in rows if r[1].startswith('Ranked independent'))
n_ftw = sum(1 for r in rows if r[1].startswith('Legal 500 firm'))
n_b = sum(1 for r in rows if r[1].startswith('Borderline'))
ws['A2'] = (f'Universe: firm-level rankings in Chambers Europe 2026 (Sweden) and Legal 500 EMEA (Sweden), incl. Legal 500 "firms to watch". Excludes international firms '
            f'and the nine largest Swedish full-service firms. Counts: {n_core} ranked independents (incl. PE-owned AGRD firms) + {n_ftw} firms to watch only = {n_core + n_ftw}; '
            f'+{n_b} borderline mid-size firms = {n_core + n_ftw + n_b}. Partner/lawyer counts: firm website counts (Oct-26) where available, else firm-supplied to Legal 500 / Chambers (2026); blank where not found.')
ws['A2'].alignment = Alignment(wrap_text=True, vertical='top'); ws.merge_cells('A2:K2'); ws.row_dimensions[2].height = 60
hdr = ['Firm', 'Category', '# Chambers rankings', '# Legal 500 rankings', '# L500 firm-to-watch', 'Best band/tier', 'Practice area(s) at best band/tier',
       'Firm-to-watch areas', 'Partners', 'Lawyers', 'Headcount source']
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
for c, h in enumerate(hdr, 1):
    x = ws.cell(4, c, h); x.font = Font(bold=True, color='FFFFFF'); x.fill = PatternFill('solid', fgColor='1F3864')
    x.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); x.border = B
fills = {'Ranked independent: PE-owned (AGRD / Axcel)': 'FFF2CC', 'Legal 500 firm to watch only': 'E2EFDA'}
for i, r in enumerate(rows, 5):
    for c, v in enumerate(r, 1):
        x = ws.cell(i, c, v if v not in ('', None) else None); x.border = B; x.alignment = Alignment(wrap_text=c in (2, 7, 8, 11), vertical='top')
        if r[1] in fills: x.fill = PatternFill('solid', fgColor=fills[r[1]])
        elif r[1].startswith('Borderline'): x.fill = PatternFill('solid', fgColor='EDEDED')
for c, w in zip(range(1, 12), (26, 30, 10, 10, 10, 9, 36, 28, 9, 9, 28)):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'B5'
end = 5 + len(rows) + 1
ws.cell(end, 1, 'Best band/tier: 1 = top (Chambers band or Legal 500 tier, whichever is better). Yellow = PE-owned (AGRD, Axcel-backed roll-up formed Aug-2025); '
                'green = Legal 500 firm to watch only; grey = borderline mid-size firms.').font = Font(italic=True, size=9)
ws.cell(end + 1, 1, 'Excluded (international): ' + ', '.join(sorted(f for f in firms if f in INTL)) + '. Excluded (large Swedish full-service): ' +
        ', '.join(sorted(f for f in firms if f in LARGE)) + '.').font = Font(italic=True, size=9)
wb.save(path)
print(n_core, n_ftw, n_b); [print(r[:6], r[8:10]) for r in rows]
