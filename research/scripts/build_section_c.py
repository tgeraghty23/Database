"""Build 'C. Benchmarking' sheet: top-30 Swedish business law firms by Affärsvärlden FY2024 fee income,
with revenue per partner / employee / lawyer and leverage ratios."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

path = sys.argv[1]
wb = load_workbook(path)
if 'C. Benchmarking' in wb.sheetnames: del wb['C. Benchmarking']
ws = wb.create_sheet('C. Benchmarking')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
N = Font(italic=True, size=9); grey = PatternFill('solid', fgColor='EDEDED')

ws['A1'] = 'Section C. Benchmarking: top 30 Swedish business law firms'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Ranking: Affärsvärlden fee income FY2024 (revised table, 3 Oct 2025; split-year firms on FY24/25). Ratios use the latest available revenue '
            '(FY2025 where published, else FY2024; basis shown) against current headcount (mostly Oct-2026), so they are indicative. '
            'Leverage (requested) = partners / total employees; partners / lawyers shown alongside. Grey = not available.')
ws['A2'].alignment = W; ws.merge_cells('A2:R2'); ws.row_dimensions[2].height = 45
hdr = ['#', 'Firm', 'Afv fee income FY2024', 'Afv FY2023', 'Growth FY24', 'Latest revenue', 'Latest revenue basis', 'Partners', 'Lawyers (incl. partners)',
       'Total employees', 'Headcount source', 'Rev / partner', 'Rev / employee', 'Rev / lawyer', 'Leverage: partners / employees', 'Leverage: partners / lawyers',
       'Employees per partner', 'Notes']
R = 4
for c, h in enumerate(hdr, 1):
    x = ws.cell(R, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); x.border = B
ws.row_dimensions[R].height = 48
WEB = 'Firm website count Oct-26'
# rank, firm, fy24, fy23, latest, basis, partners, lawyers, employees, hc source, notes
D = [
 (1, 'Vinge', 2085, 1751, 2036, 'FY2025 group (annual reports, DJ)', 71, 389, 529, WEB + '; employees: 3 ABs 2025 accounts', 'Lawyers incl. ~46 on leave'),
 (2, 'Mannheimer Swartling', 1911, 1807, 2002, 'FY2025 (annual report)', 92, 430, 649, 'DJ/Chambers 2026; employees 2025 accounts', ''),
 (3, 'Roschier (Sweden)', 1254, 1026, 1304, 'FY Jun25-May26 net sales (AB)', 26, 181, 220, WEB + ' (Stockholm); employees AB accounts', 'Split year; Afv FY24 = FY Jun24-May25'),
 (4, 'Setterwalls', 918, 899, 1000, 'FY2025 (Afv ">1bn")', 38, 177, 306, WEB + '; employees 3 entities 2025', 'Revenue/employees incl. Malmö (left Sep-26); partners/lawyers exclude it'),
 (5, 'Lindahl', 737, 707, 737, 'FY2024 (Afv)', 68, 220, 300, 'Legal 500 / Chambers (firm-supplied)', ''),
 (6, 'Gernandt & Danielsson', 676, 601, 676, 'FY24/25 (Afv)', 20, 91, 150, WEB + '; employees firm website', 'KB - no public accounts'),
 (7, 'DLA Piper (Sweden)', 602, 660, 602, 'FY24/25 (Afv)', 37, 148, None, WEB, 'Employees: allabolag band 200-499 only'),
 (8, 'CMS Wistrand', 558, 529, 558, 'FY2024 (Afv)', 51, 150, 220, 'Firm statement / Chambers 2026 (">220")', ''),
 (9, 'Delphi', 516, 505, 516, 'FY2024 (Afv)', 49, 150, 220, 'Chambers 2026 / firm website', ''),
 (10, 'Cederquist', 500, 472, 500, 'FY2024 (Afv)', 26, 124, None, 'Legal 500 (firm-supplied)', 'KB - no public accounts; Chambers 120 lawyers'),
 (11, 'Schjødt (Sweden)', 427, 343, 534, 'FY2025 est. (">0.5bn", +25%)', 20, 65, None, WEB + ' (Stockholm)', 'May include leavers to BAHR; branch band 50-99 employees'),
 (12, 'MAQS', 417, 346, 417, 'FY2024 (Afv)', 48, 144, 180, WEB + '; employees Chambers 2026', 'Splitting from 1 Jan 2027: Malmö (MAQS KB) vs Sthlm/Gbg/Sundsvall (MAQS AB)'),
 (13, 'Cirio', 363, 317, 431, 'FY2025 net sales (AB)', 27, 76, 127, 'Legal 500; employees 2025 accounts', ''),
 (14, 'Snellman (Sweden)', 295, 300, 279, 'FY2025 net sales (AB)', 17, None, 107, 'Legal 500 (stale); employees 2025 accounts', 'Partner count unverified'),
 (15, 'Foyen', 277, 263, 277, 'FY2024 (Afv)', 23, 73, 106, WEB + '; employees AB accounts', ''),
 (16, 'AG Advokat', 198, 228, 198, 'FY2024 (Afv)', None, 41, None, 'Legal 500', 'Lost 12-person team to Cirio (Jan-25)'),
 (17, 'Glimstedt', 196, 174, 196, 'FY2024 (Afv)', 41, 78, None, WEB + ' (Sweden)', '~200 staff incl. Baltics (not used)'),
 (18, 'TM & Partners', 188, 175, 188, 'FY2024 (Afv)', 18, 63, None, WEB, 'Part of AGRD since 2025; AB (73 staff) is partial'),
 (19, 'WSA', 165, 164, 165, 'FY2024 (Afv)', None, None, None, '', 'Headcount not researched'),
 (20, 'Morris Law', 163, 151, 163, 'FY2024 (Afv)', None, 56, None, 'Legal 500', 'Part of AGRD since 2025'),
 (21, 'Hammarskiöld', 158, 133, 178, 'FY2025 net sales (AB)', None, 47, 63, 'Chambers 2026; employees 2025 accounts', ''),
 (22, 'Wigge & Partners', 146, 122, 146, 'FY2024 (Afv)', 11, 40, None, WEB, 'Service AB (53 staff) is partial; joined AGRD Jan-26'),
 (23, 'Gulliksson', 141, 117, 141, 'FY2024 (Afv)', None, 41, None, 'Legal 500', ''),
 (24, 'Eversheds Sutherland', 140, 148, 166, 'FY2025 net sales (AB)', None, None, 59, '2025 accounts', ''),
 (25, 'Fylgia', 131, 120, 131, 'FY2024 (Afv)', 21, 35, None, WEB, 'KB band 20-49 employees'),
 (26, 'Kilpatrick Townsend', 130, 114, 130, 'FY2024 (Afv)', None, None, None, '', 'AB (49 staff) is partial'),
 (27, 'Nordia Law', 126, 126, 126, 'FY2024 (Afv)', None, None, 96, 'Legal 500 ("number of staff")', ''),
 (28, 'Kaiding', 122, 117, 122, 'FY24/25 (Afv)', None, None, None, '', 'Headcount not researched'),
 (29, 'Magnusson (Sweden)', 112, 124, 125, 'FY2025 net sales (AB)', 13, 47, 46, WEB + '; employees 2025 accounts', 'Merged with Gro Advokatbyrå 1 Oct 2026 (~SEK 20m)'),
 (30, 'Moll Wendén', 108, 89, 101, 'FY2025 net sales (AB)', 8, 37, 44, WEB + '; employees 2025 accounts', 'Joined AGRD May-26'),
 ('n/r', 'Baker McKenzie (Stockholm)', 334, 323, 334, 'FY23/24 (Afv Apr-25; not in revised table)', 12, 73, 125, WEB + '; "~125 professionals" (Mar-26)', 'Would rank ~14'),
 ('n/r', 'Linklaters (Stockholm)', 195, 199, 195, 'FY23/24 (Afv Apr-25; not in revised table)', None, None, 75, 'AB accounts FY24/25', 'AB net sales 376 (basis unclear)'),
]
r = R
for d in D:
    r += 1
    rk, firm, fy24, fy23, latest, basis, p, l, e, src, note = d
    vals = [rk, firm, fy24, fy23, None, latest, basis, p, l, e, src]
    for c, v in enumerate(vals, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c in (2, 7, 11) else C
        if c in (3, 4, 6): x.number_format = '#,##0'
        if c in (8, 9, 10) and v is None: x.fill = grey
    x = ws.cell(r, 5, f'=C{r}/D{r}-1'); x.number_format = '0%'; x.border = B; x.alignment = C
    calcs = {12: (p, f'=F{r}/H{r}', '0.0'), 13: (e, f'=F{r}/J{r}', '0.00'), 14: (l, f'=F{r}/I{r}', '0.00'),
             15: (p and e, f'=H{r}/J{r}', '0%'), 16: (p and l, f'=H{r}/I{r}', '0%'), 17: (p and e, f'=J{r}/H{r}', '0.0')}
    for c, (ok, f, fmt) in calcs.items():
        x = ws.cell(r, c); x.border = B; x.alignment = C
        if ok: x.value = f; x.number_format = fmt
        else: x.fill = grey
    x = ws.cell(r, 18, note); x.border = B; x.alignment = W
last = r
r += 1
ws.cell(r, 2, 'Median (top 30)').font = Font(bold=True)
for c in range(12, 18):
    col = get_column_letter(c)
    x = ws.cell(r, c, f'=MEDIAN({col}{R+1}:{col}{R+30})'); x.font = Font(bold=True); x.alignment = C
    x.number_format = ws.cell(R + 2, c).number_format
r += 1
ws.cell(r, 2, 'Top-30 total fee income FY2024').font = Font(bold=True)
ws.cell(r, 3, f'=SUM(C{R+1}:C{R+30})').number_format = '#,##0'
ws.cell(r, 4, f'=SUM(D{R+1}:D{R+30})').number_format = '#,##0'
ws.cell(r, 5, f'=C{r}/D{r}-1').number_format = '0%'
ws.cell(r + 1, 2, 'Share of Afv panel total (SEK 15,455m)')
ws.cell(r + 1, 3, f'=C{r}/15455').number_format = '0%'
r += 3
for t in [
    'Units: SEK m; ratios in SEK m per head. Revenue per partner/employee/lawyer uses "Latest revenue" (column F).',
    'Sources: Affärsvärlden 3 Oct 2025 https://www.affarsvarlden.se/artikel/affarsjuristerna-gar-starkt-finns-fler-bolag-i-ipo-pipelinen (FY2024 table) and 12 Apr 2025 (original); '
    'company accounts via allabolag.se (Bolagsverket filings); firm websites (counted Oct-2026); Legal 500 and Chambers Europe 2026 firm-supplied figures.',
    'Gaps: total employees not found for DLA Piper, Cederquist, Schjødt (Sweden), Glimstedt (Sweden only), TM & Partners, Wigge, Fylgia, AG Advokat; partner counts missing for 9 firms. '
    'Limited partnerships (KB) file no public accounts.',
]:
    ws.cell(r, 1, t).font = N; r += 1
for c, w in zip(range(1, 19), (5, 26, 11, 10, 9, 10, 26, 9, 10, 10, 30, 10, 10, 10, 12, 12, 10, 40)):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'C5'
order = ['A. Market size', 'A. Firms', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw',
         'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('ok', wb.sheetnames)
