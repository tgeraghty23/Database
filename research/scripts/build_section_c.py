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

ws['A1'] = 'Section C. Benchmarking: top 30 Swedish business law firms (replaces former A. Firms tab)'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Ranking: Affärsvärlden fee income FY2024; FY2025 revenue from filed company accounts / press where available (revised table, 3 Oct 2025; split-year firms on FY24/25). Ratios use the latest available revenue '
            '(FY2025 where published, else FY2024; basis shown) against current headcount (mostly Oct-2026), so they are indicative. '
            'Leverage (requested) = partners / total employees; partners / lawyers shown alongside. Grey = not available.')
ws['A2'].alignment = W; ws.merge_cells('A2:U2'); ws.row_dimensions[2].height = 45
hdr = ['#', 'Firm', 'Afv fee income FY2024', 'Revenue FY2025', 'YoY FY25 vs FY24 (like-for-like)', 'FY2025 source / coverage', 'Afv FY2023', 'Growth FY24', 'Latest revenue', 'Latest revenue basis', 'Partners', 'Lawyers (incl. partners)',
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
# FY2025 revenue (SEK m): (FY2025, FY2024 on the same basis, source / coverage). YoY is like-for-like on the same basis.
FY25 = {
 'Vinge': (2036, 2083, 'Group annual reports (DJ via Realtid): full firm'),
 'Mannheimer Swartling': (2002, 1916, 'Annual report (AB): full firm'),
 'Roschier (Sweden)': (1304, 1254, 'AB accounts FY Jun25-May26 vs FY24/25: full firm'),
 'White & Case (Stockholm)': (1108, 859, 'AB accounts: full firm (Stockholm)'),
 'Setterwalls': (1056, 996, 'Parent AB accounts (pass-through, incl. Malmö); Afv ">1bn"'),
 'Lindahl': (438, 406, 'PARTIAL: Malmö, Stockholm, Göteborg ABs only'),
 'Gernandt & Danielsson': (None, None, 'Not filed (KB)'),
 'DLA Piper (Sweden)': (None, None, 'Not filed (KB)'),
 'CMS Wistrand': (None, None, 'Not filed (KB)'),
 'Delphi': (356, 355, 'PARTIAL: Stockholm, Göteborg, Malmö ABs'),
 'Cederquist': (555, 500, 'Press snippet (Afv FY2025), unverified; KB not filed'),
 'Schjødt (Sweden)': (534, 427, 'Press estimate (Afv: ">0.5bn", +25%)'),
 'Cirio': (431, 373, 'AB accounts: full firm'),
 'Snellman (Sweden)': (279, 299, 'AB accounts: full firm'),
 'Foyen': (172, 164, 'PARTIAL: main Swedish AB'),
 'AG Advokat': (None, None, 'Not filed (KB)'),
 'Glimstedt': (None, None, 'Not collected (multiple regional ABs)'),
 'TM & Partners': (103, 92, 'PARTIAL: TM & Partners Law AB'),
 'WSA': (None, None, 'Not collected'),
 'Morris Law': (None, None, 'Not collected'),
 'Hammarskiöld': (178, 160, 'AB accounts: full firm'),
 'Wigge & Partners': (None, None, 'Service AB only (not comparable)'),
 'Gulliksson': (None, None, 'Not collected'),
 'Eversheds Sutherland': (166, 160, 'AB accounts: full firm'),
 'Fylgia': (None, None, 'Not filed (KB)'),
 'Kilpatrick Townsend': (None, None, 'Service AB only (not comparable)'),
 'Nordia Law': (None, None, 'Not collected'),
 'Kaiding': (None, None, 'Not collected'),
 'Magnusson (Sweden)': (125, 112, 'AB accounts: full firm'),
 'Moll Wendén': (101, 108, 'AB accounts: full firm'),
 'Baker McKenzie (Stockholm)': (None, None, 'Not filed (KB)'),
 'Linklaters (Stockholm)': (376, 422, 'AB accounts FY May24-Apr25 vs prior year; basis unclear'),
 'MAQS': (324, 301, 'MAQS AB accounts (excl. Malmö KB)'),
}
r = R
for d in D:
    r += 1
    rk, firm, fy24, fy23, latest, basis, p, l, e, src, note = d
    f25, f24s, f25src = FY25.get(firm, (None, None, 'Not collected'))
    vals = [rk, firm, fy24, f25, None, f25src, fy23, None, latest, basis, p, l, e, src]
    for c, v in enumerate(vals, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c in (2, 6, 10, 14) else C
        if c in (3, 4, 7, 9): x.number_format = '#,##0'
        if c in (4, 11, 12, 13) and v is None: x.fill = grey
        if c == 6 and ('PARTIAL' in f25src or 'Press' in f25src or 'unclear' in f25src): x.fill = PatternFill('solid', fgColor='FFF2CC')
    x = ws.cell(r, 5); x.border = B; x.alignment = C
    if f25 and f24s: x.value = f'=D{r}/{f24s}-1'; x.number_format = '+0%;-0%'
    else: x.fill = grey
    x = ws.cell(r, 8, f'=C{r}/G{r}-1'); x.number_format = '0%'; x.border = B; x.alignment = C
    calcs = {15: (p, f'=I{r}/K{r}', '0.0'), 16: (e, f'=I{r}/M{r}', '0.00'), 17: (l, f'=I{r}/L{r}', '0.00'),
             18: (p and e, f'=K{r}/M{r}', '0%'), 19: (p and l, f'=K{r}/L{r}', '0%'), 20: (p and e, f'=M{r}/K{r}', '0.0')}
    for c, (ok, f, fmt) in calcs.items():
        x = ws.cell(r, c); x.border = B; x.alignment = C
        if ok: x.value = f; x.number_format = fmt
        else: x.fill = grey
    x = ws.cell(r, 21, note); x.border = B; x.alignment = W
last = r
r += 1
ws.cell(r, 2, 'Median (top 30)').font = Font(bold=True)
for c in [5] + list(range(15, 21)):
    col = get_column_letter(c)
    x = ws.cell(r, c, f'=MEDIAN({col}{R+1}:{col}{R+30})'); x.font = Font(bold=True); x.alignment = C
    x.number_format = ws.cell(R + 2, c).number_format
r += 1
ws.cell(r, 2, 'Top-30 total fee income FY2024').font = Font(bold=True)
ws.cell(r, 3, f'=SUM(C{R+1}:C{R+30})').number_format = '#,##0'
ws.cell(r, 7, f'=SUM(G{R+1}:G{R+30})').number_format = '#,##0'
ws.cell(r, 8, f'=C{r}/G{r}-1').number_format = '0%'
ws.cell(r + 1, 2, 'Share of Afv panel total (SEK 15,455m)')
ws.cell(r + 1, 3, f'=C{r}/15455').number_format = '0%'
ws.cell(r + 2, 2, 'Top-5 share of Afv panel total')
ws.cell(r + 2, 3, f'=SUM(C{R+1}:C{R+5})/15455').number_format = '0%'
ws.cell(r + 3, 2, 'Top-10 share of Afv panel total')
ws.cell(r + 3, 3, f'=SUM(C{R+1}:C{R+10})/15455').number_format = '0%'
r += 5
for t in [
    'Units: SEK m; ratios in SEK m per head. Revenue per partner/employee/lawyer uses "Latest revenue" (column I). FY2025 YoY is like-for-like on the same basis (company accounts vs prior-year company accounts; group vs group); yellow = partial entity coverage, press figure or unclear basis. The full Afv FY2025 table is paywalled; limited partnerships (KB) do not file.',
    'Sources: Affärsvärlden 3 Oct 2025 https://www.affarsvarlden.se/artikel/affarsjuristerna-gar-starkt-finns-fler-bolag-i-ipo-pipelinen (FY2024 table) and 12 Apr 2025 (original); '
    'company accounts via allabolag.se (Bolagsverket filings); firm websites (counted Oct-2026); Legal 500 and Chambers Europe 2026 firm-supplied figures.',
    'Gaps: total employees not found for DLA Piper, Cederquist, Schjødt (Sweden), Glimstedt (Sweden only), TM & Partners, Wigge, Fylgia, AG Advokat; partner counts missing for 9 firms. '
    'Limited partnerships (KB) file no public accounts.',
]:
    ws.cell(r, 1, t).font = N; r += 1
r += 1
ws.cell(r, 1, 'C2. Utilisation benchmarks: billable-hour targets and Nordic norms').font = Font(bold=True, size=12)
r += 1
for c, h in enumerate(['', 'Firm / basis', 'Metric', '', '', 'Value', 'Period', 'Reliability / source'], 1):
    x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = C; x.border = B
UT = [
    ('Vinge', 'Associate billable-hour target', '1,500 h (incl. 100 h qualifying internal)', 'c.2012-14', 'Indicative: anonymous crowd-sourced table (juristlon.wordpress.com, 2014)'),
    ('Mannheimer Swartling', 'Associate billable-hour target', '1,500 h (none in year 1)', 'c.2012-14', 'Indicative: as above'),
    ('Cederquist', 'Bonus ladder', '0.75 month at 1,500 h up to 4.25 months at 1,800 h', 'c.2012-14', 'Indicative: as above'),
    ('White & Case (Stockholm)', 'Associate billable-hour target', '1,700 h', 'c.2012-14', 'Indicative: as above'),
    ('Lindahl', 'Associate billable-hour target', '1,350 h (up to 200 h internal); bonus from 1,450 h', 'c.2012-14', 'Indicative: as above'),
    ('Fylgia', 'Target / bonus', '1,400 h; bonus at 1,400 / 1,500 h', 'c.2012-14', 'Indicative: as above'),
    ('Delphi', 'Bonus threshold', '1,500 h', 'c.2012-14', 'Indicative: as above'),
    ('Baker McKenzie (Stockholm)', 'Bonus threshold', 'from 1,400 billed h', 'c.2012-14', 'Indicative: as above'),
    ('Bird & Bird (Stockholm)', 'Ramp-up target', '1,000 / 1,200 / 1,450 h (years 1 / 2 / 3+)', '2007', 'Indicative: as above'),
    ('Swedish business-law firms', 'Billable-hour requirement', '1,400-1,600 h p.a.', '2018', 'Advokaten no. 7 2018 (snippet)'),
    ('Swedish law firms', 'Share of worked time billed', '40-90%, median just over 70%', '2012', 'Advokaten 2012 (snippet)'),
    ('Norwegian law firms', 'Billing rate (faktureringsgrad)', 'just over 72% (~1,220 h on 1,695 h)', '2016', 'Advokatforeningen industry survey (memo Sep-2017)'),
    ('Norwegian associates', 'Actual vs contracted hours', '46 h/week actual (48 h at 50+ lawyer firms) vs 39 h contracted; 27% under-record hours', 'Nov-2025', 'Juristforbundet survey (n=279)'),
    ('Danish associates / lawyers', 'Fixed billable-hour requirement', '54% / 55% have one; 47 / 46 h worked per week', '2018', 'Djøf Advokat survey'),
]
for firm, metric, val, per, src in UT:
    r += 1
    for c, v in ((2, firm), (3, metric), (6, val), (7, per), (8, src)):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
r += 1
ws.cell(r, 1, 'Reading: Nordic norm ~1,200-1,300 billable hours (70-75% of a ~1,650-1,700 h year); large business-law firms target ~1,400-1,600 h for associates. '
              'No current published targets found for named Nordic firms or for Finland. See the charging rates briefing for international benchmarks (PwC UK, Thomson Reuters).').font = N
r += 2
for c, w in zip(range(1, 22), (5, 26, 11, 10, 10, 28, 10, 9, 10, 26, 9, 10, 10, 30, 10, 10, 10, 12, 12, 10, 40)):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'C5'
order = ['A. Market size', 'A. Nordic market', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'B. Top lawyers', 'B. Partner pipeline', 'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts', 'G. MAQS key data']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('ok', wb.sheetnames)
