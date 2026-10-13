"""Build 'A. Nordic market' sheet: FX (Riksbank) and Nordic legal / business-law market comparison (public data only)."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
path = sys.argv[1]
wb = load_workbook(path)
if 'A. Nordic market' in wb.sheetnames: del wb['A. Nordic market']
ws = wb.create_sheet('A. Nordic market')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864'); S = Font(bold=True, size=11, color='1F3864'); N = Font(italic=True, size=9)
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
ws['A1'] = 'A5. Nordic legal and business-law market comparison (SE, NO, DK, FI)'; ws['A1'].font = Font(bold=True, size=14)

def hdr(r, h):
    for c, v in enumerate(h, 1):
        x = ws.cell(r, c, v); x.font = H; x.fill = HF; x.alignment = C; x.border = B

# FX
r = 3; ws.cell(r, 1, 'FX: Riksbank fixing, 12 Oct 2026 (latest published; SEK per unit)').font = S; r += 1
hdr(r, ['Currency', 'SEK per unit', 'Source'])
FX = {'SEK': 1.0, 'USD': 10.01386, 'EUR': 11.1965, 'NOK': 1.04684, 'DKK': 1.4979}
fxrow = {}
for cur in ['USD', 'EUR', 'NOK', 'DKK']:
    r += 1; fxrow[cur] = r
    for c, v in enumerate([cur, FX[cur], 'https://api.riksbank.se/swea/v1/Observations/Latest/SEK' + cur + 'PMI (cross-checked vs ECB)'], 1):
        x = ws.cell(r, c, v); x.border = B
    ws.cell(r, 2).number_format = '0.0000'
def fx(cur): return '1' if cur == 'SEK' else f'$B${fxrow[cur]}'

# Main comparison
r += 2; ws.cell(r, 1, 'Market size, 2024 (local currency m and SEK m at today\'s FX)').font = S; r += 1
hdr(r, ['Measure', 'Country', 'Currency', 'Local (m)', 'SEK m', 'Year', 'Source / note'])
ROWS = [
 ['Legal activities (NACE 69.1), official', 'Sweden', 'SEK', 38921, '2024', 'SCB Företagens ekonomi'],
 ['Legal activities (NACE 69.1), official', 'Norway', 'NOK', 27518, '2024', 'SSB table 08206 / 12817'],
 ['Legal activities (NACE 69.10), official', 'Denmark', 'DKK', 19480, '2024', 'Danmarks Statistik GF11'],
 ['Legal activities (TOL 691), official', 'Finland', 'EUR', 1177, '2024', 'Statistics Finland 13w3 (2025: EUR 1,276m; incl. patent offices ~EUR 109m)'],
 ['Business law: official / best estimate', 'Sweden', 'SEK', 16000, '2025', 'Afv panel of 56 largest firms (floor); industry estimate >SEK 25bn'],
 ['Business law: official / best estimate', 'Norway', 'NOK', 15604, '2024', 'SSB service split "forretningsrett" (57% of 69.1)'],
 ['Business law: official / best estimate', 'Denmark', 'DKK', 13050, '2024e', 'Estimate: business clients ~67% of sector turnover (2018, Økonomistyrelsen expert group) x 2024 turnover'],
 ['Business law: official / best estimate', 'Finland', 'EUR', 600, '2025e', 'Estimate EUR 0.55-0.65bn: top-20 firms ~EUR 581m + tail'],
]
first = r + 1
for m, ctry, cur, loc, yr, src in ROWS:
    r += 1
    for c, v in enumerate([m, ctry, cur, loc, None, yr, src], 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c in (1, 7) else C
    ws.cell(r, 4).number_format = '#,##0'
    ws.cell(r, 5, f'=D{r}*{fx(cur)}').number_format = '#,##0'
r += 1
ws.cell(r, 1, 'Nordic total, legal activities').font = Font(bold=True)
ws.cell(r, 5, f'=SUM(E{first}:E{first+3})').number_format = '#,##0'; ws.cell(r, 5).font = Font(bold=True)
r += 1
ws.cell(r, 1, 'Nordic total, business law (estimate)').font = Font(bold=True)
ws.cell(r, 5, f'=SUM(E{first+4}:E{first+7})').number_format = '#,##0'; ws.cell(r, 5).font = Font(bold=True)
ws.cell(r, 7, 'Sweden at panel floor (SEK 16bn); using SEK 20-25bn for Sweden gives ~SEK 64-69bn').alignment = W

# Structure
r += 2; ws.cell(r, 1, 'Structure and growth (Eurostat SBS M69.1, 2024, EUR; consistent basis)').font = S; r += 1
hdr(r, ['Country', 'Turnover EUR m', 'Turnover SEK m', 'Enterprises', 'Employees', 'Turnover / employee EUR k', 'Local CAGR', 'Lawyers (Bar)'])
ES = [['Sweden', 3404.44, 6965, 17521, '~8% p.a. 2022-24 (SCB new series); 5.8% 2010-21', '6,712 members (5,963 practising), end-2025'],
      ['Norway', 2366.31, 2214, 9264, '9.3% p.a. 2019-24 (SSB)', '10,756 Advokatforeningen members incl. trainees (2025)'],
      ['Denmark', 2610.51, 1912, 13158, '4.5% p.a. 2019-24 (DST)', '~7,760 lawyers incl. in-house (end-2025)'],
      ['Finland', 1177.38, 2726, 6509, '4.9% p.a. 2019-25 (StatFin)', '2,443 asianajajat + 1,778 licensed counsel (end-2025)']]
s0 = r + 1
for ctry, t, e, emp, g, law in ES:
    r += 1
    for c, v in enumerate([ctry, t, None, e, emp, None, g, law], 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c in (1, 7, 8) else C
    ws.cell(r, 2).number_format = '#,##0'; ws.cell(r, 4).number_format = '#,##0'; ws.cell(r, 5).number_format = '#,##0'
    ws.cell(r, 3, f'=B{r}*{fx("EUR")}').number_format = '#,##0'
    ws.cell(r, 6, f'=B{r}/E{r}*1000').number_format = '0'
r += 1
ws.cell(r, 1, 'Total').font = Font(bold=True)
for c in (2, 3, 4, 5):
    col = get_column_letter(c); ws.cell(r, c, f'=SUM({col}{s0}:{col}{r-1})').number_format = '#,##0'

# Leading firms / concentration
r += 2; ws.cell(r, 1, 'Leading firms and concentration').font = S; r += 1
hdr(r, ['Country', 'Largest firm (revenue)', 'Top-5 revenue', 'Concentration', 'Revenue per employee, large firms', 'Partner / hourly rate evidence', 'Ownership rules for external capital'])
LF = [
 ['Sweden', 'Vinge SEK 2,036m (2025)', 'Top-5 ~46% of Afv panel (FY2024)', 'Top-10 ~66% of panel', 'SEK 3.1-3.9m (MSA, Vinge, 2025)', 'Partner list SEK 6,000-10,000/h (2025-26)', 'Bar firms lawyer-owned only; route = leave the Bar (AGRD: 9+ firms since 2025); Bar reviewing rules 2026'],
 ['Norway', 'Schjødt NOK 2,498m (2025, consolidated incl. SE/DK/UK)', 'NOK 9.4bn (2025): Schjødt, Wikborg Rein, Thommessen, BAHR, Wiersholm', 'Top-5 ~31% of SSB 69.1 (2024); top-100 NOK 22.2bn', 'NOK 4.7-5.5m (Thommessen, Wiersholm, BAHR, 2025)', 'Blended ~NOK 5,500/h (Wiersholm, court ruling)', 'Lawyers Act 2025: owners must work in the firm, lawyer board majority; passive / PE ownership not permitted for advokat firms'],
 ['Denmark', 'Gorrissen Federspiel DKK 1,311m (2025)', 'Gorrissen, Plesner, Bech-Bruun, Accura, DLA Piper ~DKK 5.5bn (2025)', 'Top-10 ~44% of sector (2022, KFST); est. 47-49% (2025)', 'DKK 2.4-2.55m per FTE (2025)', 'Partner DKK 2,057-5,400/h (2023, mid-market); Kammeradvokat DKK 2,710', 'Law-firm shares only for practising lawyers / staff (BEK 1427/2007); advocates\' court monopoly'],
 ['Finland', 'Roschier EUR 99m (FY05/2026)', 'EUR ~349m: Roschier, Krogerus, Hannes Snellman, Borenius, Castrén & Snellman', 'Top-5 ~27% of TOL 691 (2025); top-5 ~55% of top-30', 'EUR 205-384k per staff (top-10)', 'No current survey (2012 median EUR 200/h)', 'Bar firms lawyer-owned; licensed-counsel route gives full rights of audience; Fondia (listed, non-Bar)'],
]
for row in LF:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W

r += 2; ws.cell(r, 1, 'Consolidation and international activity, 2020-2026').font = S; r += 1
hdr(r, ['Country', 'Activity'])
for row in [
 ['Sweden', 'AGRD (Axcel) roll-up of 9+ firms since Aug-2025 incl. Setterwalls Malmö; CMS-Wistrand (2024); BAHR Stockholm (2026); Schjødt-Hamilton (2020); Magnusson-Gro, Walthon Ramberg (2026); Hannes Snellman split (2024); MAQS split (2027)'],
 ['Norway', 'Schjødt-Hamilton SE (2020), Copenhagen (2022), Norrbom Vinding DK (Mar-2026); BAHR opens Stockholm and Copenhagen (2025-26) after hiring ~18 Schjødt partners; Ræder Bing (2023), Lippestad / Bache-Wiig and Tofte Hald (2024); Cedra buys Big-4 regional law offices (Dec-2025)'],
 ['Denmark', 'HortenDahl merger (Jan-2026); DLA Piper takes 49 from Lundgrens (Mar-2025); Schjødt-Norrbom Vinding (Mar-2026); Kromann Reumert converts to P/S (Nov-2025); proposal to open Kammeradvokat mandate (~DKK 0.7bn) to competition'],
 ['Finland', 'Hannes Snellman split (2024); DLA Piper-Project Law (2023); Merilampi-I&O Partners (2025, single source); no PE investment found; Fondia (listed) shrinking'],
]:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W
r += 2
for t in ['Sources: SCB; SSB tables 08206 / 12817; Danmarks Statistik GF11 / REGN1; Statistics Finland 13w3; Eurostat sbs_ovw_act (M691); Advokatbladet / Kapital JUS (NO league tables); Proff.dk filed accounts and Økonomisk Ugebrev (DK); Asiakastieto / DM Faster (FI); Økonomistyrelsen expert group Dec-2024 (DK); national Bar associations.',
          'Caveats: official statistics measure all legal activities, not business law only (Norway is the only country publishing a business-law split). Firm revenues use different fiscal years and entity scopes (e.g. Schjødt consolidated). Business-law figures for Sweden, Denmark and Finland are estimates. FX: Riksbank 12 Oct 2026.']:
    ws.cell(r, 1, t).font = N; r += 1
for c, w in zip(range(1, 9), (34, 30, 28, 28, 24, 30, 50, 34)):
    ws.column_dimensions[get_column_letter(c)].width = w
order = ['A. Market size', 'A. Nordic market', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'B. Top lawyers', 'B. Partner pipeline',
         'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts', 'G. MAQS key data']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('ok', wb.sheetnames)
