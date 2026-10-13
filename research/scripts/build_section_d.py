"""Build 'D. Rate cards' sheet: hourly rate comparison across Swedish business law firms."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

path = sys.argv[1]
wb = load_workbook(path)
if 'D. Rate cards' in wb.sheetnames: del wb['D. Rate cards']
ws = wb.create_sheet('D. Rate cards')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864'); SUB = PatternFill('solid', fgColor='D9E1F2')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='center', wrap_text=True)
N = Font(italic=True, size=9)

def hdr(r, vals, fill=HF, font=H):
    for c, v in enumerate(vals, 1):
        x = ws.cell(r, c, v); x.font = font; x.fill = fill; x.alignment = C; x.border = B

ws['A1'] = 'Section D. Rate card comparison: Swedish business law firms (SEK per hour, excl. VAT)'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Rates are only comparable when dated: Swedish firms typically revise rates annually on 1 January. Public-sector price annexes are usually '
            'confidential (Kammarrätten i Stockholm, 2017, Vinge / E-hälsomyndigheten), so firm-level rate cards are scarce. Source type matters: '
            'list price (rate card) vs realised (actual billed) rates. Nordic business-law data only, 2023 onwards.')
ws['A2'].alignment = W; ws.merge_cells('A2:Q2'); ws.row_dimensions[2].height = 45

# ---- D1 list-price rate cards by level
r = 4
ws.cell(r, 1, 'D1. Rate cards by level (list prices)').font = Font(bold=True, size=11)
r += 1
top = ['Firm', 'Rate date', 'Source', 'Partner', '', 'Counsel / Senior associate', '', 'Associate', '', 'Junior associate', '', 'Paralegal',
       'Partner mid', 'Junior mid', 'Partner / junior (x)', 'Level definitions in source', 'Notes']
hdr(r, top)
for a, b in ((4, 5), (6, 7), (8, 9), (10, 11)):
    ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
r += 1
hdr(r, ['', '', '', 'Low', 'High', 'Low', 'High', 'Low', 'High', 'Low', 'High', '', '', '', '', '', ''], SUB, Font(bold=True, size=9))
for c in (1, 2, 3, 12, 13, 14, 15, 16, 17):
    ws.merge_cells(start_row=r - 1, start_column=c, end_row=r, end_column=c)
ws.row_dimensions[r - 1].height = 42
CARDS = [
    ['Mannheimer Swartling', 'Not stated (provided by colleague, Oct-26)', 'Firm rate card (client-provided snip)',
     6000, 10000, 4550, 8000, 3650, 4150, 2750, 3350, None,
     'Junior Assoc 0-24 months; Associates 25-59 months; Senior Assoc / PDL / Specialist Counsel 60+ months; Partners / Senior Advisor',
     'Counsel band (4,550-8,000) overlaps partner band. Date to be confirmed.'],
    ['Cederquist', '2025 (revised annually 1 Jan)', 'Firm cost estimate / fee structure (client-provided snip)',
     6500, 8000, 4500, 5800, 3600, 4100, 2300, 3100, 1500,
     'Partners; Senior Associates; Associates; Junior Associates; Paralegals',
     'Time-incurred basis; open to alternative fee arrangements per project.'],
    ['MAQS', 'Sell-side model (Oct-2026)', 'Sell-side model (vendor financial model)',
     5014.83020687709, 5014.83020687709, None, None, 2753.06754681269, 2753.06754681269, None, None, None,
     'Partner; Associate (single rates)',
     'Single model rates, not a published rate card range; confirm whether list or realised (effective) rates.'],
]
first = r + 1
for card in CARDS:
    r += 1
    for c, v in enumerate(card[:12], 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c <= 3 else C
        if c >= 4: x.number_format = '#,##0'
    ws.cell(r, 13, f'=AVERAGE(D{r}:E{r})')
    if card[9]:
        ws.cell(r, 14, f'=AVERAGE(J{r}:K{r})'); ws.cell(r, 15, f'=M{r}/N{r}')
    for c, fmt in ((13, '#,##0'), (14, '#,##0'), (15, '0.0')):
        x = ws.cell(r, c); x.number_format = fmt; x.border = B; x.alignment = C
    for c, v in ((16, card[12]), (17, card[13])):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W
    ws.row_dimensions[r].height = 60
r += 1
ws.cell(r, 1, 'Difference: Mannheimer Swartling vs Cederquist (mid-points)').font = Font(bold=True)
maqs_row = first + 2
for col in ('D', 'F', 'H', 'J'):
    nxt = get_column_letter(ws[col + '1'].column + 1)
    ws[f'{col}{r}'] = f'=AVERAGE({col}{first}:{nxt}{first})/AVERAGE({col}{first+1}:{nxt}{first+1})-1'
    ws[f'{col}{r}'].number_format = '+0%;-0%'; ws[f'{col}{r}'].alignment = C
r += 1
ws.cell(r, 1, 'MAQS (sell-side model) vs Cederquist mid-point').font = Font(bold=True)
for col in ('D', 'H'):
    nxt = get_column_letter(ws[col + '1'].column + 1)
    ws[f'{col}{r}'] = f'={col}{maqs_row}/AVERAGE({col}{first+1}:{nxt}{first+1})-1'
    ws[f'{col}{r}'].number_format = '+0%;-0%'; ws[f'{col}{r}'].alignment = C
r += 1
ws.cell(r, 1, 'MAQS (sell-side model) vs Mannheimer Swartling mid-point').font = Font(bold=True)
for col in ('D', 'H'):
    nxt = get_column_letter(ws[col + '1'].column + 1)
    ws[f'{col}{r}'] = f'={col}{maqs_row}/AVERAGE({col}{first}:{nxt}{first})-1'
    ws[f'{col}{r}'].number_format = '+0%;-0%'; ws[f'{col}{r}'].alignment = C
r += 2

# ---- D2 indicative firm-wide ranges
ws.cell(r, 1, 'D2. Indicative firm-wide hourly ranges (third-party directory, low reliability)').font = Font(bold=True, size=11)
r += 1
hdr(r, ['Firm', 'Accessed', 'Source', 'Low (~associate)', 'High (~partner)', 'Mid'])
SK = [('Mannheimer Swartling', 4500, 8000), ('Roschier', 4200, 7000), ('Vinge', 4000, 7000), ('DLA Piper', 4000, 7000), ('Schjødt', 3800, 6500),
      ('Cederquist', 3800, 6500), ('Setterwalls', 3500, 6000), ('Bird & Bird', 3500, 6000), ('Delphi', 3200, 5500), ('Lindahl', 3000, 5500)]
for f, lo, hi in SK:
    r += 1
    for c, v in enumerate([f, 'Oct-2026 (undated on site)', 'skatterätt.se firm pages (tax practice; no level split, no source)', lo, hi], 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c <= 3 else C
        if c >= 4: x.number_format = '#,##0'
    x = ws.cell(r, 6, f'=AVERAGE(D{r}:E{r})'); x.number_format = '#,##0'; x.border = B; x.alignment = C
r += 1
ws.cell(r, 1, 'Example: https://www.xn--skattertt-12a.se/byra/vinge-skatt/ (Vinge: "Timarvode 4 000 till 7 000 kr"). Ranges may be estimates; use only for relative positioning.').font = N
r += 2

# ---- D3 dated data points
ws.cell(r, 1, 'D3. Dated Nordic rate data points, 2023 onwards (realised rates and market commentary)').font = Font(bold=True, size=11)
r += 1
hdr(r, ['Firm', 'Rate date', 'Level (as stated)', 'Mapped level', 'Rate low', 'Rate high', 'Context', 'Source type', 'Relevance', 'Source'])
P = [
    ['Mannheimer Swartling', 'Nov-2024 to Mar-2025', 'Average across team (~6,000 h, SEK 32m)', 'Blended (realised)', 4000, 4000, 'Northvolt restructuring (Swedish counsel in US Chapter 11); lowest average of all advisers', 'Court-approved fee filing (press)', 'Only published realised blended rate for a top Nordic firm: ~4,000 vs list partner 6,000-10,000', 'https://www.affarsvarlden.se/artikel/sa-dyrt-blev-northvolts-raddningsforsok-i-usa-debiteras-47-000-timmar'],
    ['Top-tier firms (Vinge, MSA, Roschier)', '2026', 'Partner', 'Partner', 7000, 9000, 'Olle Flygt (ex-Vinge/MSA partner): partner bills 7-9k with 3-5 associates beneath', 'Market commentary', 'Current top-tier partner rate level', 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
    ['Top-tier firms (Vinge, MSA, Roschier)', '2026', 'Biträdande jurist (associate)', 'Associate', 3000, 3000, 'As above: associates bill ~3,000 and cost at most a third of that', 'Market commentary', 'Associate rate and margin (leverage)', 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
    ['Mannheimer Swartling', 'Final years before 2025 (not specified)', 'Partner (Olle Flygt)', 'Partner', 7000, 8000, 'Own billing rate', 'Market commentary', 'Individual partner rate at MSA', 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/'],
    ['Law firm (unnamed)', 'Apr-2026', 'Hourly rate held reasonable', 'Blended', 4000, 4000, 'Luleå tingsrätt cost ruling', 'Court ruling (press)', 'Courts accept ~4,000/h as reasonable for business-law counsel', 'https://www.realtid.se/juridik/ratten-om-timarvoden-4-000-kronor-i-timmen-ar-ok/'],
]
for row in P:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c not in (5, 6) else C
        if c in (5, 6): x.number_format = '#,##0'
r += 1
ws.cell(r, 1, 'Scope: Nordic business-law firms only, 2023 onwards.').font = N
r += 2

# ---- D4 benchmarks
ws.cell(r, 1, 'D4. Benchmarks (Nordic business law)').font = Font(bold=True, size=11)
r += 1
hdr(r, ['Benchmark', 'Period', 'Rate', 'Currency', 'Notes', 'Source'])
BM = [
    ['Affärsvärlden fee income per lawyer (panel)', '2024', 4130000, 'SEK p.a.', 'Implies ~SEK 2,580/h at 1,600 billable hours or ~2,950/h at 1,400 (illustrative)', 'https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
    ['Affärsvärlden fee income per lawyer (panel)', '2023', 4220000, 'SEK p.a.', '', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-om-trenden-pe-bolagen-kommer-vara-motorn'],
    ['Norway: commercial litigation rates', '2026', 3000, 'NOK/h (earlier level)', 'Norwegian Government Attorney: rates of ~NOK 3,000/h have "doubled in a short time"', 'https://www.realtid.se/juridik/regeringsadvokaten-slar-larm-advokatkostnaderna-skenar/'],
]
for row in BM:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c != 3 else C
        if c == 3: x.number_format = '#,##0'
r += 2
for t in ['Gaps: no published rate cards found for Vinge, White & Case, Setterwalls, Lindahl, Gernandt & Danielsson, CMS Wistrand, Cirio, Baker McKenzie, Snellman, Foyen, Glimstedt, Eversheds, Fylgia, Hammarskiöld, Magnusson, Kahn Pedersen, Wigge, Kastell, Westerberg or TM & Partners.',
          'Routes to fill: client invoices (e.g. via Axo, requested); expert calls with Nordic GCs; confirm whether MAQS sell-side rates are list or realised.']:
    ws.cell(r, 1, t).font = N; r += 1

for c, w in zip(range(1, 18), (26, 20, 28, 10, 10, 12, 12, 10, 10, 10, 10, 10, 10, 10, 10, 40, 40)):
    ws.column_dimensions[get_column_letter(c)].width = w
wb.save(path)
print('ok', wb.sheetnames)
