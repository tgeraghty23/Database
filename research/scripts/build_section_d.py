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
            'list price (rate card) > court/Chapter 11 filings (actual billed) > public tender (discounted).')
ws['A2'].alignment = W; ws.merge_cells('A2:R2'); ws.row_dimensions[2].height = 45

# ---- D1 list-price rate cards by level
r = 4
ws.cell(r, 1, 'D1. Rate cards by level (list prices)').font = Font(bold=True, size=11)
r += 1
top = ['Firm', 'Rate date', 'Source', 'Partner', '', 'Counsel / Senior associate', '', 'Associate', '', 'Junior associate', '', 'Paralegal',
       'Partner mid', 'Junior mid', 'Partner / junior (x)', 'Partner mid vs legal-aid norm 2025 (x)', 'Level definitions in source', 'Notes']
hdr(r, top)
for a, b in ((4, 5), (6, 7), (8, 9), (10, 11)):
    ws.merge_cells(start_row=r, start_column=a, end_row=r, end_column=b)
r += 1
hdr(r, ['', '', '', 'Low', 'High', 'Low', 'High', 'Low', 'High', 'Low', 'High', '', '', '', '', '', '', ''], SUB, Font(bold=True, size=9))
for c in (1, 2, 3, 12, 13, 14, 15, 16, 17, 18):
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
]
first = r + 1
for card in CARDS:
    r += 1
    for c, v in enumerate(card[:12], 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c <= 3 else C
        if c >= 4: x.number_format = '#,##0'
    ws.cell(r, 13, f'=AVERAGE(D{r}:E{r})'); ws.cell(r, 14, f'=AVERAGE(J{r}:K{r})')
    ws.cell(r, 15, f'=M{r}/N{r}'); ws.cell(r, 16, f'=M{r}/1586')
    for c, fmt in ((13, '#,##0'), (14, '#,##0'), (15, '0.0'), (16, '0.0')):
        x = ws.cell(r, c); x.number_format = fmt; x.border = B; x.alignment = C
    for c, v in ((17, card[12]), (18, card[13])):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W
    ws.row_dimensions[r].height = 60
r += 1
ws.cell(r, 1, 'Difference: Mannheimer Swartling vs Cederquist (mid-points)').font = Font(bold=True)
for col in ('D', 'F', 'H', 'J'):
    nxt = get_column_letter(ws[col + '1'].column + 1)
    ws[f'{col}{r}'] = f'=AVERAGE({col}{first}:{nxt}{first})/AVERAGE({col}{first+1}:{nxt}{first+1})-1'
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
ws.cell(r, 1, 'D3. Dated rate data points: invoices, court filings, tenders, market commentary').font = Font(bold=True, size=11)
r += 1
hdr(r, ['Firm', 'Rate date', 'Level (as stated)', 'Mapped level', 'Rate low', 'Rate high', 'Context / client', 'Source type', 'Discounted / public?', 'Source'])
P = [
    ['Mannheimer Swartling', 'Nov-2010', 'Delägare', 'Partner', 3150, 3150, 'SL (Stockholm public transport), C30 metro procurement', 'Invoice (press)', 'Yes: SL framework + 10% invoice discount', 'https://www.realtid.se/bors-finans/mannheimers-taxa-till-sl-1906-krh/'],
    ['Mannheimer Swartling', 'Nov-2010', 'Biträdande jurist (advokat)', 'Associate / Senior associate', 2150, 2150, 'SL', 'Invoice (press)', 'Yes', 'https://www.realtid.se/bors-finans/mannheimers-taxa-till-sl-1906-krh/'],
    ['Mannheimer Swartling', 'Nov-2010', 'Junior biträdande jurist', 'Junior associate', 1550, 1550, 'SL; blended 2,118 (1,906 after discount)', 'Invoice (press)', 'Yes', 'https://www.realtid.se/bors-finans/mannheimers-taxa-till-sl-1906-krh/'],
    ['Mannheimer Swartling', 'Oct-2009', 'Partner (chairman)', 'Partner', 3200, 3200, 'Försvarsmakten, Projekt PRIO', 'Invoice (press)', 'Public client; discount unknown', 'https://www.realtid.se/sa-hog-ar-mannheimers-timtaxa-mot-forsvarsmakten/'],
    ['Mannheimer Swartling', 'Jan-2010', 'Not stated (partner-led team)', 'Partner / blended', 3800, 3800, 'Riksgälden, Carnegie valuation dispute', 'Invoice (press)', 'Public client; direct award', 'https://www.realtid.se/bors-finans/bo-lundgrens-advokatnota-21-mkr/'],
    ['DLA Piper Sweden', '2012-2017 (claimed Jun-2017)', 'Ombud (counsel team)', 'Partner / senior', 3400, 3900, 'HQ AB v. former directors, Stockholm District Court', 'Court cost claim', 'No', 'https://www.realtid.se/bors-finans/hq-processens-prislapp-441-miljoner-kronor/'],
    ['Roschier', '2012-2017 (claimed 2017)', 'Implied blended (SEK 17.5m / 5,207 h)', 'Blended', 3360, 3360, 'HQ AB (claimant)', 'Court cost claim (calc.)', 'No', 'https://www.realtid.se/bors-finans/hq-processens-prislapp-441-miljoner-kronor/'],
    ['Mannheimer Swartling', 'Nov-2024 to Mar-2025', 'Average across team (~6,000 h, SEK 32m)', 'Blended', 4000, 4000, 'Northvolt US Chapter 11 (Swedish counsel); lowest average of advisers', 'Chapter 11 filing (press)', 'No', 'https://www.affarsvarlden.se/artikel/sa-dyrt-blev-northvolts-raddningsforsok-i-usa-debiteras-47-000-timmar'],
    ['Kirkland & Ellis (comparison)', 'Nov-Dec-2024', 'Average / top rate', 'Blended / top partner', 14300, 25500, 'Northvolt Chapter 11 (US lead counsel)', 'Chapter 11 filing (press)', 'No', 'https://www.affarsvarlden.se/artikel/northvolts-nota-over-100-miljoner-i-manaden-pa-advokater-och-radgivare'],
    ['Front Advokater', '2016-2020', 'Junior / Jurist / Senior (>=2/5/10 yrs)', 'Junior / Associate / Senior', 953, 1588, 'City of Gothenburg legal services framework (area A: 953/1,270/1,588; area B: 833/1,110/1,388)', 'Public tender (court judgment)', 'Yes: aggressive tender pricing', 'https://upphandling24.se/wp-content/uploads/2024/03/goteborgs-tr-t-20148-21-dom-2024-02-21.pdf'],
    ['7 unnamed firms', '2007', 'Tender price', 'Blended', 1800, 2800, 'State tender: advisers for divesting state-owned companies', 'Public tender (press)', 'Yes', 'https://www.realtid.se/bors-finans/det-blir-13000-kronor-tack-per-timme/'],
    ['Unnamed external advokat', 'Dec-2024', 'Advokat', 'Partner', 4500, 4500, 'Botkyrka municipality', 'Press', 'Public client', 'https://www.svt.se/nyheter/lokalt/stockholm/botkyrka-tar-in-extern-advokat-for-4-500-kronor-i-timmen'],
    ['Law firm (unnamed)', 'Apr-2026', 'Hourly rate held reasonable', 'n/a', 4000, 4000, 'Luleå tingsrätt cost ruling', 'Court ruling (press)', 'n/a', 'https://www.realtid.se/juridik/ratten-om-timarvoden-4-000-kronor-i-timmen-ar-ok/'],
    ['Top-tier firms (market)', '2026', 'Partner', 'Partner', 7000, 9000, 'Olle Flygt (ex-Vinge/MSA partner): partners bill 7,000-9,000; associates ~3,000', 'Market commentary', 'No', 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
    ['Top-tier firms (market)', '2026', 'Associate', 'Associate', 3000, 3000, 'As above', 'Market commentary', 'No', 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
    ['Mannheimer Swartling', 'Final years before 2025 (not specified)', 'Partner (Olle Flygt)', 'Partner', 7000, 8000, 'Own billing rate', 'Market commentary', 'No', 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/'],
]
for row in P:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c not in (5, 6) else C
        if c in (5, 6): x.number_format = '#,##0'
r += 2

# ---- D4 benchmarks
ws.cell(r, 1, 'D4. Benchmarks').font = Font(bold=True, size=11)
r += 1
hdr(r, ['Benchmark', 'Period', 'Rate', 'Currency', 'Notes', 'Source'])
BM = [
    ['Legal-aid hourly norm (timkostnadsnorm), F-skatt', '2025', 1586, 'SEK', 'SFS 2024:913', 'https://www.svenskforfattningssamling.se/sites/default/files/sfs/2024-10/SFS2024-913.pdf'],
    ['Legal-aid hourly norm (timkostnadsnorm), F-skatt', '2026', 1626, 'SEK', 'SFS 2025:919 (+2.5%). Bar Association: business lawyers bill "at least three times" this', 'https://xn--svenskfrfattningssamling-roc.se/sites/default/files/sfs/2025-10/SFS2025-919.pdf'],
    ['Affärsvärlden fee income per lawyer (56-63 firms)', '2024', 4130000, 'SEK p.a.', '~SEK 2,580/h at 1,600 billable hours (illustrative)', 'https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
    ['Affärsvärlden fee income per lawyer', '2023', 4220000, 'SEK p.a.', '', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-om-trenden-pe-bolagen-kommer-vara-motorn'],
    ['Denmark: Kammeradvokaten state agreement, partner', 'From Jan-2024', 2790, 'DKK', 'Advokat >3 yrs 2,263; <3 yrs 1,891; fuldmægtig 1,395 (snippet only)', 'https://ft.dk/samling/20241/almdel/reu/bilag/88/2953293.pdf'],
    ['Norway: public legal-aid rate (salærsats)', '2026', 1375, 'NOK', '2025: NOK 1,315', 'https://www.advokatforeningen.no/contentassets/770460b5483c4bd0acb76ff98d2ad3fa/vedlegg-1--salarradets-anbefaling-for-2026.pdf'],
    ['UK top-10 firms average hourly rate', '2024', 449, 'GBP', '2019: GBP 321; firms 11-25: GBP 325 (PwC, snippet only)', 'https://www.legalcheek.com/2025/04/the-2000-per-hour-solicitor'],
]
for row in BM:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c != 3 else C
        if c == 3: x.number_format = '#,##0'
r += 2
for t in ['Gaps: no published rate cards found for Vinge, White & Case, Setterwalls, Lindahl, Gernandt & Danielsson, CMS Wistrand, Cirio, MAQS, Baker McKenzie, Snellman, Foyen, Glimstedt, Eversheds, Fylgia, Hammarskiöld, Magnusson, Kahn Pedersen, Wigge, Kastell, Westerberg or TM & Partners.',
          'Routes to fill: (1) US Chapter 11 fee statements (PACER) for SAS (S.D.N.Y. 22-10925) and Northvolt (S.D. Tex. 24-90577) list per-timekeeper rates for Mannheimer Swartling 2022-25; (2) FOI requests to public buyers (price annexes often withheld); (3) client invoices (e.g. Axo, requested).']:
    ws.cell(r, 1, t).font = N; r += 1

for c, w in zip(range(1, 19), (26, 20, 28, 10, 10, 12, 12, 10, 10, 10, 10, 10, 10, 10, 10, 13, 40, 40)):
    ws.column_dimensions[get_column_letter(c)].width = w
wb.save(path)
print('ok', wb.sheetnames)
