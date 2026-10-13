"""Build Section A sheets (market size, firm concentration, growth drivers) into the workbook."""
import json, sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

path = sys.argv[1]
wb = load_workbook(path)
for n in ('A. Market size', 'A. Growth drivers'):  # 'A. Firms' merged into 'C. Benchmarking'
    if n in wb.sheetnames: del wb[n]

H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
T = Font(bold=True, size=14); S = Font(bold=True, size=11); N = Font(italic=True, size=9)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')

def table(ws, r, headers, rows, widths=None):
    for c, h in enumerate(headers, 1):
        x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center'); x.border = BOX
    for row in rows:
        r += 1
        for c, v in enumerate(row, 1):
            x = ws.cell(r, c, v); x.border = BOX; x.alignment = WRAP
    if widths:
        for c, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(c)].width = w
    return r + 2

# ---------------- A. Market size ----------------
ws = wb.create_sheet('A. Market size', 0)
ws['A1'] = 'Section A. Swedish (corporate) legal market: size'; ws['A1'].font = T
ws['A3'] = 'A1. Headline estimates'; ws['A3'].font = S
r = table(ws, 4, ['Scope', 'Figure', 'Year', 'Source'], [
    ['Business law: fee income, 56 largest firms (Affärsvärlden panel)', '~SEK 16bn (+5.3%)', '2025', 'https://www.affarsvarlden.se/artikel/affarsjuridiken-vaxer-igen-men-jattarna-tappar-mark'],
    ['Same panel, prior year (revised Oct-25, split FYs)', 'SEK 15.4bn (+8.6%); SEK 4.13m fee per lawyer', '2024', 'https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
    ['Business law market, total (industry estimate, method not given)', '>SEK 25bn p.a.', 'n/a', 'https://www.affarsvarlden.se/debatt/okad-konkurrens-starker-svensk-affarsjuridik'],
    ['All legal services (MarketLine)', 'USD 3.1bn (CAGR 5.9% 2019-24)', '2024', 'https://store.marketline.com/report/legal-services-in-sweden/'],
    ['All legal activities (IBISWorld)', 'EUR 3.5bn (CAGR -0.5% 2020-25)', '2026e', 'https://www.ibisworld.com/sweden/industry/legal-activities/200283/'],
    ['All law firms (advokatbyråer), estimate from SCB', '~SEK 29bn', '2024e', '74.5% share of SNI 69.1 in 2021 applied to 2024 (SCB, below)'],
], [58, 40, 9, 70])

ws.cell(r, 1, 'A2. Official statistics: SCB Företagens ekonomi, net turnover (SEK m)').font = S
rows = [[2010, 18738, 13904, 1956, 2879], [2015, 23986, 17936, 2750, 3300], [2019, 31237, 22277, 4194, 4765],
        [2020, 32628, 23298, 4371, 4958], [2021, 34693, 25864, 4558, 4270],
        ['2022*', 33276, 'n/p', 'n/p', 'n/p'], ['2023*', 36525, 'n/p', 'n/p', 'n/p'], ['2024*', 38921, 'n/p', 'n/p', 'n/p']]
r = table(ws, r + 1, ['Year', '69.1 Legal activities (total)', '69.101 Advokatbyråer', '69.102 Other legal offices', '69.103 Patent agents'], rows)
for rr in range(r - len(rows) - 2, r - 1):
    for c in range(2, 6): ws.cell(rr, c).number_format = '#,##0'
for t in ['* New SCB series from 2022 (NV0109P): 2022 onward is not directly comparable to 2021; sub-industry splits not published (n/p).',
          '69.1 in 2024: 6,965 companies, 13,446 employees; production value SEK 39.1bn, value added SEK 25.3bn, EBITDA margin 28.2%.',
          '69.101 turnover CAGR 2010-21: ~5.8%. Source: https://api.scb.se/OV0104/v1/doris/sv/ssd/NV/NV0109/ (BasfaktaFEngs07, NSEBasfaktaFEngs07, NSEBasfaktaVEngs07N).']:
    ws.cell(r - 1, 1, t).font = N; r += 1
ws.column_dimensions['B'].width = 40; ws.column_dimensions['C'].width = 20; ws.column_dimensions['D'].width = 24; ws.column_dimensions['E'].width = 20

# ---------------- A. Growth drivers ----------------
ws = wb.create_sheet('A. Growth drivers', 1)
ws['A1'] = 'A3. Fastest growers and growth drivers'; ws['A1'].font = T
ws['A2'] = 'Driver tags: Lateral hire / M&A / New office / Organic. Only drivers found in a cited source are listed.'; ws['A2'].font = N
G = [
    ['Carthiel', '+155%', '2024', 'M&A', 'Merged with Grey Advokatbyrå (operating from Feb-2024), adding a disputes team. Small base (~10 people).', 'https://www.carthiel.se/post/advokatfirman-carthiel-g%C3%A5r-ihop-med-grey-advokatbyr%C3%A5'],
    ['White & Case (Stockholm)', '+29% (AB net sales 859 to 1,108)', '2025', 'Organic', 'Fastest grower in 2023 (+24%) on PE deals and selective hiring; "doubled revenue" coverage Nov-25. Lost Magnus Wennerhorn to Linklaters (Sep-24).', 'https://www.realtid.se/juridik/white-case-dubblade-omsattningen-var-styrka/'],
    ['Schjødt (Sweden)', '+24% / ~+25%', '2024 / 2025', 'Lateral hire', 'Tom Wehtje (M&A) joined 2024 from Mannheimer Swartling after 15 years as partner; Karl Klackenberg (M&A/PE) reportedly from Vinge May-24 (unconfirmed).', 'https://www.legal500.com/firms/10502-advokatfirmaet-schjodt-as/c-sweden/lawyers/2101480-tom-wehtje'],
    ['', '', '', 'Organic', 'Demand for pan-Scandinavian cross-border offering, FDI and regulatory work.', 'https://www.affarsvarlden.se/artikel/vinge-gar-om-msa-sa-gar-det-for-affarsjuristerna'],
    ['', '', '', 'M&A (historic)', 'Entered Sweden via merger with Hamilton (2019/20). Risk: 54 people (23 partners, group-wide) left for BAHR, which opened Stockholm Jan-26.', 'https://www.realtid.se/juridik/oppnar-stockholmskontor-efter-massvarvning-fran-rivalen/'],
    ['Roschier (Sweden)', '+23% / +22%', 'FY23/24 / FY24/25', 'Organic', 'Swedish transaction market pick-up, larger mandates, high disputes demand, closer Finland-Sweden integration; Northvolt Expansion Ett restructuring. Stockholm now larger than Helsinki.', 'https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
    ['MAQS', '+21%', '2024', 'Organic', 'Strong demand in M&A, capital markets and disputes (CEO Katrin Troedsson).', 'https://www.affarsvarlden.se/artikel/vinge-gar-om-msa-sa-gar-det-for-affarsjuristerna'],
    ['', '', '', 'Lateral hire', 'Strategic sustainability hires, e.g. Anna Domander (ex head of legal, Systembolaget); date not given.', 'https://www.affarsvarlden.se/artikel/vinge-gar-om-msa-sa-gar-det-for-affarsjuristerna'],
    ['Vinge', '+19%', '2024', 'Organic', 'Broad client activity and full-service model; accounted for ~25% of total market growth. Then -2% in 2025 ("2024 exceptional").', 'https://www.affarsvarlden.se/artikel/vinge-gar-om-msa-sa-gar-det-for-affarsjuristerna'],
    ['Gernandt & Danielsson', '+19% / +12%', 'FY23/24 / FY24/25', 'Lateral hire', 'Linklaters Stockholm real estate team: partners Magnus Lidman and Anna Eriksson + 5 associates (effective 1 Jun 2024, G&D press release).', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-gar-starkt-finns-fler-bolag-i-ipo-pipelinen'],
    ['', '', '', 'Organic', 'Strong bond and M&A teams.', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-gar-starkt-finns-fler-bolag-i-ipo-pipelinen'],
    ['Cirio', '~+15%', '2025', 'Lateral hire', '12-person real estate and construction team from AG Advokat from Jan-25, incl. partners Johan Lindberg, Johan Rosén, Måns Derk; new construction disputes offering.', 'https://www.realtid.se/juridik/cirio-expanderar-inom-fastighets-och-entreprenadjuridik/'],
    ['', '', '', 'Organic', 'Investment in real estate transactions and insolvency.', 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/'],
    ['Setterwalls', '+2% / ~+9% (passed SEK 1bn)', '2024 / 2025', 'Organic', 'Large complex mandates: public M&A and PE take-privates (MedHelp, Integrum, Norva24), bank finance, disputes; lawyers +10% to 235 in 2024. Malmö office left for AGRD in Sep-26.', 'https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
    ['Fylgia', 'n/a (record 2025)', '2025', 'Organic', 'Transactions, construction/real estate, public procurement, restructuring/insolvency; four new partners appointed.', 'https://www.fylgia.se/en/news/following-a-record-breaking-year-fylgia-appoints-four-new-partners'],
    ['AGRD Partners (Axcel)', 'n/a (~SEK 660m combined)', '2025', 'M&A', 'PE-backed roll-up of six firms in Aug-25: Allié, Born, Morris Law, Next, Synch, TM & Partners (~130 lawyers gave up advokat title). Added Setterwalls Malmö (Ascente Law) Sep-26.', 'https://www.realtid.se/juridik/setterwalls-malmokontor-gar-till-riskkapitalagda-agrd/'],
    ['DLA Piper (Sweden)', '-9%', 'FY24/25', 'Decline', 'Consolidation year: generational shift and hesitant deal market; restructuring group busy (Northvolt, Oscar Properties, SAS).', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-gar-starkt-finns-fler-bolag-i-ipo-pipelinen'],
]
table(ws, 4, ['Firm', 'Growth', 'Period', 'Driver', 'Detail', 'Source'], G, [24, 22, 16, 14, 80, 60])
ws.cell(6 + len(G), 1, 'Market context: 2025 growth came from international transactions, disputes and regulatory work; IPOs were the weakest segment. Firms expect +6.8% in 2026 (Affärsvärlden/Realtid).').font = N

wb.move_sheet('A. Market size', -wb.index(wb['A. Market size']))
order = ['A. Market size', 'A. Nordic market', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'B. Top lawyers', 'B. Partner pipeline', 'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts', 'G. MAQS key data']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.active = 0
wb.save(path)
print('ok', wb.sheetnames)
