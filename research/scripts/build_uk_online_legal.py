"""Build the UK online / subscription legal services market map workbook.

Usage: python3 research/scripts/build_uk_online_legal.py research/uk_online_legal_services_market_map.xlsx
Reads research/data/uk_online_legal_*.json (compiled 2026-10-12).
"""
import json, sys, os, re
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

out = sys.argv[1]
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data')
def load(n):
    with open(os.path.join(D, n), encoding='utf-8') as f: return json.load(f)

SME = load('uk_online_legal_sme_subscription.json')
HR = load('uk_online_legal_hr_employment.json')
CON = load('uk_online_legal_consumer_documents.json')
CTX = load('uk_online_legal_context.json')

H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
T = Font(bold=True, size=14); S = Font(bold=True, size=11); N = Font(italic=True, size=9)
thin = Side(style='thin', color='BFBFBF'); BOX = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical='top')
TIER_FILL = {'Core': PatternFill('solid', fgColor='E2EFDA'), 'Adjacent': PatternFill('solid', fgColor='FFF2CC'),
             'Defunct / exited': PatternFill('solid', fgColor='F8CBAD'), 'Out of scope': PatternFill('solid', fgColor='EDEDED')}

def j(v):
    if v is None: return None
    if isinstance(v, list): return '\n'.join(str(x) for x in v if x is not None)
    return v

def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

def first_int(s):
    """First integer in a free-text headcount string (e.g. 'Avg 65 (62 staff...)')."""
    if s is None: return None
    if isinstance(s, (int, float)): return int(s)
    m = re.search(r'(\d[\d,]*)', str(s))
    return int(m.group(1).replace(',', '')) if m else None

def table(ws, r, headers, rows, widths=None, freeze=False, tier_col=None):
    for c, h in enumerate(headers, 1):
        x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center'); x.border = BOX
    ws.row_dimensions[r].height = 32
    if freeze: ws.freeze_panes = ws.cell(r + 1, 3)
    for row in rows:
        r += 1
        for c, v in enumerate(row, 1):
            x = ws.cell(r, c, j(v)); x.border = BOX; x.alignment = WRAP
            if num(v) is not None:
                x.number_format = '0.0%' if (headers[c-1].endswith('%') and abs(v) < 20) else ('#,##0.0' if isinstance(v, float) else '#,##0')
            if tier_col is not None and c == tier_col and v in TIER_FILL: x.fill = TIER_FILL[v]
    if widths:
        for c, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(c)].width = w
    return r + 2

def title(ws, text, sub=None):
    ws['A1'] = text; ws['A1'].font = T
    if sub: ws['A2'] = sub; ws['A2'].font = N

def growth(cur, prior):
    cur, prior = num(cur), num(prior)
    return (cur / prior - 1) if cur and prior else None

# ---- Relevance tiers vs LegalVision (hand-assigned) ----------------------------------------
TIER = {
 # core SME subscription legal
 'LegalVision UK': ('Core', 'Subscription law firm (benchmark)'),
 'Harper James': ('Core', 'Subscription law firm'),
 'Lawyerly (extra)': ('Core', 'Subscription law firm'),
 'Blackmont Legal (extra)': ('Core', 'Subscription law firm'),
 '360 Business Law / 360 Law Group': ('Core', 'Subscription legal dept (consultant lawyers)'),
 'Law 365': ('Core', 'Subscription law firm (hours-based)'),
 'LegalEdge': ('Core', 'Fractional GC retainer'),
 'Sprintlaw UK (extra)': ('Core', 'Fixed-fee online + membership (unregulated)'),
 'Rocket Lawyer UK': ('Core', 'Legal plan membership (docs + lawyer Q&A)'),
 'LawBite (defunct)': ('Defunct / exited', 'Freemium platform + ABS (administration 2024)'),
 'Linkilaw': ('Defunct / exited', 'Exited SME segment'),
 'SeedLegals': ('Adjacent', 'Startup legal SaaS + unlimited support'),
 'Genie AI': ('Adjacent', 'AI contract SaaS (self-serve)'),
 'Docue (UK)': ('Adjacent', 'Contract SaaS (self-serve)'),
 'Sparqa Legal': ('Adjacent', 'Self-serve docs (dormant)'),
 'Farillio': ('Adjacent', 'Self-serve docs, B2B2C via insurers'),
 'Pocketlaw / Miramis (UK)': ('Adjacent', 'Pivoted to enterprise CLM'),
 'Lawhive': ('Adjacent', 'Fixed-fee marketplace/ABS (consumer-led)'),
 'Summize (CLM SaaS - adjacent)': ('Out of scope', 'Enterprise CLM'),
 'Juro (CLM SaaS only)': ('Out of scope', 'Enterprise CLM'),
 'Legl (law-firm SaaS - adjacent)': ('Out of scope', 'Law-firm SaaS'),
 'Other subscription offers noted (not full profiles)': ('Adjacent', 'Long tail'),
 # HR / employment law
 'Peninsula Group (Peninsula Business Services)': ('Core', 'Employment-law/HR subscription (insurance-backed)'),
 'Citation Group': ('Core', 'Employment-law/HR/H&S subscription (PE-owned)'),
 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC': ('Core', 'Employment-law/HR subscription (PE-owned)'),
 'Avensure': ('Core', 'Employment-law/HR subscription'),
 'Croner (Peninsula Group)': ('Core', 'Employment-law/HR subscription'),
 'Wirehouse Employer Services (WorkNest / Axiom GRC)': ('Core', 'Employment-law/HR subscription'),
 'Quest Cover (Quest Business Services, formerly Qdos Consulting)': ('Core', 'HR/H&S advice lines, white-label'),
 'NatWest Mentor (being sold to Empowering People Group)': ('Core', 'Employment-law/HR subscription (bank channel)'),
 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)': ('Core', 'ABS running B2B2C legal helplines'),
 'Lawgistics (motor trade legal subscription)': ('Core', 'Vertical legal membership'),
 'The HR Dept (franchise network)': ('Adjacent', 'Franchised HR retainers'),
 'Stallard Kane (Ardonagh Advisory)': ('Adjacent', 'H&S/HR consultancy in broker'),
 'BrightHR (Peninsula Group)': ('Adjacent', 'HR SaaS with advice lines'),
 'Croner-i (Peninsula Group)': ('Adjacent', 'Compliance content subscription'),
 'Health Assured (Peninsula Group)': ('Out of scope', 'EAP'),
 'ELAS (Employment Law Advisory Services) - now part of Citation': ('Defunct / exited', 'Absorbed into Citation'),
 'FSB Legal Hub (Federation of Small Businesses member legal benefits)': ('Adjacent', 'Distribution channel'),
 'Moorepay (Zellis Group) - payroll with HR/employment law advice': ('Adjacent', 'Payroll-led'),
 'ARAG UK (incl. former DAS) / Howden - legal expenses insurance with SME legal helplines (brief)': ('Adjacent', 'Legal expenses insurance'),
 'Empowering People Group (AdviserPlus / Halborns / Learning Nexus; acquirer of NatWest Mentor)': ('Adjacent', 'PE-backed HR/employment-law group'),
 'Safe Workplace (not identified)': ('Out of scope', 'Not identified'),
 # consumer / documents
 'Farewill': ('Core', 'Digital wills/probate (acquired by Dignity)'),
 'Octopus Legacy (formerly Guardian Angel)': ('Core', 'Digital wills/probate (Octopus)'),
 'amicable': ('Core', 'Online divorce (Octopus)'),
 'Divorce-Online': ('Core', 'Online divorce + ABS'),
 'Wenup': ('Core', 'Online prenups (acquired by Blackfinch)'),
 'Which? Legal': ('Core', 'Consumer legal advice subscription'),
 'JustAnswer UK': ('Adjacent', 'Expert Q&A subscription (US)'),
 'Co-op Legal Services': ('Adjacent', 'Scaled fixed-fee probate/wills ABS'),
 'Legal Utopia': ('Defunct / exited', 'In liquidation 2026'),
 'Slater and Gordon UK (brief)': ('Out of scope', 'National consumer law firm'),
 'ARAG Law (formerly DAS Law) (brief)': ('Adjacent', 'Captive LEI law firm'),
 'Juno (conveyancing)': ('Adjacent', 'Digital conveyancing'),
 'Muve (Connect2Law)': ('Adjacent', 'Digital conveyancing (Alcuin)'),
 'Kinherit': ('Adjacent', 'Wills + digital vault, IFA channel'),
 'Make a Will Online': ('Adjacent', 'Online wills'),
 'Beyond (beyond.life)': ('Defunct / exited', 'Dissolved 2024'),
 'Bequest (unverified)': ('Out of scope', 'Unverified'),
 'Settled (status unclear)': ('Out of scope', 'Status unclear'),
 'Simply-Docs': ('Core', 'Document subscription (sold Jul 2026)'),
 'Net Lawman': ('Adjacent', 'Document templates'),
 'Lawpack': ('Adjacent', 'DIY legal kits'),
 'Legalo': ('Adjacent', 'Document templates'),
 'LawDepot UK': ('Adjacent', 'Document subscription (Canadian)'),
 'Wonder.Legal (UK)': ('Adjacent', 'Document builder (French)'),
 'Contractbook (UK)': ('Out of scope', 'CLM (acquired by Scrive)'),
 'Summize': ('Out of scope', 'Enterprise CLM'),
 'Juro': ('Out of scope', 'Enterprise CLM'),
 'Avvoka': ('Out of scope', 'Document automation for law firms'),
 'Lexoo': ('Defunct / exited', 'Absorbed into Kalexius'),
}
def tier(name): return TIER.get(name, ('Adjacent', ''))

# Numeric overrides where the research JSON holds prose instead of a number (all from the same filed sources)
PRIOR = {'Harper James': 12221570}            # FY Sep-24 turnover (Companies House)
REV = {'Lawhive': 26000000}                   # company-reported >$35m run-rate (Feb 2026), ~GBP 26m
REV_FY = {'Lawhive': 'Run-rate, Feb 2026 (company-reported >$35m, global, mostly consumer)'}
NO_RPE = {'Lawhive'}                          # 2024 headcount vs 2026 run-rate revenue delivered by ~500 consultant lawyers: not comparable

def period_is_12m(text):
    m = re.search(r'(\d+)[- ]?months?', str(text or ''))
    return not m or m.group(1) == '12'


wb = Workbook(); wb.remove(wb.active)

# ---------------- Company sheet builder ----------------
COLS = ['Company', 'Relevance vs LegalVision', 'Model tag', 'Legal entity / Companies House no.', 'HQ', 'Founded', 'Ownership / backers',
        'Funding & deal history', 'Revenue (GBP)', 'Revenue period', 'Prior-year revenue / notes', 'Growth %', 'EBITDA / profit (GBP)',
        'Revenue basis', 'Employees (avg, Companies House unless stated)', 'Employee count notes / year', 'Revenue per employee (GBP)',
        'Pricing model', 'Pricing detail', 'Services', 'Practice areas', 'Regulatory status', 'Customers / reviews', 'Tech angle', 'Valuation notes', 'Sources']
W = [26, 14, 22, 30, 24, 12, 36, 40, 14, 18, 40, 9, 16, 30, 12, 34, 13, 34, 44, 50, 28, 30, 32, 30, 30, 50]

def company_rows(records, seg_key=None):
    rows = []
    for x in records:
        t, tag = tier(x['name'])
        rev = REV.get(x['name'], num(x.get('revenue_gbp'))); prior = PRIOR.get(x['name'], num(x.get('revenue_prior')))
        fy = REV_FY.get(x['name'], x.get('revenue_fy'))
        g = growth(rev, prior) if period_is_12m(fy) else None
        emp = x.get('employees'); emp_n = first_int(emp)
        emp_text = x.get('employees_year')
        if not isinstance(emp, (int, float)) and emp: emp_text = f"{emp} | {emp_text or ''}"
        ebitda = x.get('ebitda')
        rows.append([
            x['name'], t, tag,
            ' | '.join(str(v) for v in [x.get('legal_entity'), x.get('ch_number')] if v),
            x.get('hq'), x.get('founded'), x.get('ownership'),
            ' | '.join(str(v) for v in [x.get('funding'), x.get('deal_history')] if v),
            rev, fy, x.get('revenue_prior'),
            g, ebitda, x.get('revenue_basis'),
            emp_n, emp_text, (round(rev / emp_n) if rev and emp_n and x['name'] not in NO_RPE else None),
            x.get('pricing_model'), x.get('pricing_detail'), x.get('services'), x.get('practice_areas'),
            x.get('regulatory_status'), x.get('customers'), x.get('tech'), x.get('valuation_notes'), x.get('sources')])
    order = {'Core': 0, 'Adjacent': 1, 'Defunct / exited': 2, 'Out of scope': 3}
    rows.sort(key=lambda r: (order[r[1]], -(r[8] or 0)))
    return rows

def company_sheet(name, heading, sub, records):
    ws = wb.create_sheet(name)
    title(ws, heading, sub)
    table(ws, 4, COLS, company_rows(records), W, freeze=True, tier_col=2)
    return ws

# ---------------- Summary sheet ----------------
ws = wb.create_sheet('Summary')
title(ws, 'UK online / subscription legal services: market map and benchmarking vs LegalVision',
      'Compiled 12 Oct 2026 from Companies House filings, company websites, press and deal announcements. Figures are GBP unless stated. See "Caveats & gaps".')

ws['A4'] = 'How to read this workbook'; ws['A4'].font = S
r = table(ws, 5, ['Sheet', 'Contents'], [
    ['1. SME subscription legal', 'Businesses closest to LegalVision: subscription/membership law firms, online legal platforms and fixed-fee providers serving UK SMEs and startups (22 records).'],
    ['2. HR & employment-law subs', 'The large, PE-heavy adjacent segment: insurance-backed employment law / HR / H&S subscriptions for SMEs (Peninsula, Citation, WorkNest, Avensure etc.) (21 records).'],
    ['3. Consumer & documents', 'Digitally native consumer legal services (wills, probate, divorce, prenups, advice lines) and legal document / contract platforms (33 records).'],
    ['4. LegalVision benchmark', 'LegalVision group and UK entity: model, pricing by country, revenue, headcount, ownership (Quadrant PE) and UK filings.'],
    ['5. Market context', 'UK legal market size, SME legal-needs data (Legal Services Board), regulation (SRA ABS regime), AI adoption.'],
    ['6. Deals & PE activity', 'PE, M&A and VC transactions in UK legal services and legaltech 2020-2026 with values and multiples where disclosed.'],
    ['7. Listed comps', 'Keystone, Knights, Gateley, LegalZoom, Rocket Lawyer and others: revenue, margins, EV/revenue, EV/EBITDA.'],
    ['Benchmark table', 'Side-by-side of the core comparables on revenue, growth, employees, revenue per employee, pricing and ownership.'],
    ['Suggested KPIs', 'Additional data points worth collecting for a PE market map and where to get them.'],
    ['Caveats & gaps', 'Reliability notes and what could not be sourced.'],
], [26, 120])

ws.cell(r, 1, 'Key takeaways').font = S; r += 1
take = [
 'Direct LegalVision analogues are few and small. Harper James (GBP 15.3m revenue FY Sep-25, +25%, 131 staff, ~70% of income from subscription plans, LDC minority stake 2025) is the only scaled UK subscription law firm. LegalVision UK filed GBP 3.7m turnover for FY Jun-24 with 41 staff in FY25. Everything else in the core segment is sub-GBP 2m or does not file turnover.',
 'The graveyard is long: LawBite (administration 2024, platform sold for GBP 1), Linkilaw (exited SMEs), Sparqa (dormant), Pocketlaw (pivoted to enterprise CLM), Rocket Lawyer UK (8 staff, GBP 5m net liabilities), Legal Utopia (liquidation 2026), Beyond (dissolved 2024), Lexoo (absorbed). Self-serve document models have struggled to scale in the UK; lawyer-delivered subscription models (Harper James, LegalVision) have grown.',
 'The PE-proven version of "subscription legal" in the UK is employment law / HR / H&S compliance: Peninsula (GBP 463m revenue, GBP 118m EBITDA, family-owned), Citation (GBP 275m, KKR/Hg, bought at ~15x EBITDA in 2020), WorkNest/Axiom GRC (Inflexion, GBP 430m carve-out at ~16x EBITDA 2024), Avensure (GBP 19m, 30% EBITDA margin, founder-owned). Multi-year contracts, per-employee pricing, insurance-backed indemnities and 24/7 helplines are the common model.',
 'Consumer digital legal is consolidating into corporates and PE: Farewill to Dignity (GBP 12.9m EV, ~1.4x revenue, well below the GBP 70m 2020 valuation), amicable and Octopus Legacy to Octopus Group, Wenup to Blackfinch (2026), Muve to Alcuin (2026), Simply-Docs change of control (2026). Lawhive is the outlier: >$35m run-rate revenue, 7x growth, $60m Series B (Feb 2026).',
 'Demand side is large and under-served: LSB data shows 32% of small businesses have a legal issue each year (1.8m firms) but only 25% use professional help and 44-46% handle it alone; only 10% see lawyers as cost-effective. UK legal market GBP 55.5bn (2025), B2B ~51%.',
 'Valuation anchors: listed UK law firms trade at 0.8-1.6x EV/revenue and 6-14x EV/EBITDA (Keystone 1.6x / 14x; Knights 1.3x / ~5-6x; Gateley 0.8x / 10x); LegalZoom ~1.1x revenue / ~5x adj. EBITDA; compliance subscription assets (Citation, Marlowe GRC) at 15-16x EBITDA. Harper James/LDC and LegalVision/Quadrant terms were not disclosed.',
 'Regulation is an enabler not a barrier: the SRA ABS regime (2012) allows non-lawyer and PE ownership of law firms; LegalVision UK, Harper James, Lawhive, Markel Law and Co-op all operate through licensed bodies. Unregulated models (Sprintlaw, 360 Business Law, HR advisory firms) avoid reserved activities and SRA cost but cannot litigate.',
]
for tk in take:
    ws.cell(r, 1, '-'); ws.cell(r, 2, tk).alignment = WRAP; ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2); ws.row_dimensions[r].height = 62; r += 1
r += 1
ws.cell(r, 1, 'Segment overview').font = S; r += 1
def seg_stats(records, want):
    names = [x['name'] for x in records if tier(x['name'])[0] in want]
    return len(names)
r = table(ws, r, ['Segment', 'Records', 'Core', 'Adjacent', 'Defunct / exited / out of scope', 'Scaled players (revenue > GBP 10m) in the segment'], [
    ['1. SME subscription legal', len(SME), seg_stats(SME, {'Core'}), seg_stats(SME, {'Adjacent'}), seg_stats(SME, {'Defunct / exited', 'Out of scope'}), 'Harper James (GBP 15.3m); Lawhive (>$35m group run-rate, consumer-led)'],
    ['2. HR & employment-law subs', len(HR), seg_stats(HR, {'Core'}), seg_stats(HR, {'Adjacent'}), seg_stats(HR, {'Defunct / exited', 'Out of scope'}), 'Peninsula (GBP 463m), Citation (GBP 275m), ARAG UK (GBP 244m income), Croner (GBP 53m), Croner-i (GBP 52m), BrightHR (GBP 41m), WorkNest (GBP 31m), Moorepay (GBP 30m), Avensure (GBP 19m), Markel Law (GBP 11m)'],
    ['3. Consumer & documents', len(CON), seg_stats(CON, {'Core'}), seg_stats(CON, {'Adjacent'}), seg_stats(CON, {'Defunct / exited', 'Out of scope'}), 'Slater & Gordon (GBP 109m), Co-op Legal Services (GBP 82m), Lawhive (~GBP 26m), ARAG Law (GBP 21m), Muve (GBP 19m), Juro (GBP 9.7m)'],
], [28, 9, 7, 9, 14, 90])

# ---------------- Company sheets ----------------
company_sheet('1. SME subscription legal', 'Section 1. UK SME / startup subscription and online legal services (closest LegalVision comparables)',
              'Relevance: Core = lawyer-delivered subscription/membership or fixed-fee legal service for SMEs; Adjacent = self-serve platform, marketplace or SaaS with legal support; sorted by relevance then revenue.', SME)
company_sheet('2. HR & employment-law subs', 'Section 2. Subscription employment law / HR / H&S advisory for SMEs (adjacent, PE-heavy)',
              'Insurance-backed, multi-year, per-employee-priced subscriptions. Group figures are consolidated; subsidiary figures are standalone Companies House filings.', HR)
company_sheet('3. Consumer & documents', 'Section 3. Consumer-facing digital legal services and legal document / contract platforms',
              'Consumer records (wills, probate, divorce, prenups, advice lines, conveyancing) followed by document and contract-automation businesses.', CON)

# ---------------- 4. LegalVision benchmark ----------------
ws = wb.create_sheet('4. LegalVision benchmark')
title(ws, 'Section 4. LegalVision: the benchmark', 'Group (AU/NZ/UK) and UK entity (LegalVision Law UK Ltd, 13704561). Source list per row.')
rows = []
for k, v in CTX['legalvision'].items():
    if isinstance(v, dict) and 'value' in v:
        val = v.get('value'); src = v.get('source')
        if isinstance(val, dict): val = '\n'.join(f'{a}: {b}' for a, b in val.items())
        rows.append([k.replace('_', ' '), val, src])
    elif isinstance(v, dict):   # e.g. peers_entering_uk: {lawpath: {value, source}, ...}
        for k2, v2 in v.items():
            rows.append([f"{k.replace('_', ' ')}: {k2}", v2.get('value') if isinstance(v2, dict) else v2, v2.get('source') if isinstance(v2, dict) else None])
    elif v is not None:
        rows.append([k.replace('_', ' '), v, None])
r = table(ws, 4, ['Item', 'Finding', 'Sources'], rows, [30, 110, 70])
ws.cell(r, 1, 'UK entity trajectory (Companies House)').font = S; r += 1
uk = CTX['legalvision'].get('uk_entity_financials_companies_house', {}).get('value', {})
r = table(ws, r, ['Period', 'Turnover (GBP)', 'Average employees', 'Note'], [
    ['26 Oct 2021 to 31 Oct 2022 (first period)', uk.get('turnover_26Oct2021_to_31Oct2022_gbp'), uk.get('avg_employees_FY2022'), 'UK trading from mid-2022; ABS licence 13 May 2022'],
    ['8 months to 30 Jun 2023', uk.get('turnover_8m_to_30Jun2023_gbp'), 15, 'Year-end changed to June'],
    ['FY to 30 Jun 2024', uk.get('turnover_FY_to_30Jun2024_gbp'), 24, f"Restated in FY25 accounts to GBP {uk.get('turnover_FY2024_restated_gbp')}" if uk.get('turnover_FY2024_restated_gbp') else None],
    ['FY to 30 Jun 2025', None, 41, 'Small-company accounts: P&L omitted; profit GBP 6.7k; intercompany loan at 11%; Quadrant vehicle became ultimate parent (May 2025)'],
], [34, 16, 16, 90])

# ---------------- 5. Market context ----------------
ws = wb.create_sheet('5. Market context')
title(ws, 'Section 5. UK market context', 'Market size, SME legal-needs data, regulation and AI adoption.')
table(ws, 4, ['Metric', 'Value', 'Year', 'Source'], [[m['metric'], m['value'], m['year'], m['source']] for m in CTX['market']], [48, 80, 10, 70])

# ---------------- 6. Deals ----------------
ws = wb.create_sheet('6. Deals & PE activity')
title(ws, 'Section 6. PE, M&A and VC activity in UK legal services and legaltech, 2020-2026', 'Values and multiples as reported; "derived" = calculated from reported inputs. Rows marked unverified came from prior knowledge, not a source fetched in this exercise.')
deals = sorted(CTX['deals'], key=lambda d: str(d.get('date') or ''), reverse=True)
r = table(ws, 4, ['Date', 'Target', 'Acquirer / investor', 'Type', 'Value', 'Multiple / valuation', 'Notes', 'Source'],
      [[d['date'], d['target'], d['acquirer_or_investor'], d['type'], d['value'], d['multiple_or_valuation'], d['notes'], d['source']] for d in deals],
      [16, 36, 34, 12, 26, 34, 60, 60])
ws.cell(r, 1, 'Additional transactions from the company sheets (this exercise)').font = S; r += 1
table(ws, r, ['Date', 'Target', 'Acquirer / investor', 'Type', 'Value', 'Multiple / valuation', 'Notes', 'Source'], [
    ['2025-06 / 2025-09', 'Harper James (subscription law firm)', 'LDC (minority ~45%)', 'PE', 'Undisclosed', "Described as a 'GBP 16m business' at the deal; FY25 revenue GBP 15.3m, op. profit GBP 2.2m", 'Investec debt at HJ Holdco; SRA approval required; target GBP 30m revenue by 2028', 'See sheet 1 sources'],
    ['2026-03-25', 'NatWest Mentor (c.220 staff, c.100k SMEs)', 'Empowering People Group (Limerston Capital)', 'M&A', 'Undisclosed', None, 'Completion expected Q3/Q4 2026; NatWest retains referral arrangement', 'See sheet 2 sources'],
    ['2024-10-28', 'Wirehouse Employer Services', 'WorkNest / Axiom GRC (Inflexion)', 'M&A (bolt-on)', 'Undisclosed', 'FY Jul-24 turnover GBP 4.5m', None, 'See sheet 2 sources'],
    ['2026-03-13', 'Muve / Connect2Law (digital conveyancing)', 'Alcuin Capital Partners (majority)', 'PE', 'Undisclosed', 'FY25 turnover GBP 19.3m, PAT GBP 3.7m', None, 'See sheet 3 sources'],
    ['2023-11-30', 'amicable (online divorce)', 'Octopus Group (majority)', 'M&A', 'Undisclosed', 'FY Apr-25 turnover GBP 3.4m (+136%)', 'Octopus also controls Octopus Legacy (ex-Guardian Angel) since Nov 2022', 'See sheet 3 sources'],
    ['2026-06-05', 'Wenup (online prenups)', 'Blackfinch Group', 'M&A', 'Undisclosed', None, None, 'See sheet 3 sources'],
    ['2026-07-24', 'Simply-Docs (document subscription)', 'Private buyer (D. Firth, 75%+ PSC)', 'M&A', 'Undisclosed', None, 'Founder exit', 'See sheet 3 sources'],
    ['2026-01', 'Summize (CLM)', 'Kennet Partners, Federated Hermes PE (+ Maven, YFM)', 'Growth equity', '$50m', None, 'ARR +97% YoY', 'See sheet 3 sources'],
    ['2026-03-09', 'Avvoka (document automation)', 'Valhalla Ventures', 'Growth equity', 'GBP 14m', None, None, 'See sheet 3 sources'],
    ['2025-06-18', 'Contractbook (CLM)', 'Scrive (Vitruvian-backed)', 'M&A', 'Undisclosed', None, None, 'See sheet 3 sources'],
    ['2025-09', 'Woodstock Legal Services (fee-share firm)', 'Lawhive', 'M&A (bolt-on)', 'Undisclosed', None, None, 'See sheet 3 sources'],
    ['2024-03-08', 'Lexoo (legal marketplace)', 'Kalexius (ALSP)', 'M&A', 'Undisclosed', 'Sub-scale exit', None, 'See sheet 3 sources'],
    ['2024-09-24', 'LawBite / Lawbit Ltd', 'Administration (Antony Batty & Co)', 'Distress', 'Platform sold for GBP 1', 'GBP 5.8m owed to 33 creditors; ~GBP 7m raised historically', 'SRA intervened in LawBriefs Oct 2024', 'See sheet 1 sources'],
], [16, 36, 34, 12, 26, 34, 60, 60])

# ---------------- 7. Listed comps ----------------
ws = wb.create_sheet('7. Listed comps')
title(ws, 'Section 7. Listed and large private comparables', 'Market data as at 12 Oct 2026 (stockanalysis.com) unless stated.')
table(ws, 4, ['Company', 'Ticker', 'Revenue', 'Revenue period', 'EBITDA / margin', 'Market cap / value', 'EV / revenue (x)', 'EV / EBITDA (x)', 'Source'],
      [[c['company'], c['ticker'], c['revenue'], c['revenue_fy'], c['ebitda_margin'], c['market_cap'], c['ev_revenue'], c['ev_ebitda'], c['source']] for c in CTX['listed_comps']],
      [26, 14, 50, 22, 50, 30, 12, 24, 60])

# ---------------- Benchmark table ----------------
ws = wb.create_sheet('Benchmark table')
title(ws, 'Benchmark: core comparables side by side', 'Latest filed or reported figures. Revenue per employee uses average employees from Companies House. Blank = not disclosed.')
bench_names = ['Harper James', 'LegalVision UK', 'Lawhive', 'SeedLegals', 'Rocket Lawyer UK', '360 Business Law / 360 Law Group', 'Law 365', 'Lawyerly (extra)', 'Sprintlaw UK (extra)', 'Genie AI',
               'Peninsula Group (Peninsula Business Services)', 'Citation Group', 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC', 'Avensure', 'Croner (Peninsula Group)', 'BrightHR (Peninsula Group)', 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)', 'Wirehouse Employer Services (WorkNest / Axiom GRC)',
               'Co-op Legal Services', 'Farewill', 'amicable', 'Octopus Legacy (formerly Guardian Angel)', 'Muve (Connect2Law)', 'Divorce-Online', 'Which? Legal', 'Simply-Docs', 'Juro']
allrec = {x['name']: ('1', x) for x in SME}
allrec.update({x['name']: ('2', x) for x in HR})
for x in CON: allrec.setdefault(x['name'], ('3', x))
rows = []
for n in bench_names:
    seg, x = allrec[n]
    rev = REV.get(n, num(x.get('revenue_gbp'))); prior = PRIOR.get(n, num(x.get('revenue_prior'))); emp = first_int(x.get('employees'))
    fy = REV_FY.get(n, x.get('revenue_fy'))
    rows.append([n, {'1': 'SME subscription', '2': 'HR / employment law', '3': 'Consumer / documents'}[seg], tier(n)[1], rev, fy, growth(rev, prior) if period_is_12m(fy) else None,
                 num(x.get('ebitda')), emp, (round(rev / emp) if rev and emp and n not in NO_RPE else None), x.get('pricing_model'),
                 (x.get('ownership') or '')[:220], x.get('regulatory_status')])
table(ws, 4, ['Company', 'Segment', 'Model', 'Revenue (GBP)', 'Period', 'Growth %', 'EBITDA (GBP)', 'Employees (avg)', 'Revenue / employee (GBP)', 'Pricing model', 'Ownership', 'Regulatory status'], rows,
      [26, 18, 30, 14, 22, 9, 14, 11, 13, 44, 50, 36], freeze=True)

# ---------------- Suggested KPIs ----------------
ws = wb.create_sheet('Suggested KPIs')
title(ws, 'Suggested additional data points for the market map', 'What a PE benchmarking exercise on subscription legal usually needs beyond revenue, FTEs and services, and where it can be found.')
table(ws, 4, ['Data point', 'Why it matters', 'Where to get it', 'Status in this workbook'], [
    ['Subscription / recurring revenue share and ARR', 'Separates true subscription models (LegalVision ~50% subscription in 2022; Harper James ~70% of income on plans; LegalZoom 65%) from fixed-fee or hourly businesses marketed as subscriptions.', 'Management, press interviews, investor decks; LegalZoom 10-K for a public reference', 'Captured where public (sheets 1, 4, 7)'],
    ['Renewal / churn rate and net revenue retention', 'LegalVision cited 85% renewal and NPS 70 in 2022; Peninsula/Citation renew clients onto new multi-year terms. Drives LTV and valuation.', 'Management only; Trustpilot complaint themes as a proxy for churn risk', 'Only anecdotal'],
    ['Contract length and auto-renewal terms', 'Peninsula 1/3/5-year with 2% uplift and auto-renewal; LegalVision 12-month minimum with 3/5-year discounts; Mentor 3-month rolling. Longer terms = higher visibility but regulatory/reputation risk (CMA/FCA scrutiny of auto-renewals).', 'Websites, T&Cs, filed accounts narrative', 'Captured (pricing columns)'],
    ['ARPU / average contract value and members count', 'Needed to triangulate revenue where turnover is not filed (e.g. LegalVision UK FY25, SeedLegals, 360). LegalVision group >4,500 clients; Harper James >1,000 Enable subscribers.', 'Press, websites; divide revenue by clients where both exist', 'Partly captured (customers column)'],
    ['Gross margin and lawyer utilisation', 'Lawyer-delivered subscriptions are capacity-constrained; margin depends on utilisation and seniority mix. LegalVision UK FY24 gross margin 64%; Juro 77% (software).', 'Full Companies House accounts (medium/large companies only)', 'Captured where filed'],
    ['Revenue per lawyer / per FTE and sales-to-lawyer ratio', 'LegalVision UK: 28 lawyers + 12 sales of ~40 staff (Mar 2025); Avensure 56 sales vs 56 consultants. Indicates go-to-market intensity.', 'Press, accounts employee notes, LinkedIn', 'Revenue per employee computed where possible'],
    ['Customer acquisition cost and channel mix', 'Direct digital vs B2B2C (banks, insurers, trade bodies: Markel/FSB, Mentor/NatWest, Farillio/insurers, Kinherit/IFAs). Channel-dependent businesses carry partner risk (e.g. Mentor sale).', 'Management; partner announcements', 'Channel captured qualitatively'],
    ['Regulatory structure (SRA ABS vs unregulated) and licence date', 'ABS status enables PE ownership of a law firm and reserved work (litigation, probate); unregulated models are cheaper but capped in scope. Changes of control need SRA approval (Harper James/LDC took ~3 months).', 'SRA register; Companies House PSC filings', 'Captured'],
    ['Insurance backing and indemnity exposure', 'HR/employment-law subscriptions sell tribunal indemnities (Citation GBP 1.5m guarantee; Avensure, Peninsula via LEI policies). Underwriter and claims history matter for diligence.', 'Websites, FCA register, accounts', 'Captured qualitatively'],
    ['Balance sheet health (net assets, cash, intercompany debt)', 'Small-company filings omit P&L but show net assets/cash: a leading indicator of distress (LawBite, Legal Utopia, Pocketlaw UK going-concern notes) or parent dependence (Rocket Lawyer UK, LegalVision UK 11% intercompany loan).', 'Companies House iXBRL', 'Captured in revenue notes'],
    ['Headcount trajectory (3-year)', 'Average employees are filed even by small companies and track growth better than press claims (Harper James 64 to 131 in 3 years; LegalVision UK 15 to 41).', 'Companies House', 'Captured for most'],
    ['AI leverage', 'Lawhive, Harper James (LDC capital earmarked for AI), Genie, Sprintlaw (Taylor AI) use AI to cut delivery cost; relevant to margin expansion thesis and to disruption risk for document businesses.', 'Websites, press', 'Captured (tech column)'],
    ['Pricing evolution', 'LegalVision UK membership rose from GBP 49-329/week (2022) to GBP 115-620+/week (2026): evidence of pricing power or of repositioning upmarket.', 'Wayback Machine snapshots of pricing pages', 'Captured for LegalVision, Harper James, 360'],
    ['Ownership, cap table and likely sellers', 'Founder-owned at scale (Peninsula, Avensure, 360, Lawgistics), PE-owned (Citation 2020 vintage, WorkNest 2024, Harper James 2025), VC-owned with stale rounds (SeedLegals 2019, Juro 2022) and corporate orphans (Rocket Lawyer UK, Mentor) point to who may transact.', 'Companies House PSC, press', 'Captured'],
    ['Valuation and deal multiples', 'Citation ~15x EBITDA (2020), Marlowe GRC 16x (2024), Keystone 14x, Knights ~5-6x, Farewill 1.4x revenue; Harper James, LegalVision/Quadrant, Lawhive undisclosed.', 'Press, RNS, PE announcements', 'Captured (sheets 6-7)'],
    ['Customer satisfaction', 'Trustpilot and NPS where public: Harper James 4.6, Lawhive 4.4, Farewill 4.9, Which? Legal 1.9, LawDepot 1.8 (auto-renewal complaints).', 'Trustpilot, Google, ReviewSolicitors', 'Captured'],
    ['Web traffic and search share', 'Proxy for digital acquisition scale; AI-referred traffic to UK law firms +310% in 2025.', 'Similarweb, Semrush', 'Not captured'],
    ['Geographic mix and international entrants', 'LegalVision, Sprintlaw (AU), Rocket Lawyer (US), Docue (FI), Pocketlaw (SE), LawDepot (CA), Wonder.Legal (FR) all entered the UK; most remain sub-scale. Lawhive and SeedLegals have gone the other way (US).', 'Accounts segmental notes', 'Captured'],
], [36, 80, 44, 30])

# ---------------- Caveats ----------------
ws = wb.create_sheet('Caveats & gaps')
title(ws, 'Caveats, reliability and gaps')
table(ws, 3, ['Topic', 'Note'], [
    ['Revenue basis', 'Most small UK companies file abridged or micro accounts without turnover. Revenue is shown only where filed (full/medium accounts), reported by the company, or stated in press; the "Revenue basis" column says which. Where turnover is missing, net assets, cash and headcount from the balance sheet are given instead.'],
    ['FTE vs employees', 'Companies House reports average monthly employees (headcount, not FTE). Consultant / fee-share lawyers (Lawhive ~500, 360 Business Law 750+, Setfords ~650) are not employees and are excluded from those counts.'],
    ['LegalVision UK revenue', 'Management told Capital Brief (Oct 2024) UK revenue was ~A$13m (~GBP 6.5m) for 2024; Companies House shows GBP 3.35-3.69m for FY Jun-24 and no turnover for FY Jun-25. Treat the A$13m figure as unverified.'],
    ['Lawhive revenue', '>$35m is a company-stated run-rate (Feb 2026), global and mostly consumer; filed FY2024 accounts omit the P&L.'],
    ['Citation EBITDA', 'FY24 reported "operating profit before D&A" of GBP 43.4m is after GBP 33.5m one-off transaction costs and GBP 4.6m deferred remuneration; underlying c.GBP 77-82m is derived, not reported.'],
    ['Corrections to common assumptions', 'No Markel/LawBite deal exists (LawBite went into administration; Markel Law is Markel\'s own ABS). Docue did not acquire Sparqa or Farillio. Setfords is Phoenix-backed (not Lonsdale). Lawfront is Blixt-backed (not Sovereign).'],
    ['Image-only filings', 'Group accounts for Summize, Legl and some Juro/Co-op/Farewill filings are scanned PDFs; figures read visually are flagged in the records (Co-op headcount 749 vs 689 is low-resolution).'],
    ['Unverified rows', 'A few deal rows (Robin AI, vLex/Clio, Luminance, Juro 2022) and third-party revenue estimates (Latka, ZoomInfo, multiples.vc) are marked unverified or low reliability and should be checked before use in a client document.'],
    ['Blocked sources', 'Law Gazette, some Legal Futures and Endole/Pomanda pages returned 403/405; LawBite detail came from RollOnFriday, administrator filings and Legal Futures print pages.'],
    ['Not yet researched', 'Definely, Robin AI, Luminance (legal AI vendors); gunnercooke, Taylor Rose, Keystone-style consultant platforms beyond Setfords; Howden-specific LEI products; Peninsula/Citation Trustpilot complaint themes; SeedLegals 2026 round; Zegal UK entity; 2025 IRN consumer-market total.'],
    ['Date', 'All data compiled 12 Oct 2026. Companies House figures are the latest filed accounts at that date.'],
], [28, 130])

wb.save(out)
print('saved', out, wb.sheetnames)
