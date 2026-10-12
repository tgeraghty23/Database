"""Build the one-page HTML summary of the UK online legal market map from research/data/uk_online_legal_*.json."""
import json, os, sys, re, html
D = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'data')
def load(n):
    with open(os.path.join(D, n), encoding='utf-8') as f: return json.load(f)
SME = load('uk_online_legal_sme_subscription.json'); HR = load('uk_online_legal_hr_employment.json')
CON = load('uk_online_legal_consumer_documents.json'); CTX = load('uk_online_legal_context.json')
REC = {x['name']: x for x in SME + HR}
for x in CON: REC.setdefault(x['name'], x)
E = lambda s: html.escape(str(s)) if s is not None else ''

def money(v, unit='m'):
    if v is None: return '<span class="na">n.d.</span>'
    return f'£{v/1e6:,.1f}m' if v >= 1e6 else f'£{v/1e3:,.0f}k'
def pct(v): return '<span class="na">–</span>' if v is None else f'{v*100:+.0f}%'
def first_int(s):
    if s is None: return None
    if isinstance(s, (int, float)): return int(s)
    m = re.search(r'(\d[\d,]*)', str(s)); return int(m.group(1).replace(',', '')) if m else None
def num(v): return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None

# (name, display, segment, prior override, revenue override, fy override, rpe allowed)
B = [
 ('Harper James', 'Harper James', 'SME', 12221570, None, None, True),
 ('LegalVision UK', 'LegalVision UK', 'SME', None, None, None, True),
 ('Lawhive', 'Lawhive', 'SME', None, 26000000, 'Run-rate Feb 2026 (company)', False),
 ('SeedLegals', 'SeedLegals', 'SME', None, None, None, True),
 ('360 Business Law / 360 Law Group', '360 Business Law', 'SME', None, None, None, True),
 ('Law 365', 'Law 365', 'SME', None, None, None, True),
 ('Rocket Lawyer UK', 'Rocket Lawyer UK', 'SME', None, None, None, True),
 ('Genie AI', 'Genie AI', 'SME', None, None, None, True),
 ('Lawyerly (extra)', 'Lawyerly', 'SME', None, None, None, True),
 ('Sprintlaw UK (extra)', 'Sprintlaw UK', 'SME', None, None, None, True),
 ('Peninsula Group (Peninsula Business Services)', 'Peninsula Group', 'HR', None, None, None, True),
 ('Citation Group', 'Citation Group', 'HR', None, None, None, True),
 ('Croner (Peninsula Group)', 'Croner', 'HR', None, None, None, True),
 ('BrightHR (Peninsula Group)', 'BrightHR', 'HR', None, None, None, True),
 ('WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC', 'WorkNest (Axiom GRC)', 'HR', None, None, None, True),
 ('Avensure', 'Avensure', 'HR', None, None, None, True),
 ('Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)', 'Markel Law', 'HR', None, None, None, True),
 ('Co-op Legal Services', 'Co-op Legal Services', 'Consumer', None, None, None, True),
 ('Muve (Connect2Law)', 'Muve', 'Consumer', None, None, None, True),
 ('Farewill', 'Farewill', 'Consumer', None, None, '17 months to Dec 2024', True),
 ('amicable', 'amicable', 'Consumer', None, None, None, True),
 ('Divorce-Online', 'Divorce-Online', 'Consumer', None, None, None, True),
 ('Octopus Legacy (formerly Guardian Angel)', 'Octopus Legacy', 'Consumer', None, None, None, True),
 ('Juro', 'Juro (CLM)', 'Docs', None, None, None, True),
]
OWN = {
 'Harper James': 'Founder majority; LDC ~45% (2025)', 'LegalVision UK': 'LegalVision group; Quadrant PE (2025)', 'Lawhive': 'VC: GV, TQ, Balderton, Mitch Rales',
 'SeedLegals': 'VC (Index, 2019); founders', '360 Business Law / 360 Law Group': 'Founder/family-owned', 'Law 365': 'Founder-owned', 'Rocket Lawyer UK': 'Rocket Lawyer Inc (US)',
 'Genie AI': 'VC: GV, Khosla (2024)', 'Lawyerly (extra)': 'Founder-owned', 'Sprintlaw UK (extra)': 'Sprintlaw Pty (AU)',
 'Peninsula Group (Peninsula Business Services)': 'Done family (no PE)', 'Citation Group': 'KKR 41% / Hg 41% / HarbourVest 15%', 'Croner (Peninsula Group)': 'Peninsula Group',
 'BrightHR (Peninsula Group)': 'Peninsula Group (98%)', 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC': 'Inflexion (Axiom GRC, 2024)', 'Avensure': 'Founder-owned (C. Garner)',
 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)': 'Markel Group (NYSE)', 'Co-op Legal Services': 'Co-operative Group', 'Muve (Connect2Law)': 'Alcuin Capital (Mar 2026)',
 'Farewill': 'Dignity / Castelnau (Jan 2025)', 'amicable': 'Octopus Group (Nov 2023)', 'Divorce-Online': 'Private (Byles, Madigan)', 'Octopus Legacy (formerly Guardian Angel)': 'Octopus Group (Nov 2022)',
 'Juro': 'VC: USV, Eight Roads, Point Nine',
}
def period_is_12m(t):
    m = re.search(r'(\d+)[- ]?months?', str(t or '')); return not m or m.group(1) == '12'

rows = []
for n, disp, seg, prior_o, rev_o, fy_o, rpe in B:
    x = REC[n]
    rev = rev_o or num(x.get('revenue_gbp')); prior = prior_o or num(x.get('revenue_prior')); fy = fy_o or x.get('revenue_fy') or ''
    g = (rev / prior - 1) if rev and prior and period_is_12m(fy) else None
    emp = first_int(x.get('employees')); eb = num(x.get('ebitda'))
    fy_short = re.sub(r'\(.*?\)', '', str(fy)).replace('FY to ', '').strip()[:28]
    own_short = OWN.get(n, re.split(r'[.;(]', x.get('ownership') or '')[0][:60])
    rows.append(f"<tr><td class='n'>{E(disp)}</td><td><span class='seg {seg.lower()}'>{seg}</span></td><td class='r'>{money(rev)}</td><td class='fy'>{E(fy_short)}</td><td class='r'>{pct(g)}</td><td class='r'>{money(eb) if eb else '<span class=na>–</span>'}</td><td class='r'>{emp if emp else '<span class=na>n.d.</span>'}</td><td class='r'>{(f'£{rev/emp/1e3:,.0f}k' if rev and emp and rpe else '<span class=na>–</span>')}</td><td class='t'>{E(own_short)}</td></tr>")
bench = '\n'.join(rows)

deals = [
 ('Mar 2025', 'LegalVision (AU/NZ/UK)', 'Quadrant PE growth stake', 'Undisclosed', '4,500+ clients, 30%+ growth'),
 ('Jun–Sep 2025', 'Harper James', 'LDC minority (~45%)', 'Undisclosed', '"£16m business"; FY25 rev £15.3m, op. profit £2.2m'),
 ('Feb 2026', 'Lawhive', 'Series B led by Mitch Rales; GV, TQ, Balderton', '$60m', 'Revenue >$35m, 7x YoY; >$100m raised'),
 ('May 2024', 'Marlowe GRC (WorkNest, VinciWorks)', 'Inflexion carve-out', '£430m EV', '16.2x adj. EBITDA'),
 ('Aug 2020', 'Citation Group', 'KKR (Hg re-invested)', '£450–525m (press)', '~15x EBITDA; Hg paid £185m in 2016'),
 ('Mar 2026', 'NatWest Mentor', 'Empowering People Group (Limerston)', 'Undisclosed', 'c.220 staff, c.100k SMEs'),
 ('Oct 2024', 'Farewill', 'Dignity (Castelnau)', '£12.9m EV', '~1.4x revenue; £70m valuation in 2020'),
 ('Nov 2023', 'amicable', 'Octopus Group (majority)', 'Undisclosed', 'FY25 rev £3.4m (+136%)'),
 ('Mar 2026', 'Muve / Connect2Law', 'Alcuin Capital (majority)', 'Undisclosed', 'FY25 rev £19.3m, PAT £3.7m'),
 ('Jun 2026', 'Wenup', 'Blackfinch Group', 'Undisclosed', 'Online prenups'),
 ('Jul 2026', 'Simply-Docs', 'Private buyer', 'Undisclosed', 'Founder exit; £38.50–125/yr subscriptions'),
 ('Sep 2024', 'LawBite', 'Administration', 'Platform sold for £1', '£5.8m owed; ~£7m raised historically'),
 ('Jul 2023', 'DWF Group', 'Inflexion take-private', '£342m', '~1.0x revenue'),
]
deal_rows = '\n'.join(f"<tr><td class='fy'>{E(a)}</td><td class='n'>{E(b)}</td><td>{E(c)}</td><td class='r'>{E(d)}</td><td class='t'>{E(e)}</td></tr>" for a, b, c, d, e in deals)

comps = [('Keystone Law (AIM)', '£97.7m FY25; TTM £129m', '1.6x', '14.1x'), ('Knights (AIM)', '£207.7m FY26; 24.8% EBITDA margin', '1.3x', '~5–6x'),
         ('Gateley (AIM)', '£194.3m FY26', '0.8x', '10.0x'), ('LegalZoom (NASDAQ)', '$756m FY25; 65% subscription; 23% adj. EBITDA', '1.1x', '~5x adj.'),
         ('Citation (private, 2020 deal)', '£275m FY24', '–', '~15x at entry'), ('Marlowe GRC (2024 carve-out)', '~£86m FY23 pro forma', '5.0x', '16.2x')]
comp_rows = '\n'.join(f"<tr><td class='n'>{E(a)}</td><td class='t'>{E(b)}</td><td class='r'>{E(c)}</td><td class='r'>{E(d)}</td></tr>" for a, b, c, d in comps)

def card(title, items):
    lis = ''.join(f'<li><b>{E(n)}</b> <span>{E(d)}</span></li>' for n, d in items)
    return f"<div class='card'><h4>{E(title)}</h4><ul>{lis}</ul></div>"
mapcards = [
 card('Subscription law firms (lawyer-delivered)', [('Harper James', '£15.3m rev, 131 staff, LDC'), ('LegalVision UK', '£3.7m FY24, 41 staff, Quadrant'), ('Lawyerly', '£379–1,579/mo tiers, 2024 launch'), ('Law 365', '£1,550–1,675/mo, tech clients'), ('Blackmont Legal', 'from £350/mo, Manchester'), ('360 Business Law', '750+ consultant lawyers; ABS + unregulated')]),
 card('Fixed-fee / marketplace', [('Lawhive', '>$35m run-rate, 7x growth, $60m Series B'), ('Sprintlaw UK', 'AU fixed-fee + membership, 5 UK staff'), ('LegalEdge', 'fractional GC retainers'), ('LawBite', 'administration 2024'), ('Linkilaw', 'exited SMEs')]),
 card('Self-serve platforms & legal plans', [('SeedLegals', '£649–2,699/yr, 155 staff, cash tight'), ('Rocket Lawyer UK', '£34.99/mo, 8 staff, £5m net liabilities'), ('Genie AI', '£13.3m Series A (GV), 24 staff'), ('Docue UK', '£29/user/mo, 5 staff'), ('Sparqa / Farillio / Pocketlaw', 'dormant / B2B2C / pivoted')]),
 card('Employment-law & HR subscriptions (PE-proven)', [('Peninsula', '£463m rev, £118m EBITDA, family-owned'), ('Citation', '£275m, KKR/Hg, 110k clients'), ('WorkNest / Axiom GRC', '£31m, Inflexion'), ('Avensure', '£19m, 30% margin, founder-owned'), ('Croner / BrightHR', '£53m / £41m (Peninsula)'), ('NatWest Mentor', 'sale to EPG 2026')]),
 card('Consumer digital legal', [('Farewill', 'sold to Dignity £12.9m EV'), ('Octopus Legacy & amicable', 'Octopus Group'), ('Divorce-Online', '£2m+, 200k clients'), ('Which? Legal', '£99/yr advice subscription'), ('Co-op Legal Services', '£82m probate/wills ABS'), ('Juno / Muve', 'digital conveyancing; Muve to Alcuin')]),
 card('Documents & contract software', [('Simply-Docs', '£38.50/folder/yr, sold Jul 2026'), ('Net Lawman / Lawpack / Legalo', 'owner-managed micro'), ('Juro', '£9.7m rev, 97 staff, CLM'), ('Summize', '$50m round 2026, CLM'), ('Avvoka', '£14m 2026'), ('LawDepot / Wonder.Legal', 'foreign, Trustpilot <2')]),
]

kpis = [
 ('Subscription share and ARR', 'LegalVision ~50% subscription (2022), Harper James ~70% of income on plans, LegalZoom 65%. Separates real recurring revenue from fixed-fee work marketed as subscription.'),
 ('Renewal, churn and NRR', 'LegalVision cited 85% renewal and NPS 70; Peninsula and Citation renew clients onto new multi-year terms. Management-only data; Trustpilot themes are the public proxy.'),
 ('Contract length and auto-renewal', 'Peninsula 1/3/5-year with 2% uplift and auto-renewal; LegalVision 12-month minimum; Mentor 3-month rolling. Longer terms mean visibility but also CMA/FCA auto-renewal risk.'),
 ('ARPU, members and ACV', 'Needed to triangulate revenue where turnover is not filed (LegalVision UK FY25, SeedLegals, 360). LegalVision group >4,500 clients; Harper James >1,000 Enable subscribers.'),
 ('Gross margin and lawyer utilisation', 'Lawyer-delivered subscriptions are capacity-constrained. LegalVision UK FY24 gross margin 64%; Juro 77% (software).'),
 ('Sales-to-lawyer ratio', 'LegalVision UK: 28 lawyers and 12 sales of ~40 staff; Avensure 56 sales vs 56 consultants. Go-to-market intensity drives CAC.'),
 ('Channel mix (direct vs B2B2C)', 'Markel/FSB, Mentor/NatWest, Farillio/insurers, Kinherit/IFAs depend on partners; the Mentor sale shows partner risk.'),
 ('Regulatory structure', 'SRA ABS status enables PE ownership and reserved work; change of control needs SRA approval (Harper James/LDC took ~3 months). Unregulated models are cheaper but cannot litigate.'),
 ('Insurance backing', 'HR/employment-law subscriptions sell tribunal indemnities (Citation £1.5m guarantee). Underwriter and claims history matter in diligence.'),
 ('Balance-sheet signals', 'Small-company filings omit P&L but show net assets, cash and intercompany debt: early warning for LawBite, Legal Utopia, Pocketlaw UK; parent dependence for Rocket Lawyer UK and LegalVision UK (11% intercompany loan).'),
 ('Headcount trajectory', 'Filed even by small companies: Harper James 64 to 131 in three years; LegalVision UK 15 to 41.'),
 ('Pricing evolution', 'LegalVision UK membership moved from £49–329/week (2022) to £115–620+/week (2026): pricing power or repositioning upmarket.'),
 ('Ownership and likely sellers', 'Founder-owned at scale (Peninsula, Avensure, 360), 2019-22 VC vintages (SeedLegals, Juro), corporate orphans (Rocket Lawyer UK, Mentor) and PE vintages (Citation 2020) indicate who may transact.'),
 ('Web traffic and AI leverage', 'Similarweb/Semrush for acquisition scale (AI-referred traffic to UK law firms +310% in 2025); AI delivery at Lawhive, Harper James, Genie.'),
]
kpi_rows = '\n'.join(f"<tr><td class='n'>{E(a)}</td><td class='t'>{E(b)}</td></tr>" for a, b in kpis)

page = f"""<title>UK Subscription Legal Market Map</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@400;500;600&family=IBM+Plex+Serif:wght@500;600&family=IBM+Plex+Mono:wght@400;500&display=swap">
<style>
/* Layout: a briefing memo. One column of running text up to 72ch, full-width data tables, a six-cell market map grid. */
:root {{ --bg:#f7f6f2; --fg:#1d2430; --muted:#5b6472; --rule:#d9d6cc; --accent:#1f5f8b; --accent-2:#b8541f; --card:#ffffff; --head:#e9eef3; --sme:#dbe8f1; --hr:#f3e4d8; --consumer:#e4ecdc; --docs:#ece6f2;
 --display:'IBM Plex Serif', Georgia, 'Times New Roman', serif; --body:'IBM Plex Sans', 'Segoe UI', Helvetica, Arial, sans-serif; --mono:'IBM Plex Mono', 'SFMono-Regular', Consolas, monospace; }}
@media (prefers-color-scheme: dark) {{ :root:not([data-theme="light"]) {{ --bg:#15191f; --fg:#e6e4dd; --muted:#a3aab5; --rule:#343b46; --accent:#7fb3d9; --accent-2:#e3905e; --card:#1d2229; --head:#262d36; --sme:#223447; --hr:#44302a; --consumer:#2b3a28; --docs:#352e40; color-scheme: dark }} }}
:root[data-theme="dark"] {{ --bg:#15191f; --fg:#e6e4dd; --muted:#a3aab5; --rule:#343b46; --accent:#7fb3d9; --accent-2:#e3905e; --card:#1d2229; --head:#262d36; --sme:#223447; --hr:#44302a; --consumer:#2b3a28; --docs:#352e40; color-scheme: dark }}
body {{ background:var(--bg); color:var(--fg); font-family:var(--body); font-size:15px; line-height:1.5; padding-block:28px 56px; padding-inline:clamp(16px, 4vw, 48px); }}
.wrap {{ max-width:1180px; margin:0 auto; }}
h1 {{ font-family:var(--display); font-weight:600; font-size:clamp(26px, 3.2vw, 36px); line-height:1.15; margin:0 0 6px; text-wrap:balance; }}
h2 {{ font-family:var(--display); font-weight:600; font-size:21px; margin:40px 0 10px; padding-top:14px; border-top:2px solid var(--rule); text-wrap:balance; }}
h4 {{ margin:0 0 8px; font-size:13px; letter-spacing:.04em; text-transform:uppercase; color:var(--accent); }}
.eyebrow {{ font-family:var(--mono); font-size:12px; letter-spacing:.06em; text-transform:uppercase; color:var(--muted); margin-bottom:10px; }}
.lede {{ max-width:72ch; font-size:16px; color:var(--fg); }}
p, li {{ max-width:72ch; }}
.stats {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(170px, 1fr)); gap:12px; margin:22px 0 6px; }}
.stat {{ background:var(--card); border:1px solid var(--rule); border-radius:4px; padding:12px 14px; }}
.stat .v {{ font-family:var(--mono); font-size:24px; font-weight:500; color:var(--accent); font-variant-numeric:tabular-nums; }}
.stat .l {{ font-size:12.5px; color:var(--muted); margin-top:2px; }}
.take {{ list-style:none; padding:0; margin:12px 0 0; display:grid; gap:10px; }}
.take li {{ max-width:none; padding-left:18px; position:relative; }}
.take li::before {{ content:''; position:absolute; left:0; top:.55em; width:8px; height:8px; background:var(--accent-2); }}
.grid {{ display:grid; grid-template-columns:repeat(auto-fit, minmax(300px, 1fr)); gap:14px; margin-top:14px; }}
.card {{ background:var(--card); border:1px solid var(--rule); border-radius:4px; padding:14px 16px; min-width:0; }}
.card ul {{ list-style:none; margin:0; padding:0; display:grid; gap:6px; }}
.card li {{ max-width:none; font-size:14px; display:flex; flex-wrap:wrap; gap:4px 8px; }}
.card li b {{ font-weight:600; }}
.card li span {{ color:var(--muted); }}
.tbl {{ overflow-x:auto; margin-top:12px; border:1px solid var(--rule); border-radius:4px; background:var(--card); }}
table {{ border-collapse:collapse; width:100%; font-size:13.5px; font-variant-numeric:tabular-nums; }}
th {{ text-align:left; background:var(--head); font-weight:600; font-size:12px; letter-spacing:.03em; text-transform:uppercase; padding:9px 10px; border-bottom:1px solid var(--rule); white-space:nowrap; }}
td {{ padding:8px 10px; border-bottom:1px solid var(--rule); vertical-align:top; }}
tr:last-child td {{ border-bottom:0; }}
td.n {{ font-weight:600; white-space:nowrap; }} td.r {{ text-align:right; font-family:var(--mono); font-size:13px; white-space:nowrap; }}
th.r {{ text-align:right; }} td.fy {{ color:var(--muted); font-size:12.5px; white-space:nowrap; }} td.t {{ color:var(--muted); min-width:220px; }}
.na {{ color:var(--muted); }}
.seg {{ display:inline-block; font-size:11px; font-weight:600; padding:2px 7px; border-radius:3px; letter-spacing:.03em; }}
.seg.sme {{ background:var(--sme); }} .seg.hr {{ background:var(--hr); }} .seg.consumer {{ background:var(--consumer); }} .seg.docs {{ background:var(--docs); }}
.note {{ font-size:12.5px; color:var(--muted); margin-top:8px; max-width:none; }}
.src {{ font-size:12.5px; color:var(--muted); }}
@media (prefers-reduced-motion: reduce) {{ * {{ scroll-behavior:auto }} }}
</style>
<div class="wrap">
<div class="eyebrow">Market map · UK · compiled 12 October 2026</div>
<h1>UK online and subscription legal services: who looks like LegalVision?</h1>
<p class="lede">Benchmarking 76 UK businesses across subscription law firms, employment-law/HR subscriptions, consumer digital legal and document platforms, with revenue and headcount taken from Companies House filings wherever they are filed. The full dataset, sources and caveats are in the accompanying workbook.</p>

<div class="stats">
 <div class="stat"><div class="v">£15.3m</div><div class="l">Harper James revenue FY Sep-25 (+25%), the only scaled UK subscription law firm</div></div>
 <div class="stat"><div class="v">£3.7m</div><div class="l">LegalVision UK turnover FY Jun-24; 41 average staff FY25; FY25 turnover not filed</div></div>
 <div class="stat"><div class="v">15–16x</div><div class="l">EBITDA paid for compliance subscriptions (Citation 2020, Marlowe GRC 2024)</div></div>
 <div class="stat"><div class="v">25%</div><div class="l">of small businesses with a legal issue who use professional help (LSB)</div></div>
 <div class="stat"><div class="v">8</div><div class="l">of the 22 SME-segment names are defunct, exited or pivoted away from SMEs</div></div>
</div>

<h2>What the map says</h2>
<ul class="take">
 <li><b>Direct LegalVision analogues are few and small.</b> Harper James (£15.3m, 131 staff, ~70% of income from subscription plans, LDC minority stake 2025) is the only one at scale. Lawyerly, Blackmont, Law 365 and 360 Business Law are sub-£2m or do not file turnover. LegalVision UK itself grew from £0.3m to £3.7m in three years but carries £4.4m of 11% intercompany debt and net liabilities.</li>
 <li><b>Self-serve legal platforms have not worked in the UK.</b> LawBite went into administration (platform sold for £1), Sparqa is dormant, Pocketlaw pivoted to enterprise CLM, Rocket Lawyer UK has 8 staff and £5m of net liabilities, Legal Utopia is in liquidation. Lawyer-delivered subscriptions and fixed-fee marketplaces (Lawhive, >$35m run-rate) are the models that have grown.</li>
 <li><b>The PE-proven version of subscription legal is employment-law/HR compliance.</b> Peninsula (£463m revenue, £118m EBITDA), Citation (£275m, KKR/Hg), WorkNest/Axiom GRC (Inflexion), Avensure (£19m at a 30% EBITDA margin) sell multi-year, per-employee, insurance-backed subscriptions with 24/7 helplines. NatWest Mentor is being sold to Limerston-backed Empowering People Group.</li>
 <li><b>Consumer digital legal is consolidating at low prices.</b> Farewill sold to Dignity at £12.9m EV (vs a £70m 2020 valuation); amicable and Octopus Legacy sit inside Octopus Group; Wenup, Muve and Simply-Docs all changed hands in 2026.</li>
 <li><b>Demand is large and under-served.</b> 32% of small businesses have a legal issue each year (1.8m firms) but only 25% get professional help and only 10% see lawyers as cost-effective. UK legal market £55.5bn (2025), about half B2B.</li>
 <li><b>Regulation enables rather than blocks.</b> The SRA alternative business structure regime allows PE ownership of law firms (LegalVision UK, Harper James, Lawhive, Markel Law, Co-op all operate through licensed bodies); changes of control need SRA approval.</li>
</ul>

<h2>Market map</h2>
<div class="grid">{''.join(mapcards)}</div>

<h2>Benchmark: core comparables</h2>
<div class="tbl"><table>
<thead><tr><th>Company</th><th>Segment</th><th class="r">Revenue</th><th>Period</th><th class="r">Growth</th><th class="r">EBITDA</th><th class="r">Avg staff</th><th class="r">Rev / head</th><th>Ownership</th></tr></thead>
<tbody>{bench}</tbody></table></div>
<p class="note">Revenue and average employees from filed accounts unless noted; n.d. = not disclosed (small-company filing omits turnover). Growth is shown only for comparable 12-month periods. Lawhive's revenue is a company-reported global run-rate delivered by ~500 consultant solicitors, so revenue per head is not meaningful. Citation EBITDA is after £33.5m one-off transaction costs (underlying c.£77–82m, derived).</p>

<h2>Transactions and valuation anchors</h2>
<div class="tbl"><table>
<thead><tr><th>Date</th><th>Target</th><th>Investor / acquirer</th><th class="r">Value</th><th>Notes</th></tr></thead>
<tbody>{deal_rows}</tbody></table></div>
<div class="tbl"><table>
<thead><tr><th>Comparable</th><th>Revenue</th><th class="r">EV / revenue</th><th class="r">EV / EBITDA</th></tr></thead>
<tbody>{comp_rows}</tbody></table></div>
<p class="note">Listed multiples at 12 Oct 2026 (stockanalysis.com). Harper James/LDC, LegalVision/Quadrant and Lawhive Series B terms were not disclosed.</p>

<h2>Other details worth collecting</h2>
<div class="tbl"><table>
<thead><tr><th>Data point</th><th>Why it matters and where it shows up</th></tr></thead>
<tbody>{kpi_rows}</tbody></table></div>

<h2>Caveats</h2>
<p>Most small UK companies file abridged or micro accounts without turnover, so revenue is shown only where filed, company-reported or in press, and the workbook states the basis for each figure. Companies House reports average headcount, not FTEs, and excludes consultant lawyers. Three common assumptions proved wrong: no Markel/LawBite deal exists (LawBite failed), Docue did not acquire Sparqa or Farillio, and LegalVision's stated A$13m UK revenue does not reconcile with its £3.4–3.7m filed turnover.</p>
<p class="src">Sources: Companies House (iXBRL and scanned accounts), SRA register, company websites and pricing pages, Legal Futures, Law Gazette, RollOnFriday, Capital Brief, Fortune, RNS/Investegate, Beauhurst, stockanalysis.com, Legal Services Board small-business and individual legal-needs surveys, IRN Legal Reports, TheCityUK. Full URLs are in the workbook.</p>
</div>
"""
out = sys.argv[1]
with open(out, 'w', encoding='utf-8') as f: f.write(page)
print('wrote', out, len(page))
