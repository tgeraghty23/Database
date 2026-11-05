// Build the Word briefing for the UK online / subscription legal services market map.
// Usage: node research/scripts/build_uk_online_legal_doc.js research/uk_online_legal_services_market_map.docx
const fs = require('fs'), path = require('path');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, LevelFormat, BorderStyle, Footer, PageNumber, PageBreak } = require('docx');

const D = path.join(__dirname, '..', 'data');
const load = n => JSON.parse(fs.readFileSync(path.join(D, n), 'utf8'));
const SME = load('uk_online_legal_sme_subscription.json'), HR = load('uk_online_legal_hr_employment.json');
const CON = load('uk_online_legal_consumer_documents.json'), CTX = load('uk_online_legal_context.json');
const REC = {}; [...SME, ...HR].forEach(x => REC[x.name] = x); CON.forEach(x => { if (!REC[x.name]) REC[x.name] = x; });

const FONT = 'Calibri';
const P = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts, children: Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: 21 })] });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun({ text: t })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun({ text: t })] });
const NOTE = t => P([new TextRun({ text: t, font: FONT, size: 18, italics: true, color: '595959' })]);
const B = (runs) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 80 }, children: runs });
const bold = t => new TextRun({ text: t, font: FONT, size: 21, bold: true });
const txt = t => new TextRun({ text: t, font: FONT, size: 21 });

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths, size = 17) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (t, w, head, i) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: '1F3864', type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 50, bottom: 50, left: 80, right: 80 },
    children: [new Paragraph({ alignment: (!head && typeof t === 'number') ? AlignmentType.RIGHT : AlignmentType.LEFT,
      children: [new TextRun({ text: fmt(t), font: FONT, size, bold: head || i === 0, color: head ? 'FFFFFF' : '000000' })] })] });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true, i)) }),
           ...rows.map(r => new TableRow({ cantSplit: true, children: r.map((t, i) => cell(t, widths[i], false, i)) }))] });
}
function fmt(v) { if (v === null || v === undefined) return 'n.d.'; if (typeof v === 'number') return v.toLocaleString('en-GB'); return String(v); }
const num = v => (typeof v === 'number' ? v : null);
const firstInt = s => { if (s == null) return null; if (typeof s === 'number') return s; const m = String(s).match(/(\d[\d,]*)/); return m ? parseInt(m[1].replace(/,/g, '')) : null; };
const money = v => v == null ? 'n.d.' : (v >= 1e6 ? `£${(v / 1e6).toFixed(1)}m` : `£${Math.round(v / 1e3)}k`);
const pct = v => v == null ? '–' : `${v > 0 ? '+' : ''}${Math.round(v * 100)}%`;
const is12m = t => { const m = String(t || '').match(/(\d+)[- ]?months?/); return !m || m[1] === '12'; };
const short = (s, n) => { if (!s) return ''; s = String(s); if (s.length <= n) return s; const cut = s.slice(0, n); const i = Math.max(cut.lastIndexOf('. '), cut.lastIndexOf('; ')); return (i > n * 0.5 ? cut.slice(0, i + 1) : cut.trimEnd() + '…'); };
const PRIOR = { 'Harper James': 12221570 }, REV = { 'Lawhive': 26000000 }, REV_FY = { 'Lawhive': 'run-rate Feb 2026 (company-reported, >$35m, global)' }, NO_RPE = new Set(['Lawhive']);

// Relevance tiers (mirrors build_uk_online_legal.py)
const TIER = {
 'LegalVision UK': ['Core', 'Subscription law firm (benchmark)'], 'Harper James': ['Core', 'Subscription law firm'], 'Lawyerly (extra)': ['Core', 'Subscription law firm'],
 'Blackmont Legal (extra)': ['Core', 'Subscription law firm'], '360 Business Law / 360 Law Group': ['Core', 'Subscription legal department (consultant lawyers)'],
 'Law 365': ['Core', 'Subscription law firm (hours-based)'], 'LegalEdge': ['Core', 'Fractional GC retainer'], 'Sprintlaw UK (extra)': ['Core', 'Fixed-fee online + membership (unregulated)'],
 'Rocket Lawyer UK': ['Core', 'Legal plan membership'], 'LawBite (defunct)': ['Defunct', 'Freemium platform + ABS (administration 2024)'], 'Linkilaw': ['Defunct', 'Exited SME segment'],
 'SeedLegals': ['Adjacent', 'Startup legal SaaS + unlimited support'], 'Genie AI': ['Adjacent', 'AI contract SaaS'], 'Docue (UK)': ['Adjacent', 'Contract SaaS'],
 'Sparqa Legal': ['Adjacent', 'Self-serve documents (dormant)'], 'Farillio': ['Adjacent', 'Self-serve documents, B2B2C via insurers'], 'Pocketlaw / Miramis (UK)': ['Adjacent', 'Pivoted to enterprise CLM'],
 'Lawhive': ['Adjacent', 'Fixed-fee marketplace / ABS (consumer-led)'], 'Summize (CLM SaaS - adjacent)': ['Out', 'Enterprise CLM'], 'Juro (CLM SaaS only)': ['Out', 'Enterprise CLM'],
 'Legl (law-firm SaaS - adjacent)': ['Out', 'Law-firm SaaS'], 'Other subscription offers noted (not full profiles)': ['Adjacent', 'Long tail'],
 'Peninsula Group (Peninsula Business Services)': ['Core', 'Employment-law/HR subscription (insurance-backed)'], 'Citation Group': ['Core', 'Employment-law/HR/H&S subscription (PE-owned)'],
 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC': ['Core', 'Employment-law/HR subscription (PE-owned)'], 'Avensure': ['Core', 'Employment-law/HR subscription'],
 'Croner (Peninsula Group)': ['Core', 'Employment-law/HR subscription'], 'Wirehouse Employer Services (WorkNest / Axiom GRC)': ['Core', 'Employment-law/HR subscription'],
 'Quest Cover (Quest Business Services, formerly Qdos Consulting)': ['Core', 'HR/H&S advice lines, white-label'], 'NatWest Mentor (being sold to Empowering People Group)': ['Core', 'Employment-law/HR subscription (bank channel)'],
 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)': ['Core', 'ABS running B2B2C legal helplines'], 'Lawgistics (motor trade legal subscription)': ['Core', 'Vertical legal membership'],
 'The HR Dept (franchise network)': ['Adjacent', 'Franchised HR retainers'], 'Stallard Kane (Ardonagh Advisory)': ['Adjacent', 'H&S/HR consultancy within a broker'], 'BrightHR (Peninsula Group)': ['Adjacent', 'HR SaaS with advice lines'],
 'Croner-i (Peninsula Group)': ['Adjacent', 'Compliance content subscription'], 'Health Assured (Peninsula Group)': ['Out', 'EAP'], 'ELAS (Employment Law Advisory Services) - now part of Citation': ['Defunct', 'Absorbed into Citation'],
 'FSB Legal Hub (Federation of Small Businesses member legal benefits)': ['Adjacent', 'Distribution channel'], 'Moorepay (Zellis Group) - payroll with HR/employment law advice': ['Adjacent', 'Payroll-led'],
 'ARAG UK (incl. former DAS) / Howden - legal expenses insurance with SME legal helplines (brief)': ['Adjacent', 'Legal expenses insurance'],
 'Empowering People Group (AdviserPlus / Halborns / Learning Nexus; acquirer of NatWest Mentor)': ['Adjacent', 'PE-backed HR/employment-law group'], 'Safe Workplace (not identified)': ['Out', 'Not identified'],
 'Farewill': ['Core', 'Digital wills/probate (Dignity)'], 'Octopus Legacy (formerly Guardian Angel)': ['Core', 'Digital wills/probate (Octopus)'], 'amicable': ['Core', 'Online divorce (Octopus)'],
 'Divorce-Online': ['Core', 'Online divorce + ABS'], 'Wenup': ['Core', 'Online prenups (Blackfinch)'], 'Which? Legal': ['Core', 'Consumer legal advice subscription'], 'JustAnswer UK': ['Adjacent', 'Expert Q&A subscription (US)'],
 'Co-op Legal Services': ['Adjacent', 'Scaled fixed-fee probate/wills ABS'], 'Legal Utopia': ['Defunct', 'In liquidation 2026'], 'Slater and Gordon UK (brief)': ['Out', 'National consumer law firm'],
 'ARAG Law (formerly DAS Law) (brief)': ['Adjacent', 'Captive LEI law firm'], 'Juno (conveyancing)': ['Adjacent', 'Digital conveyancing'], 'Muve (Connect2Law)': ['Adjacent', 'Digital conveyancing (Alcuin)'],
 'Kinherit': ['Adjacent', 'Wills + digital vault, IFA channel'], 'Make a Will Online': ['Adjacent', 'Online wills'], 'Beyond (beyond.life)': ['Defunct', 'Dissolved 2024'], 'Bequest (unverified)': ['Out', 'Unverified'],
 'Settled (status unclear)': ['Out', 'Status unclear'], 'Simply-Docs': ['Core', 'Document subscription (sold Jul 2026)'], 'Net Lawman': ['Adjacent', 'Document templates'], 'Lawpack': ['Adjacent', 'DIY legal kits'],
 'Legalo': ['Adjacent', 'Document templates'], 'LawDepot UK': ['Adjacent', 'Document subscription (Canadian)'], 'Wonder.Legal (UK)': ['Adjacent', 'Document builder (French)'], 'Contractbook (UK)': ['Out', 'CLM (acquired by Scrive)'],
 'Summize': ['Out', 'Enterprise CLM'], 'Juro': ['Out', 'Enterprise CLM'], 'Avvoka': ['Out', 'Document automation for law firms'], 'Lexoo': ['Defunct', 'Absorbed into Kalexius'],
};
const tier = n => TIER[n] || ['Adjacent', ''];
const clean = n => n.replace(/ \((extra|defunct|brief|unverified|status unclear|not identified)\)/, '').replace(/ - now part of Citation| \(CLM SaaS - adjacent\)| \(CLM SaaS only\)| \(law-firm SaaS - adjacent\)/, '');

function profile(x) {
  const [t, tag] = tier(x.name);
  const rev = REV[x.name] || num(x.revenue_gbp), prior = PRIOR[x.name] || num(x.revenue_prior), fy = REV_FY[x.name] || x.revenue_fy;
  const emp = firstInt(x.employees);
  const facts = [];
  if (rev) facts.push(`Revenue ${money(rev)} (${short(fy, 60)}${prior && is12m(fy) ? `, ${pct(rev / prior - 1)}` : ''})`);
  else facts.push(`Revenue not filed${x.revenue_prior ? ': ' + short(x.revenue_prior, 160) : ''}`);
  if (num(x.ebitda)) facts.push(`EBITDA ${money(x.ebitda)}`);
  if (emp) facts.push(`${emp} average employees (${short(x.employees_year, 60)})`);
  const runs = [bold(clean(x.name)), txt(` (${tag}). `), txt(short(x.services, 330) + ' ')];
  runs.push(new TextRun({ text: facts.join('; ').replace(/[.…]+$/, '') + '. ', font: FONT, size: 21, color: '1F3864' }));
  if (x.pricing_model) runs.push(txt('Pricing: ' + short(x.pricing_model, 200) + ' '));
  if (x.ownership) runs.push(txt('Ownership: ' + short(x.ownership, 200) + ' '));
  if (x.regulatory_status) runs.push(txt('Regulation: ' + short(x.regulatory_status, 110)));
  return new Paragraph({ spacing: { after: 140 }, children: runs });
}
const order = { Core: 0, Adjacent: 1, Defunct: 2, Out: 3 };
const sortRecs = recs => [...recs].sort((a, b) => (order[tier(a.name)[0]] - order[tier(b.name)[0]]) || ((REV[b.name] || num(b.revenue_gbp) || 0) - (REV[a.name] || num(a.revenue_gbp) || 0)));

function mapTable(recs) {
  const rows = sortRecs(recs).filter(x => tier(x.name)[0] !== 'Out').map(x => {
    const rev = REV[x.name] || num(x.revenue_gbp), emp = firstInt(x.employees);
    return [clean(x.name), tier(x.name)[0] === 'Defunct' ? 'Defunct / exited' : tier(x.name)[0], tier(x.name)[1], rev ? money(rev) : 'n.d.', emp ? String(emp) : 'n.d.', short((x.ownership || '').split(/[.;(]/)[0], 48)];
  });
  return table(['Company', 'Relevance', 'Model', 'Revenue', 'Staff', 'Ownership'], rows, [1900, 1050, 2400, 950, 650, 2076], 16);
}

const c = [];
c.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'UK online and subscription legal services', font: FONT, bold: true, size: 40, color: '1F3864' })] }));
c.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'Market map and benchmarking against LegalVision', font: FONT, size: 28, color: '1F3864' })] }));
c.push(NOTE('Briefing compiled 12 October 2026 from Companies House filings, SRA register, company websites, trade press and deal announcements. Amounts in GBP unless stated. "n.d." = not disclosed (small-company accounts omit turnover). Full dataset, source URLs and caveats: uk_online_legal_services_market_map.xlsx.'));

c.push(H1('1. Summary'));
c.push(P('The brief was to find UK digitally native businesses selling online, subscription-style legal services in the mould of LegalVision, with revenue, headcount and a description of services. The map covers 76 businesses in three segments: subscription and online legal services for SMEs (22), subscription employment-law, HR and compliance advisory for SMEs (21), and consumer digital legal services and document platforms (33). Each is tagged Core, Adjacent, Defunct/exited or Out of scope relative to LegalVision.'));
[
 [bold('Direct LegalVision analogues are few and small. '), txt('Harper James is the only scaled UK subscription law firm: £15.3m revenue in FY Sep-25 (+25%), £2.2m operating profit, 131 average employees, about 70% of income from subscription plans, and an LDC minority stake completed in September 2025. LegalVision UK filed £3.7m turnover for FY Jun-24 (restated £3.35m) with 41 average staff in FY25; FY25 turnover was not filed. Management\'s stated A$13m UK revenue does not reconcile with the filings. Lawyerly, Blackmont, Law 365 and 360 Business Law are sub-£2m or do not file turnover.')],
 [bold('Self-serve platforms have not scaled in the UK. '), txt('LawBite entered administration in September 2024 (platform sold for £1 after raising about £7m), Sparqa is dormant, Pocketlaw rebranded as Miramis and pivoted to enterprise contract software, Rocket Lawyer UK is down to 8 staff with £5m of net liabilities, Legal Utopia is in liquidation, Beyond was dissolved and Lexoo was absorbed. Lawyer-delivered subscriptions (Harper James, LegalVision) and the fixed-fee marketplace model (Lawhive, >$35m run-rate, $60m Series B in February 2026) are the models that have grown.')],
 [bold('The PE-proven version of subscription legal is employment-law and HR compliance. '), txt('Peninsula (£463m revenue, £118m EBITDA, family-owned), Citation (£275m, KKR and Hg, acquired at about 15x EBITDA in 2020), WorkNest within Axiom GRC (Inflexion paid £430m, 16x EBITDA, for the Marlowe division in 2024) and Avensure (£19m revenue at a 30% EBITDA margin, founder-owned) sell multi-year, per-employee, insurance-backed subscriptions with 24/7 helplines. NatWest Mentor (c.220 staff, c.100,000 SMEs) is being sold to Limerston-backed Empowering People Group.')],
 [bold('Consumer digital legal is consolidating at low prices. '), txt('Farewill sold to Dignity at £12.9m enterprise value, about 1.4x revenue and well below its £70m 2020 valuation. Octopus Group owns both amicable and Octopus Legacy. Wenup (Blackfinch), Muve (Alcuin) and Simply-Docs all changed hands in 2026.')],
 [bold('Demand is large and under-served. '), txt('Legal Services Board research shows 32% of small businesses have a legal issue each year (1.8m firms, 15.4m issues), but only 25% use professional help, 44-46% handle it alone and only 10% see lawyers as cost-effective. The UK legal market was £55.5bn in 2025 (IRN), about half of it business-to-business.')],
 [bold('Valuation anchors. '), txt('Listed UK law firms trade at 0.8-1.6x EV/revenue and roughly 5-14x EV/EBITDA (Keystone 1.6x and 14x; Knights 1.3x and 5-6x; Gateley 0.8x and 10x). LegalZoom trades at about 1.1x revenue and 5x adjusted EBITDA with 65% subscription revenue. Compliance subscription businesses have sold at 15-16x EBITDA. Terms for Harper James/LDC, LegalVision/Quadrant and Lawhive were not disclosed.')],
 [bold('Regulation enables rather than blocks. '), txt('The SRA alternative business structure regime allows non-lawyer and PE ownership of law firms; LegalVision UK, Harper James, Lawhive, Markel Law and Co-op Legal Services all operate through licensed bodies. Changes of control need SRA approval (about three months for Harper James). Unregulated models avoid that cost but cannot carry out reserved work such as litigation.')],
].forEach(r => c.push(B(r)));

c.push(H1('2. Market map'));
c.push(P('Relevance: Core = lawyer-delivered subscription, membership or fixed-fee legal service for SMEs or consumers; Adjacent = self-serve platform, marketplace, software, channel or insurance-led substitute. Out-of-scope names (enterprise contract software, EAP, national law firms) are profiled in the workbook only.'));
c.push(H2('2.1 SME subscription and online legal services'));
c.push(mapTable(SME.filter(x => !x.name.startsWith('Other subscription'))));
c.push(NOTE('Long tail of subscription offers from traditional firms (not profiled): LawPlan (£49.50-99/month, unregulated), Doyle Clayton (startup plan from £100/month), EMW Certitude, Steven Mather VIP Legal (£15-25k/year), CP Law Associates (£75/month), Keith (AI-native firm, £2m seed 2026).'));
c.push(H2('2.2 Subscription employment law, HR and compliance advisory'));
c.push(mapTable(HR.filter(x => !/Safe Workplace|ELAS/.test(x.name))));
c.push(H2('2.3 Consumer digital legal and document platforms'));
c.push(mapTable(CON.filter(x => !/Bequest|Settled/.test(x.name))));

c.push(new Paragraph({ children: [new PageBreak()] }));
c.push(H1('3. Benchmark: core comparables'));
const BN = ['Harper James', 'LegalVision UK', 'Lawhive', 'SeedLegals', '360 Business Law / 360 Law Group', 'Law 365', 'Rocket Lawyer UK', 'Genie AI', 'Lawyerly (extra)', 'Sprintlaw UK (extra)',
 'Peninsula Group (Peninsula Business Services)', 'Citation Group', 'Croner (Peninsula Group)', 'BrightHR (Peninsula Group)', 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC', 'Avensure', 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)', 'Wirehouse Employer Services (WorkNest / Axiom GRC)',
 'Co-op Legal Services', 'Muve (Connect2Law)', 'Farewill', 'amicable', 'Divorce-Online', 'Octopus Legacy (formerly Guardian Angel)', 'Simply-Docs', 'Juro'];
const OWN = { 'Harper James': 'Founder majority; LDC ~45%', 'LegalVision UK': 'LegalVision group; Quadrant PE', 'Lawhive': 'VC (GV, TQ, Balderton, Rales)', 'SeedLegals': 'VC (Index 2019); founders', '360 Business Law / 360 Law Group': 'Founder/family', 'Law 365': 'Founder', 'Rocket Lawyer UK': 'Rocket Lawyer Inc (US)', 'Genie AI': 'VC (GV, Khosla)', 'Lawyerly (extra)': 'Founder', 'Sprintlaw UK (extra)': 'Sprintlaw Pty (AU)',
 'Peninsula Group (Peninsula Business Services)': 'Done family', 'Citation Group': 'KKR / Hg / HarbourVest', 'Croner (Peninsula Group)': 'Peninsula Group', 'BrightHR (Peninsula Group)': 'Peninsula Group', 'WorkNest (formerly Ellis Whittam / Law At Work) - Axiom GRC': 'Inflexion (Axiom GRC)', 'Avensure': 'Founder', 'Markel Law / Markel UK legal services (incl. FSB Legal Hub provider)': 'Markel Group', 'Wirehouse Employer Services (WorkNest / Axiom GRC)': 'Axiom GRC (Inflexion)',
 'Co-op Legal Services': 'Co-operative Group', 'Muve (Connect2Law)': 'Alcuin Capital', 'Farewill': 'Dignity / Castelnau', 'amicable': 'Octopus Group', 'Divorce-Online': 'Private', 'Octopus Legacy (formerly Guardian Angel)': 'Octopus Group', 'Simply-Docs': 'Private (sold Jul 2026)', 'Juro': 'VC (USV, Eight Roads)' };
const SEG = n => SME.some(x => x.name === n) ? 'SME' : HR.some(x => x.name === n) ? 'HR / emp. law' : 'Consumer / docs';
c.push(table(['Company', 'Segment', 'Revenue', 'Period', 'Growth', 'EBITDA', 'Staff', 'Rev / head', 'Ownership'], BN.map(n => {
  const x = REC[n]; const rev = REV[n] || num(x.revenue_gbp), prior = PRIOR[n] || num(x.revenue_prior), fy = REV_FY[n] || x.revenue_fy || ''; const emp = firstInt(x.employees);
  const fyS = String(fy).replace(/\(.*?\)/g, '').replace('FY to ', '').trim().slice(0, 24);
  return [clean(n), SEG(n), rev ? money(rev) : 'n.d.', fyS || '–', (rev && prior && is12m(fy)) ? pct(rev / prior - 1) : '–', num(x.ebitda) ? money(x.ebitda) : '–', emp ? String(emp) : 'n.d.', (rev && emp && !NO_RPE.has(n)) ? `£${Math.round(rev / emp / 1e3)}k` : '–', OWN[n]];
}), [1700, 1000, 800, 1150, 800, 750, 550, 750, 1526], 16));
c.push(NOTE('Revenue and average employees from filed accounts unless noted. Growth shown only for comparable 12-month periods. Lawhive revenue is a company-reported global run-rate delivered by about 500 consultant solicitors, so revenue per head is not meaningful. Citation EBITDA (£43.4m) is after £33.5m one-off transaction costs and £4.6m deferred remuneration; underlying c.£77-82m is derived. Farewill revenue covers 17 months to December 2024 (c.£6.7m annualised).'));

c.push(H1('4. LegalVision: the benchmark'));
const lv = CTX.legalvision, uk = lv.uk_entity_financials_companies_house.value;
[
 ['Model', 'Subscription ("membership") law firm founded in Sydney in 2012 by Lachlan McKnight, Evan Tait-Styles and Ursula Hogben; began as an online document service, became a law practice in 2014 and adopted the membership model in 2019. Members get unlimited business-as-usual legal work (consultations, drafting, reviews, trade marks, employment/HR) for a fixed weekly fee billed monthly; disputes, corporate transactions and fundraising are billed hourly at a 50% member discount. Minimum 12-month term with 3- and 5-year discounts. Revenue mix in 2022: about 50% subscription, 25% hourly, 25% fixed fee.'],
 ['Pricing', 'UK: LV Pro £115 to £620+ per week plus VAT by business size and term (website, October 2026); £115-399 in March 2025; £49-329 at the 2022 launch. Australia: LV Pro A$219 to A$1,350+ per week; LV Essentials A$119-149 per week. New Zealand launched January 2021 at NZ$199 per month.'],
 ['Scale', 'Management told Capital Brief (October 2024) that revenue was about A$50m in Australia and New Zealand plus A$13m in the UK, with a target of A$100m by 2027. More than 4,500 active clients group-wide and 30%+ annual growth (Quadrant, March 2025); about 40% of new business comes from the UK. 2022: about 100 new members a month, 85% renewal, NPS 70. Australia 200+ staff; UK about 40 (28 lawyers, 12 sales) with a target of 100.'],
 ['Ownership', 'Founder-owned through the McKnight Family Trust until Quadrant Private Equity invested in March 2025 (terms undisclosed; first deal from its A$660m Growth Fund No.3, which writes A$50-100m cheques). UK filings show the ultimate parent changed to a Quadrant vehicle (QPE Epsilon 1 Pty Ltd) and Quadrant partners became persons with significant control from May 2025, consistent with Quadrant control.'],
 ['UK entity', `LegalVision Law UK Ltd (13704561), Manchester, incorporated 26 October 2021, SRA licensed body 837339 from 13 May 2022. Turnover: £${uk.turnover_26Oct2021_to_31Oct2022_gbp.toLocaleString('en-GB')} (first period to Oct 2022), £${uk.turnover_8m_to_30Jun2023_gbp.toLocaleString('en-GB')} (8 months to Jun 2023), £${uk.turnover_FY_to_30Jun2024_gbp.toLocaleString('en-GB')} (FY Jun-24, restated £${uk.turnover_FY2024_restated_gbp.toLocaleString('en-GB')}); FY Jun-25 small-company accounts omit the income statement (profit £6.7k). Average employees 9 (FY22), 15, 24, 41 (FY25). FY24 gross margin 64%, operating profit £101k, net liabilities £0.58m funded by an intercompany loan at 11%. The A$13m UK revenue claim (~£6.5m) does not reconcile with filed turnover of £3.35-3.69m.`],
 ['Australian peers in the UK', 'Sprintlaw (Sydney fixed-fee and membership firm) launched in England and Wales in June 2021; UK entity has 5 staff and net liabilities of £0.66m; membership launched at £399 a year; 3,000+ members across UK, AU and NZ (2024). Lawpath (Sydney SaaS legal platform, A$10m from Westpac in February 2025, 500,000+ Australian businesses) has not entered the UK.'],
].forEach(([k, v]) => c.push(P([bold(k + '. '), txt(v)])));

c.push(H1('5. Transactions and valuation anchors'));
c.push(table(['Date', 'Target', 'Investor / acquirer', 'Value', 'Notes'], [
 ['Mar 2025', 'LegalVision (AU/NZ/UK)', 'Quadrant PE growth stake', 'Undisclosed', '4,500+ clients, 30%+ growth; Quadrant vehicle now ultimate parent of UK entity'],
 ['Jun-Sep 2025', 'Harper James', 'LDC minority (~45%)', 'Undisclosed', '"£16m business" at the deal; Investec debt; target £30m revenue by 2028'],
 ['Feb 2026', 'Lawhive', 'Series B led by Mitch Rales; GV, TQ, Balderton', '$60m', 'Revenue >$35m, 7x year on year; >$100m raised in total'],
 ['May 2024', 'Marlowe GRC division (WorkNest, VinciWorks)', 'Inflexion', '£430m EV', '16.2x adjusted EBITDA; Marlowe had paid £61m for Ellis Whittam in 2020'],
 ['Aug 2020', 'Citation Group', 'KKR (Hg re-invested; HarbourVest)', '£450-525m (press)', '~15x EBITDA on £30-35m; Hg had paid £185m in 2016; recapitalised Dec 2024'],
 ['Mar 2026', 'NatWest Mentor', 'Empowering People Group (Limerston Capital)', 'Undisclosed', 'c.220 staff, c.100,000 SMEs; completion expected late 2026'],
 ['Oct 2024', 'Farewill', 'Dignity (Castelnau)', '£12.9m EV', '~1.4x annualised revenue; £70m valuation in 2020; ~£30m VC raised'],
 ['Nov 2023', 'amicable', 'Octopus Group (majority)', 'Undisclosed', 'FY Apr-25 revenue £3.4m (+136%), loss £1.6m'],
 ['Nov 2022', 'Octopus Legacy (ex-Guardian Angel)', 'Octopus Group (majority)', 'Undisclosed', 'FY Apr-25 revenue £1.3m, loss £4.0m; debt-funded by Octopus Capital'],
 ['Mar 2026', 'Muve / Connect2Law', 'Alcuin Capital Partners (majority)', 'Undisclosed', 'FY25 revenue £19.3m (+42%), PAT £3.7m'],
 ['Jun 2026', 'Wenup', 'Blackfinch Group', 'Undisclosed', 'Online prenups; 2 staff'],
 ['Jul 2026', 'Simply-Docs', 'Private buyer', 'Undisclosed', 'Founder exit; £38.50-125 a year subscriptions; 10 staff'],
 ['Jan 2026', 'Summize', 'Kennet Partners, Federated Hermes PE', '$50m', 'ARR +97%; enterprise CLM'],
 ['Sep 2024', 'LawBite (Lawbit Ltd)', 'Administration', 'Platform sold for £1', '£5.8m owed to 33 creditors; ~£7m raised from 400+ investors'],
 ['Jul 2023', 'DWF Group', 'Inflexion take-private', '£342m', '~1.0x revenue; first top-100 UK firm taken private by PE'],
 ['Sep 2024', 'Stowe Family Law', 'Investcorp (from Livingbridge)', 'Undisclosed', 'Revenue >£37m'],
 ['May 2026', 'Setfords', 'Phoenix Equity Partners running a sale (Baird)', 'n/a', 'Revenue £67.7m (+35%), ~650 consultant lawyers'],
], [950, 2100, 2050, 1150, 2776], 16));
c.push(H2('Listed and reference comparables (12 October 2026)'));
c.push(table(['Comparable', 'Revenue', 'Margin', 'EV / revenue', 'EV / EBITDA'], [
 ['Keystone Law (AIM)', '£97.7m FY Jan-25 (+11%); TTM £129m', 'Adj. PBT 13%', '1.6x', '14.1x'],
 ['Knights (AIM)', '£207.7m FY Apr-26 (+28%)', 'Underlying EBITDA 24.8%', '1.3x', '~5-6x'],
 ['Gateley (AIM)', '£194.3m FY Apr-26 (+8%)', 'Adj. operating 11.1%', '0.8x', '10.0x'],
 ['LegalZoom (NASDAQ)', '$756m FY25 (+11%); 65% subscription', 'Adj. EBITDA 23%', '1.1x', '~5x adjusted'],
 ['Rocket Lawyer (private, US)', 'Not disclosed', 'n.d.', 'Series E 2021: $223m at ~$530m pre-money', '–'],
 ['Citation (private)', '£275m FY24', 'Underlying EBITDA c.28-30% (derived)', '–', '~15x at 2020 entry'],
 ['Marlowe GRC (carve-out)', '~£86m FY23 pro forma', 'Adj. EBITDA 37%', '5.0x', '16.2x'],
], [2000, 2700, 1800, 1300, 1226], 16));

c.push(H1('6. Market context'));
c.push(table(['Metric', 'Value', 'Year'], CTX.market.filter(m => !/Global online|IBISWorld|Law Society economic|Financial Benchmarking|exports|AI-referred|AI adoption - UK SMEs/.test(m.metric)).map(m => [short(m.metric, 90), short(m.value, 230), String(m.year)]), [2700, 5600, 726], 16));

c.push(H1('7. Other details worth collecting'));
c.push(P('Beyond revenue, headcount and service descriptions, these are the data points a PE benchmarking exercise on subscription legal usually turns on. The workbook records where each is captured and where it is not.'));
[
 ['Subscription share of revenue and ARR', 'Separates true recurring revenue from fixed-fee or hourly work marketed as subscription. LegalVision was about 50% subscription in 2022, Harper James about 70% on plans, LegalZoom 65%.'],
 ['Renewal rate, churn and net revenue retention', 'LegalVision cited 85% renewal and NPS 70; Peninsula and Citation renew clients onto new multi-year terms. Management-only data; Trustpilot complaint themes are the public proxy.'],
 ['Contract length and auto-renewal terms', 'Peninsula 1/3/5-year terms with auto-renewal and a 2% uplift; LegalVision 12-month minimum with 3/5-year discounts; Mentor 3-month rolling. Longer terms give visibility but carry CMA and FCA auto-renewal scrutiny.'],
 ['Members, ARPU and average contract value', 'Needed to triangulate revenue where turnover is not filed (LegalVision UK FY25, SeedLegals, 360 Business Law). LegalVision group has 4,500+ clients; Harper James 1,000+ Enable subscribers.'],
 ['Gross margin and lawyer utilisation', 'Lawyer-delivered subscriptions are capacity-constrained; margin depends on utilisation and seniority mix. LegalVision UK FY24 gross margin 64%; Juro (software) 77%.'],
 ['Sales-to-lawyer ratio and revenue per lawyer', 'LegalVision UK: 28 lawyers and 12 sales staff out of about 40; Avensure 56 sales versus 56 consultants. Indicates customer acquisition intensity.'],
 ['Channel mix', 'Direct digital versus B2B2C through banks, insurers, trade bodies and IFAs (Markel/FSB, Mentor/NatWest, Farillio/insurers, Kinherit/IFAs). Partner-dependent businesses carry concentration risk, as the Mentor sale shows.'],
 ['Regulatory structure and licence date', 'SRA licensed-body status allows PE ownership and reserved work; changes of control need SRA approval. Unregulated models are cheaper but capped in scope.'],
 ['Insurance backing and indemnity exposure', 'Employment-law subscriptions sell tribunal indemnities (Citation £1.5m guarantee; Avensure and Peninsula via legal expenses policies). Underwriter and claims history matter in diligence.'],
 ['Balance-sheet signals', 'Small-company filings omit the P&L but show net assets, cash and intercompany debt: an early warning for LawBite, Legal Utopia and Pocketlaw UK, and a measure of parent dependence for Rocket Lawyer UK and LegalVision UK.'],
 ['Three-year headcount trajectory', 'Filed even by small companies and a better growth signal than press claims: Harper James 64 to 131; LegalVision UK 15 to 41.'],
 ['Pricing evolution', 'LegalVision UK membership moved from £49-329 a week (2022) to £115-620+ a week (2026): pricing power or repositioning upmarket. Wayback Machine snapshots give the history.'],
 ['Ownership and likely sellers', 'Founder-owned at scale (Peninsula, Avensure, 360), stale VC vintages (SeedLegals 2019, Juro 2022), corporate orphans (Rocket Lawyer UK, Mentor) and PE vintages (Citation 2020, WorkNest 2024, Harper James 2025) indicate who may transact.'],
 ['Customer satisfaction', 'Trustpilot and NPS where public: Harper James 4.6, Lawhive 4.4, Farewill 4.9, Which? Legal 1.9, LawDepot 1.8 (auto-renewal complaints).'],
 ['Web traffic and AI leverage', 'Similarweb or Semrush for acquisition scale (AI-referred traffic to UK law firms rose 310% in 2025); AI-assisted delivery at Lawhive, Harper James (LDC capital earmarked for AI), Genie and Sprintlaw.'],
].forEach(([k, v]) => c.push(B([bold(k + '. '), txt(v)])));

c.push(new Paragraph({ children: [new PageBreak()] }));
c.push(H1('Appendix A. Company profiles'));
c.push(NOTE('Core and adjacent businesses, sorted by relevance then revenue. Figures in blue are from filed accounts or company statements as labelled. Full fields and source URLs are in the workbook.'));
const prof = (recs, title) => { c.push(H2(title)); sortRecs(recs).filter(x => ['Core', 'Adjacent'].includes(tier(x.name)[0]) && !/^Other subscription|Safe Workplace|FSB Legal Hub|Bequest|Settled/.test(x.name)).forEach(x => c.push(profile(x))); };
prof(SME, 'A1. SME subscription and online legal services');
prof(HR, 'A2. Subscription employment law, HR and compliance advisory');
prof(CON.filter(x => x.segment === 'consumer'), 'A3. Consumer digital legal services');
prof(CON.filter(x => x.segment === 'documents'), 'A4. Legal document and contract platforms');
c.push(H2('A5. Defunct, exited or out of scope'));
c.push(P('LawBite (administration September 2024; platform sold for £1), Linkilaw (repositioned to UHNW disputes), Legal Utopia (creditors\' voluntary liquidation July 2026), Beyond (dissolved August 2024), Lexoo (absorbed into Kalexius 2024), ELAS (absorbed into Citation 2021). Out of scope but profiled in the workbook: Summize, Juro, Legl, Avvoka, Contractbook (enterprise contract software), Health Assured (EAP), Slater and Gordon (national consumer law firm).'));

c.push(H1('Appendix B. Caveats'));
[
 'Revenue: most small UK companies file abridged or micro accounts without turnover. Figures are shown only where filed, company-reported or in press, and the workbook states the basis for each. Where turnover is missing, net assets, cash and headcount from the balance sheet are given instead.',
 'Employees, not FTEs: Companies House reports average monthly headcount. Consultant and fee-share lawyers (Lawhive ~500, 360 Business Law 750+, Setfords ~650) are not employees.',
 'Three common assumptions proved wrong: there was no Markel acquisition of LawBite (Markel Law is Markel\'s own ABS); Docue did not acquire Sparqa or Farillio; Setfords is Phoenix-backed and Lawfront is Blixt-backed.',
 'Some filings are image-only PDFs (Summize, Legl, parts of Co-op, Farewill and Juro); figures read visually are flagged in the workbook. A few deal rows (Robin AI, vLex/Clio, Luminance) and third-party revenue estimates are marked unverified.',
 'Not researched: Definely, Robin AI, Luminance; gunnercooke and Taylor Rose; Howden legal-expenses products; any 2026 SeedLegals round.',
].forEach(t => c.push(B([txt(t)])));

const doc = new Document({
  creator: 'Research', title: 'UK online and subscription legal services: market map',
  styles: { default: { document: { run: { font: FONT, size: 21 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 28, bold: true, color: '1F3864' }, paragraph: { spacing: { before: 320, after: 120 }, keepNext: true, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true, run: { font: FONT, size: 23, bold: true, color: '1F3864' }, paragraph: { spacing: { before: 220, after: 80 }, keepNext: true, outlineLevel: 1 } },
    ] },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 240 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1250, bottom: 1250, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: 'UK online and subscription legal services market map, compiled 12 October 2026. Page ', font: FONT, size: 16, color: '595959' }), new TextRun({ children: [PageNumber.CURRENT], font: FONT, size: 16, color: '595959' })] })] }) },
    children: c }],
});
Packer.toBuffer(doc).then(b => { fs.writeFileSync(process.argv[2], b); console.log('wrote', process.argv[2], b.length); });
