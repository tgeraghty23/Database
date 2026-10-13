// Builds the market size summary and data book guide (Word).
// Usage: node build_market_summary_doc.js <out.docx> [--extras <module.js>]  (extras: optional private add-on content, kept outside the repo)
const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, LevelFormat, BorderStyle, Footer, PageNumber } = require('docx');
const FONT = 'Calibri';
const xi = process.argv.indexOf('--extras');
const X = xi > 0 ? require(require('path').resolve(process.argv[xi + 1])) : {};
const FX = { USD: 10.01386, EUR: 11.1965, NOK: 1.04684, DKK: 1.4979 }; // Riksbank fixing 12 Oct 2026, SEK per unit
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (t, o = {}) => new Paragraph({ spacing: { after: 120 }, ...o, children: Array.isArray(t) ? t : [run(t)] });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, children: [run(t, { bold: true, size: 28, color: '1F3864' })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 }, children: [run(t, { bold: true, size: 23, color: '1F3864' })] });
const Bl = (label, text) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 80 }, children: [run(label ? label + ' ' : '', { bold: true }), run(text)] });
const NOTE = t => P([run(t, { size: 18, italics: true, color: '595959' })]);
const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths, shadeRow) {
  const cell = (t, w, head, shade) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: '1F3864', type: ShadingType.CLEAR, color: 'auto' } : shade ? { fill: shade, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(String(t), { size: 18, bold: head, color: head ? 'FFFFFF' : '000000' })] })] });
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) }),
      ...rows.map((r, j) => new TableRow({ children: r.map((t, i) => cell(t, widths[i], false, shadeRow && shadeRow(j))) }))] });
}
const c = [];
if (X.banner) c.push(X.banner({ P, run }));
c.push(new Paragraph({ spacing: { after: 60 }, children: [run('Swedish and Nordic business-law market: size and data book guide', { bold: true, size: 40, color: '1F3864' })] }));
c.push(P([run('Summary of the market-size evidence for Sweden and the Nordics, and a guide to each tab of the legal market data book. Compiled October 2026. Amounts in SEK; other currencies converted at the Riksbank fixing of 12 Oct 2026 (USD 10.01, EUR 11.20, NOK 1.047, DKK 1.498 SEK).', { size: 20, italics: true, color: '595959' })]));

c.push(H1('1. Market size at a glance'));
c.push(P('There is no official definition of the "business-law" (affärsjuridik) market. The estimates below differ mainly because they measure different scopes:'));
const sek = (v, cur) => `SEK ${(v * FX[cur]).toFixed(1)}bn`;
const rows = [
  ['Business law: fee income of the 56 largest firms (Affärsvärlden panel)', '~SEK 16bn (+5.3%)', 'SEK 16bn', '2025', 'Most-cited figure; large firms only (floor)'],
  ['Same panel, prior year (revised)', 'SEK 15.4bn (+8.6%); SEK 4.13m per lawyer', 'SEK 15.4bn', '2024', 'Afv, 3 Oct 2025'],
  ['Business-law market, total (industry estimate)', '>SEK 25bn p.a.', '>SEK 25bn', 'n/a', 'AGRD Partners op-ed; method not given'],
  ['All law firms (advokatbyråer), estimate', '~SEK 29bn', 'SEK 29bn', '2024e', '2021 SCB share (74.5%) applied to 2024 total (ceiling: incl. non-business work)'],
  ['All legal activities, SCB SNI 69.1 (official)', 'SEK 38.9bn', 'SEK 38.9bn', '2024', '6,965 companies; 13,446 employees; +6.6% vs 2023'],
  ['All legal services (MarketLine)', 'USD 3.1bn', sek(3.1, 'USD'), '2024', 'Incl. taxes; CAGR 5.9% 2019-24'],
  ['All legal activities (IBISWorld)', 'EUR 3.5bn', sek(3.5, 'EUR'), '2026e', 'CAGR -0.5% 2020-25 (conflicts with other sources)'],
];
(X.marketRows || []).forEach(r => rows.push(r));
c.push(table(['Scope', 'Figure (as reported)', 'SEK equivalent', 'Year', 'Comment'], rows, [2900, 1800, 1300, 800, 2226], j => j >= rows.length - (X.marketRows || []).length ? 'FCE4D6' : null));
c.push(P(''));
c.push(P([run('Working range: ', { bold: true }), run('SEK ~16bn (large business-law firms) to ~29bn (all law firms), with the Swedish business-law market most often put at SEK 16-25bn.')]));

c.push(H2('Official statistics (SCB, net turnover SEK m)'));
c.push(table(['Year', 'SNI 69.1 total', '69.101 Advokatbyråer', '69.102 Other legal', '69.103 Patent agents'], [
  ['2010', '18,738', '13,904', '1,956', '2,879'], ['2015', '23,986', '17,936', '2,750', '3,300'], ['2019', '31,237', '22,277', '4,194', '4,765'],
  ['2021', '34,693', '25,864', '4,558', '4,270'], ['2022*', '33,276', 'n/p', 'n/p', 'n/p'], ['2023*', '36,525', 'n/p', 'n/p', 'n/p'], ['2024*', '38,921', 'n/p', 'n/p', 'n/p'],
], [1300, 1900, 2000, 1900, 1926]));
c.push(NOTE('* New SCB series from 2022, not directly comparable with 2021; sub-industry splits not published. Advokatbyrå turnover grew ~5.8% p.a. 2010-21. 2024 value added SEK 25.3bn; EBITDA margin 28%.'));

c.push(H2('Concentration and growth'));
[['Concentration:', 'top 5 firms ~46% and top 10 ~66% of the Affärsvärlden panel (FY2024). Vinge and Mannheimer Swartling each exceed SEK 2bn; five firms exceed SEK 1bn.'],
 ['Growth:', 'panel fee income +8.6% (2024) and +5.3% (2025); firms expect +6.8% in 2026. 2025 growth came from international transactions, disputes and regulatory work; IPOs were weakest.'],
 ['Supply:', '6,712 Bar members (5,963 practising) and 2,525 associates at end-2025 (Advokatsamfundet).'],
 ['Consolidation:', 'AGRD (Axcel) has combined 11 firms since Aug-2025; Nordic entrants (BAHR, CMS) and domestic mergers continue (data book Section F).'],
].forEach(([a, b]) => c.push(Bl(a, b)));

c.push(H1('2. Nordic comparison: Sweden, Norway, Denmark, Finland'));
c.push(P('Official statistics cover all legal activities; only Norway publishes a business-law split. Business-law figures for Sweden, Denmark and Finland are estimates. SEK at the Riksbank fixing of 12 Oct 2026.'));
c.push(table(['', 'Sweden', 'Norway', 'Denmark', 'Finland'], [
  ['Legal activities, official (2024)', 'SEK 38.9bn', 'NOK 27.5bn (SEK 28.8bn)', 'DKK 19.5bn (SEK 29.2bn)', 'EUR 1.18bn (SEK 13.2bn); 2025: EUR 1.28bn'],
  ['Business law (official / estimate)', 'SEK 16bn (panel floor) to >25bn', 'NOK 15.6bn official (SEK 16.3bn), 57% of total', '~DKK 13bn est. (SEK 19.5bn), ~67% business clients', '~EUR 0.55-0.65bn est. (SEK 6-7bn)'],
  ['Growth (local currency)', '~8% p.a. 2022-24; 5.8% p.a. 2010-21', '9.3% p.a. 2019-24; business law +14.9% in 2024', '4.5% p.a. 2019-24', '4.9% p.a. 2019-25; +8.4% in 2025'],
  ['Enterprises / employees (Eurostat 2024)', '6,965 / 17,521', '2,214 / 9,264', '1,912 / 13,158', '2,726 / 6,509'],
  ['Turnover per employee (EUR k, 2024)', '194', '255', '198', '181'],
  ['Lawyers', '6,712 Bar members (5,963 practising)', '10,756 Bar members incl. trainees', '~7,760 incl. in-house', '2,443 asianajajat + 1,778 licensed counsel'],
  ['Largest firm', 'Vinge SEK 2.04bn (2025)', 'Schjødt NOK 2.50bn (2025, consolidated)', 'Gorrissen Federspiel DKK 1.31bn (2025)', 'Roschier EUR 99m (FY05/2026)'],
  ['Concentration', 'Top-5 ~46%, top-10 ~66% of Afv panel', 'Top-5 ~31% of sector (NOK 9.4bn, 2025)', 'Top-10 ~44% (2022) to ~47-49% (2025e)', 'Top-5 ~27% of sector; ~55% of top-30'],
  ['Revenue per employee, large firms', 'SEK 3.1-3.9m', 'NOK 4.7-5.5m (SEK 4.9-5.8m)', 'DKK 2.4-2.55m (SEK 3.6-3.8m)', 'EUR 205-384k (SEK 2.3-4.3m)'],
  ['External / PE ownership', 'Only by leaving the Bar (AGRD precedent)', 'Lawyers Act 2025: no passive ownership of advokat firms', 'Lawyer-only ownership; advocates\' court monopoly', 'Bar firms lawyer-owned; licensed-counsel route; Fondia listed'],
], [2000, 1750, 1750, 1750, 1776]));
c.push(P(''));
c.push(P([run('Nordic total: ', { bold: true }), run('legal activities ~SEK 110bn (2024; Eurostat basis ~SEK 107bn); business law ~SEK 59-68bn, with Sweden the largest business-law market (~30-35%), followed by Denmark, Norway and Finland.')]));
[['Norway', 'is the most profitable per head (EUR 255k turnover per employee) and fastest growing, with strong Norwegian firms expanding into Sweden and Denmark (Schjødt, BAHR).'],
 ['Denmark', 'is the most concentrated market among the large firms and has the strictest framework (lawyer-only ownership and a court-representation monopoly).'],
 ['Finland', 'is the smallest, with a very profitable top tier (EBIT margins ~37-54% at Roschier, Krogerus, Avance, Borenius) and a fragmented mid-tier of firms of EUR 6-38m revenue.'],
 ['Sweden', 'is the largest and most fragmented business-law market, and the only one with an established PE roll-up (AGRD).'],
].forEach(([a, b]) => c.push(Bl(a, b)));
c.push(H2('Consolidation and international activity, 2020-2026'));
c.push(table(['Country', 'Activity'], [
  ['Sweden', 'AGRD (Axcel) PE roll-up of 9+ firms since Aug-2025 (Allié, Born, Morris Law, Next, Synch, TM & Partners; then Wigge, Bokwall Rislund, Moll Wendén, Setterwalls Malmö as Ascente Law; Synch-Qap Legal); CMS-Wistrand (2024); BAHR Stockholm (2026); Schjødt-Hamilton (2020); Magnusson-Gro and Walthon Ramberg (2026); splits: Hannes Snellman (2024), MAQS (2027), Glimstedt Stockholm/Göteborg (2021)'],
  ['Norway', 'Schjødt-Hamilton SE (2020), Copenhagen office (2022), Norrbom Vinding DK (Mar-2026); BAHR opens Stockholm and Copenhagen (2025-26) after hiring ~18 Schjødt partners; Ræder Bing (2023), Lippestad / Bache-Wiig and Tofte Hald (2024); Cedra buys Big-4 regional law offices (Dec-2025)'],
  ['Denmark', 'HortenDahl merger (Jan-2026); DLA Piper takes 49 people from Lundgrens (Mar-2025); Schjødt-Norrbom Vinding (Mar-2026); Kromann Reumert converts to P/S (Nov-2025)'],
  ['Finland', 'Hannes Snellman split into Finnish and Swedish firms (2024); DLA Piper-Project Law (2023); Merilampi-I&O Partners (2025, single source); no PE investment found; Fondia (listed, non-Bar) shrinking'],
], [1300, 7726]));
c.push(P(''));
c.push(P([run('Swedish team lifts and partner moves (2+ partners): ', { bold: true }), run('Lindahl from Westermark Anjou (2 partners, 2020); MAQS from Delphi (3 M&A partners, 2024); Gernandt & Danielsson from Linklaters (2 partners + 5 associates, real estate, 2024); Cirio from AG Advokat (3 partners + 9, real estate/construction, 2025); BAHR from Schjødt Sweden (8 partners, 2025-26); Setterwalls Malmö to AGRD (12 partners, ~100 people, 2026).')]));
[['PE consolidation:', 'Sweden is the only Nordic market with an established PE roll-up (AGRD: 11 firms in ~13 months); firms must leave the Bar. Norway and Denmark restrict external ownership; no PE deals found in Finland.'],
 ['Cross-border:', 'Norwegian firms (Schjødt, BAHR) are building pan-Scandinavian platforms through mergers and team lifts; CMS and DLA Piper consolidate through network membership and team hires.'],
 ['Net partner flows (Sweden, 2020-26):', 'largest net gainers BAHR (+10), MAQS (+5), Lindahl, Gernandt & Danielsson and Cirio (+2 each); largest net losers Setterwalls (-13, Malmö exit), Schjødt (-7), Snellman (-4), Delphi and AG Advokat (-4 each).'],
 ['Domestic mergers:', 'mostly small and mid-sized combinations (Ræder Bing, HortenDahl, Magnusson-Gro, Walthon Ramberg), consistent with scale pressure from technology costs and talent competition.'],
].forEach(([a, b]) => c.push(Bl(a, b)));
if (X.sections) X.sections({ c, H1, H2, P, Bl, NOTE, run, table }).forEach(x => c.push(x));
c.push(H1((X.sections ? '4' : '3') + '. Guide to the data book'));
const tabs = [
  ['A. Market size', 'Headline estimates and SCB official statistics (above).'],
  ['A. Nordic market', 'FX (Riksbank) and Sweden / Norway / Denmark / Finland comparison: official statistics, business-law estimates, structure, leading firms, ownership rules, consolidation.'],
  ['A. Growth drivers', 'Fastest growers and drivers, tagged lateral hire / M&A / new office / organic / price.'],
  ['A. Niche firms', 'All Chambers / Legal 500 ranked Swedish firms: 9 full-service, ranked independents, PE-owned (AGRD), firms to watch; revenue and revenue per partner.'],
  ['B. Practice rankings', 'Firm x practice-area matrix of Chambers bands and Legal 500 tiers; MAQS boxed, Magnusson / Delphi / Lindahl boxed as a group.'],
  ['B. Rankings raw', 'Full firm-level ranking extract.'],
  ['B. Top lawyers', 'Chambers Bands 1-3 and Legal 500 top categories in Corporate/M&A, PE, Banking & finance, Disputes; count per firm.'],
  ['B. Partner pipeline', '172 ranked lawyers scored A/B/C as lateral-hire targets, with approach angles.'],
  ['C. Benchmarking', 'Top 30 firms: revenue per partner / employee / lawyer, leverage (partners / employees and partners / lawyers), medians.'],
  ['D. Rate cards', 'Mannheimer Swartling, Cederquist and MAQS rate cards; indicative ranges; Nordic realised rates since 2023; benchmarks.'],
  ['E. Partner compensation', 'Profit pool and dividend per partner from filed accounts; press reporting on partner incomes (firm level); SCB lawyer salaries.'],
  ['F. M&A & team lifts', 'Mergers, PE roll-ups, market entries, splits, team lifts and net partner flows, 2020-26.'],
];
(X.tabs || []).forEach(t => tabs.push(t));
c.push(table(['Tab', 'Contents'], tabs, [2400, 6626], j => j >= tabs.length - (X.tabs || []).length ? 'FCE4D6' : null));
c.push(P(''));
c.push(P('Companion document: "Charging rates in Nordic business law" (Word briefing on rate levels, pricing trends, AI impact and international indicators).'));

c.push(H1((X.sections ? '5' : '4') + '. Key caveats'));
[['Definitions:', 'no official business-law market definition; SCB classifies by firm type, not client type.'],
 ['Revenue bases:', 'Affärsvärlden reports fee income; company accounts report net sales of single legal entities; some firms use split financial years.'],
 ['Headcount:', 'partner and lawyer counts are mostly firm-website counts (Oct-2026) or firm-supplied to directories; no FTE data is published, so employees are used as a proxy.'],
 ['Timing:', 'headcounts are current while revenue is FY2024/25, so per-head ratios are indicative.'],
 ['Gaps:', 'Affärsvärlden\'s full FY2025 table and Dagens Juridik\'s 2025 tables are paywalled; limited partnerships (KB) file no public accounts.'],
].forEach(([a, b]) => c.push(Bl(a, b)));
c.push(Bl('Nordic comparability:', 'official statistics measure all legal activities; firm revenues use different fiscal years and entity scopes; Denmark and Finland business-law figures are estimates.'));
(X.caveats || []).forEach(([a, b]) => c.push(Bl(a, b)));

const doc = new Document({ creator: 'Research', title: 'Swedish business-law market: size and data book guide',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: ['Page ', PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children: c }] });
Packer.toBuffer(doc).then(b => fs.writeFileSync(process.argv[2], b));
