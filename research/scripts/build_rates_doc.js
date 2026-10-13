const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, LevelFormat, BorderStyle, ExternalHyperlink, Footer, PageNumber } = require('docx');
const FONT = 'Calibri';
const run = (t, o = {}) => new TextRun({ text: t, font: FONT, size: 21, ...o });
const P = (t, o = {}) => new Paragraph({ spacing: { after: 120 }, ...o, children: Array.isArray(t) ? t : [run(t)] });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, children: [run(t, { bold: true, size: 28, color: '1F3864' })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 }, children: [run(t, { bold: true, size: 23, color: '1F3864' })] });
const NOTE = t => P([run(t, { size: 18, italics: true, color: '595959' })]);
const B = (text, date, src, url) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 80 }, children: [
  run(text), run(date ? ` [${date}]` : '', { size: 19, bold: true, color: '595959' }),
  ...(src ? [run(' Source: ', { size: 18, italics: true, color: '595959' }),
    url ? new ExternalHyperlink({ link: url, children: [run(src, { size: 18, italics: true, style: 'Hyperlink' })] }) : run(src, { size: 18, italics: true, color: '595959' })] : [])] });
const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths, hl) {
  const cell = (t, w, head, shade) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: '1F3864', type: ShadingType.CLEAR, color: 'auto' } : shade ? { fill: shade, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [run(String(t), { size: 18, bold: head, color: head ? 'FFFFFF' : '000000' })] })] });
  return new Table({ width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) }),
      ...rows.map(r => new TableRow({ children: r.map((t, i) => cell(t, widths[i], false, hl && hl(i, t))) }))] });
}
const RT = 'Realtid', AV = 'Affärsvärlden';
const U = {
  flygt: 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/',
  flygt2: 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/',
  delphi: 'https://www.realtid.se/juridik/timarvoden-for-jurister-en-modell-pa-vag-ut/',
  aidisc: 'https://www.realtid.se/juridik/kunderna-vill-ha-ai-rabatt-advokatbyraerna-hojer-priset/',
  banks: 'https://www.realtid.se/juridik/storbankerna-kraver-ai-rabatt-av-advokatbyraerna/',
  tull: 'https://www.realtid.se/bors-finans/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/',
  av25: 'https://www.affarsvarlden.se/artikel/vinge-gar-om-msa-sa-gar-det-for-affarsjuristerna',
  av24: 'https://www.affarsvarlden.se/artikel/rekordar-vantas-2024-byraerna-har-blivit-annu-battre-pa-att-ta-betalt',
  av23: 'https://www.affarsvarlden.se/artikel/affarsjuristerna-om-trenden-pe-bolagen-kommer-vara-motorn',
  hb: 'https://www.realtid.se/bors-finans/nu-satter-handelsbanken-press-pa-hoga-advokatarvoden/',
  no: 'https://www.realtid.se/bors-finans/norska-protester-mot-dyra-advokater-nu-kraver-bankerna-fast-pris/',
  noreg: 'https://www.realtid.se/juridik/regeringsadvokaten-slar-larm-advokatkostnaderna-skenar/',
  regi: 'https://www.realtid.se/bors-finans/advokatbyraernas-prismodeller-forandras/',
  code: 'https://www.realtid.se/debatt/rimligt-att-ersattningen-innehaller-ett-moment-av-reglering',
  maqs: 'https://www.realtid.se/juridik/advokatbyra-delas-malmo-bryter-sig-loss/',
  lx: 'https://www.lexisnexis.com/community/amp-pressroom/lexisnexis-counsellink-releases-2026-trends-report-as-rising-rates-and-big-law-share-of-wallet-continue-to-climb',
  wsj: 'https://www.realtid.se/juridik/toppjuristers-arvoden-exploderar-samtidigt-hotar-ai-botten/',
  uk: 'https://www.legalcheek.com/2025/04/the-2000-per-hour-solicitor',
  legora: 'https://www.legaltechmonitor.com/2026/04/legora-raises-50m-series-d-extension-atlassian-and-nvidia-fund-join-as-investors/',
  skatt: 'https://www.xn--skattertt-12a.se/byra/vinge-skatt/',
  dist: 'https://www.realtid.se/juridik/vinge-tappar-mark-mannheimer-swartling-drar-ifran/',
  nv: 'https://www.affarsvarlden.se/artikel/sa-dyrt-blev-northvolts-raddningsforsok-i-usa-debiteras-47-000-timmar',
  lulea: 'https://www.realtid.se/juridik/ratten-om-timarvoden-4-000-kronor-i-timmen-ar-ok/',
};
const c = [];
c.push(new Paragraph({ spacing: { after: 60 }, children: [run('Charging rates in Nordic business law', { bold: true, size: 40, color: '1F3864' })] }));
c.push(P([run('Market data briefing for a Nordic business-law investment: rate levels, pricing trends, AI impact and international leading indicators. Compiled 13 October 2026. Amounts are SEK per hour excluding VAT unless stated; each data point shows the period it relates to in [brackets]. Sections 1–5 cover the Nordic market; section 6 covers international (mainly US/UK) trends and whether they have reached the Nordics.', { size: 20, italics: true, color: '595959' })]));

c.push(H1('Key takeaways'));
[
 ['Top-tier Swedish partner list rates are SEK 6,000–10,000/h, associates SEK 3,600–4,150/h, junior associates SEK 2,300–3,350/h.', '2025–2026', 'Rate cards; Realtid (Olle Flygt), 31 Mar 2026', U.flygt],
 ['MAQS (sell-side model) prices partners at ~SEK 5,015/h and associates at ~SEK 2,753/h: roughly 30–37% below Mannheimer Swartling and Cederquist mid-points, consistent with a mid-market positioning (to be confirmed whether model rates are list or realised).', '2026', 'Sell-side model', ''],
 ['Swedish firms have shown pricing power: 2024 market growth of 8.4% with only 3.7% headcount growth, and fee income per lawyer of SEK 4.1–4.2m.', '2023–2024', 'Affärsvärlden, 5 Oct 2024', U.av24],
 ['AI is changing the model rather than cutting prices so far: Nordic firms report more work, not less, and some (Delphi) argue senior rates rise as AI removes junior hours; subscriptions and fixed fees are emerging.', '2025–2026', 'Realtid, 15 Apr 2026', U.delphi],
 ['US leading indicators (7%+ annual rate rises, USD 3,000+/h elite partners, bank "AI discounts") have largely not yet reached the Nordics; fixed-fee pressure from Nordic banks is the main exception.', '2021–2026', 'See section 6', ''],
].forEach(x => c.push(B(...x)));

c.push(H1('1. Nordic rate levels by seniority'));
c.push(P('Rate cards provided directly (Mannheimer Swartling, Cederquist) and the MAQS sell-side model:'));
c.push(table(['Level', 'Mannheimer Swartling (date not stated)', 'Cederquist (2025)', 'MAQS (sell-side model)'], [
  ['Partner / Senior Advisor', '6,000 – 10,000', '6,500 – 8,000', '5,015'],
  ['Senior Associate / Counsel (MSA: 60+ months, incl. PDL, Specialist Counsel)', '4,550 – 8,000', '4,500 – 5,800', 'n/a'],
  ['Associate (MSA: 25–59 months)', '3,650 – 4,150', '3,600 – 4,100', '2,753'],
  ['Junior Associate (MSA: 0–24 months)', '2,750 – 3,350', '2,300 – 3,100', 'n/a'],
  ['Paralegal', 'n/a', '1,500', 'n/a'],
], [3300, 2050, 1800, 1876], (i) => i === 3 ? 'FFF2CC' : null));
c.push(P(''));
c.push(table(['Comparison (mid-points)', 'Partner', 'Associate'], [
  ['Mannheimer Swartling vs Cederquist', '+10%', '+1%'],
  ['MAQS (sell-side model) vs Cederquist', '−31%', '−28%'],
  ['MAQS (sell-side model) vs Mannheimer Swartling', '−37%', '−29%'],
], [4626, 2200, 2200]));
c.push(P(''));
c.push(NOTE('MAQS rates are taken from the sell-side model and are single rates rather than published ranges. If they are realised (effective) rates after discounts and write-offs, the gap to list-rate cards overstates the true price difference; if they are list rates, MAQS is clearly positioned below the top tier. The Mannheimer Swartling card is undated.'));
c.push(H2('Market commentary on Nordic rate levels'));
[
 ['Partners at top Swedish firms bill SEK 7,000–9,000/h, with 3–5 associates per partner billing ~SEK 3,000/h who "cost at most a third" of what they bill (Olle Flygt, ex-Vinge and Mannheimer Swartling partner).', '2026', `${RT}, 31 Mar 2026`, U.flygt],
 ['Flygt billed SEK 7,000–8,000/h in his final years at Mannheimer Swartling and argues advice "has become too expensive".', 'pre-2025', `${RT}, 26 Apr 2026`, U.flygt2],
 ['Indicative firm-wide ranges (third-party directory, tax practices, low reliability): MSA 4,500–8,000; Roschier 4,200–7,000; Vinge and DLA Piper 4,000–7,000; Schjødt and Cederquist 3,800–6,500; Setterwalls and Bird & Bird 3,500–6,000; Delphi 3,200–5,500; Lindahl 3,000–5,500.', 'accessed Oct 2026', 'skatterätt.se', U.skatt],
 ['Norway: commercial litigation rates that were ~NOK 3,000/h have "doubled in a short time" according to the Norwegian Government Attorney.', '2026', `${RT}, 5 Jun 2026`, U.noreg],
].forEach(x => c.push(B(...x)));

c.push(H2('Top-firm economics: profit distribution, rates and leverage (Vinge and Mannheimer Swartling)'));
c.push(P('Based on Realtid\'s interview series with Olle Flygt (advokat with three decades at Vinge and Mannheimer Swartling), published 1 April 2026, supplemented with FY2025 annual-report data:'));
[
 ['Business model: "hours times price". A partner bills SEK 7,000–9,000/h; beneath each partner sit three to five associates billing ~SEK 3,000/h who "cost at most a third" of that. The result is "a very efficient billing machine" that takes firms such as Roschier, Vinge and Mannheimer Swartling to SEK 1bn+ revenue.', '2026', `${RT}, 1 Apr 2026`, U.flygt],
 ['Leverage: "It is built on leverage. The more people under you, the better." And: "the bigger you make the matters, the bigger the invoices."', '2026', `${RT}, 1 Apr 2026`, U.flygt],
 ['Profit distribution differs: Vinge rewards the individual partner\'s importance and contribution; Mannheimer Swartling shares profit equally (lockstep) with a deduction for junior partners. Some partners could switch firm and earn double, but it rarely happens; the two firms watch each other closely.', '2026', `${RT}, 1 Apr 2026`, U.flygt],
 ['Profit distributed (FY2025): Mannheimer Swartling proposed SEK 617m (up from 582m), ~SEK 6.8m per partner across 92 partners; Vinge Stockholm SEK 273m (down from 286m), ~SEK 5.5m each across 49 partners; Vinge Malmö SEK 32m to 14 partners.', 'FY2025', `${RT} (Dagens Juridik), 7 Aug 2026`, U.dist],
 ['Ownership: law firms can only be owned by their partners (no listing, unlike e.g. EQT), which brings privileges such as protection from police searches and confidentiality, and makes profit distribution central.', '2026', `${RT}, 1 Apr 2026`, U.flygt],
 ['Conflict: "There is a conflict between the law firm\'s demand for high profitability and what is the best advice." Flygt calls big firms a "golden cage", financially attractive but hard to leave.', '2026', `${RT}, 1 Apr and 26 Apr 2026`, U.flygt2],
].forEach(x => c.push(B(...x)));
c.push(P([run('Illustrative leverage economics: ', { bold: true }), run('an associate billing SEK 3,000/h at a cost of at most ~SEK 1,000/h contributes ≥SEK 2,000 per billed hour; at 1,400–1,600 billed hours that is ~SEK 2.8–3.2m per associate per year, or ~SEK 8–16m per partner with three to five associates, before overheads. This is why leverage (and its exposure to AI, see section 4) drives partner profit.')]));
c.push(H2('Realised rates, 2023 onwards'));
[
 ['Mannheimer Swartling\'s realised blended rate on the Northvolt restructuring was ~SEK 4,000/h (~6,000 hours, SEK 32m), the lowest of all advisers; this compares with list partner rates of SEK 6,000–10,000/h and illustrates the gap between list and realised rates.', 'Nov 2024 – Mar 2025', `${AV}, 2 May 2025`, U.nv],
 ['Luleå tingsrätt held SEK 4,000/h reasonable for law-firm counsel in a cost ruling.', 'Apr 2026', `${RT}, 25 Apr 2026`, U.lulea],
].forEach(x => c.push(B(...x)));
c.push(P('Both points sit close to the ~SEK 2,600–2,950/h blended realised rate implied by industry fee income per lawyer (section 5), suggesting realised blended rates of roughly SEK 3,000–4,000/h at top firms.'));

c.push(H1('2. Nordic rate trends and client pushback'));
[
 ['Affärsvärlden: market growth 8.4% and operating margin 13.5% while lawyer headcount grew only 3.7%; firms are "better at charging".', '2024', `${AV}, 5 Oct 2024`, U.av24],
 ['Rates are revised annually on 1 January (Cederquist); no Swedish firm was found publicly announcing a specific percentage increase for 2025 or 2026.', '2025–2026', 'Cederquist fee structure', ''],
 ['Firms expect fee income to grow 6.8% in 2026; IPO work was the weakest segment in 2025.', '2026', `${RT}, 26 Apr 2026`, U.flygt2],
 ['Handelsbanken (Sweden) shares Norwegian banks\' concerns over fee levels and is reviewing alternative pricing models.', '2021', `${RT}, 1 Dec 2021`, U.hb],
 ['Norwegian banks: DNB required at least 50% of mandates on fixed fees in its previous framework; Handelsbanken Norway wants fixed fees.', '2021', `${RT}, 25 Nov 2021`, U.no],
 ['Client surveys (Regi): value for money among the lowest-rated criteria since 2008; price is the most common reason for switching firm.', '2008–2017', `${RT}, 15 Dec 2017`, U.regi],
].forEach(x => c.push(B(...x)));

c.push(H1('3. Nordic billing models and rules'));
[
 ['Fees must be reasonable (skäligt) under the Swedish Bar\'s code (rule 4.1.1); the assessment may weigh the agreement, scope, difficulty, importance, skill and result (4.1.2). Outcome-based elements rarely matter in practice; exact success-fee restrictions to be verified.', '2019', `${RT} op-ed, 29 Mar 2019`, U.code],
 ['Time-based billing remains the default (Cederquist: "time incurred basis", open to alternative arrangements per project).', '2025', 'Cederquist fee structure', ''],
 ['Firms owned by external investors must leave the Bar (Advokatsamfundet bars external ownership of advokat firms); AGRD\'s member firms did so, which also takes them outside the Bar\'s fee rules.', '2025–2026', 'See workbook section F', ''],
].forEach(x => c.push(B(...x)));

c.push(H1('4. AI impact'));
c.push(H2('4.1 Impact on the business model'));
[
 ['Leverage model at risk: top-firm economics rely on associates billing ~SEK 3,000/h at a third of cost; AI automates much of the junior work that drives this margin.', '2026', `${RT}, 31 Mar 2026`, U.flygt],
 ['Roschier: AI has not reduced the volume of work; output can be produced faster but the amount of searchable data is almost unlimited and everything must be source-checked.', '2025', `${RT}, 3 Oct 2025`, U.tull],
 ['Setterwalls: international clients increasingly require firms to use generative AI and show how it improves advice, as a condition for winning work.', '2025', `${RT}, 3 Oct 2025`, U.tull],
 ['Investment by leading firms: Mannheimer Swartling runs an AI incubator (MSA Innovation Lab, where Legora was founded) and has bought an AI company; Vinge partners with Harvey AI; Gernandt & Danielsson uses Legora (as a "second opinion" in disputes).', '2025', `${AV}, 12 Apr 2025; ${RT}, 3 Oct 2025`, U.av25],
 ['Legora (Swedish legal-AI platform) raised a USD 550m Series D at ~USD 5.5bn valuation, extended by USD 50m in April 2026, signalling capital flowing into tools that substitute for associate hours.', 'Mar–Apr 2026', 'Legal Tech Monitor', U.legora],
 ['MAQS split: the Malmö practice (keeping the MAQS name from 1 Jan 2027) plans an "AI-forward premium" firm, while the Stockholm/Göteborg/Sundsvall offices (Keelion Law) pursue an integrated model.', 'Aug 2026', `${RT}, 20 Aug 2026`, U.maqs],
].forEach(x => c.push(B(...x)));
c.push(H2('4.2 Impact on pricing and charging rates'));
[
 ['Delphi (Agnes Hammarstrand): "Colleagues tell me that what we did in a working day now takes five minutes. So how do we charge?" She argues hourly billing is "almost an outdated way to charge", favours fixed monthly subscriptions, and sees AI efficiency pushing senior rates up as junior hours disappear.', '2026', `${RT}, 15 Apr 2026`, U.delphi],
 ['Baker McKenzie (which has a Stockholm office) has made a "strategic decision to largely abandon the billable-hour model".', '2026', `${RT}, 15 Apr 2026`, U.delphi],
 ['Client expectation of AI-driven fee reductions is international, not yet Nordic-specific: in a European/US GC survey, 79% say AI should lower fees, 84% have seen no reduction and 72% believe firms keep the savings.', '2026', `${RT}, 3 Oct 2026`, U.aidisc],
 ['No Nordic client or bank has publicly demanded explicit "AI discounts" yet (US banks have; see section 6).', 'Oct 2026', `${RT}, 1 Sep 2026`, U.banks],
].forEach(x => c.push(B(...x)));
c.push(P([run('Implications for an investor: ', { bold: true }), run('(i) revenue per partner may hold or rise as senior rates increase, but associate-hour volumes and leverage are at risk; (ii) firms with fixed-fee, subscription or value-based offerings are better placed if clients push for AI savings; (iii) diligence should test the share of revenue from junior-heavy, automatable work (due diligence, document review, standard drafting) and the firm\'s AI adoption and pricing roadmap.')]));

c.push(H1('5. Revenue per lawyer and implied realised rates (Sweden)'));
[
 ['Affärsvärlden panel fee income per lawyer: SEK 4.22m (2023) and SEK 4.13m (2024); White & Case Stockholm has the highest fee per lawyer.', '2023–2024', `${AV}, 29 Apr 2024; ${RT}, 3 Oct 2025`, U.av23],
 ['Illustrative: SEK 4.13m per lawyer implies a blended realised rate of ~SEK 2,580/h at 1,600 billable hours or ~SEK 2,950/h at 1,400 hours, well below list rates, reflecting leverage, write-offs and discounts. Notably close to the MAQS sell-side associate rate.', '2024', 'Calculation', ''],
].forEach(x => c.push(B(...x)));

c.push(H1('6. International pricing trends (leading indicators)'));
c.push(P('Leading indicators from the US and UK markets, and whether each has reached the Nordic market:'));
c.push(table(['Indicator', 'Market / period', 'Reached the Nordics?', 'Comment'], [
  ['Annual rate increases of ~7% (US firms +7.3% avg; average top-100 lawyer > USD 1,000/h)', 'US, 2025 (Thomson Reuters / Georgetown)', 'Partly', 'Swedish firms show pricing power (8.4% market growth vs 3.7% headcount, 2024) but no announced rate rises of this size'],
  ['Partner rates +5.1%; M&A partner rates +8.8%', 'US, 2025 invoices (LexisNexis CounselLink)', 'Not evidenced', 'No Nordic rate-inflation data published'],
  ['Elite partner rates USD 3,000–3,400/h; top-lawyer rates up 16% in a year', 'US, 2025–26 (WSJ)', 'No', 'Top Swedish partners at SEK 7,000–10,000/h (~USD 700–1,000), a fraction of US elite levels'],
  ['UK top-10 firms average GBP 449/h (2019: GBP 321)', 'UK, 2024 (PwC)', 'Partly', 'Swedish top-tier blended rates are below London; Norway reports litigation rates doubling'],
  ['Clients/banks demanding explicit "AI discounts" (Goldman Sachs, Morgan Stanley, Citi)', 'US, 2026', 'Not yet', 'No public demand by Nordic banks; Nordic banks\' pushback so far is on fixed fees (DNB, Handelsbanken)'],
  ['GC expectation that AI should lower fees (79%) but no reductions seen (84%)', 'Europe/US, 2026 (Juro)', 'Emerging', 'Setterwalls: international clients require AI use and evidence of benefit'],
  ['Shift from billable hour to fixed fees, subscriptions, outcome-based pricing (90% of US spend still hourly)', 'US/global, 2026 (Thomson Reuters)', 'Emerging', 'Delphi subscriptions; Baker McKenzie abandoning the billable hour; DNB 50% fixed-fee requirement (Norway)'],
  ['External / PE capital into law firms (MSO structures)', 'US, 2025–26', 'Yes', 'AGRD (Axcel) roll-up in Sweden since 2025; firms must leave the Bar'],
], [2700, 1900, 1200, 3226], (i, t) => i === 2 ? ({ 'Yes': 'E2EFDA', 'Partly': 'FFF2CC', 'Emerging': 'FFF2CC', 'Not yet': 'FCE4D6', 'No': 'FCE4D6', 'Not evidenced': 'FCE4D6' })[t] : null));
c.push(P(''));
[
 ['US: firms raised rates 7.3% on average in 2025; average top-100 lawyer above USD 1,000/h.', '2025', `${RT}, 3 Oct 2026`, U.aidisc],
 ['US: LexisNexis CounselLink partner rates +5.1%, M&A partner rates +8.8%.', '2025', 'LexisNexis, 22 Apr 2026', U.lx],
 ['US: elite partners up to USD 3,000–3,400/h; top-lawyer rates at top-50 firms +16% in a year.', '2025–2026', `${RT} (WSJ), 19 Feb 2026`, U.wsj],
 ['US banks demand AI discounts in 2026 panel reviews.', '2026', `${RT}, 1 Sep 2026`, U.banks],
 ['UK: top-10 firms average GBP 449/h; firms 11–25 average GBP 325 (PwC; snippet only).', '2024', 'Legal Cheek, Apr 2025', U.uk],
].forEach(x => c.push(B(...x)));

c.push(H1('7. Data gaps and next steps'));
[
 ['Confirm whether MAQS sell-side model rates are list or realised rates, and obtain the rate split for senior associates/counsel and juniors.', '', '', ''],
 ['Confirm the date of the Mannheimer Swartling rate card.', '', '', ''],
 ['No published rate cards found for Vinge, Setterwalls, Lindahl, Gernandt & Danielsson, Delphi, Cirio or other top-20 Nordic firms; client invoices (e.g. via Axo, requested) would add realised-rate evidence.', '', '', ''],
 ['No Nordic data on annual rate increases or AI-related discounts; consider expert calls with Nordic GCs/procurement leads.', '', '', ''],
].forEach(x => c.push(B(...x)));

const doc = new Document({ creator: 'Research', title: 'Charging rates in Nordic business law',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: ['Page ', PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children: c }] });
Packer.toBuffer(doc).then(b => fs.writeFileSync(process.argv[2], b));
