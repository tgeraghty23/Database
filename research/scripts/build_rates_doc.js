const fs = require('fs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel, Table, TableRow, TableCell, WidthType, ShadingType,
        AlignmentType, LevelFormat, BorderStyle, ExternalHyperlink, Footer, PageNumber } = require('docx');

const FONT = 'Calibri';
const P = (text, opts = {}) => new Paragraph({ spacing: { after: 120 }, ...opts, children: Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: 21 })] });
const H1 = t => new Paragraph({ heading: HeadingLevel.HEADING_1, spacing: { before: 280, after: 120 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 28, color: '1F3864' })] });
const H2 = t => new Paragraph({ heading: HeadingLevel.HEADING_2, spacing: { before: 200, after: 80 }, children: [new TextRun({ text: t, font: FONT, bold: true, size: 23, color: '1F3864' })] });
// bullet: [text, date, source label, url]
const B = (text, date, src, url) => new Paragraph({ numbering: { reference: 'b', level: 0 }, spacing: { after: 80 }, children: [
  new TextRun({ text, font: FONT, size: 21 }),
  new TextRun({ text: date ? ` [${date}]` : '', font: FONT, size: 19, bold: true, color: '595959' }),
  ...(src ? [new TextRun({ text: ' Source: ', font: FONT, size: 18, italics: true, color: '595959' }),
             url ? new ExternalHyperlink({ link: url, children: [new TextRun({ text: src, font: FONT, size: 18, italics: true, style: 'Hyperlink' })] })
                 : new TextRun({ text: src, font: FONT, size: 18, italics: true, color: '595959' })] : []) ] });

const border = { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' };
const borders = { top: border, bottom: border, left: border, right: border };
function table(headers, rows, widths) {
  const total = widths.reduce((a, b) => a + b, 0);
  const cell = (t, w, head) => new TableCell({ borders, width: { size: w, type: WidthType.DXA },
    shading: head ? { fill: '1F3864', type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 60, bottom: 60, left: 100, right: 100 },
    children: [new Paragraph({ children: [new TextRun({ text: String(t), font: FONT, size: 18, bold: head, color: head ? 'FFFFFF' : '000000' })] })] });
  return new Table({ width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    rows: [new TableRow({ tableHeader: true, children: headers.map((h, i) => cell(h, widths[i], true)) }),
           ...rows.map(r => new TableRow({ children: r.map((t, i) => cell(t, widths[i], false)) }))] });
}
const RT = 'Realtid';
const c = [];
c.push(new Paragraph({ spacing: { after: 60 }, children: [new TextRun({ text: 'Charging rates in Swedish business law', font: FONT, bold: true, size: 40, color: '1F3864' })] }));
c.push(P([new TextRun({ text: 'Market data briefing: hourly rates, rate trends, billing models and benchmarks. Compiled 9 October 2026. All amounts are SEK per hour excluding VAT unless stated. Each data point shows the period it relates to in [brackets].', font: FONT, size: 20, italics: true, color: '595959' })]));

c.push(H1('Key takeaways'));
[
 ['Top-tier partner list rates are around SEK 6,000–10,000/h; associates SEK 3,600–4,150/h; junior associates SEK 2,300–3,350/h.', '2025–2026', 'Mannheimer Swartling and Cederquist rate cards; Realtid (Olle Flygt), 31 Mar 2026', 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
 ['Partner rates have roughly doubled in 15 years: Mannheimer Swartling billed partners at SEK 3,150–3,800/h to public clients in 2009–10.', '2009–2010 vs 2025–26', 'Realtid, 3 Feb 2011', 'https://www.realtid.se/bors-finans/mannheimers-taxa-till-sl-1906-krh/'],
 ['Business lawyers bill "at least three times" the legal-aid norm of SEK 1,626/h (2026).', '2026', 'Advokatsamfundet response to SOU 2025:111, 27 Mar 2026', 'https://www.regeringen.se/contentassets/d498fc7a919d45fbae9ef7b5c02978d8/sveriges-advokatsamfund.pdf'],
 ['Firm-specific rate cards are scarce: public-sector price annexes are usually confidential, so most data comes from invoices, court cost claims and US Chapter 11 filings.', '2017–2026', 'Realtid, 30 Aug 2017', 'https://www.realtid.se/bors-finans/vinges-timpriser-sekretessbelagda/'],
 ['Clients are pushing for AI-driven discounts and fixed fees, but rates keep rising: US firms raised rates ~7% in 2025.', '2025–2026', 'Realtid, 3 Oct 2026', 'https://www.realtid.se/juridik/kunderna-vill-ha-ai-rabatt-advokatbyraerna-hojer-priset/'],
].forEach(x => c.push(B(...x)));

c.push(H1('1. Rate cards by seniority (list prices)'));
c.push(P('The two firm rate cards available (provided directly) are set out below. The Mannheimer Swartling card is undated; Cederquist states 2025 rates, revised annually on 1 January.'));
c.push(table(['Level', 'Mannheimer Swartling (date not stated)', 'Cederquist (2025)', 'MSA vs Cederquist (mid-point)'], [
  ['Partner / Senior Advisor', '6,000 – 10,000', '6,500 – 8,000', '+10%'],
  ['Senior Associate / Counsel (MSA: 60+ months, incl. PDL, Specialist Counsel)', '4,550 – 8,000', '4,500 – 5,800', '+22%'],
  ['Associate (MSA: 25–59 months)', '3,650 – 4,150', '3,600 – 4,100', '+1%'],
  ['Junior Associate (MSA: 0–24 months)', '2,750 – 3,350', '2,300 – 3,100', '+13%'],
  ['Paralegal', 'n/a', '1,500', 'n/a'],
], [3400, 2200, 1900, 1526]));
c.push(P(''));
c.push(P('Partner mid-points (MSA SEK 8,000; Cederquist SEK 7,250) are ~2.6–2.7x junior mid-points and ~4.6–5.0x the 2025 legal-aid norm. Comparison is indicative until the MSA card date is confirmed.'));
c.push(H2('Market commentary on rate levels'));
[
 ['Partners at top firms bill SEK 7,000–9,000/h, supported by 3–5 associates billing ~SEK 3,000/h who "cost at most a third" of what they bill (Olle Flygt, ex-Vinge and Mannheimer Swartling partner).', '2026', `${RT}, 31 Mar 2026`, 'https://www.realtid.se/perfectweekend/affarsjuridikens-karna-det-har-blivit-for-mycket-fokus-pa-pengar/'],
 ['Flygt billed SEK 7,000–8,000/h in his last years at Mannheimer Swartling.', 'final years before 2025', `${RT}, 26 Apr 2026`, 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/'],
 ['Delphi (Agnes Hammarstrand): AI efficiency pushes senior rates up; favours fixed monthly subscriptions.', '2026', `${RT}, 15 Apr 2026`, 'https://www.realtid.se/juridik/timarvoden-for-jurister-en-modell-pa-vag-ut/'],
 ['Third-party directory (skatterätt.se, tax practices; no level split, low reliability): MSA 4,500–8,000; Roschier 4,200–7,000; Vinge, DLA Piper 4,000–7,000; Schjødt, Cederquist 3,800–6,500; Setterwalls, Bird & Bird 3,500–6,000; Delphi 3,200–5,500; Lindahl 3,000–5,500.', 'accessed Oct 2026, undated', 'skatterätt.se', 'https://www.xn--skattertt-12a.se/byra/vinge-skatt/'],
 ['Consumer guide range for business-law advice: SEK 1,500–5,000/h; fixed prices for simple agreements SEK 3,000–15,000 (low reliability).', '2026', 'Offerta, 25 Sep 2026', 'https://offerta.se/guider/ekonomi-och-juridik/vad-kostar-en-affarsjurist'],
].forEach(x => c.push(B(...x)));

c.push(H1('2. Billed rates: invoices, court filings and Chapter 11 cases'));
c.push(table(['Firm', 'Period', 'Level', 'SEK/h', 'Context'], [
  ['Mannheimer Swartling', 'Nov 2010', 'Partner / Associate / Junior', '3,150 / 2,150 / 1,550', 'SL framework; blended 2,118, 1,906 after 10% discount'],
  ['Mannheimer Swartling', 'Oct 2009', 'Partner (chairman)', '3,200', 'Försvarsmakten, Projekt PRIO'],
  ['Mannheimer Swartling', 'Jan 2010', 'Partner-led team', '3,800', 'Riksgälden, Carnegie valuation'],
  ['DLA Piper Sweden', '2012–2017', 'Counsel team', '3,400 – 3,900', 'HQ AB litigation cost claim'],
  ['Roschier', '2012–2017', 'Implied blended', '~3,360', 'HQ AB (SEK 17.5m / 5,207 h)'],
  ['Mannheimer Swartling', 'Nov 2024 – Mar 2025', 'Team average', '~4,000', 'Northvolt Ch. 11; ~6,000 h, SEK 32m; lowest of all advisers'],
  ['Kirkland & Ellis', 'Nov – Dec 2024', 'Average / top', '14,300 / 25,500', 'Northvolt Ch. 11 (USD converted at 11)'],
  ['Front Advokater', '2016 – 2020', 'Junior / Associate / Senior', '953 / 1,270 / 1,588', 'City of Gothenburg framework (tender)'],
  ['7 unnamed firms', '2007', 'Tender price', '1,800 – 2,800', 'State divestment adviser tender'],
  ['Unnamed advokat', 'Dec 2024', 'Advokat', '4,500', 'Botkyrka municipality'],
  ['Unnamed firm', 'Apr 2026', 'Court-approved', '4,000', 'Luleå tingsrätt: 4,000/h reasonable'],
], [1900, 1400, 1700, 1500, 2526]));
c.push(P(''));
c.push(P([new TextRun({ text: 'Sources: Realtid 3 Feb 2011, 2009 (Försvarsmakten), 2010 (Riksgälden), 2017 (HQ), 2007 (state tender), 25 Apr 2026 (Luleå); Affärsvärlden 4 Feb and 2 May 2025 (Northvolt); Göteborgs tingsrätt T 20148-21, 21 Feb 2024 (Front); SVT Dec 2024 (Botkyrka). Full links in the workbook, sheet "D. Rate cards".', font: FONT, size: 18, italics: true, color: '595959' })]));

c.push(H1('3. Rate increases, inflation and client pushback'));
[
 ['Affärsvärlden market report: market growth 8.4% with operating margin 13.5% while lawyer headcount grew only 3.7%, i.e. firms "better at charging".', '2024', 'Affärsvärlden, 5 Oct 2024', 'https://www.affarsvarlden.se/artikel/rekordar-vantas-2024-byraerna-har-blivit-annu-battre-pa-att-ta-betalt'],
 ['US firms raised rates 7.3% on average; average top-100 US lawyer now above USD 1,000/h (Thomson Reuters/Georgetown).', '2025', `${RT}, 3 Oct 2026`, 'https://www.realtid.se/juridik/kunderna-vill-ha-ai-rabatt-advokatbyraerna-hojer-priset/'],
 ['LexisNexis CounselLink: partner rates +5.1%; M&A partner rates +8.8%.', '2025', 'LexisNexis, 22 Apr 2026', 'https://www.lexisnexis.com/community/amp-pressroom/lexisnexis-counsellink-releases-2026-trends-report-as-rising-rates-and-big-law-share-of-wallet-continue-to-climb'],
 ['Juro survey (~130 GCs): 79% say AI should lower fees; 84% have seen no reduction; 72% think firms keep the AI savings.', '2026', `${RT}, 3 Oct 2026`, 'https://www.realtid.se/juridik/kunderna-vill-ha-ai-rabatt-advokatbyraerna-hojer-priset/'],
 ['Goldman Sachs, Morgan Stanley and Citi demand lower bills due to AI; explicit "AI discounts" appearing in panel reviews. No equivalent public demand by Swedish banks found.', '2026', `${RT}, 1 Sep 2026`, 'https://www.realtid.se/juridik/storbankerna-kraver-ai-rabatt-av-advokatbyraerna/'],
 ['Handelsbanken (Sweden) shares Norwegian banks\' concern over fee levels and is reviewing alternative pricing models.', '2021', `${RT}, 1 Dec 2021`, 'https://www.realtid.se/bors-finans/nu-satter-handelsbanken-press-pa-hoga-advokatarvoden/'],
 ['Regi client surveys: value for money among the lowest-rated criteria since 2008; price the most common reason for switching firm.', '2008–2017', `${RT}, 15 Dec 2017`, 'https://www.realtid.se/bors-finans/advokatbyraernas-prismodeller-forandras/'],
].forEach(x => c.push(B(...x)));
c.push(P('No Swedish firm was found publicly announcing a specific percentage increase for 1 January 2025 or 2026; Cederquist states rates are revised annually on 1 January.'));

c.push(H1('4. Billing models and rules'));
[
 ['Fees must be reasonable (skäligt) under the Bar\'s code (rule 4.1.1); the assessment may consider the agreement, scope, difficulty, importance, skill and result (4.1.2). Outcome-based elements rarely matter in practice. Exact wording of success-fee restrictions not verified.', '2019', `${RT} op-ed (P. Danowsky), 29 Mar 2019`, 'https://www.realtid.se/debatt/rimligt-att-ersattningen-innehaller-ett-moment-av-reglering'],
 ['Baker McKenzie: "strategic decision to largely abandon the billable-hour model".', '2026', `${RT}, 15 Apr 2026`, 'https://www.realtid.se/juridik/timarvoden-for-jurister-en-modell-pa-vag-ut/'],
 ['Norwegian banks: DNB required at least 50% of mandates on fixed fees in its previous framework; Handelsbanken Norway wants fixed fees.', '2021', `${RT}, 25 Nov 2021`, 'https://www.realtid.se/bors-finans/norska-protester-mot-dyra-advokater-nu-kraver-bankerna-fast-pris/'],
 ['Thomson Reuters proposes fixed fees, outcome-based pricing, subscriptions and gain-sharing; 90% of US legal spend is still hourly.', '2026', `${RT}, 3 Oct 2026`, 'https://www.realtid.se/juridik/kunderna-vill-ha-ai-rabatt-advokatbyraerna-hojer-priset/'],
 ['Setterwalls: international clients require AI use and evidence of improvement; Roschier: AI has not reduced workload.', '2025', `${RT}, 3 Oct 2025`, 'https://www.realtid.se/bors-finans/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/'],
 ['47% of corporate legal departments have generative AI access (Thomson Reuters 2026 State of the Corporate Law Department).', '2026', `${RT}, 26 Apr 2026`, 'https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/'],
].forEach(x => c.push(B(...x)));

c.push(H1('5. Benchmarks and public-sector pricing'));
c.push(table(['Benchmark', 'Period', 'Rate', 'Note'], [
  ['Legal-aid norm (timkostnadsnorm), F-skatt', '2025', 'SEK 1,586', 'SFS 2024:913'],
  ['Legal-aid norm (timkostnadsnorm), F-skatt', '2026', 'SEK 1,626', 'SFS 2025:919; +2.5%'],
  ['Legal-aid norm, not F-skatt', '2026', 'SEK 1,237', '2025: SEK 1,207'],
  ['Denmark, defence counsel rate', '2026', 'DKK 2,055', '"almost double" Sweden (Advokatsamfundet)'],
  ['Denmark, Kammeradvokaten state agreement', 'from 2024', 'Partner DKK 2,790', 'Advokat >3y 2,263; <3y 1,891 (snippet)'],
  ['Norway, legal-aid rate (salærsats)', '2026', 'NOK 1,375', '2025: NOK 1,315'],
], [3300, 1300, 1700, 2726]));
c.push(P(''));
[
 ['The legal-aid norm assumes 72.5% billable time on a 1,656-hour year (~1,200 billable hours). SOU 2025:111 proposes 80% of norm for non-advokat associates; the Bar opposes.', 'Nov 2025 – Mar 2026', 'SOU 2025:111; Advokatsamfundet, 27 Mar 2026', 'https://data.riksdagen.se/dokument/HDB3111'],
 ['Public payments to advokat firms SEK 4.7bn, of which courts (legal aid) SEK 3.8bn; "real" public market ~SEK 900m. Largest suppliers: Mannheimer Swartling SEK 115.7m, Kahn Pedersen 96.2m, Lindahl 39.6m, Setterwalls 24.7m, Hannes Snellman 20m, Vinge 19.4m.', '2020', `${RT}, 10 Feb 2022`, 'https://www.realtid.se/juridik/mannheimer-swartling-toppar-statens-advokatnota/'],
 ['No current central state (Kammarkollegiet) or Adda framework for general legal services found; agencies appear to procure individually, and price annexes are typically confidential (Kammarrätten i Stockholm, Vinge / E-hälsomyndigheten).', '2017–2026', `${RT}, 30 Aug 2017`, 'https://www.realtid.se/bors-finans/vinges-timpriser-sekretessbelagda/'],
].forEach(x => c.push(B(...x)));

c.push(H1('6. Nordic and international context'));
[
 ['Norway: commercial litigation rates that were ~NOK 3,000/h have "doubled in a short time" (Government Attorney).', '2026', `${RT}, 5 Jun 2026`, 'https://www.realtid.se/juridik/regeringsadvokaten-slar-larm-advokatkostnaderna-skenar/'],
 ['UK: top-10 firms average GBP 449/h (2019: GBP 321); firms 11–25 average GBP 325 (PwC; snippet only).', '2024', 'Legal Cheek, Apr 2025', 'https://www.legalcheek.com/2025/04/the-2000-per-hour-solicitor'],
 ['US: elite partners up to USD 3,000–3,400/h; top-lawyer rates at top-50 firms up 16% in a year (WSJ).', '2025–2026', `${RT}, 19 Feb 2026`, 'https://www.realtid.se/juridik/toppjuristers-arvoden-exploderar-samtidigt-hotar-ai-botten/'],
].forEach(x => c.push(B(...x)));

c.push(H1('7. Revenue per lawyer and implied rates'));
[
 ['Affärsvärlden panel fee income per lawyer: SEK 4.22m (2023) and SEK 4.13m (2024). White & Case Stockholm highest fee per lawyer (2023).', '2023–2024', 'Affärsvärlden 29 Apr 2024; Realtid 3 Oct 2025', 'https://www.affarsvarlden.se/artikel/affarsjuristerna-om-trenden-pe-bolagen-kommer-vara-motorn'],
 ['Illustrative only: SEK 4.13m per lawyer implies a blended realised rate of ~SEK 2,580/h at 1,600 billable hours or ~SEK 2,950/h at 1,400 hours, well below list rates, consistent with leverage, write-offs and discounts.', '2024', 'Calculation', ''],
].forEach(x => c.push(B(...x)));

c.push(H1('8. Data gaps and next steps'));
[
 ['No published rate cards found for Vinge, White & Case, Setterwalls, Lindahl, Gernandt & Danielsson, CMS Wistrand, Cirio, MAQS, Baker McKenzie, Snellman and other top-20 firms.', '', '', ''],
 ['US Chapter 11 monthly fee statements (PACER) for SAS (S.D.N.Y. 22-10925) and Northvolt (S.D. Tex. 24-90577) list per-timekeeper rates for Mannheimer Swartling, 2022–2025.', '', '', ''],
 ['Client invoices (e.g. via Axo, requested) and FOI requests to public buyers (e.g. Riksbanken 2024 legal services framework, Göteborgs Stad) could add firm-level data, though price annexes are often withheld.', '', '', ''],
 ['Confirm the date of the Mannheimer Swartling rate card.', '', '', ''],
].forEach(x => c.push(B(...x)));

const doc = new Document({
  creator: 'Research', title: 'Charging rates in Swedish business law',
  styles: { default: { document: { run: { font: FONT, size: 21 } } } },
  numbering: { config: [{ reference: 'b', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 360, hanging: 260 } } } }] }] },
  sections: [{ properties: { page: { margin: { top: 1200, bottom: 1200, left: 1300, right: 1300 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.RIGHT, children: [new TextRun({ children: ['Page ', PageNumber.CURRENT], font: FONT, size: 16, color: '808080' })] })] }) },
    children: c }],
});
Packer.toBuffer(doc).then(b => fs.writeFileSync(process.argv[2], b));
