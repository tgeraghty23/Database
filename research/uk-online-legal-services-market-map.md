# UK online / subscription legal services: market map and benchmarking vs LegalVision

The main deliverable is `research/uk_online_legal_services_market_map.xlsx`. This file summarises it. Data was compiled on 2026-10-12 from Companies House filings (iXBRL and scanned PDFs), company websites, trade press (Legal Futures, Law Gazette, RollOnFriday, Capital Brief) and deal announcements. Raw records: `research/data/uk_online_legal_*.json`.

## Scope and segmentation
Businesses are grouped into three segments and tagged by relevance to LegalVision (Core / Adjacent / Defunct or exited / Out of scope):

- **1. SME subscription legal (22 records):** the closest analogues. Lawyer-delivered subscription or membership law firms (Harper James, LegalVision UK, Lawyerly, Blackmont, Law 365, 360 Business Law, LegalEdge), legal-plan memberships (Rocket Lawyer UK), fixed-fee online firms (Sprintlaw UK), self-serve platforms (SeedLegals, Genie AI, Docue, Sparqa, Farillio, Pocketlaw) and the marketplace model (Lawhive). Enterprise CLM vendors (Summize, Juro, Legl) are listed as out of scope.
- **2. HR and employment-law subscriptions (21 records):** the PE-proven adjacent segment. Insurance-backed, multi-year, per-employee-priced subscriptions for SMEs: Peninsula group (Peninsula, Croner, Croner-i, BrightHR, Health Assured), Citation (KKR/Hg), WorkNest and Wirehouse (Axiom GRC, Inflexion), Avensure, Quest, NatWest Mentor (being sold to Empowering People Group), Markel Law / FSB Legal Hub, Moorepay, ARAG, HR Dept, Stallard Kane, Lawgistics.
- **3. Consumer and documents (33 records):** digital wills/probate (Farewill, Octopus Legacy, Kinherit, Make a Will Online, Beyond), online divorce and prenups (amicable, Divorce-Online, Wenup), advice subscriptions (Which? Legal, JustAnswer), scaled fixed-fee ABSs (Co-op Legal Services, Slater and Gordon, ARAG Law), digital conveyancing (Juno, Muve) and document / contract platforms (Simply-Docs, Net Lawman, Lawpack, Legalo, LawDepot, Wonder.Legal, Genie AI, Docue, Contractbook, Summize, Juro, Avvoka, Lexoo).

## Headline findings
- **Few scaled direct analogues.** Harper James is the only UK subscription law firm at scale: £15.3m revenue FY Sep-25 (+25%), £2.2m operating profit, 131 average employees, ~70% of income from subscription plans, >1,000 Enable subscribers, LDC minority stake (exchanged Jun 2025, completed Sep 2025 after SRA approval). LegalVision UK filed £3.69m turnover for FY Jun-24 (restated £3.35m) and 41 average employees in FY25; FY25 turnover was not filed. Management's A$13m UK revenue claim (Oct 2024) does not reconcile with the filings.
- **High attrition in the segment.** LawBite (administration Sep 2024, platform sold for £1, ~£7m raised), Linkilaw (exited SMEs), Sparqa (dormant), Pocketlaw (rebranded Miramis, pivoted to enterprise CLM), Rocket Lawyer UK (8 staff, £5m net liabilities), Legal Utopia (liquidation Jul 2026), Beyond (dissolved 2024), Lexoo (absorbed by Kalexius). Self-serve document models have not scaled in the UK; lawyer-delivered subscriptions have.
- **The PE money is in employment-law/HR subscriptions.** Peninsula £463m revenue and £118m EBITDA (FY Mar-25, family-owned); Citation £275m (FY24, KKR/Hg/HarbourVest, bought at ~15x EBITDA in 2020); Marlowe GRC (WorkNest, VinciWorks) carved out to Inflexion at £430m EV, 16.2x adjusted EBITDA (2024); Avensure £19.2m revenue at a 30% EBITDA margin (founder-owned); WorkNest £30.8m; NatWest Mentor (c.220 staff, c.100k SMEs) agreed sale to Limerston-backed Empowering People Group (Mar 2026).
- **Consumer digital legal is consolidating.** Farewill sold to Dignity at £12.9m EV (~1.4x annualised revenue, vs a £70m 2020 valuation); amicable (Nov 2023) and Octopus Legacy (Nov 2022) are Octopus Group companies; Wenup to Blackfinch (Jun 2026); Muve to Alcuin (Mar 2026); Simply-Docs changed control (Jul 2026). Lawhive is the growth outlier: >$35m run-rate, 7x year-on-year, $60m Series B (Feb 2026), >$100m raised.
- **Demand is large and under-served.** LSB small-business survey: 32% of small businesses have a legal issue each year (1.8m firms, 15.4m issues) but only 25% use professional help, 44-46% handle it alone, and 10% see lawyers as cost-effective. UK legal market £55.5bn in 2025 (IRN), B2B about 51%.
- **Valuation anchors.** Listed UK law firms: Keystone 1.6x EV/revenue and 14x EV/EBITDA; Knights 1.3x and ~5-6x; Gateley 0.8x and 10x. LegalZoom ~1.1x revenue, ~5x adjusted EBITDA, 65% subscription. Compliance subscriptions: Citation ~15x, Marlowe GRC 16x EBITDA. Harper James/LDC, LegalVision/Quadrant and Lawhive terms undisclosed.

## Workbook structure
- **Summary:** how to read, key takeaways, segment overview.
- **1-3. Company sheets:** one row per company with entity and Companies House number, HQ, founding, ownership, funding and deals, revenue (latest, period, prior year, growth), EBITDA where filed, revenue basis, average employees, revenue per employee, pricing model and detail, services, practice areas, regulatory status, customers and reviews, tech angle, valuation notes and sources. Rows are colour-coded by relevance and sorted by relevance then revenue.
- **4. LegalVision benchmark:** group model, pricing in UK/AU/NZ, revenue claims, headcount, Quadrant PE investment (Mar 2025; Quadrant vehicle became ultimate parent of the UK entity May 2025), SRA ABS status, UK entity turnover and headcount trajectory, and the Australian peers that also entered the UK (Sprintlaw) or did not (Lawpath).
- **5. Market context:** 28 metrics with sources.
- **6. Deals & PE activity:** 28 deals from the context research plus 13 transactions surfaced in the company research.
- **7. Listed comps:** Keystone, Knights, Gateley, RBG, Frenkel Topping, DWF, LegalZoom, Rocket Lawyer.
- **Benchmark table:** 27 core comparables side by side.
- **Suggested KPIs:** 18 additional data points worth collecting for a PE market map, why they matter and where to source them.
- **Caveats & gaps.**

## Rebuild
- `python3 research/scripts/build_uk_online_legal.py research/uk_online_legal_services_market_map.xlsx` rebuilds the workbook from `research/data/uk_online_legal_*.json`.
- `python3 research/scripts/build_uk_online_legal_html.py <out.html>` builds the one-page summary from the same data.

## Key caveats
- **Revenue:** most small UK companies file abridged or micro accounts without turnover; the "Revenue basis" column says whether a figure is filed, company-reported or estimated. Where turnover is missing, the notes give net assets, cash and headcount from the balance sheet.
- **Employees, not FTEs:** Companies House reports average monthly headcount. Consultant and fee-share lawyers (Lawhive ~500, 360 Business Law 750+) are not employees.
- **Citation EBITDA:** reported £43.4m FY24 is after £33.5m one-off transaction costs and £4.6m deferred remuneration; underlying c.£77-82m is derived.
- **Corrections:** no Markel/LawBite deal exists; Docue did not acquire Sparqa or Farillio; Setfords is Phoenix-backed; Lawfront is Blixt-backed.
- **Unverified rows:** some deal rows (Robin AI, vLex/Clio, Luminance) and third-party revenue estimates are flagged as unverified.
- **Not researched:** Definely, Robin AI, Luminance; gunnercooke, Taylor Rose; Howden LEI products; SeedLegals 2026 round.
