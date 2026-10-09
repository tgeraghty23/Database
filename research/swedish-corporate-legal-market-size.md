# Swedish corporate legal market: research notes

The main deliverable is `research/swedish_legal_market.xlsx`. This file summarises it. Data was compiled on 2026-10-08.

## Section A. Market size and firms (sheets "A. Market size", "A. Firms", "A. Growth drivers")
- **A1. Headline estimates:** ~SEK 16bn in 2025 fee income for the 56 largest business-law firms (Affärsvärlden). An industry estimate puts the business-law market above SEK 25bn. All advokatbyråer turned over ~SEK 29bn in 2024 (estimated from SCB).
- **A2. Official statistics (SCB Företagens ekonomi):** SNI 69.1 legal activities had SEK 38.9bn net turnover in 2024. Advokatbyråer (SNI 69.101) had SEK 25.9bn in 2021, the last year SCB published that split.
- **A3. Concentration:** 25 firms ranked by latest revenue, with partners, lawyers and employees (used as an FTE proxy). Revenue per partner, per employee and per lawyer is calculated wherever both inputs exist.
- **A4. Growth drivers:** each firm's growth is tagged as lateral hire, M&A, new office or organic, with a source for each.

## Section B. Chambers & Legal 500 practice rankings (sheets "B. Practice rankings", "B. Rankings raw")
- **Matrix:** firm × practice area, showing the Chambers Europe 2026 band and the current Legal 500 EMEA tier. MAQS is boxed on its own, and Magnusson, Delphi and Lindahl are boxed as one group.

## Rebuild
- `python3 research/scripts/build_rankings.py research/swedish_legal_market.xlsx` builds Section B from `research/data/*.json`.
- `python3 research/scripts/build_section_a.py research/swedish_legal_market.xlsx` builds Section A.

## Key caveats
- **Revenue bases:** Affärsvärlden reports fee income, while company accounts report the net sales of a single legal entity. Some firms report on split financial years.
- **Partner counts:** these are thin and partly firm-supplied to Chambers. Vinge's count covers only Stockholm and Malmö partners. White & Case's count is implied from Dagens Juridik.
- **Headcount:** no FTE data is published, so employee headcount is used instead. Company-account headcounts usually exclude partners who bill through their own companies.
- **Paywalls:** Affärsvärlden's full FY2025 table and Dagens Juridik's 2025 tables are paywalled. Those would fill the remaining FY2025 gaps.

## Section D. Rate card comparison (sheet "D. Rate cards")
- **D1:** list-price rate cards by level. Mannheimer Swartling is undated (supplied by a colleague, date to be confirmed); Cederquist is 2025. Includes mid-points, the partner/junior multiple and the multiple of the legal-aid hourly norm.
- **D2:** indicative firm-wide hourly ranges from a third-party directory (low reliability).
- **D3:** dated rates from invoices, court cost claims, Chapter 11 filings, public tenders and market commentary.
- **D4:** benchmarks, including the legal-aid hourly norm (timkostnadsnorm), fee income per lawyer, and Danish, Norwegian and UK rates.
- **Rebuilding it:** run `python3 research/scripts/build_section_d.py research/swedish_legal_market.xlsx`.

## Market data briefing (Word)
- **File:** `research/swedish_legal_charging_rates.docx`, a dated and sourced briefing on charging rates in Swedish business law.
- **Rebuilding it:** run `node research/scripts/build_rates_doc.js research/swedish_legal_charging_rates.docx`.
