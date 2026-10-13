# Swedish corporate legal market: research notes

The main deliverable is `research/swedish_legal_market.xlsx`. This file summarises it. Data was compiled on 2026-10-08.

## Section A. Market size (sheets "A. Market size", "A. Growth drivers", "A. Niche firms")
- **A1. Headline estimates:** ~SEK 16bn in 2025 fee income for the 56 largest business-law firms (Affärsvärlden). An industry estimate puts the business-law market above SEK 25bn. All advokatbyråer turned over ~SEK 29bn in 2024 (estimated from SCB).
- **A2. Official statistics (SCB Företagens ekonomi):** SNI 69.1 legal activities had SEK 38.9bn net turnover in 2024. Advokatbyråer (SNI 69.101) had SEK 25.9bn in 2021, the last year SCB published that split.
- **Firm benchmarking:** the former "A. Firms" tab has been merged into "C. Benchmarking" (top 30).
- **A3. Growth drivers:** each firm's growth is tagged as lateral hire, M&A, new office or organic, with a source for each.

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
- **D1:** rate cards by level. Mannheimer Swartling is undated (supplied by a colleague, date to be confirmed); Cederquist is 2025; MAQS rates come from the sell-side model. Includes mid-points, the partner/junior multiple and the multiple of the legal-aid hourly norm.
- **D2:** indicative firm-wide hourly ranges from a third-party directory (low reliability).
- **D3:** Nordic business-law data from 2023 onwards only: realised rates (Mannheimer Swartling on Northvolt, a court ruling) and market commentary.
- **D4:** Nordic business-law benchmarks: fee income per lawyer and Norwegian litigation rates.
- **Rebuilding it:** run `python3 research/scripts/build_section_d.py research/swedish_legal_market.xlsx`.

## Market data briefing (Word)
- **File:** `research/swedish_legal_charging_rates.docx`, a dated and sourced briefing on charging rates in Nordic business law: rate levels, pricing trends, AI impact on the business model and pricing, and international (US/UK) leading indicators with whether each has reached the Nordics.
- **Rebuilding it:** run `node research/scripts/build_rates_doc.js research/swedish_legal_charging_rates.docx`.

## Section E. Partner compensation (sheet "E. Partner compensation")
- **Basis:** profit pool per partner, calculated from annual reports filed with Bolagsverket (read via allabolag.se). Profit before tax plus disclosed board/CEO salaries, divided by partners, gives a proxy for average equity-partner pay before tax. Also shows dividend per partner and an illustrative senior-partner figure (1.4x the average, an assumption). Raw data: `research/data/allabolag_firm_entity_accounts.json`.
- **Context:** SCB salary distribution for salaried lawyers (SSYK 2611/2614), 2025.
- **Rebuilding it:** run `python3 research/scripts/build_section_e.py research/swedish_legal_market.xlsx`.

## Section C. Benchmarking (sheet "C. Benchmarking")
- **Coverage:** the top 30 firms by Affärsvärlden FY2024 fee income, plus Baker McKenzie and Linklaters, which are missing from the revised table. Shows revenue per partner, per employee and per lawyer, two leverage measures (partners / total employees and partners / lawyers), employees per partner, and medians across the 30.
- **Rebuilding it:** run `python3 research/scripts/build_section_c.py research/swedish_legal_market.xlsx`.

## Section F. M&A & team lifts (sheet "F. M&A & team lifts")
- **F1:** mergers, acquisitions, PE roll-ups (AGRD), market entries and splits, 2020 to Oct 2026.
- **F2:** team lifts of 2+ partners or 5+ lawyers.
- **F3:** notable single-partner moves.
- **F4:** net partner flow per firm.
- **F5:** key takeaways.
- **Rebuilding it:** run `python3 research/scripts/build_section_f.py research/swedish_legal_market.xlsx`.

## Section A4. Niche firms (sheet "A. Niche firms")
- **Coverage:** every Swedish firm ranked by Chambers or Legal 500: 9 full-service firms, ranked independents, PE-owned (AGRD) firms and Legal 500 firms to watch. Includes Affärsvärlden FY2024 revenue and revenue per partner where available.

## Section B3. Top lawyers (sheet "B. Top lawyers")
- **Contents:** individually ranked lawyers in Corporate / M&A, Private equity, Banking & finance and Dispute resolution. Includes Chambers Europe 2026 Bands 1-3 and Legal 500's Hall of Fame, Leading partners and Next Generation Partners, with a summary count per firm.
- **Rebuilding it:** run `python3 research/scripts/build_lawyers.py research/swedish_legal_market.xlsx`.
