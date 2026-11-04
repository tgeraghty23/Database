# Verification brief for AIM screen candidates (shared by all research agents)

Governing objective: identify AIM-quoted companies that could be investable for Corsair Capital's buyouts platform
(financial-services specialist PE; verticals: Payments [integrated, international, B2B]; Software [vertical & enterprise,
risk & compliance, HR & finance tools]; Business services [wealth & asset management, insurance services, professional
services, healthcare services, trust & corporate services]; enterprise value USD 150m-750m, larger with co-invest;
platform, add-on or minority; control or real influence; modest leverage). Date: 4 November 2026.

For EACH company assigned to you, find and return:
1. listing_status: still on AIM / moved to Main Market / delisted or taken over (with date and acquirer).
2. business: 1-2 sentences, what it actually does and for whom.
3. corsair_vertical: which Corsair vertical/sub-sector it maps to, or "weak fit" with reason.
4. market_cap_gbp_m with source and as-of date (use LSE page, company site, or a data site; state which).
5. Latest full-year financials (state FY end): revenue_gbp_m, adj_ebitda_gbp_m (or PBT if EBITDA not disclosed), net_cash_or_debt_gbp_m.
6. ownership: holders above ~10%, and whether founders/management/a strategic hold a blocking stake; free float if stated.
7. corporate_activity: any bid approach, strategic review, formal sale process, activist, or take-private offer 2024-2026 (dates, bidder, outcome, premium).
8. flags: regulatory (FCA/PRA permissions, Lloyd's), balance sheet (bank/lender), litigation, concentration, anything that would stop a PE buyer.
9. sources: list of URLs actually consulted.
10. confidence: high / medium / low, with the main reason.

Rules: prefer primary sources (RNS on londonstockexchange.com, investegate.co.uk, company annual reports) over aggregators.
If a number cannot be found, write null and say so; never estimate silently. Do not pad with generic text.
Write your return as a JSON array (one object per company, keys exactly as the field names above) to the output file
named in your task, then reply with a 10-line summary: which names look like genuine fits, which are gone from AIM,
and anything that challenges the assignment.
