# Verification brief for European screen agents (shared)

Read /home/user/Database/research/europe-screen/00_brief.md first. Then, for your market:

STEP 1 - Candidate generation. Start from the seed names in your task. Add any other listed company in your market that
plausibly sits in a Corsair vertical with market cap roughly EUR 80m to 900m (wide net; the EV test comes later). Scan
at least one exchange or index sector list (for example the exchange's financials / technology / business-services
sector pages, or a data site's country list) and say which you scanned. Explicitly note names you know that are gone
(taken private, acquired, delisted) in 2024-2026, with acquirer and date, because they are precedents.

STEP 2 - Verification, for each candidate (aim for 10-18 names, prioritising the in-band ones):
1. listing_status: exchange/segment; still listed; any live offer period.
2. business: 1-2 sentences, what it does and for whom.
3. corsair_vertical: mapping, or "weak fit" with reason.
4. market_cap_eur_m with source and as-of date.
5. Latest full-year: fy_end, revenue_eur_m, adj_ebitda_eur_m (or operating profit), net_cash_or_debt_eur_m. State
   local currency and the rate used.
6. ownership: holders above ~10%; founder/family/strategic blocks; free float.
7. corporate_activity: bid approaches, strategic reviews, sale processes, activists, take-private offers 2024-2026.
8. flags: regulatory (bank/insurer licences, supervisory change-of-control), leverage, concentration, litigation,
   squeeze-out thresholds relevant to a take-private in that jurisdiction.
9. sources: URLs actually consulted. 10. confidence: high/medium/low with reason.

Rules: prefer company reports, exchange pages and regulator filings. Write null when not found; never estimate silently.
Budget: the web-search budget is shared across seven agents; use AT MOST 25 WebSearch calls, and prefer WebFetch of
company investor pages, exchange pages and data sites (stockanalysis.com, marketscreener.com) for figures.
Write a JSON array to the output file named in your task, then reply with a 10-line summary: genuine fits, names gone
from the market since 2024, what you scanned, and anything that challenges the assignment.
