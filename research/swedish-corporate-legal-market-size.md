# Swedish corporate legal market: size sources

Compiled 2026-10-08 from web search results. The full reports and paywalled articles were not opened, so the figures below come from news summaries and search snippets. Check them against the primary source before citing.

## Headline estimates

| Scope | Figure | Year | Source |
|---|---|---|---|
| Corporate law: fee income of the 56 largest firms | ~SEK 16bn (+5.3% YoY) | 2025 | Affärsvärlden annual branschrapport (2026 ed.) |
| Same panel, prior year | SEK 15.4bn (+8.6%) | 2024 | Affärsvärlden / Realtid |
| Same panel, earlier | >SEK 14bn | ~2022 | Affärsvärlden |
| Corporate law market, total (industry estimate, method not given) | >SEK 25bn/yr | n/a | AGRD Partners op-ed, Affärsvärlden |
| All legal services (MarketLine, including taxes) | USD 3.1bn (CAGR 5.9% 2019–24) | 2024 | MarketLine "Legal Services in Sweden" |
| All legal activities (NACE 69.1) | EUR 3.5bn (CAGR −0.5% 2020–25) | 2026e | IBISWorld |
| Advokat sector, total turnover | SEK 18bn (~4bn export) | 2016 | Tidningen Advokaten (2018) |
| Smaller firms: 735 firms filing digital annual reports | SEK 4.1bn combined | 2025 | Företagsbevakning |

## Official statistics: SCB Företagens ekonomi (pulled 2026-10-08)

Net turnover (nettoomsättning) is in SEK m and covers all companies classified to each industry code. Raw JSON is in `research/data/`.

| Year | 69.1 Legal activities, total | 69.101 Advokatbyråer | 69.102 Other legal offices | 69.103 Patent agents |
|---|---|---|---|---|
| 2010 | 18,738 | 13,904 | 1,956 | 2,879 |
| 2015 | 23,986 | 17,936 | 2,750 | 3,300 |
| 2019 | 31,237 | 22,277 | 4,194 | 4,765 |
| 2020 | 32,628 | 23,298 | 4,371 | 4,958 |
| 2021 | 34,693 | 25,864 | 4,558 | 4,270 |
| 2022* | 33,276 | n/p | n/p | n/p |
| 2023* | 36,525 | n/p | n/p | n/p |
| 2024* | 38,921 | n/p | n/p | n/p |

\* SCB started a new series in 2022 (table NV0109P), so 2021 and 2022 are not directly comparable. In the new tables, SCB shows the sub-industry figures as "..", meaning they are not published.

Other figures:
- **Company count, 69.1:** 6,965 companies with 13,446 employees in 2024. In 2021, 69.101 alone had 3,836 companies and 10,242 employees.
- **69.101 growth:** advokatbyrå turnover grew about 5.8% a year from 2010 to 2021.
- **Activity-unit data, 69.1, 2024:** production value SEK 39.1bn, value added SEK 25.3bn, EBITDA margin 28.2%.
- **69.101 estimate for 2024:** about SEK 29bn. This applies the 2021 share (74.5% of 69.1) to the 2024 total of SEK 38.9bn, so it is an estimate, not a published figure.

How this relates to the press figures: SCB's ~SEK 29bn for all advokatbyråer in 2024 sits above Affärsvärlden's ~SEK 16bn for the 56 largest business-law firms in 2025. That puts the large business-law firms at roughly 50–55% of total advokatbyrå turnover. The rest goes mainly to criminal, family and small-business practices.

## Bolagsverket and company filings
- **Bolagsverket API:** the free API for high-value datasets (värdefulla datamängder), at gw.api.bolagsverket.se, needs OAuth credentials from registration, so filings could not be pulled programmatically. The firms' own annual reports can be ordered through Bolagsverket's e-service.
- **Företagsbevakning:** its data, built from digital annual reports, covers 766 firms with SEK 4.4bn total turnover, and the largest of them has only SEK 104m. The big business-law firms are missing because they file on paper or as partnerships, or operate through multi-entity structures. Register data therefore understates the top of the market.
- **Dagens Juridik 2025 revenue tables:** these are based on the firms' year-end accounts but sit behind a paywall (Plus subscription).

## Concentration and firm data
- In 2025, five firms had revenue above SEK 1bn. Vinge and Mannheimer Swartling each had more than SEK 2bn, and Setterwalls passed SEK 1bn (Dagens Juridik, Affärsvärlden).
- 2022 fee income from Affärsvärlden's 2023 survey: MSA SEK 1,738m, Vinge SEK 1,609m, Setterwalls SEK 803m, followed by Roschier, White & Case and Lindahl. Lindahl reported SEK 637m for 2022.
- Fastest growers in 2025: Schjødt (~+25%) and Cirio (~+15%). IPO work was the weakest segment. Firms expect +6.8% growth in 2026.

## Supply side (Advokatsamfundet)
- 31 Dec 2025: 6,712 members (5,963 active advokater) and 2,525 associates (biträdande jurister).
- In 2018 there were 1,959 advokatbyråer, of which 1,249 were one-person firms.
- Nordamicus (Feb 2024) counted 134 large firms, employing 5,288 advokater and 2,996 associates.

## Caveats
- "Corporate legal market" has no official definition. The Affärsvärlden panel (~SEK 16bn) is the most-cited figure, but it leaves out smaller firms and in-house counsel.
- The MarketLine and IBISWorld figures cover all legal services, and the two sources disagree on the growth trend.

## Sources
- SCB API: https://api.scb.se/OV0104/v1/doris/sv/ssd/NV/NV0109/ (tables BasfaktaFEngs07, NSEBasfaktaFEngs07, NSEBasfaktaVEngs07N)
- https://www.affarsvarlden.se/artikel/affarsjuridiken-vaxer-igen-men-jattarna-tappar-mark
- https://www.realtid.se/juridik/affarsjuridiken-vaxer-men-sprickorna-syns/
- https://www.realtid.se/affarer/154-miljarder-sa-klarade-affarsjuristerna-tullchocken/
- https://www.affarsvarlden.se/artikel/stor-kartlaggning-av-affarsjuristerna-har-ar-byraerna-som-drar-in-mest
- https://www.affarsvarlden.se/debatt/okad-konkurrens-starker-svensk-affarsjuridik
- https://www.dagensjuridik.se/nyheter/sa-mycket-vaxer-de-storsta-advokatbyraerna-i-sverige/
- https://www.dagensjuridik.se/nyheter/advokatjattarnas-2025-vinge-backar-msa-satter-rekord/
- https://www.advokaten.se/tidigare-nummer/2018/nr-2-2018-argang-84/att-halla-tva-tankar-i-huvudet-samtidigt/
- https://store.marketline.com/report/legal-services-in-sweden/
- https://www.ibisworld.com/sweden/industry/legal-activities/200283/
- https://www.advokatsamfundet.se/Pressrum/fakta-om-advokatsamfundet2/
- https://www.advokatsamfundet.se/globalassets/advokatsamfundet_sv/advokatsamfundet/advokatsamfundets_verksamhetsberattelse_2018_digital_vb2.pdf
- https://nordamicus.se/2024/02/19/sveriges-storsta-advokatbyraer-2024/
- https://foretagsbevakning.se/branscher/advokatbyraer
- https://lindahl.se/en/latest-news/news/2023/continued-growth-and-success-for-lindahl
- https://snisok.scb.se/69101

## Section B: Chambers & Legal 500 practice rankings
- **Where it is:** `research/swedish_legal_market.xlsx`, sheet "B. Practice rankings". It is a firm × practice-area matrix showing the Chambers Europe 2026 band and the current Legal 500 EMEA tier for each firm. MAQS is boxed on its own, and Magnusson, Delphi and Lindahl are boxed as one group. Sheet "B. Rankings raw" has the full 650+ row extract.
- **Rebuilding it:** run `python3 research/scripts/build_rankings.py research/swedish_legal_market.xlsx`. The script reads `research/data/chambers_europe_2026_sweden_firm_bands.json` and `research/data/legal500_emea_sweden_firm_tiers.json`.
