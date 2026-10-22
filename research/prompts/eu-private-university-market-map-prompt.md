# Claude Code prompt: European private university operator market map

Copy everything below the line into Claude Code. Fill in the bracketed fields first.

---

You are a senior investment analyst supporting a PE associate at Corsair Capital. Build a complete, sourced market map of privately owned higher-education operators in Europe and assess which are realistically investable for a private equity fund.

## Our parameters
- Target EV range: [e.g. EUR 150m to 1bn]
- Deal type: [control / significant minority / either]
- Equity cheque: [e.g. EUR 75m to 300m]
- Hold period and return hurdle: [e.g. 5 years, 2.5x / 20%+ IRR]
- Geographic priority: [e.g. UK, DACH, Iberia, France, Italy, Benelux, Nordics, Poland]
- Hard exclusions: [e.g. K-12 schools, pure edtech software, student housing, test prep]

## Scope
In scope: degree-granting or accredited private higher-education providers in Europe (EU27, UK, Switzerland, Norway) that are owned for profit or could be, including campus universities, universities of applied sciences, business schools, online and hybrid universities, specialist schools (health, hospitality, design, engineering, law), pathway and foundation providers attached to a degree brand, and multi-brand platforms that roll these up. Include listed companies with European higher-ed assets.

Out of scope: state universities, church or foundation universities with no realistic route to private ownership (note them only where they are a competitor or acquisition target), K-12, language schools, corporate training, and pure software.

## Deliverables
Write all outputs to ./output/.

1. `operators.xlsx` - one row per operator, with the fields listed below. Target 60 to 100 operators. Add a second sheet for individual brands under each platform.
2. `precedents.xlsx` - every PE, strategic, or IPO transaction in European private higher education from 2015 onward you can verify: date, target, buyer, seller, EV, revenue and EBITDA multiples where disclosed, source.
3. `country_matrix.xlsx` - one row per country: accreditation regime, who grants degree-awarding powers, whether for-profit ownership of universities is permitted, limits on tuition, state funding or student loan access for private providers, international student visa rules and recent changes, recent regulatory shocks, and an overall attractiveness rating.
4. `market_map.md` - a segmentation of the market along two axes (delivery model and geographic footprint) with operators placed, plus a ranked shortlist of the 10 to 15 most investable targets and why.
5. `memo.md` - a 4 to 6 page investment memo: market size and growth by country, demand drivers, consolidation history, who the active sponsors are, what returns prior sponsors made, the key risks, and a recommended sourcing angle for a fund with our parameters.
6. `sources.md` - every source used, with URL and the date accessed.

## Fields per operator
Name, HQ country, countries of operation, brand list, ownership type (PE-backed / founder or family / listed / strategic subsidiary / foundation), current owner and entry year, prior owners, estimated revenue and EBITDA with year and source, student count, mix of domestic vs international students, mix of online vs campus, mix of undergraduate vs postgraduate vs executive, accreditation status and bodies, tuition price points, known acquisitions, known financing (debt packages, rating agency reports), management, latest known valuation or transaction, likely sale trigger (fund life, founder age, listed and underperforming, carve-out), and an investability score.

## Investability score
Score each operator 1 to 5 on: size fit with our EV range, ownership route to a deal within 3 years, regulatory durability of its licence in its home market, revenue quality (recurring enrolment, retention, pricing power), growth runway (online, international, bolt-ons), and reputational risk (student outcomes, regulator findings, press). Weight the dimensions and explain the weights. Flag every score that rests on an estimate rather than a disclosed figure.

## Research process
- Start by anchoring on known platforms and their sponsors, then work outward via each platform's acquisition history, competitor mentions, and accreditation registers. Examples to verify, not to assume: Galileo Global Education, Global University Systems, IU Group, Universidad Europea, Universidad Alfonso X el Sabio, Proeduca / UNIR, Multiversity, Omnes Education, AD Education, Sommet Education, SRH, Hochschule Fresenius, BPP, Arden, University of Law, Nebrija, UCAM, NCOI, Vistula, Studium Group, Ilerna, Planeta Formacion.
- Use national accreditation and quality-agency registers (for example OfS register in the UK, Wissenschaftsrat and HRK in Germany, ANECA and regional registers in Spain, HCERES and CEFDG in France, ANVUR and MUR in Italy, NVAO in the Netherlands, PKA in Poland) to find operators the deal press has missed.
- Use company registries and filed accounts (Companies House, Bundesanzeiger, Registro Mercantil, Infogreffe, Kamer van Koophandel) for revenue and EBITDA where available. Prefer filed accounts over press estimates.
- Use sponsor websites, bond prospectuses, rating agency reports, and trade press (PE Hub, Unquote, Mergermarket summaries, Times Higher Education, local business press) for transactions.
- Do not fabricate a number. If a figure is unavailable, write "n/a" and note what source would have it. Mark every estimate as an estimate with the basis for it.
- Cite a source for every figure in every file.

## Working style
- Before researching, ask me up to five clarifying questions on the parameters above, then proceed with reasonable assumptions for anything unanswered and state them at the top of memo.md.
- Work country by country and save after each country so partial progress survives.
- At the end, give me a short summary in chat: how many operators mapped, the top 10 shortlist with one line each, the three biggest open data gaps, and which countries you were least confident on.
