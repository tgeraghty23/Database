"""Build 'F. M&A & team lifts' sheet: consolidation and lateral team moves in Swedish business law, 2020 - Oct 2026."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

path = sys.argv[1]
wb = load_workbook(path)
if 'F. M&A & team lifts' in wb.sheetnames: del wb['F. M&A & team lifts']
ws = wb.create_sheet('F. M&A & team lifts')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
N = Font(italic=True, size=9); S = Font(bold=True, size=11)
FILL = {'PE roll-up': 'FFF2CC', 'Market entry': 'DDEBF7', 'Split / exit': 'FCE4D6'}

def table(r, headers, rows, typecol=None):
    for c, h in enumerate(headers, 1):
        x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); x.border = B
    for row in rows:
        r += 1
        for c, v in enumerate(row, 1):
            x = ws.cell(r, c, v); x.border = B; x.alignment = W
            if typecol is not None:
                for k, col in FILL.items():
                    if str(row[typecol]).startswith(k): x.fill = PatternFill('solid', fgColor=col)
    return r + 2

ws['A1'] = 'Section F. Consolidation and team lifts: Swedish business law, Jan 2020 - Oct 2026'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Team lift = 2+ partners or 5+ lawyers moving together from one firm within ~12 months. Dates from article metadata (search-index dates are unreliable). '
            'Sources: Realtid, Dagens Juridik (paywalled: headline/lead only), Affärsvärlden, Advokatbladet, Legal 500 2026 firm pages, firm press releases. '
            '"1-src" = single source. Coverage of 2020-2023 is thinner (Legal 500 reflects ~2024-26; international legal press paywalled).')
ws['A2'].alignment = W; ws.merge_cells('A2:J2'); ws.row_dimensions[2].height = 45

r = 4
ws.cell(r, 1, 'F1. Mergers, acquisitions, PE roll-ups, market entries and splits').font = S
M = [
 ['Oct-2019 / Jan-2020', 'Market entry (merger)', 'Schjødt (Norway)', 'Hamilton Advokatbyrå', '', '~65 lawyers', 'Full service, Stockholm', 'Hamilton 2018 revenue SEK 215m; Norwegian-Swedish one-stop shop', 'realtid.se/bors-finans/hamilton-gar-ihop-med-norska-byran-schjodt/'],
 ['Jan-2020', 'Merger', 'Time Danowsky', 'Time Advokatbyrå + Danowsky & Partners', '', '27 staff combined', 'Tech, media', '1-src', 'realtid.se/bors-finans/advokatbyran-time-danowsky-bildas-efter-sammanslagning/'],
 ['Nov-2021 / Jan-2022', 'Split / exit', 'Born Advokater (Sthlm); Norma Advokater (Gbg)', 'Glimstedt Stockholm and Göteborg offices', '', '', 'Banking & finance, real estate, M&A', 'Offices left Glimstedt brand; Born later a founding AGRD firm. 1-src', 'realtid.se/bors-finans/glimstedt-renodlar-verksamheten/'],
 ['Nov-2023 / Mar-2024', 'Market entry (membership)', 'CMS', 'Wistrand (now CMS Wistrand)', '', '~150 lawyers', 'Stockholm, Göteborg; full service', 'Joins 5,800-lawyer network; cross-border deal growth', 'dagensjuridik.se/nyheter/wistrand-blir-en-del-av-cms-gar-fran-150-jurister-till-5-800-varlden-over/'],
 ['Feb-2024', 'Merger', 'Carthiel', 'Grey Advokatbyrå', '', '~10 staff combined', 'Disputes added', 'Carthiel grew +155% (FY2024). Firm release only (1-src)', 'carthiel.se/post/advokatfirman-carthiel-går-ihop-med-grey-advokatbyrå'],
 ['Feb / Mar-2024', 'Split / exit', 'Hannes Snellman Sweden (now "Snellman")', 'Hannes Snellman group (Finland)', '', '>80 lawyers (Sthlm)', 'Stockholm', 'Helsinki and Stockholm strategies diverged', 'dagensjuridik.se/nyheter/snellmans-nya-mp-stora-strategiska-fragor-for-branschen/'],
 ['Mar-2024', 'Split / exit', '(bankruptcy)', 'Front Advokater', '', '', 'Göteborg', 'Bankrupt after losing case vs City of Göteborg. 1-src', 'dagensjuridik.se/nyheter/efter-forlusten-i-domstol-front-advokater-i-konkurs/'],
 ['Aug-2025', 'PE roll-up (formation)', 'AGRD Partners (Axcel)', 'Allié, Born, Morris Law, Next, Synch, TM & Partners', '', '~250 staff, 8 offices; ~100-130 advokater left the Bar', 'Business law', 'Revenue ~SEK 660m (top-10). Firms keep names. CEO Maria-Pia Hope (ex-Vinge)', 'realtid.se/juridik/massavhopp-fran-advokatsamfundet-nar-ny-jattebyra-bildas/'],
 ['Oct-2025 / Jan-2026', 'Market entry (team lift)', 'BAHR (Norway)', 'Schjødt (Oslo and Stockholm)', '8 (Sweden); 23 group-wide', '~20 of ~30 Stockholm opening staff; 54 people group-wide', 'M&A, banking, capital markets', 'Targets 100+ lawyers in Stockholm; ~50 Stockholm staff by Feb-2026', 'dagensjuridik.se/nyheter/schjodt-skakas-av-avhopp-atta-i-stockholm-lamnar/'],
 ['Jan-2026', 'PE roll-up (add-on)', 'AGRD Partners', 'Wigge & Partners; Bokwall Rislund', '', '', 'Transactions/PE; competition & procurement', 'First add-ons; both left the Bar', 'dagensjuridik.se/nyheter/agrd-vaxer-tva-nya-byraer-gar-med/'],
 ['May-2026', 'PE roll-up (add-on)', 'AGRD Partners', 'BBA//Fjeldco (Iceland)', '', '', 'Iceland', 'First non-Swedish investment (context)', 'dagensjuridik.se/nyheter/agrd-expanderar-till-island/'],
 ['May-2026', 'PE roll-up (add-on)', 'AGRD Partners', 'Moll Wendén (Malmö)', '8', '~47 staff', 'Malmö, 15 practice areas', 'First deal in southern Sweden; Öresund strategy', 'dagensjuridik.se/nyheter/moll-wenden-gar-in-i-agrd/'],
 ['Aug-2026 / Jan-2027', 'Split / exit', 'MAQS KB (Malmö, keeps brand) / Keelion Law (Sthlm, Gbg, Sundsvall)', 'MAQS', '', '', 'Malmö vs northern offices', 'Amicable; Malmö plans AI-forward premium firm', 'realtid.se/juridik/advokatbyra-delas-malmo-bryter-sig-loss/'],
 ['Aug-2026', 'PE roll-up (add-on)', 'Synch (AGRD)', 'Qap Legal Advisors', '', 'All Qap staff', 'Stockholm boutique', 'Synch "open to more acquisitions". 1-src', 'dagensjuridik.se/nyheter/synch-och-qap-legal-gar-ihop/'],
 ['Aug-2026', 'Merger', 'MAQS Malmö', 'Fredersen Advokatbyrå', '', '', 'Malmö', 'Ahead of MAQS separation. 1-src', 'dagensjuridik.se/nyheter/fredersen-gar-ihop-med-maqs-i-malmo/'],
 ['Sep-2026', 'PE roll-up (add-on)', 'AGRD Partners (as Ascente Law)', 'Setterwalls Malmö office', '12', '~100 people', 'Malmö; corporate, energy, patent litigation', 'AGRD\'s largest deal; office leaves the Bar (~SEK 305m revenue in 2025)', 'realtid.se/juridik/setterwalls-malmokontor-gar-till-riskkapitalagda-agrd/'],
 ['Sep / Oct-2026', 'Merger', 'Magnusson', 'Gro Advokatbyrå', '', '<10 staff, ~SEK 20m revenue', 'VC, tech, data protection; Stockholm', 'Positioned as a commitment to the advokat model (vs AGRD)', 'realtid.se/juridik/byraer-gar-samman-fortsatter-satsa-pa-advokatmodellen/'],
 ['Oct-2026 / Jan-2027', 'Merger', 'Walthon Ramberg', 'Walthon Advokater + Ramberg Advokater', '', '~40 lawyers', 'Transactions, disputes, insolvency; Stockholm', '1-src outlet', 'dagensjuridik.se/nyheter/walthon-och-ramberg-gar-ihop-samlar-40-jurister/'],
]
r = table(r + 1, ['Date (announced / effective)', 'Type', 'Acquirer / surviving firm', 'Target / source', 'Partners', 'People', 'Practice / office', 'Notes', 'Source'], M, typecol=1)
ws.cell(r - 1, 1, 'Colour: yellow = PE roll-up; blue = market entry; orange = split / exit. Context: Schjødt merged with Norrbom Vinding (Denmark, Jan-2026); Advokatsamfundet reviewing its ownership rules after AGRD.').font = N

r += 1
ws.cell(r, 1, 'F2. Team lifts (2+ partners or 5+ lawyers)').font = S
T = [
 ['Feb-2020', 'Lindahl (new Stockholm office)', 'Westermark Anjou', 2, 2, 'Capital markets / public M&A', 'Monica Lagercrantz, Dennis Westermark', 'Medium (firm release)', 'dagensjuridik.se/pressrum/lindahl/advokatfirman-lindahl-vaxer-i-stockholm-tva-nya-delagare-och-tva-jurister-2/'],
 ['Jan-Mar-2020', 'Fylgia', 'Setterwalls', '1+', 3, 'Insolvency / restructuring', 'Mathias Winge (+3, titles unstated)', 'Medium (1-src; borderline)', 'realtid.se/bors-finans/fylgia-tar-over-obestandsavdelningen-fran-setterwalls/'],
 ['Jan-2024', 'MAQS', 'Delphi', 3, '', 'M&A', 'Anders Lindström, Per Lagerkvist, Fredrik Mörner', 'Medium (1-src)', 'realtid.se/juridik/tre-nya-delagare-inom-ma-ansluter-till-maqs/'],
 ['Apr-2024', 'RE:FI STHLM (new boutique)', 'Cederquist', '1+ (4 partners at new firm)', '3', 'Banking / finance / restructuring', 'Johan Lundberg (ex-MP)', 'Medium', 'dagensjuridik.se/nyheter/finansieringsexpert-pa-cederquist-lanserar-egen-specialistbyra/'],
 ['Jun-2024', 'Gernandt & Danielsson', 'Linklaters Stockholm', 2, 5, 'Real estate transactions', 'Magnus Lidman, Anna Eriksson', 'High', 'mynewsdesk.com/se/gernandtdanielsson'],
 ['Jan-2025', 'Cirio', 'AG Advokat', 3, 9, 'Real estate / construction disputes', 'Johan Lindberg, Johan Rosén, Måns Derk', 'High', 'realtid.se/juridik/cirio-expanderar-inom-fastighets-och-entreprenadjuridik/'],
 ['Oct-2025 / Feb-2026', 'BAHR (new Stockholm office)', 'Schjødt Sweden', 8, '~12', 'M&A, banking, capital markets', 'Ola Hansson, Eric Halvarsson, Emil Hedberg, Peter Svanqvist, Kristoffer Stråth, Frida Lantoft, Karl Klackenberg, Tom Wehtje', 'High (count); medium (names)', 'dagensjuridik.se/nyheter/schjodt-skakas-av-avhopp-atta-i-stockholm-lamnar/'],
 ['Sep-2026', 'Ascente Law (AGRD)', 'Setterwalls Malmö', 12, '~88', 'Full service', 'Marc Tullgren (MP); others not named', 'High (event)', 'realtid.se/juridik/setterwalls-malmokontor-gar-till-riskkapitalagda-agrd/'],
]
r = table(r + 1, ['Date', 'Hiring firm', 'Source firm', 'Partners', 'Other lawyers / staff', 'Practice', 'Partner names', 'Confidence', 'Source'], T)
ws.cell(r - 1, 1, 'Below threshold: AG Advokat -> MAQS (Elin Schedin + 1, Jan-2025); Lindahl -> Foyen (Jan Berg + 3, Dec-2025).').font = N

r += 1
ws.cell(r, 1, 'F3. Notable single partner moves at top firms').font = S
P = [
 ['Jun-2020', 'Hannes Snellman', 'Roschier', 'Ola Åhman', 'ECM / public M&A'], ['Apr-2021', 'Bird & Bird', 'MAQS', 'Olof Wetterling', 'Finance'],
 ['Mar-2023', 'Skierfe', 'Cederquist', 'Christian Johard', 'Real estate'], ['May-2023', 'Lindahl', 'Bird & Bird', 'Jan Byström', 'M&A'],
 ['Aug-2024', 'Schjødt', 'Mannheimer Swartling', 'Tom Wehtje', 'M&A (later to BAHR)'], ['Sep-2024', 'Linklaters', 'White & Case', 'Magnus Wennerhorn', 'Banking / finance'],
 ['Dec-2024', 'DLA Piper', 'Snellman', 'Joel Montin', 'Financial regulation'], ['Dec-2024', 'Hammarskiöld', 'Snellman', 'Ola Åhman', 'Corporate / M&A'],
 ['Mar-2025', 'DLA Piper', 'Hammarskiöld', 'Carl-Johan Pousette', 'Capital markets'], ['Mar-2025', 'Bird & Bird', 'Lindahl', 'Henrik Nobel', 'M&A'],
 ['Mar-2025', 'Bird & Bird', 'Westerberg & Partners', 'Wendela Hårdemark', 'IP / patent litigation'], ['May-2025', 'Gernandt & Danielsson', 'Snellman', 'Peter Forsberg', 'Competition'],
 ['Jun-2025 / Feb-2026', 'Schjødt then BAHR', 'Roschier then Schjødt', 'Peder Boström', 'Employment'], ['Oct-2025', 'BAHR', 'Snellman', 'Johanna Wärnberg', 'M&A'],
 ['Nov-2025', 'MAQS', 'Baker McKenzie', 'Markus Mueller', 'Banking / finance'], ['Jan-2026', 'Lindahl', 'Vinge', 'Henrik Schön', 'Financial services'],
 ['2026', 'MAQS', 'Delphi', 'Karin Roberts', 'Defence / security'], ['Apr-2026', 'Schjødt', '(ex-Linklaters)', 'Patrik Björklund', 'Banking / finance'],
 ['Jul-2026', 'DLA Piper', 'Baker McKenzie', 'Olof Forssell', 'Trade / sanctions'], ['2026', 'Gernandt & Danielsson', 'Snellman', 'Pontus Ewerlöf', 'Disputes'],
]
r = table(r + 1, ['Date', 'Hiring firm', 'Source firm', 'Partner', 'Practice'], P)

ws.cell(r, 1, 'F4. Net partner flow 2020 - Oct 2026 (named partners in F2-F3 plus Legal 500 departures)').font = S
NF = [['BAHR', 10, 0], ['MAQS', 6, 1], ['Lindahl', 4, 2], ['Gernandt & Danielsson', 4, 2], ['Cirio', 3, 1], ['Bird & Bird', 3, 2], ['DLA Piper', 3, 2],
      ['Hammarskiöld', 1, 1], ['Linklaters', 1, 2], ['Mannheimer Swartling', 0, 2], ['Baker McKenzie', 0, 2], ['Roschier', 0, 3], ['Cederquist', 0, 3],
      ['Snellman', 1, 5], ['Delphi', 0, 4], ['AG Advokat', 0, 4], ['Schjødt', 2, 9], ['Setterwalls', 0, 13]]
hr = r + 1
for c, h in enumerate(['Firm', 'Partner gains', 'Partner losses', 'Net'], 1):
    x = ws.cell(hr, c, h); x.font = H; x.fill = HF; x.border = B; x.alignment = C
for i, (f, g, l) in enumerate(NF, hr + 1):
    for c, v in enumerate([f, g, l], 1):
        x = ws.cell(i, c, v); x.border = B; x.alignment = C if c > 1 else W
    x = ws.cell(i, 4, f'=B{i}-C{i}'); x.border = B; x.alignment = C; x.font = Font(bold=True)
r = hr + len(NF) + 2
ws.cell(r, 1, 'Schjødt gains: Wehtje, Boström, Björklund (counted in losses again when they left). Setterwalls losses include 12 Malmö partners to Ascente/AGRD. Snellman: Montin partner status unconfirmed.').font = N

r += 2
ws.cell(r, 1, 'F5. Key takeaways').font = S
for t in [
    'Consolidation accelerated from 2025: AGRD (Axcel) has assembled 11 firms in ~13 months (incl. Setterwalls Malmö, Wigge, Moll Wendén) and says more acquisitions are planned. AGRD firms must leave the Bar (no external ownership of advokat firms).',
    'Nordic firms enter Sweden via combinations and lifts: Schjødt (Hamilton, 2020), CMS (Wistrand, 2024), BAHR (8 Schjødt partners, 2026).',
    'Team lifts concentrate in real estate (G&D from Linklaters; Cirio from AG Advokat) and M&A/finance (MAQS from Delphi; BAHR from Schjødt).',
    'Net losers of partners: Setterwalls (Malmö exit), Schjødt, Snellman, Delphi and AG Advokat. The Big 3 (MSA, Vinge, Roschier) show no team losses.',
    'Splits are rising: Hannes Snellman (2024), MAQS (2026/27) and Glimstedt\'s Stockholm/Göteborg offices (2021).',
]:
    r += 1; ws.cell(r, 1, '- ' + t).alignment = Alignment(wrap_text=False)
r += 2
ws.cell(r, 1, 'Not verified / gaps: Hannes Snellman Sweden rename to "Snellman" (no formal notice found); no qualifying team moves found out of MSA, Vinge or Roschier; 2020-23 coverage thinner. All URLs prefixed https://www.').font = N
for c, w in zip(range(1, 10), (18, 22, 26, 28, 14, 22, 26, 40, 50)):
    ws.column_dimensions[get_column_letter(c)].width = w
order = ['A. Market size', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw',
         'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('ok', wb.sheetnames)
