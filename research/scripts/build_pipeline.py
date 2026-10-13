"""Build 'B. Partner pipeline' sheet: ranked partners in Corporate/M&A, Private equity, Banking & finance and Dispute resolution,
scored for likelihood to move and fit, as a lateral-hire target list. Inputs: research/data/*.json (directory rankings, profiles)."""
import json, re, sys, unicodedata
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
sys.argv, path = [sys.argv[0], '/dev/null'], sys.argv[1]
exec(open('research/scripts/build_rankings.py').read().split('AREAS = [')[0])  # CANON / canon()
D = 'research/data/'
lawyers = json.load(open(D + 'top_lawyers_chambers2026_legal500_sweden.json'))
seniority = json.load(open(D + 'chambers_lawyer_seniority.json'))
profiles = json.load(open(D + 'legal500_lawyer_profiles.json'))

def slug(s):
    s = unicodedata.normalize('NFKD', s).encode('ascii', 'ignore').decode().lower()
    return re.sub(r'[^a-z]+', '-', s).strip('-')
prof_by = {re.sub(r'^\d*-', '', k): v for k, v in profiles.items()}

EXCLUDE_LEVELS = {'Associates to watch', 'Leading associates'}
STAGE = {'Band 1': -1, 'Hall of Fame': -1, 'Eminent Practitioners': -2, 'Senior Statespeople': -2, 'Band 2': 0, 'Leading partners': 0,
         'Band 3': 1, 'Band 4': 1, 'Band 5': 1, 'Next Generation Partners': 1, 'Up and Coming': 1}
LEVEL_ORDER = ['Eminent Practitioners', 'Senior Statespeople', 'Hall of Fame', 'Band 1', 'Band 2', 'Leading partners', 'Band 3',
               'Next Generation Partners', 'Band 4', 'Band 5', 'Up and Coming']
# Firm context: (pay level, pressure score, signal) - from Sections C, E and F of this workbook
F = {
 'Mannheimer Swartling': ('High', 0, 'Lockstep: profit pool ~SEK 8.7m/partner; 2 partner exits since 2024'),
 'Vinge': ('High', 0, 'Profit pool ~SEK 7m/partner; individual-contribution model'),
 'Setterwalls': ('High', 2, 'Malmö office (12 partners) left for AGRD, Sep-2026'),
 'Roschier': ('High', 1, '3 partner exits 2020-26'),
 'White & Case': ('High', 0, 'Global LLP; ~SEK 69m revenue/partner'),
 'Gernandt & Danielsson': ('High', 0, 'Net partner gainer (team lifts 2024-26)'),
 'Linklaters': ('High', 2, 'No new partners for 3 years; Stockholm staff down from 88 to 75'),
 'Kilpatrick Townsend & Stockton Advokat KB': ('Medium', 0, ''),
 'Cederquist': ('Medium', 2, '3 partner exits 2023-25 incl. ex-MP (RE:FI STHLM spin-out)'),
 'Cirio': ('Medium', 0, 'Growing (+15% 2025); net partner gainer'),
 'Snellman': ('Medium', 2, 'Separated from Hannes Snellman (2024); ~5 partner exits since 2024 incl. heads of competition and disputes'),
 'Baker McKenzie': ('Medium', 2, '2 partner exits 2025-26; moving away from billable hour'),
 'DLA Piper': ('Medium', 0, 'Revenue -9% FY24/25; net partner gainer'),
 'CMS Wistrand': ('Medium', 0, 'Joined CMS network 2024'),
 'Hammarskiöld': ('Medium', 0, ''),
 'Schjødt': ('Medium', 2, 'Lost 8 Stockholm partners to BAHR (2025-26)'),
 'BAHR': ('Medium', -2, 'New Stockholm office (2026); partners just moved'),
 'Bird & Bird': ('Medium', 0, ''),
 'Delphi': ('Lower', 2, 'Lost 4 partners to MAQS 2024-26; profit pool ~SEK 1.7m/partner'),
 'Lindahl': ('Lower', 0, ''), 'MAQS': ('Lower', 0, 'Target firm'),
 'AG Advokat': ('Lower', 2, 'Lost 4 partners Jan-2025 (Cirio, MAQS)'),
 'Magnusson': ('Lower', 0, 'Merged with Gro (Oct-2026)'),
 'Foyen': ('Lower', 0, ''), 'Westerberg & Partners': ('Lower', 0, ''), 'Kastell': ('Lower', 0, ''), 'Fylgia': ('Lower', 0, ''),
 'Kahn Pedersen': ('Lower', 0, ''), 'Norburg & Scherp': ('Lower', 0, ''), 'Gulliksson': ('Lower', 0, ''), 'Hellström': ('Lower', 0, ''),
 'RE:FI STHLM': ('Lower', 0, 'Boutique founded 2024 (ex-Cederquist)'), 'Andulf': ('Lower', 0, ''), 'Sandart & Partners': ('Lower', 0, ''),
 'Kanter': ('Lower', 0, ''), 'Nordia Law': ('Lower', 0, ''), 'Norma': ('Lower', 0, ''),
}
AGRD = {'Morris Law', 'Synch', 'TM & Partners', 'Born', 'Next Law', 'Allié', 'Wigge & Partners', 'Bokwall Rislund', 'Moll Wendén'}
for f in AGRD:
    F[f] = ('Lower', 1, 'AGRD (PE-owned) member: lawyers left the Bar; partners wishing to remain advokater may move')
# Known moves 2020-26 (Section F): recent (2025-26) = just moved; earlier = has moved before
RECENT = {'Peter Forsberg', 'Peder Boström', 'Johanna Wärnberg', 'Markus Mueller', 'Henrik Schön', 'Karin Roberts', 'Patrik Björklund',
          'Olof Forssell', 'Pontus Ewerlöf', 'Carl-Johan Pousette', 'Henrik Nobel', 'Wendela Hårdemark', 'Johan Lindberg', 'Johan Rosén',
          'Måns Derk', 'Ola Hansson', 'Eric Halvarsson', 'Emil Hedberg', 'Peter Svanqvist', 'Kristoffer Stråth', 'Frida Lantoft',
          'Karl Klackenberg', 'Tom Wehtje', 'Marc Tullgren'}
EARLIER = {'Ola Åhman', 'Magnus Wennerhorn', 'Joel Montin', 'Jan Byström', 'Johan Lundberg', 'Magnus Lidman', 'Anna Eriksson',
           'Christian Johard', 'Olof Wetterling', 'Anders Lindström', 'Per Lagerkvist', 'Fredrik Mörner', 'Monica Lagercrantz', 'Dennis Westermark'}

people = defaultdict(lambda: {'areas': defaultdict(list), 'firms': set()})
for src, area, level, name, firm in lawyers:
    if level in EXCLUDE_LEVELS: continue
    p = people[name]; p['areas'][area].append(f"{'CH' if src == 'Chambers' else 'L500'}: {level}"); p['firms'].add(canon(firm)); p.setdefault('levels', []).append(level)

rows = []
for name, p in people.items():
    firm = sorted(p['firms'])[0]
    if firm == 'MAQS': continue
    pay, pressure, signal = F.get(firm, ('Lower' if firm not in ('Vinge',) else 'High', 0, ''))
    best = min(p['levels'], key=lambda l: LEVEL_ORDER.index(l) if l in LEVEL_ORDER else 99)
    stage = STAGE.get(best, 0)
    yrs = seniority.get(name, {}).get('years'); head = seniority.get(name, {}).get('head', False)
    sen = (1 if yrs is not None and yrs <= 6 else -1 if yrs is not None and yrs >= 15 else 0)
    afford = {'High': -1, 'Medium': 0, 'Lower': 1}[pay]
    move = -1 if name in RECENT else (1 if name in EARLIER else 0)
    score = pressure + stage + sen + afford + move
    prio = 'A' if score >= 3 else 'B' if score >= 1 else 'C'
    pr = prof_by.get(slug(name), {})
    angles = []
    if pressure >= 2: angles.append('Firm under pressure: ' + signal)
    elif pressure == 1 and signal: angles.append(signal)
    if stage >= 1: angles.append('Mid-career / next-generation: room to lead and build a practice')
    if stage <= -1: angles.append('Senior market leader: hard to move; possible senior adviser / team anchor')
    if pay == 'High': angles.append('Likely pay cut vs current profit share: needs equity / upside story')
    if name in EARLIER: angles.append('Has moved firm before (2020-24)')
    if name in RECENT: angles.append('Moved firm in 2024-26: unlikely to move again soon')
    if len(p['areas']) >= 2: angles.append('Ranked in ' + ' & '.join(sorted(p['areas'])))
    if pressure == -2: angles.append(signal)
    rows.append([prio, score, name, firm, '; '.join(sorted(p['areas'])),
                 ' | '.join(f"{a}: {', '.join(v)}" for a, v in sorted(p['areas'].items())),
                 pr.get('position') or ('Department head' if head else ''), yrs, pay, signal, '; '.join(angles)])
rows.sort(key=lambda r: ('ABC'.index(r[0]), -r[1], r[3], r[2]))

wb = load_workbook(path)
if 'B. Partner pipeline' in wb.sheetnames: del wb['B. Partner pipeline']
ws = wb.create_sheet('B. Partner pipeline')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
PF = {'A': 'C6E0B4', 'B': 'FFF2CC', 'C': 'F2F2F2'}
ws['A1'] = 'B4. Partner pipeline: ranked lawyers in Corporate / M&A, Private equity, Banking & finance and Dispute resolution'; ws['A1'].font = Font(bold=True, size=14)
nA = sum(r[0] == 'A' for r in rows); nB = sum(r[0] == 'B' for r in rows); nC = sum(r[0] == 'C' for r in rows)
ws['A2'] = (f'Long list: {len(rows)} individually ranked lawyers (Chambers Europe 2026 all bands; Legal 500 Hall of Fame, Leading partners, Next Generation Partners) '
            f'in the four practice areas where MAQS has no ranked lawyers. Priority A = {nA}, B = {nB}, C = {nC}. '
            'Score = firm pressure (-2 to +2: partner losses, restructuring, PE ownership; recently joined = -2) + career stage (+1 Band 3-5 / Next Gen / Up and Coming; '
            '0 Band 2 / Leading partner; -1 Band 1 / Hall of Fame; -2 Eminent / Senior Statespeople) + seniority (+1 if ranked <=6 years, -1 if >=15) '
            '+ affordability vs MAQS economics (+1 lower-paying firm, -1 high-paying firm) + mover history (+1 moved 2020-24, -1 moved 2024-26). '
            'A >= 3; B = 1-2; C <= 0. Scores are indicative screening aids, not assessments of individuals.')
ws['A2'].alignment = W; ws.merge_cells('A2:K2'); ws.row_dimensions[2].height = 90
hdr = ['Priority', 'Score', 'Lawyer', 'Current firm (per directory)', 'Practice area(s)', 'Rankings', 'Position (Legal 500 profile)', 'Chambers years ranked',
       'Current firm pay level', 'Firm signal', 'Approach angle / notes']
for c, h in enumerate(hdr, 1):
    x = ws.cell(4, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); x.border = B
ws.row_dimensions[4].height = 36
r = 4
for row in rows:
    r += 1
    for c, v in enumerate(row, 1):
        x = ws.cell(r, c, v if v != '' else None); x.border = B; x.alignment = C if c in (1, 2, 8, 9) else W
        x.fill = PatternFill('solid', fgColor=PF[row[0]])
ws.auto_filter.ref = f'A4:K{r}'
r += 2
for t in [
    'Firm pay level = profit pool / revenue per partner from Sections C and E (High: MSA, Vinge, Setterwalls, Roschier, White & Case, G&D, Linklaters; Lower: Delphi, Lindahl, AGRD firms, smaller independents). MAQS profit pool ~SEK 2m/partner, so High-pay partners would need an equity/upside case.',
    'Firm signals from Section F (2020-26 moves). Firm shown is as listed by the directory and may be out of date; verify current firm, title and office (Stockholm vs Göteborg/Malmö: MAQS Malmö vs Keelion Law post-split) before any approach.',
    'Next steps: (1) agree target office(s) and practice priorities; (2) validate A/B names via firm websites and LinkedIn; (3) check non-compete / notice terms (Swedish partner agreements often 6-12 months); (4) prioritise team lifts (partner + associates) over individuals.',
    'Source data: research/data/top_lawyers_chambers2026_legal500_sweden.json; chambers_lawyer_seniority.json; legal500_lawyer_profiles.json (extracted 13 Oct 2026).',
]:
    ws.cell(r, 1, t).font = Font(italic=True, size=9); r += 1
for c, w in zip(range(1, 12), (8, 7, 22, 24, 24, 42, 18, 10, 11, 40, 60)):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'D5'
order = ['A. Market size', 'A. Nordic market', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'B. Top lawyers', 'B. Partner pipeline', 'C. Benchmarking', 'D. Rate cards', 'E. Partner compensation', 'F. M&A & team lifts', 'G. MAQS key data']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print(len(rows), nA, nB, nC)
for row in rows[:25]: print(row[0], row[1], row[2], '|', row[3], '|', row[4])
missing = sorted({r[3] for r in rows if r[3] not in F})
print('firms without context:', missing)
