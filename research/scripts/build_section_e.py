"""Build 'E. Partner compensation' sheet from company accounts (Bolagsverket filings via allabolag.se) and SCB wage statistics."""
import sys
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

path = sys.argv[1]
wb = load_workbook(path)
if 'E. Partner compensation' in wb.sheetnames: del wb['E. Partner compensation']
ws = wb.create_sheet('E. Partner compensation')
H = Font(bold=True, color='FFFFFF'); HF = PatternFill('solid', fgColor='1F3864')
thin = Side(style='thin', color='BFBFBF'); B = Border(left=thin, right=thin, top=thin, bottom=thin)
W = Alignment(wrap_text=True, vertical='top'); C = Alignment(horizontal='center', vertical='top', wrap_text=True)
N = Font(italic=True, size=9); grey = PatternFill('solid', fgColor='EDEDED'); amber = PatternFill('solid', fgColor='FFF2CC')

ws['A1'] = 'Section E. Partner compensation: profit pool per partner from filed annual reports'; ws['A1'].font = Font(bold=True, size=14)
ws['A2'] = ('Source: annual reports filed with Bolagsverket, read via allabolag.se (Oct-2026); amounts SEK m. Swedish law firms pay partners mainly through the profit of the '
            'firm company (dividends to partners or their holding companies) plus any salary. "Profit pool per partner" = profit before tax (+ disclosed board/CEO salaries, '
            'typically partners) / partners: a proxy for AVERAGE equity-partner compensation before personal tax. It understates pay where partners are also on the payroll '
            '(salaries are not split out) and overstates local partner pay where profits flow to an international or cross-border partnership (amber rows). '
            'Limited partnerships (KB: Gernandt & Danielsson, DLA Piper, Baker McKenzie, CMS Wistrand, Magnusson) do not file public accounts.')
ws['A2'].alignment = W; ws.merge_cells('A2:P2'); ws.row_dimensions[2].height = 75

hdr = ['Firm', 'Entities included (org.nr)', 'FY', 'Revenue', 'Profit before tax', 'Board / CEO salaries', 'Proposed dividend', 'Partners', 'Partner source',
       'Profit pool / partner', 'Dividend / partner', 'Illustrative senior partner (1.4x avg)', 'Employees', 'Staff salaries / employee', 'Reliability', 'Notes']
R = 4
for c, h in enumerate(hdr, 1):
    x = ws.cell(R, c, h); x.font = H; x.fill = HF; x.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); x.border = B
ws.row_dimensions[R].height = 45
# firm, entities, FY, revenue, PBT, board sal, dividend, partners, partner src, employees, staff salaries, reliability, notes, amber
D = [
 ['Mannheimer Swartling', '5563994499', '2025', 2002.2, 797.4, None, 617, 92, 'DJ via Realtid (dividend split)', 649, 496.0, 'High', 'Pure lockstep. Dividend SEK 617m per Dagens Juridik/Realtid (allabolag shows none). 2024: PBT 789.6, dividend 582.', False],
 ['Vinge', '5566907100, 5566866108, 5566886551', '2025', 1446.6, 445.3, 52.3, 294.0, 71, 'Firm website, Oct-26', 529, 775.9, 'Medium', 'Stockholm + Malmö + Göteborg ABs. Realtid: Stockholm dividend 273m to 49 partners (SEK 5.5m each), Malmö 32m to 14. Partners may also draw salary (not split out).', False],
 ['Roschier (Sweden)', '5566865670', 'FY Jun25-May26', 1304.1, 646.4, 6.0, 508.9, 26, 'Firm website (Stockholm), Oct-26', 220, 234.9, 'Low', 'Profit likely pooled with Finnish practice (group 52 partners): SEK 12.4m per group partner on the same basis.', True],
 ['White & Case (Stockholm)', '5562323922', '2025', 1107.8, 239.0, 0.8, 198.1, 16, 'Implied (DJ)', 120, 105.3, 'Low', 'Global LLP: entity profit likely flows to the international partnership, not equal to local partner pay.', True],
 ['Setterwalls', '5568743230, 5565943221', '2025', 705.0, 272.6, None, 215.2, 38, 'Firm website, Oct-26', 217, 193.6, 'Medium', 'Stockholm + Göteborg operating ABs (parent AB is a pass-through). Malmö office excluded (left Sep-26).', False],
 ['Cederquist', '5569891277', '2025', None, 134.1, 0, 105.1, 26, 'Legal 500 (firm-supplied)', None, None, 'Low', 'AB is a partner in the operating KB (profit share); may not capture all partners.', False],
 ['Cirio', '5569530008', '2025', 430.6, 124.8, 6.5, 98.0, 27, 'Legal 500 (firm-supplied)', 127, 121.7, 'Medium', '2024: PBT 122.9, dividend 96.7.', False],
 ['Snellman (Sweden)', '5567572101', '2025', 279.4, 88.8, 4.2, 74.3, 17, 'Legal 500 (stale)', 107, 76.3, 'Low', 'Partner count unverified.', False],
 ['Delphi', '5566623293, 5566438163, 5563262913', '2025', 355.8, 82.9, None, 49.0, 49, 'Chambers (firm-supplied)', 126, 70.9, 'Low', 'Stockholm, Göteborg, Malmö ABs only (Östergötland missing). Low PBT relative to revenue suggests partners are paid largely via salary.', False],
 ['MAQS', '5569507733', '2025', 324.3, 96.4, None, 0, 48, 'Firm website, Oct-26', 121, 97.2, 'Low', 'Main AB only (~78% of Afv revenue). 2024 dividend 73.7. Partners likely partly salaried.', False],
 ['Hammarskiöld', '5565471637', '2025', 178.3, 49.8, None, 38.4, None, '', 63, 58.6, 'n/a', 'Partner count not found.', False],
 ['Foyen', '5565837134', '2025', 172.5, 46.1, 6.0, 35.9, 23, 'Firm website, Oct-26', 104, 72.3, 'Medium', 'Main Swedish AB.', False],
 ['Moll Wendén', '5566487939', '2025', 101.3, 30.2, None, 23.8, 8, 'Firm website, Oct-26', 44, None, 'Medium', '', False],
 ['TM & Partners', '5566947163', '2025', 102.5, 30.1, 4.6, 0, 18, 'Firm website, Oct-26', 73, 54.2, 'Low', 'Part of AGRD (PE-owned) since 2025; post-deal economics differ.', False],
 ['Westerberg & Partners', '5591623268', '2025', 106.5, 24.7, 5.0, 19.3, 12, 'Firm website, Oct-26', 35, 22.7, 'Medium', '', False],
 ['Harvest', '5590700224', '2025', 82.3, 22.9, None, 17.8, 5, 'Firm website, Oct-26', 27, None, 'Medium', '', False],
 ['Synch', '5569556656', '2025', 81.4, 20.3, None, 0, 9, 'Firm website, Oct-26', 38, None, 'Low', 'Part of AGRD (PE-owned) since 2025.', False],
 ['Kastell', '5568116494', '2025', 79.6, 16.9, None, 15.2, 9, 'Firm website, Oct-26', 21, None, 'Medium', '', False],
 ['Eversheds Sutherland', '5568782774', '2025', 165.5, 21.1, None, 16.5, None, '', 59, 49.4, 'n/a', 'Partner count not found.', False],
]
r = R
for d in D:
    r += 1
    firm, ents, fy, rev, pbt, bsal, div, p, psrc, emp, ssal, rel, note, am = d
    vals = [firm, ents, fy, rev, pbt, bsal, div, p, psrc]
    for c, v in enumerate(vals, 1):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W if c in (1, 2, 9) else C
        if c in (4, 5, 6, 7): x.number_format = '#,##0.0'
    if p:
        ws.cell(r, 10, f'=(E{r}+N(F{r}))/H{r}').number_format = '0.0'
        ws.cell(r, 11, f'=G{r}/H{r}').number_format = '0.0'
        ws.cell(r, 12, f'=J{r}*1.4').number_format = '0.0'
    for c in (10, 11, 12):
        x = ws.cell(r, c); x.border = B; x.alignment = C
        if not p: x.fill = grey
    x = ws.cell(r, 13, emp); x.border = B; x.alignment = C
    x = ws.cell(r, 14); x.border = B; x.alignment = C
    if ssal and emp:
        x.value = f'={ssal}/M{r}'; x.number_format = '0.00'
    else:
        x.fill = grey
    for c, v in ((15, rel), (16, note)):
        x = ws.cell(r, c, v); x.border = B; x.alignment = W
    if am:
        for c in range(1, 17):
            if ws.cell(r, c).fill != grey: ws.cell(r, c).fill = amber
last = r
r += 2
for t in [
    'Reading the table: e.g. Mannheimer Swartling SEK 797m profit before tax / 92 partners = SEK 8.7m average per partner before corporate and personal tax; dividend SEK 6.7m per partner (after 20.6% corporate tax).',
    'Senior partner: Mannheimer Swartling operates a pure lockstep (profit shares rise with seniority to a plateau). Individual point scales are not published. The illustrative column assumes a top-of-ladder partner earns ~1.4x the average (consistent with a 1:2.5 entry-to-top lockstep ladder); this is an assumption, not reported data.',
    'Tax treatment: partners owning <4% of the firm company have nearly all of their profit share taxed as salary income (rule since 2014, per Advokaten), so profit-pool figures are pre-personal-tax.',
    'Individual taxed income is public in Sweden (Skatteverket), but compiling named individuals\' income is personal data processing under GDPR; this sheet uses firm-level filed accounts only.',
    'Staff salaries / employee = "Löner övriga" / average employees (SEK m): an average across associates and support staff (and partners where on payroll).',
]:
    ws.cell(r, 1, t).font = N; r += 1

r += 1
ws.cell(r, 1, 'Context: SCB salary statistics, private-sector employees, monthly salary (SEK), 2025').font = Font(bold=True, size=11)
r += 1
for c, h in enumerate(['Occupation (SSYK 2012)', 'Mean', 'P10', 'P25', 'Median', 'P75', 'P90', 'P90 annualised (x12)'], 1):
    x = ws.cell(r, c, h); x.font = H; x.fill = HF; x.alignment = C; x.border = B
for occ, vals in (('2611 Advokater', (70200, 47200, 54800, 66000, 80000, 100000)), ('2614 Business & corporate lawyers', (73000, 40700, 50000, 67800, 86400, 109100))):
    r += 1
    ws.cell(r, 1, occ).border = B
    for c, v in enumerate(vals, 2):
        x = ws.cell(r, c, v); x.number_format = '#,##0'; x.border = B; x.alignment = C
    x = ws.cell(r, 8, f'=G{r}*12'); x.number_format = '#,##0'; x.border = B; x.alignment = C
r += 1
ws.cell(r, 1, 'SCB table AM0110A/LoneSpridSektYrk4AN (salaried employees only; equity partners paid via profit shares are not captured). 2611 mean rose from 58,600 (2023) to 70,200 (2025).').font = N

for c, w in zip(range(1, 17), (24, 22, 12, 10, 11, 11, 11, 9, 22, 11, 11, 13, 10, 11, 10, 50)):
    ws.column_dimensions[get_column_letter(c)].width = w
ws.freeze_panes = 'B5'
order = ['A. Market size', 'A. Firms', 'A. Growth drivers', 'A. Niche firms', 'B. Practice rankings', 'B. Rankings raw', 'D. Rate cards', 'E. Partner compensation']
wb._sheets = [wb[n] for n in order if n in wb.sheetnames] + [w for w in wb._sheets if w.title not in order]
wb.save(path)
print('ok', last)
