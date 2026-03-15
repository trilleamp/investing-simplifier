import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter

wb = Workbook()

# Color constants
BLUE = Font(name='Arial', color='0000FF', size=10)
BLUE_BOLD = Font(name='Arial', color='0000FF', size=10, bold=True)
BLACK = Font(name='Arial', color='000000', size=10)
BLACK_BOLD = Font(name='Arial', color='000000', size=10, bold=True)
GREEN = Font(name='Arial', color='008000', size=10)
WHITE_BOLD = Font(name='Arial', color='FFFFFF', size=11, bold=True)
HEADER_FONT = Font(name='Arial', color='FFFFFF', size=10, bold=True)
TITLE_FONT = Font(name='Arial', color='FFFFFF', size=14, bold=True)
SECTION_FONT = Font(name='Arial', color='1F4E79', size=11, bold=True)

DARK_BLUE_FILL = PatternFill('solid', fgColor='1F4E79')
MED_BLUE_FILL = PatternFill('solid', fgColor='2E75B6')
LIGHT_BLUE_FILL = PatternFill('solid', fgColor='D6E4F0')
LIGHT_GREEN_FILL = PatternFill('solid', fgColor='C6EFCE')
LIGHT_RED_FILL = PatternFill('solid', fgColor='FFC7CE')
YELLOW_FILL = PatternFill('solid', fgColor='FFFF00')
LIGHT_GRAY_FILL = PatternFill('solid', fgColor='F2F2F2')

CENTER = Alignment(horizontal='center', vertical='center')
LEFT = Alignment(horizontal='left', vertical='center')
RIGHT = Alignment(horizontal='right', vertical='center')
WRAP = Alignment(horizontal='center', vertical='center', wrap_text=True)

THIN_BORDER = Border(
    left=Side(style='thin'), right=Side(style='thin'),
    top=Side(style='thin'), bottom=Side(style='thin')
)

PCT_FMT = '0.0%'
PCT_FMT2 = '0.00%'
MONEY_FMT = '$#,##0'
MONEY_FMT1 = '$#,##0.0'
NUM_FMT = '#,##0.0'

# Company data (current as of March 13, 2026)
companies = [
    ['Veeva (VEEV)', 'FY2026 (Jan 2026)', 1438, 1121, 6500, 0, 8500, 0.22, 1.08, 0.0, 691.4, 5832, 0, 0.223],
    ['ServiceNow (NOW)', 'FY2025 (Dec 2025)', 1824, 1748, 12964, 1491, 26038, 0.227, 0.97, 0.0153, 1364, 9609, 2278, 0.18],
    ['Intuit (INTU)', 'FY2025 (Jul 2025)', 4923, 3869, 19710, 5973, 36958, 0.20, 1.28, 0.048, 3630, 18436, 6038, 0.17],
    ['Adobe (ADBE)', 'FY2025 (Nov 2025)', 8706, 7130, 11623, 6210, 29496, 0.184, 1.53, 0.0424, 6741, 14105, 5628, 0.198],
    ['Shopify (SHOP)', 'FY2025 (Dec 2025)', 1468, 1231, 13473, 188, 15189, 0.184, 2.82, 0.045, 1075, 11558, 918, 0.094],
    ['Axon (AXON)', 'FY2025 (Dec 2025)', -62.1, 124.7, 3243, 1811, 7000, 0.21, 1.52, 0.052, 58.5, 2328, 690, 0.012],
    ['BillionToOne (BLLN)', 'FY2025 (Dec 2025)', 16.0, 2.9, 480.1, 52.5, 600, 0.04, 1.34, 0.08, -47.15, 168, 109, 0.04],
    ['Intuitive Surgical (ISRG)', 'FY2025 (Dec 2025)', 2949, 2900, 17942, 0, 20459, 0.215, 1.69, 0.0, 1767, 13397, 0, 0.1258],
    ['Eli Lilly (LLY)', 'FY2025 (Dec 2025)', 28300, 20600, 26535, 42510, 112476, 0.19, 0.43, 0.026, 12899, 14272, 33644, 0.165],
    ['Palantir (PLTR)', 'FY2025 (Dec 2025)', 1414, 1625, 7500, 0, 8900, 0.015, 1.69, 0.0, 310.4, 5094, 0, 0.0435],
    ['Tesla (TSLA)', 'FY2025 (Dec 2025)', 4355, 3794, 82900, 1800, 137806, 0.27, 2.07, 0.025, 8891, 72913, 7878, 0.204],
    ['Constellation Software (CSU)', 'FY2025 (Dec 2025)', 1895, 512, 4267, 4554, 16171, 0.376, 0.72, 0.052, 1463, 3288, 4166, 0.241],
    ['Samsara (IOT)', 'FY2026 (Jan 2026)', -52.6, -9.1, 1420, 72.8, 2541, 0.0, 1.62, 0.0, -190.0, 1069, 0, 0.0],
    ['Comfort Systems (FIX)', 'FY2025 (Dec 2025)', 1315, 1023, 2449, 145, 6441, 0.209, 1.58, 0.062, 749.4, 1705, 68.3, 0.2162],
  ['Amphenol (APH)', 'FY2025 (Dec 2025)', 5869, 4270.3, 13510, 15502, 36237, 0.231, 1.21, 0.0237, 3157, 9856, 6886, 0.189],
]

RFR = 0.0425
MRP = 0.055

# ===================================================================
# SHEET 1: INPUTS
# ===================================================================
ws1 = wb.active
ws1.title = 'Inputs'
ws1.sheet_properties.tabColor = '1F4E79'

# Title
ws1.merge_cells('A1:P1')
ws1['A1'] = 'EVA & COMPOUNDER ANALYSIS - INPUT DATA'
ws1['A1'].font = TITLE_FONT
ws1['A1'].fill = DARK_BLUE_FILL
ws1['A1'].alignment = CENTER

# Assumptions
ws1['A3'] = 'CAPM ASSUMPTIONS'
ws1['A3'].font = SECTION_FONT
ws1['A4'] = 'Risk-Free Rate'
ws1['B4'] = RFR
ws1['B4'].font = BLUE
ws1['B4'].number_format = PCT_FMT2
ws1['B4'].fill = YELLOW_FILL
ws1['A5'] = 'Market Risk Premium'
ws1['B5'] = MRP
ws1['B5'].font = BLUE
ws1['B5'].number_format = PCT_FMT2
ws1['B5'].fill = YELLOW_FILL
for r in range(3, 6):
    ws1.cell(row=r, column=1).font = BLACK_BOLD if r == 3 else BLACK

# Headers row 7
headers = ['Company', 'Period', 'EBIT ($M)', 'Net Income ($M)', 'Equity ($M)',
           'Debt ($M)', 'Total Assets ($M)', 'Tax Rate', 'Beta', 'Cost of Debt',
           'Prev EBIT ($M)', 'Prev Equity ($M)', 'Prev Debt ($M)', 'Prev Tax Rate']

for col, h in enumerate(headers, 1):
    cell = ws1.cell(row=7, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = MED_BLUE_FILL
    cell.alignment = WRAP
    cell.border = THIN_BORDER

# Data rows
for i, c in enumerate(companies):
    row = 8 + i
    for col, val in enumerate(c, 1):
        cell = ws1.cell(row=row, column=col, value=val)
        cell.border = THIN_BORDER
        if col <= 2:
            cell.font = BLACK_BOLD if col == 1 else BLACK
            cell.alignment = LEFT
        elif col in [8, 9, 10, 14]:
            cell.font = BLUE
            cell.number_format = PCT_FMT2
            cell.alignment = CENTER
        else:
            cell.font = BLUE
            cell.number_format = MONEY_FMT1 if abs(val) < 100 else MONEY_FMT
            cell.alignment = RIGHT
        if row % 2 == 0:
            cell.fill = LIGHT_GRAY_FILL

# Column widths
widths = [30, 22, 14, 16, 14, 14, 16, 12, 10, 13, 15, 15, 14, 14]
for i, w in enumerate(widths):
    ws1.column_dimensions[get_column_letter(i+1)].width = w

# ===================================================================
# SHEET 2: CALCULATIONS
# ===================================================================
ws2 = wb.create_sheet('Calculations')
ws2.sheet_properties.tabColor = '2E75B6'

ws2.merge_cells('A1:O1')
ws2['A1'] = 'EVA & COMPOUNDER CALCULATIONS'
ws2['A1'].font = TITLE_FONT
ws2['A1'].fill = DARK_BLUE_FILL
ws2['A1'].alignment = CENTER

calc_headers = ['Company', 'Cost of Equity', 'Capital Invested ($M)',
                'Equity Weight', 'Debt Weight', 'WACC', 'NOPAT ($M)',
                'ROIC', 'ROE', 'ROA', 'Prev NOPAT ($M)', 'Prev Capital ($M)',
                'Delta NOPAT ($M)', 'Delta Capital ($M)', 'ROIIC']

for col, h in enumerate(calc_headers, 1):
    cell = ws2.cell(row=3, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = MED_BLUE_FILL
    cell.alignment = WRAP
    cell.border = THIN_BORDER

for i in range(len(companies)):
    row = 4 + i
    r_in = 8 + i  # row in Inputs sheet

    # A: Company name
    ws2.cell(row=row, column=1, value=f"=Inputs!A{r_in}").font = GREEN
    ws2.cell(row=row, column=1).border = THIN_BORDER

    # B: Cost of Equity = RFR + Beta * MRP
    ws2.cell(row=row, column=2, value=f"=Inputs!$B$4+Inputs!I{r_in}*Inputs!$B$5").font = BLACK
    ws2.cell(row=row, column=2).number_format = PCT_FMT2
    ws2.cell(row=row, column=2).border = THIN_BORDER

    # C: Capital Invested = Equity + Debt
    ws2.cell(row=row, column=3, value=f"=Inputs!E{r_in}+Inputs!F{r_in}").font = BLACK
    ws2.cell(row=row, column=3).number_format = MONEY_FMT
    ws2.cell(row=row, column=3).border = THIN_BORDER

    # D: Equity Weight
    ws2.cell(row=row, column=4, value=f"=IF(C{row}=0,0,Inputs!E{r_in}/C{row})").font = BLACK
    ws2.cell(row=row, column=4).number_format = PCT_FMT
    ws2.cell(row=row, column=4).border = THIN_BORDER

    # E: Debt Weight
    ws2.cell(row=row, column=5, value=f"=IF(C{row}=0,0,Inputs!F{r_in}/C{row})").font = BLACK
    ws2.cell(row=row, column=5).number_format = PCT_FMT
    ws2.cell(row=row, column=5).border = THIN_BORDER

    # F: WACC = Eq_W * Ke + Dt_W * Kd * (1 - tax)
    ws2.cell(row=row, column=6, value=f"=D{row}*B{row}+E{row}*Inputs!J{r_in}*(1-Inputs!H{r_in})").font = BLACK
    ws2.cell(row=row, column=6).number_format = PCT_FMT2
    ws2.cell(row=row, column=6).border = THIN_BORDER

    # G: NOPAT = EBIT * (1 - tax)
    ws2.cell(row=row, column=7, value=f"=Inputs!C{r_in}*(1-Inputs!H{r_in})").font = BLACK
    ws2.cell(row=row, column=7).number_format = MONEY_FMT
    ws2.cell(row=row, column=7).border = THIN_BORDER

    # H: ROIC = NOPAT / Capital
    ws2.cell(row=row, column=8, value=f"=IF(C{row}=0,0,G{row}/C{row})").font = BLACK
    ws2.cell(row=row, column=8).number_format = PCT_FMT2
    ws2.cell(row=row, column=8).border = THIN_BORDER

    # I: ROE = Net Income / Equity
    ws2.cell(row=row, column=9, value=f"=IF(Inputs!E{r_in}=0,0,Inputs!D{r_in}/Inputs!E{r_in})").font = BLACK
    ws2.cell(row=row, column=9).number_format = PCT_FMT2
    ws2.cell(row=row, column=9).border = THIN_BORDER

    # J: ROA = Net Income / Total Assets
    ws2.cell(row=row, column=10, value=f"=IF(Inputs!G{r_in}=0,0,Inputs!D{r_in}/Inputs!G{r_in})").font = BLACK
    ws2.cell(row=row, column=10).number_format = PCT_FMT2
    ws2.cell(row=row, column=10).border = THIN_BORDER

    # K: Prev NOPAT = Prev EBIT * (1 - Prev Tax)
    ws2.cell(row=row, column=11, value=f"=Inputs!K{r_in}*(1-Inputs!N{r_in})").font = BLACK
    ws2.cell(row=row, column=11).number_format = MONEY_FMT
    ws2.cell(row=row, column=11).border = THIN_BORDER

    # L: Prev Capital = Prev Equity + Prev Debt
    ws2.cell(row=row, column=12, value=f"=Inputs!L{r_in}+Inputs!M{r_in}").font = BLACK
    ws2.cell(row=row, column=12).number_format = MONEY_FMT
    ws2.cell(row=row, column=12).border = THIN_BORDER

    # M: Delta NOPAT
    ws2.cell(row=row, column=13, value=f"=G{row}-K{row}").font = BLACK
    ws2.cell(row=row, column=13).number_format = MONEY_FMT
    ws2.cell(row=row, column=13).border = THIN_BORDER

    # N: Delta Capital
    ws2.cell(row=row, column=14, value=f"=C{row}-L{row}").font = BLACK
    ws2.cell(row=row, column=14).number_format = MONEY_FMT
    ws2.cell(row=row, column=14).border = THIN_BORDER

    # O: ROIIC = Delta NOPAT / Delta Capital
    ws2.cell(row=row, column=15, value=f'=IF(ABS(N{row})<10,"N/M",M{row}/N{row})').font = BLACK
    ws2.cell(row=row, column=15).number_format = PCT_FMT
    ws2.cell(row=row, column=15).border = THIN_BORDER

    if row % 2 == 0:
        for col in range(1, 16):
            ws2.cell(row=row, column=col).fill = LIGHT_GRAY_FILL

calc_widths = [30, 14, 18, 13, 12, 10, 14, 10, 10, 10, 15, 15, 16, 16, 10]
for i, w in enumerate(calc_widths):
    ws2.column_dimensions[get_column_letter(i+1)].width = w

# ===================================================================
# SHEET 3: EVA DASHBOARD
# ===================================================================
ws3 = wb.create_sheet('EVA Dashboard')
ws3.sheet_properties.tabColor = '00B050'

ws3.merge_cells('A1:J1')
ws3['A1'] = 'EVA & COMPOUNDER DASHBOARD'
ws3['A1'].font = TITLE_FONT
ws3['A1'].fill = DARK_BLUE_FILL
ws3['A1'].alignment = CENTER

ws3['A2'] = 'Data as of March 13, 2026 | Most recent fiscal year for each company'
ws3['A2'].font = Font(name='Arial', color='666666', size=9, italic=True)

dash_headers = ['Company', 'ROIC', 'WACC', 'ROIC-WACC', 'ROIIC', 'ROE', 'ROA',
                'EVA ($M)', 'Score', 'Assessment']

for col, h in enumerate(dash_headers, 1):
    cell = ws3.cell(row=4, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = DARK_BLUE_FILL
    cell.alignment = WRAP
    cell.border = THIN_BORDER

for i in range(len(companies)):
    row = 5 + i
    r_calc = 4 + i

    # A: Company
    ws3.cell(row=row, column=1, value=f"=Calculations!A{r_calc}").font = GREEN
    ws3.cell(row=row, column=1).border = THIN_BORDER

    # B: ROIC
    ws3.cell(row=row, column=2, value=f"=Calculations!H{r_calc}").font = BLACK
    ws3.cell(row=row, column=2).number_format = PCT_FMT
    ws3.cell(row=row, column=2).border = THIN_BORDER

    # C: WACC
    ws3.cell(row=row, column=3, value=f"=Calculations!F{r_calc}").font = BLACK
    ws3.cell(row=row, column=3).number_format = PCT_FMT
    ws3.cell(row=row, column=3).border = THIN_BORDER

    # D: ROIC - WACC spread
    ws3.cell(row=row, column=4, value=f"=B{row}-C{row}").font = BLACK_BOLD
    ws3.cell(row=row, column=4).number_format = '+0.0%;-0.0%'
    ws3.cell(row=row, column=4).border = THIN_BORDER

    # E: ROIIC
    ws3.cell(row=row, column=5, value=f"=Calculations!O{r_calc}").font = BLACK
    ws3.cell(row=row, column=5).number_format = PCT_FMT
    ws3.cell(row=row, column=5).border = THIN_BORDER

    # F: ROE
    ws3.cell(row=row, column=6, value=f"=Calculations!I{r_calc}").font = BLACK
    ws3.cell(row=row, column=6).number_format = PCT_FMT
    ws3.cell(row=row, column=6).border = THIN_BORDER

    # G: ROA
    ws3.cell(row=row, column=7, value=f"=Calculations!J{r_calc}").font = BLACK
    ws3.cell(row=row, column=7).number_format = PCT_FMT
    ws3.cell(row=row, column=7).border = THIN_BORDER

    # H: EVA = NOPAT - (Capital * WACC)
    ws3.cell(row=row, column=8, value=f"=Calculations!G{r_calc}-(Calculations!C{r_calc}*Calculations!F{r_calc})").font = BLACK_BOLD
    ws3.cell(row=row, column=8).number_format = '$#,##0;($#,##0);"-"'
    ws3.cell(row=row, column=8).border = THIN_BORDER

    # I: Score (out of 5)
    # 1 pt: ROIC > WACC, 1 pt: ROIC > 15%, 1 pt: ROIIC > WACC (if numeric), 1 pt: ROE > 15%, 1 pt: EVA > 0
    ws3.cell(row=row, column=9, value=(
        f'=IF(B{row}>C{row},1,0)'
        f'+IF(B{row}>0.15,1,0)'
        f'+IF(AND(ISNUMBER(E{row}),E{row}>C{row}),1,0)'
        f'+IF(F{row}>0.15,1,0)'
        f'+IF(H{row}>0,1,0)'
    )).font = BLACK_BOLD
    ws3.cell(row=row, column=9).alignment = CENTER
    ws3.cell(row=row, column=9).border = THIN_BORDER

    # J: Assessment
    ws3.cell(row=row, column=10, value=(
        f'=IF(I{row}>=4,"COMPOUNDER",IF(I{row}>=2,"POTENTIAL COMPOUNDER","NOT YET A COMPOUNDER"))'
    )).font = BLACK_BOLD
    ws3.cell(row=row, column=10).alignment = CENTER
    ws3.cell(row=row, column=10).border = THIN_BORDER

dash_widths = [30, 10, 10, 12, 10, 10, 10, 14, 8, 26]
for i, w in enumerate(dash_widths):
    ws3.column_dimensions[get_column_letter(i+1)].width = w

# ===================================================================
# SHEET 4: VEEVA PROJECTION
# ===================================================================
ws4 = wb.create_sheet('Veeva Projection')
ws4.sheet_properties.tabColor = 'FF6600'

ws4.merge_cells('A1:H1')
ws4['A1'] = 'VEEVA SYSTEMS (VEEV) - PROJECTED EVA: FY2025 to FY2030'
ws4['A1'].font = TITLE_FONT
ws4['A1'].fill = DARK_BLUE_FILL
ws4['A1'].alignment = CENTER

ws4['A3'] = 'Assumptions:'
ws4['A3'].font = SECTION_FONT
assumptions = [
    'Salesforce royalties (~$120M/yr) fully eliminated by FY2027',
    'AI agent revenue begins FY2028, scales through FY2030',
    'Management target: $6B revenue by 2030, 35%+ operating margin',
    'Beta declines as business matures (1.10 -> 1.00)',
    'Debt remains at $0 (Veeva has no debt)',
]
for j, a in enumerate(assumptions):
    ws4.cell(row=4+j, column=1, value=f'  - {a}').font = Font(name='Arial', size=9, color='444444')

proj_headers = ['Year', 'Revenue ($M)', 'EBIT Margin', 'EBIT ($M)', 'Tax Rate',
                'NOPAT ($M)', 'Equity ($M)', 'Beta', 'WACC', 'ROIC',
                'ROIC-WACC', 'EVA ($M)', 'Verdict']

for col, h in enumerate(proj_headers, 1):
    cell = ws4.cell(row=10, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = MED_BLUE_FILL
    cell.alignment = WRAP
    cell.border = THIN_BORDER

veeva_proj = [
    ['FY2025 (Actual)', 2756, 0.251, 0.223, 5832, 1.10],
    ['FY2026 (Actual)', 3195, 0.45, 0.22, 6500, 1.08],
    ['FY2027 (Guided)', 3590, 0.44, 0.22, 7200, 1.06],
    ['FY2028 (Projected)', 4100, 0.40, 0.22, 8000, 1.04],
    ['FY2029 (Projected)', 5000, 0.37, 0.22, 9000, 1.02],
    ['FY2030 (Target)', 6000, 0.35, 0.22, 10000, 1.00],
]

for i, p in enumerate(veeva_proj):
    row = 11 + i
    # A: Year
    ws4.cell(row=row, column=1, value=p[0]).font = BLACK_BOLD
    ws4.cell(row=row, column=1).border = THIN_BORDER

    # B: Revenue
    ws4.cell(row=row, column=2, value=p[1]).font = BLUE
    ws4.cell(row=row, column=2).number_format = MONEY_FMT
    ws4.cell(row=row, column=2).border = THIN_BORDER
    ws4.cell(row=row, column=2).fill = YELLOW_FILL

    # C: EBIT Margin
    ws4.cell(row=row, column=3, value=p[2]).font = BLUE
    ws4.cell(row=row, column=3).number_format = PCT_FMT
    ws4.cell(row=row, column=3).border = THIN_BORDER
    ws4.cell(row=row, column=3).fill = YELLOW_FILL

    # D: EBIT = Revenue * Margin
    ws4.cell(row=row, column=4, value=f"=B{row}*C{row}").font = BLACK
    ws4.cell(row=row, column=4).number_format = MONEY_FMT
    ws4.cell(row=row, column=4).border = THIN_BORDER

    # E: Tax Rate
    ws4.cell(row=row, column=5, value=p[3]).font = BLUE
    ws4.cell(row=row, column=5).number_format = PCT_FMT
    ws4.cell(row=row, column=5).border = THIN_BORDER

    # F: NOPAT = EBIT * (1 - Tax)
    ws4.cell(row=row, column=6, value=f"=D{row}*(1-E{row})").font = BLACK
    ws4.cell(row=row, column=6).number_format = MONEY_FMT
    ws4.cell(row=row, column=6).border = THIN_BORDER

    # G: Equity (= Capital since no debt)
    ws4.cell(row=row, column=7, value=p[4]).font = BLUE
    ws4.cell(row=row, column=7).number_format = MONEY_FMT
    ws4.cell(row=row, column=7).border = THIN_BORDER
    ws4.cell(row=row, column=7).fill = YELLOW_FILL

    # H: Beta
    ws4.cell(row=row, column=8, value=p[5]).font = BLUE
    ws4.cell(row=row, column=8).number_format = '0.00'
    ws4.cell(row=row, column=8).border = THIN_BORDER
    ws4.cell(row=row, column=8).fill = YELLOW_FILL

    # I: WACC = Cost of Equity (no debt) = RFR + Beta * MRP
    ws4.cell(row=row, column=9, value=f"=Inputs!$B$4+H{row}*Inputs!$B$5").font = BLACK
    ws4.cell(row=row, column=9).number_format = PCT_FMT2
    ws4.cell(row=row, column=9).border = THIN_BORDER

    # J: ROIC = NOPAT / Capital
    ws4.cell(row=row, column=10, value=f"=IF(G{row}=0,0,F{row}/G{row})").font = BLACK
    ws4.cell(row=row, column=10).number_format = PCT_FMT
    ws4.cell(row=row, column=10).border = THIN_BORDER

    # K: ROIC - WACC
    ws4.cell(row=row, column=11, value=f"=J{row}-I{row}").font = BLACK_BOLD
    ws4.cell(row=row, column=11).number_format = '+0.0%;-0.0%'
    ws4.cell(row=row, column=11).border = THIN_BORDER

    # L: EVA = NOPAT - (Capital * WACC)
    ws4.cell(row=row, column=12, value=f"=F{row}-(G{row}*I{row})").font = BLACK_BOLD
    ws4.cell(row=row, column=12).number_format = '$#,##0;($#,##0)'
    ws4.cell(row=row, column=12).border = THIN_BORDER

    # M: Verdict
    ws4.cell(row=row, column=13, value=f'=IF(L{row}>0,"CREATING VALUE","DESTROYING VALUE")').font = BLACK_BOLD
    ws4.cell(row=row, column=13).border = THIN_BORDER

    if row % 2 == 0:
        for col in range(1, 14):
            if not ws4.cell(row=row, column=col).fill or ws4.cell(row=row, column=col).fill.fgColor.rgb == '00000000':
                ws4.cell(row=row, column=col).fill = LIGHT_GRAY_FILL

proj_widths = [22, 14, 12, 12, 10, 14, 14, 8, 10, 10, 12, 14, 20]
for i, w in enumerate(proj_widths):
    ws4.column_dimensions[get_column_letter(i+1)].width = w

# ===================================================================
# SHEET 5: FORMULAS GUIDE
# ===================================================================
ws5 = wb.create_sheet('Formulas Guide')
ws5.sheet_properties.tabColor = '7030A0'

ws5.merge_cells('A1:C1')
ws5['A1'] = 'FORMULAS & DEFINITIONS GUIDE'
ws5['A1'].font = TITLE_FONT
ws5['A1'].fill = DARK_BLUE_FILL
ws5['A1'].alignment = CENTER

formulas = [
    ['NOPAT', 'Net Operating Profit After Tax', 'EBIT x (1 - Tax Rate)'],
    ['Capital Invested', 'Total capital used to run business', 'Equity + Interest-Bearing Debt'],
    ['WACC', 'Weighted Average Cost of Capital', '(Eq% x Cost of Equity) + (Debt% x Cost of Debt x (1-Tax))'],
    ['Cost of Equity', 'Return equity investors demand (CAPM)', 'Risk-Free Rate + Beta x Market Risk Premium'],
    ['ROIC', 'Return on Invested Capital', 'NOPAT / Capital Invested'],
    ['EVA', 'Economic Value Added', 'NOPAT - (Capital Invested x WACC)'],
    ['ROE', 'Return on Equity', 'Net Income / Shareholders Equity'],
    ['ROA', 'Return on Assets', 'Net Income / Total Assets'],
    ['ROIIC', 'Return on Incremental Invested Capital', 'Change in NOPAT / Change in Capital Invested'],
    ['ROIC-WACC Spread', 'Value creation signal', 'ROIC - WACC (positive = creating value)'],
]

guide_headers = ['Metric', 'Description', 'Formula']
for col, h in enumerate(guide_headers, 1):
    cell = ws5.cell(row=3, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = MED_BLUE_FILL
    cell.border = THIN_BORDER

for i, f in enumerate(formulas):
    for j, val in enumerate(f):
        cell = ws5.cell(row=4+i, column=j+1, value=val)
        cell.font = BLACK_BOLD if j == 0 else BLACK
        cell.border = THIN_BORDER
        if (4+i) % 2 == 0:
            cell.fill = LIGHT_GRAY_FILL

ws5.column_dimensions['A'].width = 20
ws5.column_dimensions['B'].width = 45
ws5.column_dimensions['C'].width = 55

# Scoring guide
ws5['A16'] = 'COMPOUNDER SCORING (5 points possible):'
ws5['A16'].font = SECTION_FONT
scoring = [
    ['1 point', 'ROIC > WACC', 'Company earns above its cost of capital'],
    ['1 point', 'ROIC > 15%', 'High absolute return on capital'],
    ['1 point', 'ROIIC > WACC', 'Reinvesting incremental capital at returns above cost'],
    ['1 point', 'ROE > 15%', 'Strong return to equity holders'],
    ['1 point', 'EVA > 0', 'Creating economic value in absolute dollars'],
]
score_headers = ['Points', 'Criteria', 'What It Means']
for col, h in enumerate(score_headers, 1):
    cell = ws5.cell(row=17, column=col, value=h)
    cell.font = HEADER_FONT
    cell.fill = MED_BLUE_FILL
    cell.border = THIN_BORDER

for i, s in enumerate(scoring):
    for j, val in enumerate(s):
        cell = ws5.cell(row=18+i, column=j+1, value=val)
        cell.font = BLACK
        cell.border = THIN_BORDER

ws5['A24'] = 'Score 4-5: COMPOUNDER'
ws5['A24'].font = Font(name='Arial', color='008000', size=10, bold=True)
ws5['A25'] = 'Score 2-3: POTENTIAL COMPOUNDER'
ws5['A25'].font = Font(name='Arial', color='FF8C00', size=10, bold=True)
ws5['A26'] = 'Score 0-1: NOT YET A COMPOUNDER'
ws5['A26'].font = Font(name='Arial', color='FF0000', size=10, bold=True)

# Source note
ws5['A28'] = 'Source: Based on "Forget ROE. This Is the Metric That Actually Matters" - Jimmy\'s Journal, July 2025'
ws5['A28'].font = Font(name='Arial', color='666666', size=9, italic=True)

# ===================================================================
# SAVE
# ===================================================================
output_path = r'C:\Users\jytte\Desktop\CLAUDE CODE WORK\EVA_Compounder_Analysis.xlsx'
wb.save(output_path)
print(f'Spreadsheet saved to: {output_path}')
print('Sheets: Inputs, Calculations, EVA Dashboard, Veeva Projection, Formulas Guide')
