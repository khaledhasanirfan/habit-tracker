"""
Personal Life Tracker - Excel Workbook Generator
Generates Personal_Life_Tracker.xlsx with 9 sheets for full year 2026.
"""

from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.formatting.rule import (
    ColorScaleRule, CellIsRule, FormulaRule, DataBarRule
)
from openpyxl.chart import BarChart, Reference, LineChart, PieChart
from openpyxl.chart.series import DataPoint
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo
from datetime import date, timedelta
import calendar

# ─────────────────────────────────────────────
# GLOBAL STYLE HELPERS
# ─────────────────────────────────────────────

def hex_fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def make_font(bold=False, size=11, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def right_align():
    return Alignment(horizontal="right", vertical="center")

# Color palette
DARK_NAVY   = "1B2A4A"
TEAL        = "00897B"
LIGHT_TEAL  = "E0F2F1"
PURPLE      = "5C35A1"
LIGHT_PURPLE= "EDE7F6"
ORANGE      = "F57C00"
LIGHT_ORANGE= "FFF3E0"
RED         = "E53935"
LIGHT_RED   = "FFEBEE"
GREEN       = "2E7D32"
LIGHT_GREEN = "E8F5E9"
YELLOW      = "F9A825"
LIGHT_YELLOW= "FFFDE7"
GRAY_BG     = "F5F5F5"
WHITE       = "FFFFFF"
MID_GRAY    = "EEEEEE"
DARK_GRAY   = "757575"
BLUE        = "1565C0"
LIGHT_BLUE  = "E3F2FD"
PINK        = "AD1457"
LIGHT_PINK  = "FCE4EC"

# ─────────────────────────────────────────────
# DATE HELPERS
# ─────────────────────────────────────────────

def all_dates_2026():
    start = date(2026, 1, 1)
    end   = date(2026, 12, 31)
    delta = end - start
    return [start + timedelta(days=i) for i in range(delta.days + 1)]

DATES = all_dates_2026()
DAYS  = ["Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"]

DAY_TYPE = {
    "Monday":    "Grind Day",
    "Tuesday":   "Grind Day",
    "Wednesday": "Grind Day",
    "Thursday":  "Rest Day",
    "Friday":    "Grind Day",
    "Saturday":  "Grind Day",
    "Sunday":    "Rest Day",
}

MONTHS = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
]

# ─────────────────────────────────────────────
# APPLY HEADER ROW STYLE
# ─────────────────────────────────────────────

def style_header_row(ws, row_num, bg_color, font_color=WHITE, cols=None):
    """Apply header styling to an entire row."""
    if cols is None:
        cols = range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row_num, column=c)
        cell.fill = hex_fill(bg_color)
        cell.font = make_font(bold=True, color=font_color, size=10)
        cell.alignment = center()
        cell.border = thin_border()

def style_data_row(ws, row_num, bg_color=WHITE, cols=None):
    if cols is None:
        cols = range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row_num, column=c)
        cell.fill = hex_fill(bg_color)
        cell.border = thin_border()
        cell.alignment = center()
        cell.font = make_font(size=10)

def set_col_width(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def add_dv_truefalse(ws, col_letter, start_row, end_row):
    dv = DataValidation(
        type="list", formula1='"TRUE,FALSE"',
        allow_blank=True, showDropDown=False
    )
    dv.sqref = f"{col_letter}{start_row}:{col_letter}{end_row}"
    ws.add_data_validation(dv)
    return dv

# ─────────────────────────────────────────────
# SHEET 2: SALAT TRACKER
# ─────────────────────────────────────────────

def create_salat_sheet(wb):
    ws = wb.create_sheet("Salat Tracker")
    ws.sheet_view.showGridLines = True

    headers = ["Date","Day","Fajr","Dhuhr","Asr","Maghrib","Isha",
               "Total Prayers","Completion %","Notes"]
    col_widths = [13, 12, 9, 9, 9, 9, 9, 14, 14, 25]

    # Title row
    ws.merge_cells("A1:J1")
    title = ws["A1"]
    title.value = "🕌  SALAT TRACKER 2026"
    title.fill = hex_fill(DARK_NAVY)
    title.font = make_font(bold=True, size=14, color=WHITE)
    title.alignment = center()
    ws.row_dimensions[1].height = 32

    # Sub-title
    ws.merge_cells("A2:J2")
    sub = ws["A2"]
    sub.value = "Track your 5 daily prayers  •  Goal: 5/5 prayers every day"
    sub.fill = hex_fill(TEAL)
    sub.font = make_font(italic=True, color=WHITE, size=10)
    sub.alignment = center()
    ws.row_dimensions[2].height = 22

    # Header row 3
    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        cell = ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, TEAL, WHITE, range(1, 11))

    # Data rows
    prayer_cols = [3, 4, 5, 6, 7]  # C, D, E, F, G
    prayer_letters = ["C", "D", "E", "F", "G"]

    for i, d in enumerate(DATES):
        row = i + 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        # Prayers default FALSE
        for c in prayer_cols:
            cell = ws.cell(row=row, column=c, value="FALSE")
        # Total prayers = COUNTIF on prayer cells
        pr = f"C{row}:G{row}"
        ws.cell(row=row, column=8,
                value=f'=COUNTIF({pr},"TRUE")')
        ws.cell(row=row, column=9,
                value=f"=H{row}/5").number_format = "0%"
        # Notes blank
        ws.cell(row=row, column=10, value="")
        style_data_row(ws, row, bg, range(1, 11))
        ws.row_dimensions[row].height = 20

    # Data validations for prayer columns
    last_row = 3 + len(DATES)
    for ltr in prayer_letters:
        add_dv_truefalse(ws, ltr, 4, last_row)

    # Conditional formatting on Completion %
    col_i = "I"
    data_range = f"{col_i}4:{col_i}{last_row}"
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=['"100%"'],
                   fill=hex_fill(LIGHT_GREEN),
                   font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="greaterThanOrEqual", formula=["0.6"],
                   fill=hex_fill(LIGHT_YELLOW),
                   font=Font(color=YELLOW, bold=True)))
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="lessThan", formula=["0.6"],
                   fill=hex_fill(LIGHT_RED),
                   font=Font(color=RED, bold=True)))
    # Full green if =1 (100%)
    ws.conditional_formatting.add(data_range,
        CellIsRule(operator="equal", formula=["1"],
                   fill=hex_fill(LIGHT_GREEN),
                   font=Font(color=GREEN, bold=True)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:J{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 3: GYM ROUTINE
# ─────────────────────────────────────────────

def create_gym_sheet(wb):
    ws = wb.create_sheet("Gym Routine")

    headers = ["Date","Day","Planned Type","Workout Done",
               "Workout Name","Body Part","Duration (min)","Intensity","Notes"]
    col_widths = [13, 12, 13, 14, 22, 16, 15, 13, 25]

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = "💪  GYM ROUTINE TRACKER 2026"
    t.fill = hex_fill(PURPLE)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    s = ws["A2"]
    s.value = "Mon/Tue/Wed/Fri/Sat = Grind Day  •  Thu/Sun = Rest Day  •  Goal: 5 workouts/week"
    s.fill = hex_fill(LIGHT_PURPLE)
    s.font = make_font(italic=True, color=PURPLE, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, PURPLE, WHITE, range(1, 10))

    intensity_options = ["Low","Medium","High","Very High"]
    body_part_options = ["Chest","Back","Shoulders","Arms","Legs","Core","Full Body","Cardio"]

    dv_done = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    dv_intensity = DataValidation(type="list",
                                   formula1='"Low,Medium,High,Very High"',
                                   allow_blank=True)
    ws.add_data_validation(dv_done)
    ws.add_data_validation(dv_intensity)

    last_row = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        day_name = d.strftime("%A")
        plan = DAY_TYPE[day_name]
        bg = LIGHT_PURPLE if plan == "Rest Day" else (WHITE if i % 2 == 0 else GRAY_BG)

        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=day_name)
        ws.cell(row=row, column=3, value=plan)
        ws.cell(row=row, column=4, value="FALSE")
        ws.cell(row=row, column=5, value="")
        ws.cell(row=row, column=6, value="")
        ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=9, value="")
        style_data_row(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 20

        # Color Rest Day differently
        if plan == "Rest Day":
            for c in range(1, 10):
                ws.cell(row=row, column=c).fill = hex_fill(LIGHT_PURPLE)
                ws.cell(row=row, column=c).font = make_font(color=PURPLE, size=10, italic=True)

    dv_done.sqref = f"D4:D{last_row}"
    dv_intensity.sqref = f"H4:H{last_row}"

    # Conditional formatting: Grind Done = green, Grind not done = light red
    ws.conditional_formatting.add(f"D4:D{last_row}",
        FormulaRule(formula=['AND(C4="Grind Day",D4="TRUE")'],
                    fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"D4:D{last_row}",
        FormulaRule(formula=['AND(C4="Grind Day",D4="FALSE")'],
                    fill=hex_fill(LIGHT_RED), font=Font(color=RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 4: SUPPLEMENTS
# ─────────────────────────────────────────────

def create_supplements_sheet(wb):
    ws = wb.create_sheet("Supplements")

    headers = ["Date","Day","Multivitamin Taken","Omega Taken",
               "Both Completed","Notes"]
    col_widths = [13, 12, 18, 14, 16, 25]

    ws.merge_cells("A1:F1")
    t = ws["A1"]
    t.value = "💊  SUPPLEMENTS TRACKER 2026"
    t.fill = hex_fill(ORANGE)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:F2")
    s = ws["A2"]
    s.value = "Daily Supplements: Multivitamin + Omega  •  Goal: Both every day"
    s.fill = hex_fill(LIGHT_ORANGE)
    s.font = make_font(italic=True, color=ORANGE, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, ORANGE, WHITE, range(1, 7))

    last_row = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value="FALSE")
        ws.cell(row=row, column=4, value="FALSE")
        ws.cell(row=row, column=5,
                value=f'=IF(AND(C{row}="TRUE",D{row}="TRUE"),"TRUE","FALSE")')
        ws.cell(row=row, column=6, value="")
        style_data_row(ws, row, bg, range(1, 7))
        ws.row_dimensions[row].height = 20

    add_dv_truefalse(ws, "C", 4, last_row)
    add_dv_truefalse(ws, "D", 4, last_row)

    ws.conditional_formatting.add(f"E4:E{last_row}",
        CellIsRule(operator="equal", formula=['"TRUE"'],
                   fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"E4:E{last_row}",
        CellIsRule(operator="equal", formula=['"FALSE"'],
                   fill=hex_fill(LIGHT_RED), font=Font(color=RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 5: STEPS TRACKER
# ─────────────────────────────────────────────

def create_steps_sheet(wb):
    ws = wb.create_sheet("Steps Tracker")

    headers = ["Date","Day","Steps","Goal","Goal Met","Difference","Notes"]
    col_widths = [13, 12, 12, 10, 11, 13, 25]

    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "👟  STEPS TRACKER 2026"
    t.fill = hex_fill(BLUE)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:G2")
    s = ws["A2"]
    s.value = "Daily Step Goal: 10,000 steps  •  Track and crush your daily movement!"
    s.fill = hex_fill(LIGHT_BLUE)
    s.font = make_font(italic=True, color=BLUE, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, BLUE, WHITE, range(1, 8))

    last_row = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value="")  # Steps to fill
        # Goal from Goals & Settings sheet
        ws.cell(row=row, column=4,
                value="='Goals & Settings'!B2").number_format = "#,##0"
        ws.cell(row=row, column=5,
                value=f'=IF(C{row}="","",IF(C{row}>=D{row},"TRUE","FALSE"))')
        ws.cell(row=row, column=6,
                value=f'=IF(C{row}="","",C{row}-D{row})').number_format = "+#,##0;-#,##0;0"
        ws.cell(row=row, column=7, value="")
        style_data_row(ws, row, bg, range(1, 8))
        ws.row_dimensions[row].height = 20

    ws.conditional_formatting.add(f"E4:E{last_row}",
        CellIsRule(operator="equal", formula=['"TRUE"'],
                   fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"E4:E{last_row}",
        CellIsRule(operator="equal", formula=['"FALSE"'],
                   fill=hex_fill(LIGHT_RED), font=Font(color=RED)))

    ws.conditional_formatting.add(f"F4:F{last_row}",
        CellIsRule(operator="greaterThanOrEqual", formula=["0"],
                   fill=hex_fill(LIGHT_GREEN)))
    ws.conditional_formatting.add(f"F4:F{last_row}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=hex_fill(LIGHT_RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:G{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 6: RUNNING TRACKER
# ─────────────────────────────────────────────

def create_running_sheet(wb):
    ws = wb.create_sheet("Running Tracker")

    headers = ["Date","Day","Ran Today","Distance (KM)",
               "Duration (min)","Pace (min/km)","Calories","Route/Location","Notes"]
    col_widths = [13, 12, 11, 14, 15, 15, 12, 20, 22]

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = "🏃  RUNNING TRACKER 2026"
    t.fill = hex_fill(GREEN)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    s = ws["A2"]
    s.value = "Log your runs  •  Pace = Duration ÷ Distance  •  Monthly KM Goal set in Goals & Settings"
    s.fill = hex_fill(LIGHT_GREEN)
    s.font = make_font(italic=True, color=GREEN, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, GREEN, WHITE, range(1, 10))

    last_row = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value="FALSE")
        ws.cell(row=row, column=4, value="")
        ws.cell(row=row, column=5, value="")
        ws.cell(row=row, column=6,
                value=f'=IFERROR(IF(AND(D{row}<>"",E{row}<>""),E{row}/D{row},""),"")')
        ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=9, value="")
        style_data_row(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 20

    add_dv_truefalse(ws, "C", 4, last_row)

    ws.conditional_formatting.add(f"C4:C{last_row}",
        CellIsRule(operator="equal", formula=['"TRUE"'],
                   fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 7: EXPENSE & SAVINGS
# ─────────────────────────────────────────────

def create_expense_sheet(wb):
    ws = wb.create_sheet("Expense & Savings")

    headers = ["Date","Day","Category","Description",
               "Income","Expense","Savings","Payment Method","Notes"]
    col_widths = [13, 12, 16, 28, 13, 13, 13, 16, 22]

    ws.merge_cells("A1:I1")
    t = ws["A1"]
    t.value = "💰  EXPENSE & SAVINGS TRACKER 2026"
    t.fill = hex_fill(PINK)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    s = ws["A2"]
    s.value = "Track income, expenses, and savings daily  •  Savings = Income - Expense"
    s.fill = hex_fill(LIGHT_PINK)
    s.font = make_font(italic=True, color=PINK, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, PINK, WHITE, range(1, 10))

    categories = ["Food","Transport","Shopping","Education","Health",
                  "Subscription","Family","Personal","Other"]
    payment_methods = ["Cash","Card","Online Transfer","Mobile Pay"]

    dv_cat = DataValidation(type="list",
                             formula1='"Food,Transport,Shopping,Education,Health,Subscription,Family,Personal,Other"',
                             allow_blank=True)
    dv_pay = DataValidation(type="list",
                             formula1='"Cash,Card,Online Transfer,Mobile Pay"',
                             allow_blank=True)
    ws.add_data_validation(dv_cat)
    ws.add_data_validation(dv_pay)

    last_row = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value="")
        ws.cell(row=row, column=4, value="")
        ws.cell(row=row, column=5, value="").number_format = "#,##0.00"
        ws.cell(row=row, column=6, value="").number_format = "#,##0.00"
        ws.cell(row=row, column=7,
                value=f"=IFERROR(IF(AND(E{row}<>\"\",F{row}<>\"\"),E{row}-F{row},IF(E{row}<>\"\",E{row},IF(F{row}<>\"\",-F{row},\"\"))),"
                      f"\"\")").number_format = "#,##0.00"
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=9, value="")
        style_data_row(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 20
        # Number formatting for income/expense/savings columns
        for c in [5, 6, 7]:
            ws.cell(row=row, column=c).number_format = "#,##0.00"

    dv_cat.sqref = f"C4:C{last_row}"
    dv_pay.sqref = f"H4:H{last_row}"

    # Positive savings = green, negative = red
    ws.conditional_formatting.add(f"G4:G{last_row}",
        CellIsRule(operator="greaterThan", formula=["0"],
                   fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN)))
    ws.conditional_formatting.add(f"G4:G{last_row}",
        CellIsRule(operator="lessThan", formula=["0"],
                   fill=hex_fill(LIGHT_RED), font=Font(color=RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last_row}"
    return ws

# ─────────────────────────────────────────────
# SHEET 8: GOALS & SETTINGS
# ─────────────────────────────────────────────

def create_goals_sheet(wb):
    ws = wb.create_sheet("Goals & Settings")

    ws.merge_cells("A1:D1")
    t = ws["A1"]
    t.value = "⚙️  GOALS & SETTINGS"
    t.fill = hex_fill(DARK_NAVY)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:D2")
    s = ws["A2"]
    s.value = "Edit the goal values below  •  All tracking sheets reference these cells automatically"
    s.fill = hex_fill(LIGHT_BLUE)
    s.font = make_font(italic=True, color=BLUE, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    # Section: Daily Goals
    ws.merge_cells("A3:D3")
    sec = ws["A3"]
    sec.value = "DAILY GOALS"
    sec.fill = hex_fill(TEAL)
    sec.font = make_font(bold=True, color=WHITE, size=11)
    sec.alignment = center()
    ws.row_dimensions[3].height = 24

    daily_goals = [
        ("Daily Steps Goal",            10000,  "B2",  "steps/day"),
        ("Daily Salat Goal",            5,      "B5",  "prayers/day"),
        ("Daily Supplements Goal",      2,      "B6",  "supplements/day"),
    ]

    settings_data = [
        ("Setting",         "Value",    "Cell Ref",     "Unit"),
        ("Daily Steps Goal",10000,      "",             "steps/day"),
        ("Daily Salat Goal",5,          "",             "prayers/day"),
        ("Daily Supplements Goal",2,    "",             "supplements/day"),
    ]

    col_widths = [30, 18, 15, 20]
    for c, w in enumerate(col_widths, 1):
        set_col_width(ws, c, w)

    header_row = 4
    ws.row_dimensions[header_row].height = 26
    for c, h in enumerate(["Setting", "Value", "Unit", "Notes"], 1):
        cell = ws.cell(row=header_row, column=c, value=h)
        cell.fill = hex_fill(DARK_NAVY)
        cell.font = make_font(bold=True, color=WHITE, size=10)
        cell.alignment = center()
        cell.border = thin_border()

    rows_data = [
        ("Daily Steps Goal",       10000, "steps/day",     "Used in Steps Tracker column D"),
        ("Daily Salat Goal",       5,     "prayers/day",   "Max is 5 prayers"),
        ("Daily Supplements Goal", 2,     "supplements/day","Multivitamin + Omega = 2"),
        ("Weekly Gym Days Goal",   5,     "days/week",     "Mon/Tue/Wed/Fri/Sat = Grind Days"),
    ]

    for i, (setting, val, unit, note) in enumerate(rows_data):
        r = 5 + i
        ws.row_dimensions[r].height = 24
        bg = WHITE if i % 2 == 0 else GRAY_BG

        c1 = ws.cell(row=r, column=1, value=setting)
        c1.fill = hex_fill(bg); c1.font = make_font(bold=True, size=10); c1.border = thin_border(); c1.alignment = left()

        c2 = ws.cell(row=r, column=2, value=val)
        c2.fill = hex_fill(LIGHT_TEAL); c2.font = make_font(bold=True, color=TEAL, size=11)
        c2.border = thin_border(); c2.alignment = center()

        c3 = ws.cell(row=r, column=3, value=unit)
        c3.fill = hex_fill(bg); c3.font = make_font(size=10, italic=True); c3.border = thin_border(); c3.alignment = center()

        c4 = ws.cell(row=r, column=4, value=note)
        c4.fill = hex_fill(bg); c4.font = make_font(size=9, color=DARK_GRAY, italic=True); c4.border = thin_border(); c4.alignment = left()

    # Monthly Goals section
    monthly_start = 10
    ws.merge_cells(f"A{monthly_start}:D{monthly_start}")
    ms = ws.cell(row=monthly_start, column=1, value="MONTHLY GOALS (EDITABLE)")
    ms.fill = hex_fill(ORANGE)
    ms.font = make_font(bold=True, color=WHITE, size=11)
    ms.alignment = center()
    ws.row_dimensions[monthly_start].height = 24

    monthly_goals = [
        ("Monthly Running KM Goal",     0,  "km/month",   "Set your target running distance"),
        ("Monthly Savings Goal",        0,  "currency",   "Set your monthly savings target"),
    ]

    for i, (s, v, u, n) in enumerate(monthly_goals):
        r = monthly_start + 1 + i
        ws.row_dimensions[r].height = 24
        bg = WHITE if i % 2 == 0 else GRAY_BG

        c1 = ws.cell(row=r, column=1, value=s)
        c1.fill = hex_fill(bg); c1.font = make_font(bold=True, size=10); c1.border = thin_border(); c1.alignment = left()
        c2 = ws.cell(row=r, column=2, value=v)
        c2.fill = hex_fill(LIGHT_ORANGE); c2.font = make_font(bold=True, color=ORANGE, size=11)
        c2.border = thin_border(); c2.alignment = center()
        c3 = ws.cell(row=r, column=3, value=u)
        c3.fill = hex_fill(bg); c3.font = make_font(size=10, italic=True); c3.border = thin_border(); c3.alignment = center()
        c4 = ws.cell(row=r, column=4, value=n)
        c4.fill = hex_fill(bg); c4.font = make_font(size=9, color=DARK_GRAY, italic=True); c4.border = thin_border(); c4.alignment = left()

    # Date range
    date_start = monthly_start + len(monthly_goals) + 2
    ws.merge_cells(f"A{date_start}:D{date_start}")
    ds = ws.cell(row=date_start, column=1, value="TRACKER DATE RANGE")
    ds.fill = hex_fill(DARK_NAVY)
    ds.font = make_font(bold=True, color=WHITE, size=11)
    ds.alignment = center()
    ws.row_dimensions[date_start].height = 24

    date_info = [
        ("Start Date", date(2026, 1, 1),  "DD-MMM-YYYY", "First tracked date"),
        ("End Date",   date(2026, 12, 31), "DD-MMM-YYYY", "Last tracked date"),
    ]

    for i, (label, val, fmt, note) in enumerate(date_info):
        r = date_start + 1 + i
        ws.row_dimensions[r].height = 24
        bg = WHITE if i % 2 == 0 else GRAY_BG

        c1 = ws.cell(row=r, column=1, value=label)
        c1.fill = hex_fill(bg); c1.font = make_font(bold=True, size=10); c1.border = thin_border(); c1.alignment = left()
        c2 = ws.cell(row=r, column=2, value=val)
        c2.number_format = fmt
        c2.fill = hex_fill(LIGHT_BLUE); c2.font = make_font(bold=True, color=BLUE, size=11)
        c2.border = thin_border(); c2.alignment = center()
        c3 = ws.cell(row=r, column=3, value=fmt)
        c3.fill = hex_fill(bg); c3.font = make_font(size=9, italic=True); c3.border = thin_border(); c3.alignment = center()
        c4 = ws.cell(row=r, column=4, value=note)
        c4.fill = hex_fill(bg); c4.font = make_font(size=9, color=DARK_GRAY, italic=True); c4.border = thin_border(); c4.alignment = left()

    ws.freeze_panes = "A5"
    return ws

# ─────────────────────────────────────────────
# SHEET 9: MONTHLY SUMMARY
# ─────────────────────────────────────────────

def create_monthly_summary_sheet(wb):
    ws = wb.create_sheet("Monthly Summary")

    headers = ["Month","Salat %","Gym %","Suppl. %","Steps Goal %",
               "Running KM","Total Income","Total Expense","Total Savings",
               "Savings Rate %","Overall Score %"]
    col_widths = [14, 12, 10, 12, 14, 14, 15, 15, 15, 15, 16]

    ws.merge_cells("A1:K1")
    t = ws["A1"]
    t.value = "📊  MONTHLY SUMMARY 2026"
    t.fill = hex_fill(DARK_NAVY)
    t.font = make_font(bold=True, size=14, color=WHITE)
    t.alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:K2")
    s = ws["A2"]
    s.value = (
        "Overall Score = Salat 25% + Gym 20% + Supplements 15% + "
        "Steps 15% + Running 10% + Savings 15%"
    )
    s.fill = hex_fill(LIGHT_BLUE)
    s.font = make_font(italic=True, color=BLUE, size=10)
    s.alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_col_width(ws, c, w)
    style_header_row(ws, 3, DARK_NAVY, WHITE, range(1, 12))

    # Salat sheet rows: 4 to 368 (row = date index + 4)
    # Steps sheet: same structure
    # For monthly SUMIF we need to reference data sheets

    month_ranges = []
    for m in range(1, 13):
        days_in_m = calendar.monthrange(2026, m)[1]
        # Find first and last row index for this month
        first_day_idx = (date(2026, m, 1) - date(2026, 1, 1)).days
        last_day_idx  = first_day_idx + days_in_m - 1
        first_row = first_day_idx + 4
        last_row  = last_day_idx  + 4
        month_ranges.append((m, first_row, last_row, days_in_m))

    # Grind days per month
    grind_per_month = []
    for m in range(1, 13):
        days_in_m = calendar.monthrange(2026, m)[1]
        count = 0
        for d in range(1, days_in_m + 1):
            wd = date(2026, m, d).strftime("%A")
            if DAY_TYPE[wd] == "Grind Day":
                count += 1
        grind_per_month.append(count)

    for i, (m, fr, lr, days_in_m) in enumerate(month_ranges):
        row = 4 + i
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.row_dimensions[row].height = 22

        # Month name
        ws.cell(row=row, column=1, value=MONTHS[m-1])

        # Salat %: COUNTIF(I:I,"TRUE") / days * 5... actually use completion avg
        # Salat completion % col I, rows fr:lr
        ws.cell(row=row, column=2,
                value=f"=IFERROR(AVERAGEIF('Salat Tracker'!I{fr}:I{lr},\">0\")*1,0)"
                ).number_format = "0.0%"
        # Simpler: average of column I (Completion %)
        ws.cell(row=row, column=2,
                value=f"=IFERROR(AVERAGE('Salat Tracker'!I{fr}:I{lr}),0)"
                ).number_format = "0.0%"

        # Gym %: count TRUE in D where C = "Grind Day" / grind_days
        grind_days = grind_per_month[i]
        ws.cell(row=row, column=3,
                value=f'=IFERROR(COUNTIFS(\'Gym Routine\'!C{fr}:C{lr},"Grind Day",\'Gym Routine\'!D{fr}:D{lr},"TRUE")/{grind_days},0)'
                ).number_format = "0.0%"

        # Supplements %: COUNTIF(E,"TRUE") / days
        ws.cell(row=row, column=4,
                value=f'=IFERROR(COUNTIF(\'Supplements\'!E{fr}:E{lr},"TRUE")/{days_in_m},0)'
                ).number_format = "0.0%"

        # Steps Goal %: COUNTIF(E,"TRUE") / days
        ws.cell(row=row, column=5,
                value=f'=IFERROR(COUNTIF(\'Steps Tracker\'!E{fr}:E{lr},"TRUE")/{days_in_m},0)'
                ).number_format = "0.0%"

        # Running KM: SUM of D col (Distance)
        ws.cell(row=row, column=6,
                value=f"=IFERROR(SUMIF('Running Tracker'!C{fr}:C{lr},\"TRUE\",'Running Tracker'!D{fr}:D{lr}),0)"
                ).number_format = "0.0"

        # Expense sheet: Income, Expense, Savings
        ws.cell(row=row, column=7,
                value=f"=IFERROR(SUM('Expense & Savings'!E{fr}:E{lr}),0)"
                ).number_format = "#,##0.00"
        ws.cell(row=row, column=8,
                value=f"=IFERROR(SUM('Expense & Savings'!F{fr}:F{lr}),0)"
                ).number_format = "#,##0.00"
        ws.cell(row=row, column=9,
                value=f"=IFERROR(G{row}-H{row},0)"
                ).number_format = "#,##0.00"

        # Savings Rate %
        ws.cell(row=row, column=10,
                value=f"=IFERROR(I{row}/G{row},0)"
                ).number_format = "0.0%"

        # Overall Score = Salat*0.25 + Gym*0.20 + Suppl*0.15 + Steps*0.15 + Running_score*0.10 + Savings_score*0.15
        # Running score: MIN(F_row / monthly goal, 1) — goal from Goals sheet B11
        # Savings score: MIN(I_row / monthly savings goal, 1) — goal from Goals sheet B12
        ws.cell(row=row, column=11,
                value=f"=B{row}*0.25+C{row}*0.20+D{row}*0.15+E{row}*0.15"
                      f"+MIN(IFERROR(F{row}/'Goals & Settings'!B11,0),1)*0.10"
                      f"+MIN(IFERROR(I{row}/'Goals & Settings'!B12,0),1)*0.15"
                ).number_format = "0.0%"

        style_data_row(ws, row, bg, range(1, 12))
        # Set number formats again (style_data_row overwrites)
        for c in [2,3,4,5,10,11]:
            ws.cell(row=row, column=c).number_format = "0.0%"
        ws.cell(row=row, column=6).number_format = "0.0"
        for c in [7,8,9]:
            ws.cell(row=row, column=c).number_format = "#,##0.00"

    # Annual totals row
    total_row = 4 + 12
    ws.row_dimensions[total_row].height = 26
    ws.cell(row=total_row, column=1, value="YEAR 2026 TOTAL/AVG")
    ws.cell(row=total_row, column=2, value=f"=AVERAGE(B4:B15)").number_format = "0.0%"
    ws.cell(row=total_row, column=3, value=f"=AVERAGE(C4:C15)").number_format = "0.0%"
    ws.cell(row=total_row, column=4, value=f"=AVERAGE(D4:D15)").number_format = "0.0%"
    ws.cell(row=total_row, column=5, value=f"=AVERAGE(E4:E15)").number_format = "0.0%"
    ws.cell(row=total_row, column=6, value=f"=SUM(F4:F15)").number_format = "0.0"
    ws.cell(row=total_row, column=7, value=f"=SUM(G4:G15)").number_format = "#,##0.00"
    ws.cell(row=total_row, column=8, value=f"=SUM(H4:H15)").number_format = "#,##0.00"
    ws.cell(row=total_row, column=9, value=f"=SUM(I4:I15)").number_format = "#,##0.00"
    ws.cell(row=total_row, column=10, value=f"=IFERROR(I{total_row}/G{total_row},0)").number_format = "0.0%"
    ws.cell(row=total_row, column=11, value=f"=AVERAGE(K4:K15)").number_format = "0.0%"
    style_header_row(ws, total_row, DARK_NAVY, WHITE, range(1, 12))

    # Conditional formatting on Overall Score
    ws.conditional_formatting.add(f"K4:K15",
        CellIsRule(operator="greaterThanOrEqual", formula=["0.8"],
                   fill=hex_fill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"K4:K15",
        CellIsRule(operator="between", formula=["0.6","0.8"],
                   fill=hex_fill(LIGHT_YELLOW), font=Font(color=YELLOW, bold=True)))
    ws.conditional_formatting.add(f"K4:K15",
        CellIsRule(operator="lessThan", formula=["0.6"],
                   fill=hex_fill(LIGHT_RED), font=Font(color=RED, bold=True)))

    ws.freeze_panes = "A4"

    # ── Bar Chart: Monthly Performance ──
    chart = BarChart()
    chart.type   = "col"
    chart.title  = "Monthly Performance Overview 2026"
    chart.y_axis.title = "Score (%)"
    chart.x_axis.title = "Month"
    chart.style  = 10
    chart.height = 14
    chart.width  = 24

    cats = Reference(ws, min_col=1, min_row=4, max_row=15)

    for col, label, color in [
        (2, "Salat",        "00897B"),
        (3, "Gym",          "5C35A1"),
        (4, "Supplements",  "F57C00"),
        (5, "Steps",        "1565C0"),
        (11,"Overall Score","1B2A4A"),
    ]:
        data = Reference(ws, min_col=col, min_row=3, max_row=15)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)

    ws.add_chart(chart, "A18")
    return ws

# ─────────────────────────────────────────────
# SHEET 1: DASHBOARD
# ─────────────────────────────────────────────

def create_dashboard(wb):
    ws = wb.active
    ws.title = "Dashboard"

    # Set column widths
    col_widths = [2, 14, 14, 14, 14, 14, 14, 14, 2]
    for c, w in enumerate(col_widths, 1):
        set_col_width(ws, c, w)

    for r in range(1, 60):
        ws.row_dimensions[r].height = 20

    # ── TITLE BANNER ──
    ws.merge_cells("B1:H1")
    title = ws["B1"]
    title.value = "PERSONAL LIFE TRACKER 2026"
    title.fill = hex_fill(DARK_NAVY)
    title.font = make_font(bold=True, size=18, color=WHITE)
    title.alignment = center()
    ws.row_dimensions[1].height = 42

    ws.merge_cells("B2:H2")
    sub = ws["B2"]
    sub.value = "Your daily companion for habits, health, and wealth  •  Stay consistent, track everything!"
    sub.fill = hex_fill(TEAL)
    sub.font = make_font(italic=True, color=WHITE, size=10)
    sub.alignment = center()
    ws.row_dimensions[2].height = 24

    # ── TODAY'S DATE ──
    ws.row_dimensions[3].height = 8  # spacer

    ws.merge_cells("B4:H4")
    date_label = ws["B4"]
    date_label.value = '=TODAY()'
    date_label.number_format = '"Today: "DDDD, DD MMMM YYYY'
    date_label.fill = hex_fill(LIGHT_TEAL)
    date_label.font = make_font(bold=True, size=12, color=TEAL)
    date_label.alignment = center()
    ws.row_dimensions[4].height = 28

    # ── KPI CARDS - ROW 1 ──
    ws.row_dimensions[5].height = 8  # spacer

    ws.merge_cells("B6:H6")
    kpi_hdr = ws["B6"]
    kpi_hdr.value = "CURRENT YEAR PERFORMANCE  —  KEY METRICS"
    kpi_hdr.fill = hex_fill(DARK_NAVY)
    kpi_hdr.font = make_font(bold=True, color=WHITE, size=11)
    kpi_hdr.alignment = center()
    ws.row_dimensions[6].height = 26

    # KPI card data: (merge cols, label, formula, fill, font_color)
    kpi_cards = [
        # Row 7-9: card labels
        ("B7:C7", "🕌 SALAT",      TEAL,   WHITE),
        ("D7:E7", "💪 GYM",        PURPLE, WHITE),
        ("F7:G7", "💊 SUPPLEMENTS",ORANGE, WHITE),
        ("H7:H7", "👟 STEPS",      BLUE,   WHITE),
    ]
    kpi_values = [
        ("B8:C8", '=IFERROR(COUNTIF(\'Salat Tracker\'!I$4:I$368,"=1")/COUNTA(\'Salat Tracker\'!A$4:A$368),0)',  "0.0%", LIGHT_TEAL,   TEAL),
        ("D8:E8", '=IFERROR(COUNTIFS(\'Gym Routine\'!C$4:C$368,"Grind Day",\'Gym Routine\'!D$4:D$368,"TRUE")/COUNTIF(\'Gym Routine\'!C$4:C$368,"Grind Day"),0)', "0.0%", LIGHT_PURPLE, PURPLE),
        ("F8:G8", '=IFERROR(COUNTIF(\'Supplements\'!E$4:E$368,"TRUE")/COUNTA(\'Supplements\'!A$4:A$368),0)', "0.0%", LIGHT_ORANGE,  ORANGE),
        ("H8:H8", '=IFERROR(COUNTIF(\'Steps Tracker\'!E$4:E$368,"TRUE")/COUNTA(\'Steps Tracker\'!A$4:A$368),0)', "0.0%", LIGHT_BLUE,   BLUE),
    ]
    kpi_subs = [
        ("B9:C9", "Prayers 100%",  LIGHT_TEAL,   DARK_GRAY),
        ("D9:E9", "Workouts Done", LIGHT_PURPLE, DARK_GRAY),
        ("F9:G9", "Both Taken",    LIGHT_ORANGE,  DARK_GRAY),
        ("H9:H9", "Goal Met Days", LIGHT_BLUE,   DARK_GRAY),
    ]

    for merge_range, label, fill_c, font_c in kpi_cards:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = label
        c.fill = hex_fill(fill_c)
        c.font = make_font(bold=True, color=font_c, size=10)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[7].height = 26

    for merge_range, formula, fmt, fill_c, font_c in kpi_values:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = formula
        c.number_format = fmt
        c.fill = hex_fill(fill_c)
        c.font = make_font(bold=True, color=font_c, size=16)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[8].height = 36

    for merge_range, label, fill_c, font_c in kpi_subs:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = label
        c.fill = hex_fill(fill_c)
        c.font = make_font(italic=True, color=font_c, size=9)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[9].height = 20

    # ── KPI ROW 2: Running / Expenses / Savings / Overall ──
    ws.row_dimensions[10].height = 8

    kpi2_cards = [
        ("B11:C11", "🏃 RUNNING",       GREEN, WHITE),
        ("D11:E11", "💰 INCOME YTD",    PINK,  WHITE),
        ("F11:G11", "📉 EXPENSE YTD",   RED,   WHITE),
        ("H11:H11", "📈 SAVINGS YTD",   TEAL,  WHITE),
    ]
    kpi2_values = [
        ("B12:C12", "=IFERROR(SUM('Running Tracker'!D$4:D$368),0)",                      "0.0 km",   LIGHT_GREEN, GREEN),
        ("D12:E12", "=IFERROR(SUM('Expense & Savings'!E$4:E$368),0)",                    "#,##0.00", LIGHT_PINK,  PINK),
        ("F12:G12", "=IFERROR(SUM('Expense & Savings'!F$4:F$368),0)",                    "#,##0.00", LIGHT_RED,   RED),
        ("H12:H12", "=IFERROR(SUM('Expense & Savings'!E$4:E$368)-SUM('Expense & Savings'!F$4:F$368),0)", "#,##0.00", LIGHT_TEAL, TEAL),
    ]
    kpi2_subs = [
        ("B13:C13", "Total KM Run",      LIGHT_GREEN, DARK_GRAY),
        ("D13:E13", "All Income",        LIGHT_PINK,  DARK_GRAY),
        ("F13:G13", "All Expenses",      LIGHT_RED,   DARK_GRAY),
        ("H13:H13", "Net Savings",       LIGHT_TEAL,  DARK_GRAY),
    ]

    for merge_range, label, fill_c, font_c in kpi2_cards:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = label
        c.fill = hex_fill(fill_c)
        c.font = make_font(bold=True, color=font_c, size=10)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[11].height = 26

    for merge_range, formula, fmt, fill_c, font_c in kpi2_values:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = formula
        c.number_format = fmt
        c.fill = hex_fill(fill_c)
        c.font = make_font(bold=True, color=font_c, size=16)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[12].height = 36

    for merge_range, label, fill_c, font_c in kpi2_subs:
        ws.merge_cells(merge_range)
        start_cell = merge_range.split(":")[0]
        c = ws[start_cell]
        c.value = label
        c.fill = hex_fill(fill_c)
        c.font = make_font(italic=True, color=font_c, size=9)
        c.alignment = center()
        c.border = thin_border()
    ws.row_dimensions[13].height = 20

    # ── OVERALL LIFE SCORE ──
    ws.row_dimensions[14].height = 8
    ws.merge_cells("B15:H15")
    overall_label = ws["B15"]
    overall_label.value = "OVERALL LIFE SCORE 2026"
    overall_label.fill = hex_fill(DARK_NAVY)
    overall_label.font = make_font(bold=True, color=WHITE, size=11)
    overall_label.alignment = center()
    ws.row_dimensions[15].height = 26

    ws.merge_cells("B16:H17")
    overall_val = ws["B16"]
    overall_val.value = "='Monthly Summary'!K16"
    overall_val.number_format = "0.0%"
    overall_val.fill = hex_fill(DARK_NAVY)
    overall_val.font = make_font(bold=True, color=WHITE, size=28)
    overall_val.alignment = center()
    ws.row_dimensions[16].height = 38
    ws.row_dimensions[17].height = 20

    ws.merge_cells("B18:H18")
    score_sub = ws["B18"]
    score_sub.value = "Salat 25%  +  Gym 20%  +  Supplements 15%  +  Steps 15%  +  Running 10%  +  Savings 15%"
    score_sub.fill = hex_fill(LIGHT_BLUE)
    score_sub.font = make_font(italic=True, color=BLUE, size=9)
    score_sub.alignment = center()
    ws.row_dimensions[18].height = 20

    # ── MONTHLY MINI TABLE ──
    ws.row_dimensions[19].height = 8

    ws.merge_cells("B20:H20")
    mt_hdr = ws["B20"]
    mt_hdr.value = "MONTHLY PERFORMANCE AT A GLANCE"
    mt_hdr.fill = hex_fill(TEAL)
    mt_hdr.font = make_font(bold=True, color=WHITE, size=11)
    mt_hdr.alignment = center()
    ws.row_dimensions[20].height = 26

    mini_headers = ["Month", "Salat %", "Gym %", "Suppl. %", "Steps %", "Running", "Overall"]
    mini_cols = ["B","C","D","E","F","G","H"]
    ws.row_dimensions[21].height = 24
    for col_l, h in zip(mini_cols, mini_headers):
        c = ws[f"{col_l}21"]
        c.value = h
        c.fill = hex_fill(DARK_NAVY)
        c.font = make_font(bold=True, color=WHITE, size=9)
        c.alignment = center()
        c.border = thin_border()

    for i, month in enumerate(MONTHS):
        r = 22 + i
        src_row = 4 + i  # Monthly Summary rows start at 4
        bg = WHITE if i % 2 == 0 else GRAY_BG
        ws.row_dimensions[r].height = 20

        cells_data = [
            ("B", month, None),
            ("C", f"='Monthly Summary'!B{src_row}", "0%"),
            ("D", f"='Monthly Summary'!C{src_row}", "0%"),
            ("E", f"='Monthly Summary'!D{src_row}", "0%"),
            ("F", f"='Monthly Summary'!E{src_row}", "0%"),
            ("G", f"='Monthly Summary'!F{src_row}", "0.0"),
            ("H", f"='Monthly Summary'!K{src_row}", "0%"),
        ]
        for col_l, val, fmt in cells_data:
            c = ws[f"{col_l}{r}"]
            c.value = val
            if fmt:
                c.number_format = fmt
            c.fill = hex_fill(bg)
            c.font = make_font(size=9)
            c.alignment = center()
            c.border = thin_border()

    # Color-scale on Overall column
    ws.conditional_formatting.add(f"H22:H33",
        ColorScaleRule(start_type='min', start_color='E53935',
                       mid_type='percentile', mid_value=50, mid_color='F9A825',
                       end_type='max', end_color='2E7D32'))

    # No freeze on dashboard, keep it scrollable
    ws.sheet_view.showGridLines = False
    return ws

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────

def main():
    wb = Workbook()

    # Create sheets in order (Dashboard will be ws.active = first)
    create_dashboard(wb)
    create_salat_sheet(wb)
    create_gym_sheet(wb)
    create_supplements_sheet(wb)
    create_steps_sheet(wb)
    create_running_sheet(wb)
    create_expense_sheet(wb)
    create_goals_sheet(wb)
    create_monthly_summary_sheet(wb)

    # Tab colors
    tab_colors = {
        "Dashboard":        "1B2A4A",
        "Salat Tracker":    "00897B",
        "Gym Routine":      "5C35A1",
        "Supplements":      "F57C00",
        "Steps Tracker":    "1565C0",
        "Running Tracker":  "2E7D32",
        "Expense & Savings":"AD1457",
        "Goals & Settings": "1B2A4A",
        "Monthly Summary":  "F9A825",
    }
    for sheet in wb.sheetnames:
        if sheet in tab_colors:
            wb[sheet].sheet_properties.tabColor = tab_colors[sheet]

    out_path = "/home/user/habit-tracker/Personal_Life_Tracker.xlsx"
    wb.save(out_path)
    print(f"✅  Workbook saved → {out_path}")
    print(f"   Sheets: {wb.sheetnames}")
    print(f"   Total dates tracked: {len(DATES)} days (2026-01-01 to 2026-12-31)")

if __name__ == "__main__":
    main()
