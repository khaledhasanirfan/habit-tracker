"""
Personal Life Tracker 2026 — v2
Checkboxes: ☑ / ☐ dropdowns
Auto-red: past uncompleted habits turn red automatically via TODAY() formula rules
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date, timedelta
import calendar

# ── Symbols ────────────────────────────────────────────────────────────────────
CHECKED   = "☑"   # done / prayed / taken / goal met
UNCHECKED = "☐"   # not done

# ── Colors ─────────────────────────────────────────────────────────────────────
DARK_NAVY    = "1B2A4A"
TEAL         = "00897B"
LIGHT_TEAL   = "E0F2F1"
PURPLE       = "5C35A1"
LIGHT_PURPLE = "EDE7F6"
ORANGE       = "F57C00"
LIGHT_ORANGE = "FFF3E0"
RED          = "E53935"
LIGHT_RED    = "FFEBEE"
GREEN        = "2E7D32"
LIGHT_GREEN  = "E8F5E9"
YELLOW       = "F9A825"
LIGHT_YELLOW = "FFFDE7"
GRAY_BG      = "F5F5F5"
WHITE        = "FFFFFF"
DARK_GRAY    = "757575"
BLUE         = "1565C0"
LIGHT_BLUE   = "E3F2FD"
PINK         = "AD1457"
LIGHT_PINK   = "FCE4EC"

# ── Style helpers ──────────────────────────────────────────────────────────────
def hfill(hex_c):
    return PatternFill("solid", fgColor=hex_c)

def mfont(bold=False, size=11, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")

def tborder():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def center():
    return Alignment(horizontal="center", vertical="center", wrap_text=True)

def left_align():
    return Alignment(horizontal="left", vertical="center", wrap_text=True)

def set_w(ws, col, width):
    ws.column_dimensions[get_column_letter(col)].width = width

def style_header(ws, row, bg, fg=WHITE, cols=None):
    cols = cols or range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill  = hfill(bg)
        cell.font  = mfont(bold=True, color=fg, size=10)
        cell.alignment = center()
        cell.border    = tborder()

def style_data(ws, row, bg=WHITE, cols=None):
    cols = cols or range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill  = hfill(bg)
        cell.border = tborder()
        cell.alignment = center()
        cell.font = mfont(size=10)

def add_dv_checkbox(ws, col_letter, start, end):
    """☑ / ☐ dropdown."""
    dv = DataValidation(
        type="list",
        formula1=f'"{CHECKED},{UNCHECKED}"',
        allow_blank=True,
        showDropDown=False
    )
    dv.sqref = f"{col_letter}{start}:{col_letter}{end}"
    ws.add_data_validation(dv)
    return dv

# ── Date data ──────────────────────────────────────────────────────────────────
def all_dates():
    s = date(2026, 1, 1)
    return [s + timedelta(days=i) for i in range(365)]

DATES = all_dates()

DAY_TYPE = {
    "Monday":    "Grind Day",
    "Tuesday":   "Grind Day",
    "Wednesday": "Grind Day",
    "Thursday":  "Rest Day",
    "Friday":    "Grind Day",
    "Saturday":  "Grind Day",
    "Sunday":    "Rest Day",
}

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

# ── Conditional formatting rules ───────────────────────────────────────────────

def cf_checked_green(ws, rng):
    """☑ → green (any date)."""
    ws.conditional_formatting.add(rng,
        FormulaRule(
            formula=[f'LEFT({rng.split(":")[0]},1)="{CHECKED}"'],
            fill=hfill(LIGHT_GREEN),
            font=Font(color=GREEN, bold=True, name="Calibri")
        )
    )

def cf_unchecked_past_red(ws, rng, date_col="$A"):
    """☐ + date in past → RED."""
    first = rng.split(":")[0]
    row   = ''.join(filter(str.isdigit, first))
    ws.conditional_formatting.add(rng,
        FormulaRule(
            formula=[f'AND({date_col}{row}<TODAY(),LEFT({first},1)="{UNCHECKED}")'],
            fill=hfill(LIGHT_RED),
            font=Font(color=RED, bold=True, name="Calibri")
        )
    )

def cf_unchecked_past_orange(ws, rng, date_col="$A"):
    """☐ + date in past → soft orange (optional habits)."""
    first = rng.split(":")[0]
    row   = ''.join(filter(str.isdigit, first))
    ws.conditional_formatting.add(rng,
        FormulaRule(
            formula=[f'AND({date_col}{row}<TODAY(),LEFT({first},1)="{UNCHECKED}")'],
            fill=hfill(LIGHT_ORANGE),
            font=Font(color=ORANGE, bold=True, name="Calibri")
        )
    )

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — SALAT TRACKER
# ══════════════════════════════════════════════════════════════════════════════

def create_salat_sheet(wb):
    ws = wb.create_sheet("Salat Tracker")
    headers    = ["Date","Day","Fajr","Dhuhr","Asr","Maghrib","Isha",
                  "Total Prayers","Completion %","Notes"]
    col_widths = [13,12,9,9,9,9,9,14,14,25]

    ws.merge_cells("A1:J1")
    ws["A1"].value     = "🕌  SALAT TRACKER 2026"
    ws["A1"].fill      = hfill(DARK_NAVY)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:J2")
    ws["A2"].value     = f"Click cell → pick {CHECKED} (prayed) or {UNCHECKED} (not prayed)   •   Past unprayed cells turn RED automatically"
    ws["A2"].fill      = hfill(TEAL)
    ws["A2"].font      = mfont(italic=True, color=WHITE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, TEAL, WHITE, range(1, 11))

    last = 3 + len(DATES)
    prayer_cols = ["C","D","E","F","G"]

    for i, d in enumerate(DATES):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        for c in range(3, 8):
            ws.cell(row=row, column=c, value=UNCHECKED)
        pr = f"C{row}:G{row}"
        ws.cell(row=row, column=8,
                value=f'=COUNTIF({pr},"{CHECKED}")')
        ws.cell(row=row, column=9,
                value=f"=H{row}/5").number_format = "0%"
        ws.cell(row=row, column=10, value="")
        style_data(ws, row, bg, range(1, 11))
        ws.row_dimensions[row].height = 22

    for ltr in prayer_cols:
        add_dv_checkbox(ws, ltr, 4, last)

    # Checkbox colour rules (applied per column range so row references stay correct)
    prayer_range = f"C4:G{last}"
    # 1. ☑ → green
    ws.conditional_formatting.add(prayer_range,
        FormulaRule(
            formula=[f'C4="{CHECKED}"'],
            fill=hfill(LIGHT_GREEN),
            font=Font(color=GREEN, bold=True, name="Calibri")
        )
    )
    # 2. ☐ + past date → RED
    ws.conditional_formatting.add(prayer_range,
        FormulaRule(
            formula=[f'AND($A4<TODAY(),C4="{UNCHECKED}")'],
            fill=hfill(LIGHT_RED),
            font=Font(color=RED, bold=True, name="Calibri")
        )
    )

    # Completion % colours
    comp = f"I4:I{last}"
    ws.conditional_formatting.add(comp,
        CellIsRule("equal",   ["1"],    fill=hfill(LIGHT_GREEN),  font=Font(color=GREEN,  bold=True)))
    ws.conditional_formatting.add(comp,
        CellIsRule("greaterThanOrEqual", ["0.6"], fill=hfill(LIGHT_YELLOW), font=Font(color=YELLOW, bold=True)))
    ws.conditional_formatting.add(comp,
        CellIsRule("lessThan",["0.6"], fill=hfill(LIGHT_RED),   font=Font(color=RED,    bold=True)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:J{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — GYM ROUTINE
# ══════════════════════════════════════════════════════════════════════════════

def create_gym_sheet(wb):
    ws = wb.create_sheet("Gym Routine")
    headers    = ["Date","Day","Planned Type","Workout Done",
                  "Workout Name","Body Part","Duration (min)","Intensity","Notes"]
    col_widths = [13,12,13,14,22,16,15,13,25]

    ws.merge_cells("A1:I1")
    ws["A1"].value     = "💪  GYM ROUTINE TRACKER 2026"
    ws["A1"].fill      = hfill(PURPLE)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    ws["A2"].value     = f"Mon/Tue/Wed/Fri/Sat = Grind Day  •  Thu/Sun = Rest Day  •  {CHECKED}=Done  {UNCHECKED}=Missed  •  Past missed Grind Days → RED"
    ws["A2"].fill      = hfill(LIGHT_PURPLE)
    ws["A2"].font      = mfont(italic=True, color=PURPLE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, PURPLE, WHITE, range(1, 10))

    dv_done = DataValidation(type="list",
                              formula1=f'"{CHECKED},{UNCHECKED}"',
                              allow_blank=True)
    dv_int  = DataValidation(type="list",
                              formula1='"Low,Medium,High,Very High"',
                              allow_blank=True)
    ws.add_data_validation(dv_done)
    ws.add_data_validation(dv_int)

    last = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row      = i + 4
        day_name = d.strftime("%A")
        plan     = DAY_TYPE[day_name]
        bg       = LIGHT_PURPLE if plan == "Rest Day" else (WHITE if i % 2 == 0 else GRAY_BG)

        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=day_name)
        ws.cell(row=row, column=3, value=plan)
        ws.cell(row=row, column=4, value=UNCHECKED)
        ws.cell(row=row, column=5, value="")
        ws.cell(row=row, column=6, value="")
        ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=9, value="")
        style_data(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 22

        if plan == "Rest Day":
            for c in range(1, 10):
                ws.cell(row=row, column=c).fill = hfill(LIGHT_PURPLE)
                ws.cell(row=row, column=c).font = mfont(color=PURPLE, size=10, italic=True)

    dv_done.sqref = f"D4:D{last}"
    dv_int.sqref  = f"H4:H{last}"

    rng = f"D4:D{last}"
    # ☑ Grind Day → green
    ws.conditional_formatting.add(rng,
        FormulaRule(
            formula=[f'AND(C4="Grind Day",D4="{CHECKED}")'],
            fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)
        )
    )
    # ☐ Grind Day + past → red
    ws.conditional_formatting.add(rng,
        FormulaRule(
            formula=[f'AND($A4<TODAY(),C4="Grind Day",D4="{UNCHECKED}")'],
            fill=hfill(LIGHT_RED), font=Font(color=RED, bold=True)
        )
    )

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — SUPPLEMENTS
# ══════════════════════════════════════════════════════════════════════════════

def create_supplements_sheet(wb):
    ws = wb.create_sheet("Supplements")
    headers    = ["Date","Day","Multivitamin","Omega","Both Completed","Notes"]
    col_widths = [13,12,16,14,16,25]

    ws.merge_cells("A1:F1")
    ws["A1"].value     = "💊  SUPPLEMENTS TRACKER 2026"
    ws["A1"].fill      = hfill(ORANGE)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:F2")
    ws["A2"].value     = f"Daily: Multivitamin + Omega   •   Past uncompleted days → RED automatically"
    ws["A2"].fill      = hfill(LIGHT_ORANGE)
    ws["A2"].font      = mfont(italic=True, color=ORANGE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, ORANGE, WHITE, range(1, 7))

    last = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value=UNCHECKED)
        ws.cell(row=row, column=4, value=UNCHECKED)
        ws.cell(row=row, column=5,
                value=f'=IF(AND(C{row}="{CHECKED}",D{row}="{CHECKED}"),"{CHECKED}","{UNCHECKED}")')
        ws.cell(row=row, column=6, value="")
        style_data(ws, row, bg, range(1, 7))
        ws.row_dimensions[row].height = 22

    add_dv_checkbox(ws, "C", 4, last)
    add_dv_checkbox(ws, "D", 4, last)

    for rng in [f"C4:C{last}", f"D4:D{last}", f"E4:E{last}"]:
        ws.conditional_formatting.add(rng,
            FormulaRule(formula=[f'C4="{CHECKED}"'],
                        fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
        ws.conditional_formatting.add(rng,
            FormulaRule(formula=[f'AND($A4<TODAY(),C4="{UNCHECKED}")'],
                        fill=hfill(LIGHT_RED), font=Font(color=RED, bold=True)))

    # Fix E column references specifically
    ws.conditional_formatting.add(f"E4:E{last}",
        FormulaRule(formula=[f'E4="{CHECKED}"'],
                    fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(f"E4:E{last}",
        FormulaRule(formula=[f'AND($A4<TODAY(),E4="{UNCHECKED}")'],
                    fill=hfill(LIGHT_RED), font=Font(color=RED, bold=True)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:F{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — STEPS TRACKER
# ══════════════════════════════════════════════════════════════════════════════

def create_steps_sheet(wb):
    ws = wb.create_sheet("Steps Tracker")
    headers    = ["Date","Day","Steps","Goal","Goal Met","Difference","Notes"]
    col_widths = [13,12,12,10,11,13,25]

    ws.merge_cells("A1:G1")
    ws["A1"].value     = "👟  STEPS TRACKER 2026"
    ws["A1"].fill      = hfill(BLUE)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:G2")
    ws["A2"].value     = f"Goal: 10,000 steps/day   •   {CHECKED}=Goal Met  {UNCHECKED}=Not Met   •   Past unmet days → RED"
    ws["A2"].fill      = hfill(LIGHT_BLUE)
    ws["A2"].font      = mfont(italic=True, color=BLUE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, BLUE, WHITE, range(1, 8))

    last = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value="")
        c_goal = ws.cell(row=row, column=4, value="='Goals & Settings'!B5")
        c_goal.number_format = "#,##0"
        ws.cell(row=row, column=5,
                value=f'=IF(C{row}="","",IF(C{row}>=D{row},"{CHECKED}","{UNCHECKED}"))')
        ws.cell(row=row, column=6,
                value=f'=IF(C{row}="","",C{row}-D{row})').number_format = "+#,##0;-#,##0;0"
        ws.cell(row=row, column=7, value="")
        style_data(ws, row, bg, range(1, 8))
        ws.row_dimensions[row].height = 22
        ws.cell(row=row, column=4).number_format = "#,##0"
        ws.cell(row=row, column=6).number_format = "+#,##0;-#,##0;0"

    # Goal Met column E
    rng = f"E4:E{last}"
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=[f'E4="{CHECKED}"'],
                    fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=[f'AND($A4<TODAY(),E4="{UNCHECKED}")'],
                    fill=hfill(LIGHT_RED), font=Font(color=RED, bold=True)))

    # Steps column: empty + past date → soft red reminder
    ws.conditional_formatting.add(f"C4:C{last}",
        FormulaRule(formula=['AND($A4<TODAY(),C4="")'],
                    fill=hfill(LIGHT_RED), font=Font(color=RED, italic=True)))
    # Positive difference → green
    ws.conditional_formatting.add(f"F4:F{last}",
        CellIsRule("greaterThanOrEqual", ["0"], fill=hfill(LIGHT_GREEN)))
    ws.conditional_formatting.add(f"F4:F{last}",
        CellIsRule("lessThan", ["0"], fill=hfill(LIGHT_RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:G{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 — RUNNING TRACKER
# ══════════════════════════════════════════════════════════════════════════════

def create_running_sheet(wb):
    ws = wb.create_sheet("Running Tracker")
    headers    = ["Date","Day","Ran Today","Distance (KM)",
                  "Duration (min)","Pace (min/km)","Calories","Route / Location","Notes"]
    col_widths = [13,12,11,14,15,15,12,22,22]

    ws.merge_cells("A1:I1")
    ws["A1"].value     = "🏃  RUNNING TRACKER 2026"
    ws["A1"].fill      = hfill(GREEN)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    ws["A2"].value     = f"Pace = Duration ÷ Distance (auto)   •   {CHECKED}=Ran  {UNCHECKED}=Rest   •   Past unrun days → amber reminder"
    ws["A2"].fill      = hfill(LIGHT_GREEN)
    ws["A2"].font      = mfont(italic=True, color=GREEN, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, GREEN, WHITE, range(1, 10))

    last = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        ws.cell(row=row, column=3, value=UNCHECKED)
        ws.cell(row=row, column=4, value="")
        ws.cell(row=row, column=5, value="")
        ws.cell(row=row, column=6,
                value=f'=IFERROR(IF(AND(D{row}<>"",E{row}<>""),E{row}/D{row},""),"")')
        ws.cell(row=row, column=7, value="")
        ws.cell(row=row, column=8, value="")
        ws.cell(row=row, column=9, value="")
        style_data(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 22

    add_dv_checkbox(ws, "C", 4, last)

    rng = f"C4:C{last}"
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=[f'C4="{CHECKED}"'],
                    fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    # Running is optional → soft orange, not hard red
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=[f'AND($A4<TODAY(),C4="{UNCHECKED}")'],
                    fill=hfill(LIGHT_ORANGE), font=Font(color=ORANGE, bold=True)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7 — EXPENSE & SAVINGS
# ══════════════════════════════════════════════════════════════════════════════

def create_expense_sheet(wb):
    ws = wb.create_sheet("Expense & Savings")
    headers    = ["Date","Day","Category","Description",
                  "Income","Expense","Savings","Payment Method","Notes"]
    col_widths = [13,12,16,28,13,13,13,16,22]

    ws.merge_cells("A1:I1")
    ws["A1"].value     = "💰  EXPENSE & SAVINGS TRACKER 2026"
    ws["A1"].fill      = hfill(PINK)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:I2")
    ws["A2"].value     = "Track income, expenses, and savings daily  •  Savings = Income − Expense"
    ws["A2"].fill      = hfill(LIGHT_PINK)
    ws["A2"].font      = mfont(italic=True, color=PINK, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, PINK, WHITE, range(1, 10))

    dv_cat = DataValidation(type="list",
                             formula1='"Food,Transport,Shopping,Education,Health,Subscription,Family,Personal,Other"',
                             allow_blank=True)
    dv_pay = DataValidation(type="list",
                             formula1='"Cash,Card,Online Transfer,Mobile Pay"',
                             allow_blank=True)
    ws.add_data_validation(dv_cat)
    ws.add_data_validation(dv_pay)

    last = 3 + len(DATES)

    for i, d in enumerate(DATES):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value=d).number_format = "DD-MMM-YYYY"
        ws.cell(row=row, column=2, value=d.strftime("%A"))
        for c in [3,4,5,6,8,9]:
            ws.cell(row=row, column=c, value="")
        ws.cell(row=row, column=7,
                value=f'=IFERROR(IF(AND(E{row}<>"",F{row}<>""),E{row}-F{row},'
                      f'IF(E{row}<>"",E{row},IF(F{row}<>"",-F{row},""))),"")').number_format = "#,##0.00"
        style_data(ws, row, bg, range(1, 10))
        ws.row_dimensions[row].height = 22
        for c in [5,6,7]:
            ws.cell(row=row, column=c).number_format = "#,##0.00"

    dv_cat.sqref = f"C4:C{last}"
    dv_pay.sqref = f"H4:H{last}"

    ws.conditional_formatting.add(f"G4:G{last}",
        CellIsRule("greaterThan", ["0"], fill=hfill(LIGHT_GREEN), font=Font(color=GREEN)))
    ws.conditional_formatting.add(f"G4:G{last}",
        CellIsRule("lessThan",    ["0"], fill=hfill(LIGHT_RED),   font=Font(color=RED)))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:I{last}"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8 — GOALS & SETTINGS
# ══════════════════════════════════════════════════════════════════════════════

def create_goals_sheet(wb):
    ws = wb.create_sheet("Goals & Settings")
    col_widths = [30,18,20,30]
    for c, w in enumerate(col_widths, 1):
        set_w(ws, c, w)

    ws.merge_cells("A1:D1")
    ws["A1"].value     = "⚙️  GOALS & SETTINGS"
    ws["A1"].fill      = hfill(DARK_NAVY)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:D2")
    ws["A2"].value     = "Edit values in the yellow cells  •  All tracking sheets reference these automatically"
    ws["A2"].fill      = hfill(LIGHT_BLUE)
    ws["A2"].font      = mfont(italic=True, color=BLUE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 22

    # Section header
    def section(row, label, color):
        ws.merge_cells(f"A{row}:D{row}")
        c = ws.cell(row=row, column=1, value=label)
        c.fill = hfill(color); c.font = mfont(bold=True, color=WHITE, size=11)
        c.alignment = center()
        ws.row_dimensions[row].height = 24

    def goal_row(row, label, value, unit, note, bg=WHITE):
        ws.row_dimensions[row].height = 26
        c1 = ws.cell(row=row, column=1, value=label)
        c1.fill=hfill(bg); c1.font=mfont(bold=True,size=10); c1.border=tborder(); c1.alignment=left_align()
        c2 = ws.cell(row=row, column=2, value=value)
        c2.fill=hfill(LIGHT_YELLOW); c2.font=mfont(bold=True,color=ORANGE,size=13)
        c2.border=tborder(); c2.alignment=center()
        c3 = ws.cell(row=row, column=3, value=unit)
        c3.fill=hfill(bg); c3.font=mfont(size=10,italic=True); c3.border=tborder(); c3.alignment=center()
        c4 = ws.cell(row=row, column=4, value=note)
        c4.fill=hfill(bg); c4.font=mfont(size=9,color=DARK_GRAY,italic=True); c4.border=tborder(); c4.alignment=left_align()

    section(3, "DAILY HABIT GOALS", TEAL)
    # Row 4 header
    for c, h in enumerate(["Setting","Value","Unit","Notes"], 1):
        cell = ws.cell(row=4, column=c, value=h)
        cell.fill=hfill(DARK_NAVY); cell.font=mfont(bold=True,color=WHITE,size=10)
        cell.alignment=center(); cell.border=tborder()
    ws.row_dimensions[4].height = 24

    # B5 = Steps Goal  (Steps Tracker references 'Goals & Settings'!B5)
    goal_row(5,  "Daily Steps Goal",          10000, "steps/day",      "Steps Tracker references this cell",    WHITE)
    goal_row(6,  "Daily Salat Goal",          5,     "prayers/day",    "Max 5 prayers",                         GRAY_BG)
    goal_row(7,  "Daily Supplements Goal",    2,     "supplements/day","Multivitamin + Omega = 2",              WHITE)
    goal_row(8,  "Weekly Gym Grind Days Goal",5,     "days/week",      "Mon Tue Wed Fri Sat = 5 Grind Days",    GRAY_BG)

    section(9, "MONTHLY GOALS (EDITABLE)", ORANGE)
    for c, h in enumerate(["Setting","Value","Unit","Notes"], 1):
        cell = ws.cell(row=10, column=c, value=h)
        cell.fill=hfill(DARK_NAVY); cell.font=mfont(bold=True,color=WHITE,size=10)
        cell.alignment=center(); cell.border=tborder()
    ws.row_dimensions[10].height = 24

    # B11 = Monthly Running Goal  (Monthly Summary references 'Goals & Settings'!B11)
    goal_row(11, "Monthly Running KM Goal",   0,     "km/month",       "Set your monthly running target",       WHITE)
    # B12 = Monthly Savings Goal  (Monthly Summary references 'Goals & Settings'!B12)
    goal_row(12, "Monthly Savings Goal",      0,     "currency",       "Set your monthly savings target",       GRAY_BG)

    section(13, "TRACKER DATE RANGE", DARK_NAVY)
    goal_row(14, "Start Date", date(2026,1,1),  "DD-MMM-YYYY", "First tracked date", WHITE)
    goal_row(15, "End Date",   date(2026,12,31),"DD-MMM-YYYY", "Last tracked date",  GRAY_BG)
    ws.cell(row=14, column=2).number_format = "DD-MMM-YYYY"
    ws.cell(row=15, column=2).number_format = "DD-MMM-YYYY"

    ws.freeze_panes = "A5"


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 9 — MONTHLY SUMMARY
# ══════════════════════════════════════════════════════════════════════════════

def create_monthly_summary_sheet(wb):
    ws = wb.create_sheet("Monthly Summary")
    headers    = ["Month","Salat %","Gym %","Suppl. %","Steps Goal %",
                  "Running KM","Total Income","Total Expense","Total Savings",
                  "Savings Rate %","Overall Score %"]
    col_widths = [14,12,10,12,14,14,15,15,15,15,16]

    ws.merge_cells("A1:K1")
    ws["A1"].value     = "📊  MONTHLY SUMMARY 2026"
    ws["A1"].fill      = hfill(DARK_NAVY)
    ws["A1"].font      = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:K2")
    ws["A2"].value     = ("Overall Score = Salat 25% + Gym 20% + Supplements 15% + "
                          "Steps 15% + Running 10% + Savings 15%")
    ws["A2"].fill      = hfill(LIGHT_BLUE)
    ws["A2"].font      = mfont(italic=True, color=BLUE, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 22

    ws.row_dimensions[3].height = 28
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, DARK_NAVY, WHITE, range(1, 12))

    # Pre-compute grind days per month
    grind_per_month = []
    for m in range(1, 13):
        dm = calendar.monthrange(2026, m)[1]
        cnt = sum(1 for d in range(1, dm+1) if DAY_TYPE[date(2026,m,d).strftime("%A")] == "Grind Day")
        grind_per_month.append(cnt)

    for i, month in enumerate(MONTHS):
        m  = i + 1
        dm = calendar.monthrange(2026, m)[1]
        fi = (date(2026, m, 1) - date(2026, 1, 1)).days
        fr = fi + 4          # data starts at row 4
        lr = fr + dm - 1
        row = 4 + i
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.row_dimensions[row].height = 22

        ws.cell(row=row, column=1, value=month)

        # Salat %: average of Completion % column I
        ws.cell(row=row, column=2,
                value=f"=IFERROR(AVERAGE('Salat Tracker'!I{fr}:I{lr}),0)"
                ).number_format = "0.0%"

        # Gym %: grind done / total grind days
        grind = grind_per_month[i]
        ws.cell(row=row, column=3,
                value=f'=IFERROR(COUNTIFS(\'Gym Routine\'!C{fr}:C{lr},"Grind Day",\'Gym Routine\'!D{fr}:D{lr},"{CHECKED}")/{grind},0)'
                ).number_format = "0.0%"

        # Supplements %
        ws.cell(row=row, column=4,
                value=f'=IFERROR(COUNTIF(\'Supplements\'!E{fr}:E{lr},"{CHECKED}")/{dm},0)'
                ).number_format = "0.0%"

        # Steps Goal %
        ws.cell(row=row, column=5,
                value=f'=IFERROR(COUNTIF(\'Steps Tracker\'!E{fr}:E{lr},"{CHECKED}")/{dm},0)'
                ).number_format = "0.0%"

        # Running KM
        ws.cell(row=row, column=6,
                value=f'=IFERROR(SUMIF(\'Running Tracker\'!C{fr}:C{lr},"{CHECKED}",\'Running Tracker\'!D{fr}:D{lr}),0)'
                ).number_format = "0.0"

        # Financials
        ws.cell(row=row, column=7,
                value=f"=IFERROR(SUM('Expense & Savings'!E{fr}:E{lr}),0)"
                ).number_format = "#,##0.00"
        ws.cell(row=row, column=8,
                value=f"=IFERROR(SUM('Expense & Savings'!F{fr}:F{lr}),0)"
                ).number_format = "#,##0.00"
        ws.cell(row=row, column=9,
                value=f"=IFERROR(G{row}-H{row},0)"
                ).number_format = "#,##0.00"
        ws.cell(row=row, column=10,
                value=f"=IFERROR(I{row}/G{row},0)"
                ).number_format = "0.0%"

        # Overall Score
        ws.cell(row=row, column=11,
                value=(f"=B{row}*0.25+C{row}*0.20+D{row}*0.15+E{row}*0.15"
                       f"+MIN(IFERROR(F{row}/'Goals & Settings'!B11,0),1)*0.10"
                       f"+MIN(IFERROR(I{row}/'Goals & Settings'!B12,0),1)*0.15")
                ).number_format = "0.0%"

        style_data(ws, row, bg, range(1, 12))
        # Re-apply number formats after style_data
        for c in [2,3,4,5,10,11]:
            ws.cell(row=row, column=c).number_format = "0.0%"
        ws.cell(row=row, column=6).number_format = "0.0"
        for c in [7,8,9]:
            ws.cell(row=row, column=c).number_format = "#,##0.00"

    # Annual totals
    tr = 16
    ws.row_dimensions[tr].height = 26
    ws.cell(row=tr, column=1, value="YEAR 2026 TOTAL / AVG")
    totals = [
        (2,  "=AVERAGE(B4:B15)", "0.0%"),
        (3,  "=AVERAGE(C4:C15)", "0.0%"),
        (4,  "=AVERAGE(D4:D15)", "0.0%"),
        (5,  "=AVERAGE(E4:E15)", "0.0%"),
        (6,  "=SUM(F4:F15)",     "0.0"),
        (7,  "=SUM(G4:G15)",     "#,##0.00"),
        (8,  "=SUM(H4:H15)",     "#,##0.00"),
        (9,  "=SUM(I4:I15)",     "#,##0.00"),
        (10, f"=IFERROR(I{tr}/G{tr},0)", "0.0%"),
        (11, "=AVERAGE(K4:K15)", "0.0%"),
    ]
    for col, formula, fmt in totals:
        c = ws.cell(row=tr, column=col, value=formula)
        c.number_format = fmt
    style_header(ws, tr, DARK_NAVY, WHITE, range(1, 12))
    # Re-apply number formats on header row
    for col, formula, fmt in totals:
        ws.cell(row=tr, column=col).number_format = fmt

    # Conditional formatting Overall Score
    ws.conditional_formatting.add("K4:K15",
        CellIsRule("greaterThanOrEqual", ["0.8"], fill=hfill(LIGHT_GREEN), font=Font(color=GREEN,  bold=True)))
    ws.conditional_formatting.add("K4:K15",
        CellIsRule("between", ["0.6","0.8"],       fill=hfill(LIGHT_YELLOW),font=Font(color=YELLOW, bold=True)))
    ws.conditional_formatting.add("K4:K15",
        CellIsRule("lessThan", ["0.6"],            fill=hfill(LIGHT_RED),  font=Font(color=RED,    bold=True)))

    ws.freeze_panes = "A4"

    # Bar chart
    chart = BarChart()
    chart.type  = "col"
    chart.title = "Monthly Performance Overview 2026"
    chart.y_axis.title = "Score (%)"
    chart.x_axis.title = "Month"
    chart.style  = 10
    chart.height = 14
    chart.width  = 26
    cats = Reference(ws, min_col=1, min_row=4, max_row=15)
    for col, label in [(2,"Salat"),(3,"Gym"),(4,"Supplements"),(5,"Steps"),(11,"Overall")]:
        data = Reference(ws, min_col=col, min_row=3, max_row=15)
        chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws.add_chart(chart, "A18")


# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

def create_dashboard(wb):
    ws = wb.active
    ws.title = "Dashboard"
    ws.sheet_view.showGridLines = False

    col_widths = [2,14,14,14,14,14,14,14,2]
    for c, w in enumerate(col_widths, 1):
        set_w(ws, c, w)

    def merge_cell(rng, value, fill_c, font_c, size=11, bold=False, italic=False, fmt=None):
        ws.merge_cells(rng)
        start = rng.split(":")[0]
        c = ws[start]
        c.value     = value
        c.fill      = hfill(fill_c)
        c.font      = mfont(bold=bold, size=size, color=font_c, italic=italic)
        c.alignment = center()
        c.border    = tborder()
        if fmt:
            c.number_format = fmt
        return c

    # ── Title ──
    merge_cell("B1:H1","PERSONAL LIFE TRACKER 2026",DARK_NAVY,WHITE,18,bold=True)
    ws.row_dimensions[1].height = 42

    merge_cell("B2:H2",
               "Your daily companion for habits, health, and wealth  •  Stay consistent, track everything!",
               TEAL, WHITE, 10, italic=True)
    ws.row_dimensions[2].height = 24

    # ── Today ──
    ws.row_dimensions[3].height = 6
    merge_cell("B4:H4", '=TODAY()', LIGHT_TEAL, TEAL, 12, bold=True,
               fmt='"Today: "DDDD, DD MMMM YYYY')
    ws.row_dimensions[4].height = 28

    # ── KPI Row 1 ──
    ws.row_dimensions[5].height = 6
    merge_cell("B6:H6","CURRENT YEAR PERFORMANCE  —  KEY METRICS",DARK_NAVY,WHITE,11,bold=True)
    ws.row_dimensions[6].height = 26

    kpi1 = [
        ("B","C", "🕌 SALAT",       TEAL,   LIGHT_TEAL,
         f'=IFERROR(COUNTIF(\'Salat Tracker\'!I$4:I$368,"=1")/COUNTA(\'Salat Tracker\'!A$4:A$368),0)',
         "0.0%", "Prayers 100%"),
        ("D","E", "💪 GYM",          PURPLE, LIGHT_PURPLE,
         f'=IFERROR(COUNTIFS(\'Gym Routine\'!C$4:C$368,"Grind Day",\'Gym Routine\'!D$4:D$368,"{CHECKED}")/COUNTIF(\'Gym Routine\'!C$4:C$368,"Grind Day"),0)',
         "0.0%", "Workouts Done"),
        ("F","G", "💊 SUPPLEMENTS",  ORANGE, LIGHT_ORANGE,
         f'=IFERROR(COUNTIF(\'Supplements\'!E$4:E$368,"{CHECKED}")/COUNTA(\'Supplements\'!A$4:A$368),0)',
         "0.0%", "Both Taken"),
        ("H","H", "👟 STEPS",        BLUE,   LIGHT_BLUE,
         f'=IFERROR(COUNTIF(\'Steps Tracker\'!E$4:E$368,"{CHECKED}")/COUNTA(\'Steps Tracker\'!A$4:A$368),0)',
         "0.0%", "Goal Met Days"),
    ]
    for c1, c2, lbl, hdr_bg, val_bg, formula, fmt, sub in kpi1:
        rng_l = f"{c1}7:{c2}7"; rng_v = f"{c1}8:{c2}8"; rng_s = f"{c1}9:{c2}9"
        merge_cell(rng_l, lbl,     hdr_bg, WHITE,   10, bold=True)
        merge_cell(rng_v, formula, val_bg, hdr_bg, 16, bold=True, fmt=fmt)
        merge_cell(rng_s, sub,     val_bg, DARK_GRAY, 9, italic=True)
    ws.row_dimensions[7].height = 26
    ws.row_dimensions[8].height = 36
    ws.row_dimensions[9].height = 20

    # ── KPI Row 2 ──
    ws.row_dimensions[10].height = 6
    kpi2 = [
        ("B","C","🏃 RUNNING",      GREEN, LIGHT_GREEN,
         "=IFERROR(SUM('Running Tracker'!D$4:D$368),0)","0.0 \"km\"","Total KM Run"),
        ("D","E","💰 INCOME YTD",   PINK,  LIGHT_PINK,
         "=IFERROR(SUM('Expense & Savings'!E$4:E$368),0)","#,##0.00","All Income"),
        ("F","G","📉 EXPENSE YTD",  RED,   LIGHT_RED,
         "=IFERROR(SUM('Expense & Savings'!F$4:F$368),0)","#,##0.00","All Expenses"),
        ("H","H","📈 SAVINGS YTD",  TEAL,  LIGHT_TEAL,
         "=IFERROR(SUM('Expense & Savings'!E$4:E$368)-SUM('Expense & Savings'!F$4:F$368),0)","#,##0.00","Net Savings"),
    ]
    for c1, c2, lbl, hdr_bg, val_bg, formula, fmt, sub in kpi2:
        rng_l = f"{c1}11:{c2}11"; rng_v = f"{c1}12:{c2}12"; rng_s = f"{c1}13:{c2}13"
        merge_cell(rng_l, lbl,     hdr_bg, WHITE,   10, bold=True)
        merge_cell(rng_v, formula, val_bg, hdr_bg, 16, bold=True, fmt=fmt)
        merge_cell(rng_s, sub,     val_bg, DARK_GRAY, 9, italic=True)
    ws.row_dimensions[11].height = 26
    ws.row_dimensions[12].height = 36
    ws.row_dimensions[13].height = 20

    # ── Overall Life Score ──
    ws.row_dimensions[14].height = 6
    merge_cell("B15:H15","OVERALL LIFE SCORE 2026",DARK_NAVY,WHITE,11,bold=True)
    ws.row_dimensions[15].height = 26
    merge_cell("B16:H17","='Monthly Summary'!K16",DARK_NAVY,WHITE,28,bold=True,fmt="0.0%")
    ws.row_dimensions[16].height = 38; ws.row_dimensions[17].height = 20
    merge_cell("B18:H18",
               "Salat 25%  +  Gym 20%  +  Supplements 15%  +  Steps 15%  +  Running 10%  +  Savings 15%",
               LIGHT_BLUE, BLUE, 9, italic=True)
    ws.row_dimensions[18].height = 20

    # ── Monthly table ──
    ws.row_dimensions[19].height = 6
    merge_cell("B20:H20","MONTHLY PERFORMANCE AT A GLANCE",TEAL,WHITE,11,bold=True)
    ws.row_dimensions[20].height = 26

    mini_h = ["Month","Salat %","Gym %","Suppl. %","Steps %","Running KM","Overall"]
    mini_c = ["B","C","D","E","F","G","H"]
    ws.row_dimensions[21].height = 24
    for col_l, h in zip(mini_c, mini_h):
        c = ws[f"{col_l}21"]
        c.value=h; c.fill=hfill(DARK_NAVY); c.font=mfont(bold=True,color=WHITE,size=9)
        c.alignment=center(); c.border=tborder()

    for i, month in enumerate(MONTHS):
        r   = 22 + i
        src = 4  + i
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.row_dimensions[r].height = 20
        data = [
            ("B", month,                              None),
            ("C", f"='Monthly Summary'!B{src}",      "0%"),
            ("D", f"='Monthly Summary'!C{src}",      "0%"),
            ("E", f"='Monthly Summary'!D{src}",      "0%"),
            ("F", f"='Monthly Summary'!E{src}",      "0%"),
            ("G", f"='Monthly Summary'!F{src}",      "0.0"),
            ("H", f"='Monthly Summary'!K{src}",      "0%"),
        ]
        for col_l, val, fmt in data:
            c = ws[f"{col_l}{r}"]
            c.value=val; c.fill=hfill(bg); c.font=mfont(size=9)
            c.alignment=center(); c.border=tborder()
            if fmt: c.number_format=fmt

    ws.conditional_formatting.add("H22:H33",
        ColorScaleRule(start_type='min',   start_color=RED,
                       mid_type='percentile', mid_value=50, mid_color=YELLOW,
                       end_type='max',    end_color=GREEN))


# ══════════════════════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    wb = Workbook()

    create_dashboard(wb)
    create_salat_sheet(wb)
    create_gym_sheet(wb)
    create_supplements_sheet(wb)
    create_steps_sheet(wb)
    create_running_sheet(wb)
    create_expense_sheet(wb)
    create_goals_sheet(wb)
    create_monthly_summary_sheet(wb)

    tab_colors = {
        "Dashboard":        DARK_NAVY,
        "Salat Tracker":    TEAL,
        "Gym Routine":      PURPLE,
        "Supplements":      ORANGE,
        "Steps Tracker":    BLUE,
        "Running Tracker":  GREEN,
        "Expense & Savings":PINK,
        "Goals & Settings": DARK_NAVY,
        "Monthly Summary":  YELLOW,
    }
    for name in wb.sheetnames:
        if name in tab_colors:
            wb[name].sheet_properties.tabColor = tab_colors[name]

    path = "/home/user/habit-tracker/Personal_Life_Tracker.xlsx"
    wb.save(path)
    print(f"✅  Saved → {path}")
    print(f"   Sheets  : {wb.sheetnames}")
    print(f"   Symbol  : {CHECKED} = done   {UNCHECKED} = not done")
    print(f"   Auto-red: past uncompleted cells turn RED via TODAY() formula")

if __name__ == "__main__":
    main()
