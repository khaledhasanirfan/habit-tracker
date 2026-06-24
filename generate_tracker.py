"""
Personal Life Tracker 2026 — v3
- Native Excel 365 checkboxes (injected via ZIP/XML post-processing)
- Boolean FALSE/TRUE cell values (no text dropdowns)
- All unchecked  → RED  (default motivation state)
- Any checked    → GREEN (instant feedback on tap/click)
- Past unchecked → RED  (already the default, stays red)
- Can retroactively check any past day
"""

import zipfile, re, os, calendar
from io import BytesIO
from datetime import date, timedelta

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation

# ── Color palette ──────────────────────────────────────────────────────────────
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
# Checkbox cell colors — baked into cell DEFAULT FILL so they work in ALL apps
CB_RED       = "FFCDD2"  # unchecked default (red, visible even without CF)
CB_RED_FONT  = "B71C1C"  # dark red text
CB_GREEN     = "C8E6C9"  # checked CF override (green)
CB_GRN_FONT  = "1B5E20"  # dark green text

# ── Style helpers ──────────────────────────────────────────────────────────────
def hfill(c):   return PatternFill("solid", fgColor=c)
def mfont(bold=False, size=11, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")
def tborder():
    s = Side(style="thin", color="DDDDDD")
    return Border(left=s, right=s, top=s, bottom=s)
def center():   return Alignment(horizontal="center", vertical="center", wrap_text=True)
def left_al():  return Alignment(horizontal="left",   vertical="center", wrap_text=True)
def set_w(ws, col, w): ws.column_dimensions[get_column_letter(col)].width = w

def style_header(ws, row, bg, fg=WHITE, cols=None):
    cols = cols or range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill = hfill(bg); cell.font = mfont(bold=True, color=fg, size=10)
        cell.alignment = center(); cell.border = tborder()

def style_data(ws, row, bg=WHITE, cols=None):
    cols = cols or range(1, ws.max_column + 1)
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill = hfill(bg); cell.border = tborder()
        cell.alignment = center(); cell.font = mfont(size=10)

# ── Checkbox cell helpers ──────────────────────────────────────────────────────
# Strategy: bake RED into the cell's default fill so it's visible in every app
# (Excel, Google Sheets, LibreOffice, mobile) even without CF.
# CF only needs to override TRUE → GREEN.

def set_cb_red(ws, row, cols):
    """Set checkbox cells to red fill by default (unchecked base state)."""
    for c in cols:
        cell = ws.cell(row=row, column=c)
        cell.fill = hfill(CB_RED)
        cell.font = mfont(size=10, color=CB_RED_FONT)
        cell.alignment = center()
        cell.number_format = ';;;'  # hide 0/1 value; color CF gives visual feedback

def cf_checkbox_green(ws, rng, first_cell):
    """CF: only 1 → GREEN  (red is the default fill, no CF rule needed for it)."""
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=[f'{first_cell}=1'],
                    fill=hfill(CB_GREEN),
                    font=Font(color=CB_GRN_FONT, bold=True, name="Calibri")))

# ── Date data ──────────────────────────────────────────────────────────────────
def all_dates():
    s = date(2026, 1, 1)
    return [s + timedelta(days=i) for i in range(365)]

DATES = all_dates()
DAY_TYPE = {
    "Monday":"Grind Day","Tuesday":"Grind Day","Wednesday":"Grind Day",
    "Thursday":"Rest Day","Friday":"Grind Day","Saturday":"Grind Day","Sunday":"Rest Day",
}
MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 2 — SALAT TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def create_salat_sheet(wb):
    ws = wb.create_sheet("Salat Tracker")
    headers    = ["Date","Day","Fajr","Dhuhr","Asr","Maghrib","Isha",
                  "Total Prayers","Completion %","Notes"]
    col_widths = [13,12,9,9,9,9,9,14,14,25]

    ws.merge_cells("A1:J1")
    ws["A1"].value="\U0001f319  SALAT TRACKER 2026"; ws["A1"].fill=hfill(DARK_NAVY)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height = 32

    ws.merge_cells("A2:J2")
    ws["A2"].value = "Tap a prayer cell to check ✓ (green) or uncheck (red)   •   All cells start RED — go green as you pray!"
    ws["A2"].fill=hfill(TEAL); ws["A2"].font=mfont(italic=True,color=WHITE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height = 24

    ws.row_dimensions[3].height = 28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,TEAL,WHITE,range(1,11))

    last = 3 + len(DATES)
    for i,d in enumerate(DATES):
        row = i+4
        bg  = WHITE if i%2==0 else GRAY_BG
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=d.strftime("%A"))
        for c in range(3,8):              # Fajr–Isha: 0 = unchecked
            ws.cell(row=row,column=c,value=0)
        ws.cell(row=row,column=8,value=f"=COUNTIF(C{row}:G{row},1)")
        ws.cell(row=row,column=9,value=f"=H{row}/5").number_format="0%"
        ws.cell(row=row,column=10,value="")
        style_data(ws,row,bg,range(1,11))
        set_cb_red(ws, row, range(3,8))     # Fajr–Isha: red by default
        ws.row_dimensions[row].height = 22

    # CF: only TRUE → GREEN (red is baked into cell fill already)
    cf_checkbox_green(ws, f"C4:G{last}", "C4")

    # Completion % CF
    comp = f"I4:I{last}"
    ws.conditional_formatting.add(comp, CellIsRule("equal",["1"],      fill=hfill(LIGHT_GREEN), font=Font(color=GREEN, bold=True)))
    ws.conditional_formatting.add(comp, CellIsRule("greaterThanOrEqual",["0.6"], fill=hfill(LIGHT_YELLOW),font=Font(color=YELLOW,bold=True)))
    ws.conditional_formatting.add(comp, CellIsRule("lessThan",["0.6"], fill=hfill(LIGHT_RED),  font=Font(color=RED,   bold=True)))

    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:J{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 3 — GYM ROUTINE
# ══════════════════════════════════════════════════════════════════════════════
def create_gym_sheet(wb):
    ws = wb.create_sheet("Gym Routine")
    headers    = ["Date","Day","Planned Type","Workout Done",
                  "Workout Name","Body Part","Duration (min)","Intensity","Notes"]
    col_widths = [13,12,13,14,22,16,15,13,25]

    ws.merge_cells("A1:I1")
    ws["A1"].value="\U0001f4aa  GYM ROUTINE TRACKER 2026"; ws["A1"].fill=hfill(PURPLE)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:I2")
    ws["A2"].value="Mon/Tue/Wed/Fri/Sat = Grind Day  •  Thu/Sun = Rest Day  •  Tap Workout Done to check green — past missed Grind Days stay RED"
    ws["A2"].fill=hfill(LIGHT_PURPLE); ws["A2"].font=mfont(italic=True,color=PURPLE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,PURPLE,WHITE,range(1,10))

    dv_int = DataValidation(type="list",formula1='"Low,Medium,High,Very High"',allow_blank=True)
    ws.add_data_validation(dv_int)

    last = 3+len(DATES)
    for i,d in enumerate(DATES):
        row=i+4; day_name=d.strftime("%A"); plan=DAY_TYPE[day_name]
        bg = LIGHT_PURPLE if plan=="Rest Day" else (WHITE if i%2==0 else GRAY_BG)
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=day_name)
        ws.cell(row=row,column=3,value=plan)
        ws.cell(row=row,column=4,value=0)   # 0 = unchecked
        for c in [5,6,7,8,9]: ws.cell(row=row,column=c,value="")
        style_data(ws,row,bg,range(1,10)); ws.row_dimensions[row].height=22
        if plan=="Rest Day":
            for c in range(1,10):
                ws.cell(row=row,column=c).fill=hfill(LIGHT_PURPLE)
                ws.cell(row=row,column=c).font=mfont(color=PURPLE,size=10,italic=True)
        else:
            set_cb_red(ws, row, [4])        # Grind Day: D column red by default

    dv_int.sqref=f"H4:H{last}"

    # CF: Grind Day + TRUE → GREEN  (red is baked into cell fill for Grind rows)
    rng = f"D4:D{last}"
    ws.conditional_formatting.add(rng,
        FormulaRule(formula=['AND(C4="Grind Day",D4=1)'],
                    fill=hfill(CB_GREEN), font=Font(color=CB_GRN_FONT,bold=True)))

    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:I{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 4 — SUPPLEMENTS
# ══════════════════════════════════════════════════════════════════════════════
def create_supplements_sheet(wb):
    ws = wb.create_sheet("Supplements")
    headers    = ["Date","Day","Multivitamin","Omega","Both Completed","Notes"]
    col_widths = [13,12,16,14,16,25]

    ws.merge_cells("A1:F1")
    ws["A1"].value="\U0001f48a  SUPPLEMENTS TRACKER 2026"; ws["A1"].fill=hfill(ORANGE)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:F2")
    ws["A2"].value="Tap each cell to check  •  Both Completed auto-calculates  •  Unchecked cells → RED by default"
    ws["A2"].fill=hfill(LIGHT_ORANGE); ws["A2"].font=mfont(italic=True,color=ORANGE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,ORANGE,WHITE,range(1,7))

    last = 3+len(DATES)
    for i,d in enumerate(DATES):
        row=i+4; bg=WHITE if i%2==0 else GRAY_BG
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=d.strftime("%A"))
        ws.cell(row=row,column=3,value=0)   # Multivitamin: 0=unchecked
        ws.cell(row=row,column=4,value=0)   # Omega: 0=unchecked
        ws.cell(row=row,column=5,value=f"=C{row}*D{row}")   # Both=1 only if both checked
        ws.cell(row=row,column=6,value="")
        style_data(ws,row,bg,range(1,7)); ws.row_dimensions[row].height=22
        set_cb_red(ws, row, [3,4,5])        # Multivitamin, Omega, Both: red by default

    # CF: TRUE → GREEN  (red is baked into cell fill)
    cf_checkbox_green(ws, f"C4:C{last}", "C4")
    cf_checkbox_green(ws, f"D4:D{last}", "D4")
    cf_checkbox_green(ws, f"E4:E{last}", "E4")

    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:F{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 5 — STEPS TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def create_steps_sheet(wb):
    ws = wb.create_sheet("Steps Tracker")
    headers    = ["Date","Day","Steps","Goal","Goal Met","Difference","Notes"]
    col_widths = [13,12,12,10,11,13,25]

    ws.merge_cells("A1:G1")
    ws["A1"].value="\U0001f45f  STEPS TRACKER 2026"; ws["A1"].fill=hfill(BLUE)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:G2")
    ws["A2"].value="Enter your step count — Goal Met turns GREEN automatically when steps ≥ goal, RED when below"
    ws["A2"].fill=hfill(LIGHT_BLUE); ws["A2"].font=mfont(italic=True,color=BLUE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,BLUE,WHITE,range(1,8))

    last = 3+len(DATES)
    for i,d in enumerate(DATES):
        row=i+4; bg=WHITE if i%2==0 else GRAY_BG
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=d.strftime("%A"))
        ws.cell(row=row,column=3,value="")
        c_goal = ws.cell(row=row,column=4,value="='Goals & Settings'!B5")
        c_goal.number_format="#,##0"
        ws.cell(row=row,column=5,value=f'=IF(C{row}="","",IF(C{row}>=D{row},1,0))')
        ws.cell(row=row,column=6,value=f'=IF(C{row}="","",C{row}-D{row})').number_format="+#,##0;-#,##0;0"
        ws.cell(row=row,column=7,value="")
        style_data(ws,row,bg,range(1,8)); ws.row_dimensions[row].height=22
        ws.cell(row=row,column=4).number_format="#,##0"
        ws.cell(row=row,column=5).number_format=';;;'   # hide 0/1; CF supplies color
        ws.cell(row=row,column=6).number_format="+#,##0;-#,##0;0"

    # Goal Met (E): CF 1→green, 0→red (cell starts empty for unfilled days)
    cf_checkbox_green(ws, f"E4:E{last}", "E4")
    ws.conditional_formatting.add(f"E4:E{last}",
        FormulaRule(formula=['AND(E4<>"",E4=0)'],
                    fill=hfill(CB_RED), font=Font(color=CB_RED_FONT, bold=True)))
    # Empty steps + past date → soft red hint
    ws.conditional_formatting.add(f"C4:C{last}",
        FormulaRule(formula=['AND($A4<TODAY(),C4="")'],
                    fill=hfill(LIGHT_RED), font=Font(color=RED,italic=True)))
    # Difference: green if positive
    ws.conditional_formatting.add(f"F4:F{last}",
        CellIsRule("greaterThanOrEqual",["0"],fill=hfill(LIGHT_GREEN)))
    ws.conditional_formatting.add(f"F4:F{last}",
        CellIsRule("lessThan",["0"],fill=hfill(LIGHT_RED)))

    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:G{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 6 — RUNNING TRACKER
# ══════════════════════════════════════════════════════════════════════════════
def create_running_sheet(wb):
    ws = wb.create_sheet("Running Tracker")
    headers    = ["Date","Day","Ran Today","Distance (KM)",
                  "Duration (min)","Pace (min/km)","Calories","Route / Location","Notes"]
    col_widths = [13,12,11,14,15,15,12,22,22]

    ws.merge_cells("A1:I1")
    ws["A1"].value="\U0001f3c3  RUNNING TRACKER 2026"; ws["A1"].fill=hfill(GREEN)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:I2")
    ws["A2"].value="Tap 'Ran Today' to check  •  Pace auto-calculates  •  Rest days stay neutral (orange reminder only)"
    ws["A2"].fill=hfill(LIGHT_GREEN); ws["A2"].font=mfont(italic=True,color=GREEN,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,GREEN,WHITE,range(1,10))

    last = 3+len(DATES)
    for i,d in enumerate(DATES):
        row=i+4; bg=WHITE if i%2==0 else GRAY_BG
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=d.strftime("%A"))
        ws.cell(row=row,column=3,value=0)   # Ran Today: 0=unchecked
        for c in [4,5,7,8,9]: ws.cell(row=row,column=c,value="")
        ws.cell(row=row,column=6,
                value=f'=IFERROR(IF(AND(D{row}<>"",E{row}<>""),E{row}/D{row},""),"")')
        style_data(ws,row,bg,range(1,10)); ws.row_dimensions[row].height=22
        set_cb_red(ws, row, [3])            # Ran Today: red by default

    # Ran Today: CF TRUE → GREEN (red baked into default fill)
    cf_checkbox_green(ws, f"C4:C{last}", "C4")

    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:I{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 7 — EXPENSE & SAVINGS
# ══════════════════════════════════════════════════════════════════════════════
def create_expense_sheet(wb):
    ws = wb.create_sheet("Expense & Savings")
    headers    = ["Date","Day","Category","Description",
                  "Income","Expense","Savings","Payment Method","Notes"]
    col_widths = [13,12,16,28,13,13,13,16,22]

    ws.merge_cells("A1:I1")
    ws["A1"].value="\U0001f4b0  EXPENSE & SAVINGS TRACKER 2026"; ws["A1"].fill=hfill(PINK)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:I2")
    ws["A2"].value="Track income, expenses, and savings daily  •  Savings = Income − Expense"
    ws["A2"].fill=hfill(LIGHT_PINK); ws["A2"].font=mfont(italic=True,color=PINK,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,PINK,WHITE,range(1,10))

    dv_cat = DataValidation(type="list",
        formula1='"Food,Transport,Shopping,Education,Health,Subscription,Family,Personal,Other"',allow_blank=True)
    dv_pay = DataValidation(type="list",
        formula1='"Cash,Card,Online Transfer,Mobile Pay"',allow_blank=True)
    ws.add_data_validation(dv_cat); ws.add_data_validation(dv_pay)

    last = 3+len(DATES)
    for i,d in enumerate(DATES):
        row=i+4; bg=WHITE if i%2==0 else GRAY_BG
        ws.cell(row=row,column=1,value=d).number_format="DD-MMM-YYYY"
        ws.cell(row=row,column=2,value=d.strftime("%A"))
        for c in [3,4,5,6,8,9]: ws.cell(row=row,column=c,value="")
        ws.cell(row=row,column=7,
                value=f'=IFERROR(IF(AND(E{row}<>"",F{row}<>""),E{row}-F{row},'
                      f'IF(E{row}<>"",E{row},IF(F{row}<>"",-F{row},""))),"")').number_format="#,##0.00"
        style_data(ws,row,bg,range(1,10)); ws.row_dimensions[row].height=22
        for c in [5,6,7]: ws.cell(row=row,column=c).number_format="#,##0.00"

    dv_cat.sqref=f"C4:C{last}"; dv_pay.sqref=f"H4:H{last}"
    ws.conditional_formatting.add(f"G4:G{last}",
        CellIsRule("greaterThan",["0"],fill=hfill(LIGHT_GREEN),font=Font(color=GREEN)))
    ws.conditional_formatting.add(f"G4:G{last}",
        CellIsRule("lessThan",["0"],fill=hfill(LIGHT_RED),font=Font(color=RED)))
    ws.freeze_panes="A4"; ws.auto_filter.ref=f"A3:I{last}"

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 8 — GOALS & SETTINGS
# ══════════════════════════════════════════════════════════════════════════════
def create_goals_sheet(wb):
    ws = wb.create_sheet("Goals & Settings")
    col_widths = [30,18,20,30]
    for c,w in enumerate(col_widths,1): set_w(ws,c,w)

    ws.merge_cells("A1:D1")
    ws["A1"].value="⚙️  GOALS & SETTINGS"; ws["A1"].fill=hfill(DARK_NAVY)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:D2")
    ws["A2"].value="Edit the yellow value cells below  •  All tracker sheets reference these automatically"
    ws["A2"].fill=hfill(LIGHT_BLUE); ws["A2"].font=mfont(italic=True,color=BLUE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=22

    def section(row, label, color):
        ws.merge_cells(f"A{row}:D{row}")
        c=ws.cell(row=row,column=1,value=label); c.fill=hfill(color)
        c.font=mfont(bold=True,color=WHITE,size=11); c.alignment=center()
        ws.row_dimensions[row].height=24

    def hrow(row):
        for c,h in enumerate(["Setting","Value","Unit","Notes"],1):
            cell=ws.cell(row=row,column=c,value=h); cell.fill=hfill(DARK_NAVY)
            cell.font=mfont(bold=True,color=WHITE,size=10); cell.alignment=center(); cell.border=tborder()
        ws.row_dimensions[row].height=24

    def goal_row(row, label, value, unit, note, bg=WHITE):
        ws.row_dimensions[row].height=26
        c1=ws.cell(row=row,column=1,value=label)
        c1.fill=hfill(bg); c1.font=mfont(bold=True,size=10); c1.border=tborder(); c1.alignment=left_al()
        c2=ws.cell(row=row,column=2,value=value)
        c2.fill=hfill(LIGHT_YELLOW); c2.font=mfont(bold=True,color=ORANGE,size=13)
        c2.border=tborder(); c2.alignment=center()
        c3=ws.cell(row=row,column=3,value=unit)
        c3.fill=hfill(bg); c3.font=mfont(size=10,italic=True); c3.border=tborder(); c3.alignment=center()
        c4=ws.cell(row=row,column=4,value=note)
        c4.fill=hfill(bg); c4.font=mfont(size=9,color=DARK_GRAY,italic=True); c4.border=tborder(); c4.alignment=left_al()

    section(3,"DAILY HABIT GOALS",TEAL); hrow(4)
    # B5 = Steps Goal (referenced by Steps Tracker)
    goal_row(5,"Daily Steps Goal",          10000,"steps/day",      "Steps Tracker references this cell",    WHITE)
    goal_row(6,"Daily Salat Goal",          5,    "prayers/day",    "Max 5 prayers",                         GRAY_BG)
    goal_row(7,"Daily Supplements Goal",    2,    "supplements/day","Multivitamin + Omega = 2",              WHITE)
    goal_row(8,"Weekly Gym Grind Days Goal",5,    "days/week",      "Mon Tue Wed Fri Sat = 5 Grind Days",    GRAY_BG)

    section(9,"MONTHLY GOALS (EDITABLE)",ORANGE); hrow(10)
    # B11 = Monthly Running KM Goal, B12 = Monthly Savings Goal
    goal_row(11,"Monthly Running KM Goal",  0,"km/month",    "Set your monthly running target",   WHITE)
    goal_row(12,"Monthly Savings Goal",     0,"currency",    "Set your monthly savings target",   GRAY_BG)

    section(13,"TRACKER DATE RANGE",DARK_NAVY)
    goal_row(14,"Start Date",date(2026,1,1), "DD-MMM-YYYY","First tracked date",WHITE)
    goal_row(15,"End Date",  date(2026,12,31),"DD-MMM-YYYY","Last tracked date", GRAY_BG)
    ws.cell(row=14,column=2).number_format="DD-MMM-YYYY"
    ws.cell(row=15,column=2).number_format="DD-MMM-YYYY"
    ws.freeze_panes="A5"

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
    ws["A1"].value="\U0001f4ca  MONTHLY SUMMARY 2026"; ws["A1"].fill=hfill(DARK_NAVY)
    ws["A1"].font=mfont(bold=True,size=14,color=WHITE); ws["A1"].alignment=center()
    ws.row_dimensions[1].height=32

    ws.merge_cells("A2:K2")
    ws["A2"].value="Overall Score = Salat 25% + Gym 20% + Supplements 15% + Steps 15% + Running 10% + Savings 15%"
    ws["A2"].fill=hfill(LIGHT_BLUE); ws["A2"].font=mfont(italic=True,color=BLUE,size=10)
    ws["A2"].alignment=center(); ws.row_dimensions[2].height=22

    ws.row_dimensions[3].height=28
    for c,(h,w) in enumerate(zip(headers,col_widths),1):
        ws.cell(row=3,column=c,value=h); set_w(ws,c,w)
    style_header(ws,3,DARK_NAVY,WHITE,range(1,12))

    grind_per_month = []
    for m in range(1,13):
        dm=calendar.monthrange(2026,m)[1]
        cnt=sum(1 for d in range(1,dm+1) if DAY_TYPE[date(2026,m,d).strftime("%A")]=="Grind Day")
        grind_per_month.append(cnt)

    for i,month in enumerate(MONTHS):
        m=i+1; dm=calendar.monthrange(2026,m)[1]
        fi=(date(2026,m,1)-date(2026,1,1)).days
        fr=fi+4; lr=fr+dm-1
        row=4+i; bg=WHITE if i%2==0 else GRAY_BG
        ws.row_dimensions[row].height=22

        ws.cell(row=row,column=1,value=month)
        # Salat %: average Completion % from column I
        ws.cell(row=row,column=2,
                value=f"=IFERROR(AVERAGE('Salat Tracker'!I{fr}:I{lr}),0)").number_format="0.0%"
        # Gym %: count TRUE in D where C=Grind Day
        grind=grind_per_month[i]
        ws.cell(row=row,column=3,
                value=f"=IFERROR(COUNTIFS('Gym Routine'!C{fr}:C{lr},\"Grind Day\",'Gym Routine'!D{fr}:D{lr},1)/{grind},0)"
                ).number_format="0.0%"
        # Supplements %: count 1 in E (Both Completed)
        ws.cell(row=row,column=4,
                value=f"=IFERROR(COUNTIF('Supplements'!E{fr}:E{lr},1)/{dm},0)").number_format="0.0%"
        # Steps Goal %: count 1 in E (Goal Met)
        ws.cell(row=row,column=5,
                value=f"=IFERROR(COUNTIF('Steps Tracker'!E{fr}:E{lr},1)/{dm},0)").number_format="0.0%"
        # Running KM: sumif C=1
        ws.cell(row=row,column=6,
                value=f"=IFERROR(SUMIF('Running Tracker'!C{fr}:C{lr},1,'Running Tracker'!D{fr}:D{lr}),0)").number_format="0.0"
        # Financials
        ws.cell(row=row,column=7,value=f"=IFERROR(SUM('Expense & Savings'!E{fr}:E{lr}),0)").number_format="#,##0.00"
        ws.cell(row=row,column=8,value=f"=IFERROR(SUM('Expense & Savings'!F{fr}:F{lr}),0)").number_format="#,##0.00"
        ws.cell(row=row,column=9,value=f"=IFERROR(G{row}-H{row},0)").number_format="#,##0.00"
        ws.cell(row=row,column=10,value=f"=IFERROR(I{row}/G{row},0)").number_format="0.0%"
        # Overall Score
        ws.cell(row=row,column=11,
                value=(f"=B{row}*0.25+C{row}*0.20+D{row}*0.15+E{row}*0.15"
                       f"+MIN(IFERROR(F{row}/'Goals & Settings'!B11,0),1)*0.10"
                       f"+MIN(IFERROR(I{row}/'Goals & Settings'!B12,0),1)*0.15")
                ).number_format="0.0%"

        style_data(ws,row,bg,range(1,12))
        for c in [2,3,4,5,10,11]: ws.cell(row=row,column=c).number_format="0.0%"
        ws.cell(row=row,column=6).number_format="0.0"
        for c in [7,8,9]: ws.cell(row=row,column=c).number_format="#,##0.00"

    # Annual totals row
    tr=16; ws.row_dimensions[tr].height=26
    ws.cell(row=tr,column=1,value="YEAR 2026 TOTAL / AVG")
    totals=[(2,"=AVERAGE(B4:B15)","0.0%"),(3,"=AVERAGE(C4:C15)","0.0%"),
            (4,"=AVERAGE(D4:D15)","0.0%"),(5,"=AVERAGE(E4:E15)","0.0%"),
            (6,"=SUM(F4:F15)","0.0"),(7,"=SUM(G4:G15)","#,##0.00"),
            (8,"=SUM(H4:H15)","#,##0.00"),(9,"=SUM(I4:I15)","#,##0.00"),
            (10,f"=IFERROR(I{tr}/G{tr},0)","0.0%"),(11,"=AVERAGE(K4:K15)","0.0%")]
    for col,formula,fmt in totals:
        ws.cell(row=tr,column=col,value=formula).number_format=fmt
    style_header(ws,tr,DARK_NAVY,WHITE,range(1,12))
    for col,formula,fmt in totals: ws.cell(row=tr,column=col).number_format=fmt

    ws.conditional_formatting.add("K4:K15",
        CellIsRule("greaterThanOrEqual",["0.8"],fill=hfill(LIGHT_GREEN),font=Font(color=GREEN,bold=True)))
    ws.conditional_formatting.add("K4:K15",
        CellIsRule("between",["0.6","0.8"],fill=hfill(LIGHT_YELLOW),font=Font(color=YELLOW,bold=True)))
    ws.conditional_formatting.add("K4:K15",
        CellIsRule("lessThan",["0.6"],fill=hfill(LIGHT_RED),font=Font(color=RED,bold=True)))

    ws.freeze_panes="A4"

    # Bar chart
    chart=BarChart(); chart.type="col"
    chart.title="Monthly Performance Overview 2026"
    chart.y_axis.title="Score (%)"; chart.x_axis.title="Month"
    chart.style=10; chart.height=14; chart.width=26
    cats=Reference(ws,min_col=1,min_row=4,max_row=15)
    for col,label in [(2,"Salat"),(3,"Gym"),(4,"Supplements"),(5,"Steps"),(11,"Overall")]:
        chart.add_data(Reference(ws,min_col=col,min_row=3,max_row=15),titles_from_data=True)
    chart.set_categories(cats); ws.add_chart(chart,"A18")

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 1 — DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════
def create_dashboard(wb):
    ws=wb.active; ws.title="Dashboard"
    ws.sheet_view.showGridLines=False
    for c,w in enumerate([2,14,14,14,14,14,14,14,2],1): set_w(ws,c,w)

    def mc(rng, value, fill_c, font_c, size=11, bold=False, italic=False, fmt=None):
        ws.merge_cells(rng)
        start=rng.split(":")[0]; c=ws[start]
        c.value=value; c.fill=hfill(fill_c); c.border=tborder()
        c.font=mfont(bold=bold,size=size,color=font_c,italic=italic); c.alignment=center()
        if fmt: c.number_format=fmt
        return c

    mc("B1:H1","PERSONAL LIFE TRACKER 2026",DARK_NAVY,WHITE,18,bold=True)
    ws.row_dimensions[1].height=42
    mc("B2:H2","Your daily companion for habits, health, and wealth  •  Tap to check, red = pending, green = done!",
       TEAL,WHITE,10,italic=True)
    ws.row_dimensions[2].height=24

    ws.row_dimensions[3].height=6
    mc("B4:H4",'=TODAY()',LIGHT_TEAL,TEAL,12,bold=True,fmt='"Today: "DDDD, DD MMMM YYYY')
    ws.row_dimensions[4].height=28
    ws.row_dimensions[5].height=6
    mc("B6:H6","CURRENT YEAR PERFORMANCE  —  KEY METRICS",DARK_NAVY,WHITE,11,bold=True)
    ws.row_dimensions[6].height=26

    kpi1=[
        ("B","C","🕌 SALAT",     TEAL,  LIGHT_TEAL,
         "=IFERROR(COUNTIF('Salat Tracker'!I$4:I$368,1)/COUNTA('Salat Tracker'!A$4:A$368),0)","0.0%","Prayers 100%"),
        ("D","E","💪 GYM",        PURPLE,LIGHT_PURPLE,
         "=IFERROR(COUNTIFS('Gym Routine'!C$4:C$368,\"Grind Day\",'Gym Routine'!D$4:D$368,1)/COUNTIF('Gym Routine'!C$4:C$368,\"Grind Day\"),0)","0.0%","Workouts Done"),
        ("F","G","💊 SUPPLEMENTS",ORANGE,LIGHT_ORANGE,
         "=IFERROR(COUNTIF('Supplements'!E$4:E$368,1)/COUNTA('Supplements'!A$4:A$368),0)","0.0%","Both Taken"),
        ("H","H","👟 STEPS",      BLUE,  LIGHT_BLUE,
         "=IFERROR(COUNTIF('Steps Tracker'!E$4:E$368,1)/COUNTA('Steps Tracker'!A$4:A$368),0)","0.0%","Goal Met Days"),
    ]
    for c1,c2,lbl,hc,vc,formula,fmt,sub in kpi1:
        mc(f"{c1}7:{c2}7",lbl,hc,WHITE,10,bold=True)
        mc(f"{c1}8:{c2}8",formula,vc,hc,16,bold=True,fmt=fmt)
        mc(f"{c1}9:{c2}9",sub,vc,DARK_GRAY,9,italic=True)
    ws.row_dimensions[7].height=26; ws.row_dimensions[8].height=36; ws.row_dimensions[9].height=20

    ws.row_dimensions[10].height=6
    kpi2=[
        ("B","C","🏃 RUNNING",     GREEN,LIGHT_GREEN,
         "=IFERROR(SUM('Running Tracker'!D$4:D$368),0)",'0.0 "km"',"Total KM Run"),
        ("D","E","💰 INCOME YTD",  PINK, LIGHT_PINK,
         "=IFERROR(SUM('Expense & Savings'!E$4:E$368),0)","#,##0.00","All Income"),
        ("F","G","📉 EXPENSE YTD", RED,  LIGHT_RED,
         "=IFERROR(SUM('Expense & Savings'!F$4:F$368),0)","#,##0.00","All Expenses"),
        ("H","H","📈 SAVINGS YTD", TEAL, LIGHT_TEAL,
         "=IFERROR(SUM('Expense & Savings'!E$4:E$368)-SUM('Expense & Savings'!F$4:F$368),0)","#,##0.00","Net Savings"),
    ]
    for c1,c2,lbl,hc,vc,formula,fmt,sub in kpi2:
        mc(f"{c1}11:{c2}11",lbl,hc,WHITE,10,bold=True)
        mc(f"{c1}12:{c2}12",formula,vc,hc,16,bold=True,fmt=fmt)
        mc(f"{c1}13:{c2}13",sub,vc,DARK_GRAY,9,italic=True)
    ws.row_dimensions[11].height=26; ws.row_dimensions[12].height=36; ws.row_dimensions[13].height=20

    ws.row_dimensions[14].height=6
    mc("B15:H15","OVERALL LIFE SCORE 2026",DARK_NAVY,WHITE,11,bold=True)
    ws.row_dimensions[15].height=26
    mc("B16:H17","='Monthly Summary'!K16",DARK_NAVY,WHITE,28,bold=True,fmt="0.0%")
    ws.row_dimensions[16].height=38; ws.row_dimensions[17].height=20
    mc("B18:H18","Salat 25%  +  Gym 20%  +  Supplements 15%  +  Steps 15%  +  Running 10%  +  Savings 15%",
       LIGHT_BLUE,BLUE,9,italic=True)
    ws.row_dimensions[18].height=20

    ws.row_dimensions[19].height=6
    mc("B20:H20","MONTHLY PERFORMANCE AT A GLANCE",TEAL,WHITE,11,bold=True)
    ws.row_dimensions[20].height=26

    for col_l,h in zip(["B","C","D","E","F","G","H"],
                        ["Month","Salat %","Gym %","Suppl.%","Steps %","Running KM","Overall"]):
        c=ws[f"{col_l}21"]; c.value=h; c.fill=hfill(DARK_NAVY)
        c.font=mfont(bold=True,color=WHITE,size=9); c.alignment=center(); c.border=tborder()
    ws.row_dimensions[21].height=24

    for i,month in enumerate(MONTHS):
        r=22+i; src=4+i; bg=WHITE if i%2==0 else GRAY_BG
        ws.row_dimensions[r].height=20
        for col_l,val,fmt in [
            ("B",month,None),
            ("C",f"='Monthly Summary'!B{src}","0%"),
            ("D",f"='Monthly Summary'!C{src}","0%"),
            ("E",f"='Monthly Summary'!D{src}","0%"),
            ("F",f"='Monthly Summary'!E{src}","0%"),
            ("G",f"='Monthly Summary'!F{src}","0.0"),
            ("H",f"='Monthly Summary'!K{src}","0%"),
        ]:
            c=ws[f"{col_l}{r}"]; c.value=val; c.fill=hfill(bg)
            c.font=mfont(size=9); c.alignment=center(); c.border=tborder()
            if fmt: c.number_format=fmt

    ws.conditional_formatting.add("H22:H33",
        ColorScaleRule(start_type='min',start_color=RED,
                       mid_type='percentile',mid_value=50,mid_color=YELLOW,
                       end_type='max',end_color=GREEN))

# ══════════════════════════════════════════════════════════════════════════════
# POST-PROCESSING: inject Excel 365 native checkboxes via ZIP/XML
# ══════════════════════════════════════════════════════════════════════════════

# Excel 365 native checkbox format:
# Cells store boolean 0/1. The sheet extLst gets an x14:dataValidation type="checkBox"
# which tells Excel to render those cells as clickable checkbox controls.

CB_URI = "{CCF9B9E7-4FED-4B87-89B2-7A4E5FC3597B}"
X14_NS = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"
XM_NS  = "http://schemas.microsoft.com/office/excel/2006/main"

def _make_cb_ext(sqref: str) -> str:
    return (
        f'<ext uri="{CB_URI}" xmlns:x14="{X14_NS}">'
        f'<x14:dataValidations>'
        f'<x14:dataValidation type="checkBox">'
        f'<xm:sqref xmlns:xm="{XM_NS}">{sqref}</xm:sqref>'
        f'</x14:dataValidation>'
        f'</x14:dataValidations>'
        f'</ext>'
    )

def inject_checkboxes(xlsx_path: str, sheet_sqref: dict):
    """
    Inject Excel 365 native checkbox DV into worksheet XML files inside the xlsx ZIP.
    sheet_sqref: {sheet_name: sqref_string}  e.g. {"Salat Tracker": "C4:G368"}
    """
    buf = BytesIO()
    with zipfile.ZipFile(xlsx_path, 'r') as zin:
        # Map sheet names → worksheet XML file paths
        wb_xml  = zin.read('xl/workbook.xml').decode('utf-8')
        rels_xml = zin.read('xl/_rels/workbook.xml.rels').decode('utf-8')

        rid_target = {}
        # Rels XML has attributes in order: Type Target Id — extract each independently
        for m in re.finditer(r'<Relationship[^>]+>', rels_xml):
            tag = m.group(0)
            mid = re.search(r'Id="([^"]+)"', tag)
            mtg = re.search(r'Target="([^"]+)"', tag)
            if mid and mtg:
                rid_target[mid.group(1)] = mtg.group(1)

        name_file = {}
        for m in re.finditer(r'<sheet[ ][^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wb_xml):
            sname, rid = m.group(1), m.group(2)
            # Decode XML entity encoding in sheet names (e.g. &amp; → &)
            sname = sname.replace('&amp;', '&').replace('&lt;', '<').replace('&gt;', '>')
            if rid in rid_target:
                t = rid_target[rid]
                # Strip leading slash — ZIP entries don't have it
                name_file[sname] = t.lstrip('/')

        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)

                matched = next((s for s,f in name_file.items()
                                if item == f and s in sheet_sqref), None)
                if matched:
                    xml = data.decode('utf-8')
                    ext_xml = _make_cb_ext(sheet_sqref[matched])
                    if '<extLst>' in xml:
                        xml = xml.replace('</extLst>', ext_xml + '</extLst>', 1)
                    else:
                        xml = xml.replace('</worksheet>',
                                          f'<extLst>{ext_xml}</extLst></worksheet>', 1)
                    data = xml.encode('utf-8')

                zout.writestr(item, data)

    with open(xlsx_path, 'wb') as f:
        f.write(buf.getvalue())

# ══════════════════════════════════════════════════════════════════════════════
# SHEET 10 — TASK TRACKER
# ══════════════════════════════════════════════════════════════════════════════
TASK_ROWS = 50   # number of task rows to pre-create

def create_task_sheet(wb):
    ws = wb.create_sheet("Task Tracker")
    headers    = ["Task Description", "Assigned To", "Check", "Status"]
    col_widths = [46, 24, 12, 18]

    # ── Title ────────────────────────────────────────────────────────────────
    ws.merge_cells("A1:D1")
    ws["A1"].value = "☑  TASK TRACKER"
    ws["A1"].fill  = hfill(DARK_NAVY)
    ws["A1"].font  = mfont(bold=True, size=14, color=WHITE)
    ws["A1"].alignment = center()
    ws.row_dimensions[1].height = 38

    # ── Subtitle ─────────────────────────────────────────────────────────────
    ws.merge_cells("A2:D2")
    ws["A2"].value = "Tap Check ✓ to mark complete  •  Status updates automatically"
    ws["A2"].fill  = hfill("ECEFF1")   # soft blue-grey — 2nd colour only
    ws["A2"].font  = mfont(italic=True, color=DARK_GRAY, size=10)
    ws["A2"].alignment = center()
    ws.row_dimensions[2].height = 22

    # ── Column headers ────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 30
    for c, (h, w) in enumerate(zip(headers, col_widths), 1):
        ws.cell(row=3, column=c, value=h)
        set_w(ws, c, w)
    style_header(ws, 3, DARK_NAVY, WHITE, range(1, 5))

    # ── Data rows ─────────────────────────────────────────────────────────────
    last = 3 + TASK_ROWS
    for i in range(TASK_ROWS):
        row = i + 4
        bg  = WHITE if i % 2 == 0 else GRAY_BG
        ws.cell(row=row, column=1, value="")    # Task Description
        ws.cell(row=row, column=2, value="")    # Assigned To
        ws.cell(row=row, column=3, value=0)     # Check: 0 = unchecked
        ws.cell(row=row, column=4,
                value=f'=IF(A{row}="","",IF(C{row}=1,"Complete","In Progress"))')
        style_data(ws, row, bg, range(1, 5))
        # Check cell: neutral default — CF turns it red/green based on content
        cell = ws.cell(row=row, column=3)
        cell.number_format = ';;;'   # hide 0/1 value; color comes from CF
        cell.alignment = center()
        ws.row_dimensions[row].height = 26

    # ── Conditional formatting ────────────────────────────────────────────────
    # Check column: red only when a task exists and is unchecked
    ws.conditional_formatting.add(f"C4:C{last}",
        FormulaRule(formula=['AND(A4<>"",C4=0)'],
                    fill=hfill(CB_RED),
                    font=Font(color=CB_RED_FONT, name="Calibri")))
    # Check column: green when checked
    ws.conditional_formatting.add(f"C4:C{last}",
        FormulaRule(formula=['C4=1'],
                    fill=hfill(CB_GREEN),
                    font=Font(color=CB_GRN_FONT, bold=True, name="Calibri")))

    # Status column: green for Complete, soft orange for In Progress
    ws.conditional_formatting.add(f"D4:D{last}",
        FormulaRule(formula=['D4="Complete"'],
                    fill=hfill(CB_GREEN),
                    font=Font(color=CB_GRN_FONT, bold=True, name="Calibri")))
    ws.conditional_formatting.add(f"D4:D{last}",
        FormulaRule(formula=['D4="In Progress"'],
                    fill=hfill("FFF3E0"),
                    font=Font(color="E65100", bold=True, name="Calibri")))

    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:D{last}"

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
    create_task_sheet(wb)

    tab_colors = {
        "Dashboard":        DARK_NAVY, "Salat Tracker":    TEAL,
        "Gym Routine":      PURPLE,    "Supplements":      ORANGE,
        "Steps Tracker":    BLUE,      "Running Tracker":  GREEN,
        "Expense & Savings":PINK,      "Goals & Settings": DARK_NAVY,
        "Monthly Summary":  YELLOW,    "Task Tracker":     DARK_GRAY,
    }
    for name in wb.sheetnames:
        if name in tab_colors:
            wb[name].sheet_properties.tabColor = tab_colors[name]

    path = "/home/user/habit-tracker/Personal_Life_Tracker.xlsx"
    wb.save(path)
    print(f"✅  Saved initial workbook → {path}")

    # Inject native Excel 365 checkboxes via ZIP post-processing
    last_row = 3 + len(DATES)   # = 368
    inject_checkboxes(path, {
        "Salat Tracker":   f"C4:G{last_row}",         # 5 prayer columns
        "Gym Routine":     f"D4:D{last_row}",         # Workout Done
        "Supplements":     f"C4:D{last_row}",         # Multivitamin + Omega
        "Running Tracker": f"C4:C{last_row}",         # Ran Today
        "Task Tracker":    f"C4:C{3 + TASK_ROWS}",   # Check column
    })
    print("✅  Injected Excel 365 native checkboxes")
    print(f"   Sheets  : {wb.sheetnames}")
    print(f"   Behavior: FALSE (unchecked) = RED  •  TRUE (checked) = GREEN")
    print(f"   Tap any prayer/supplement/gym/run cell to toggle instantly!")

if __name__ == "__main__":
    main()
