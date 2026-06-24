# Personal Life Tracker 2026

A complete personal habit & finance tracking workbook for the full year 2026, built with Python + openpyxl.

## Sheets

| Sheet | What it tracks |
|---|---|
| **Dashboard** | KPI cards, monthly mini-table, live year-to-date totals, color-scale scoring |
| **Salat Tracker** | 5 daily prayers (Fajr, Dhuhr, Asr, Maghrib, Isha) with completion % |
| **Gym Routine** | 5 Grind Days / 2 Rest Days per week, workout name, duration, intensity |
| **Supplements** | Multivitamin + Omega — both-completed auto-formula |
| **Steps Tracker** | Daily step count vs 10,000-step goal, deficit/surplus |
| **Running Tracker** | Distance, duration, auto-calculated pace (min/km), monthly totals |
| **Expense & Savings** | Income / Expense / Savings per day, category & payment method dropdowns |
| **Goals & Settings** | Single place to edit all goals — referenced by every other sheet |
| **Monthly Summary** | Per-month aggregates + weighted Overall Life Score, bar chart |
| **Task Tracker** | 50-row task list with auto-status (Complete / In Progress) |

## Overall Life Score Formula

```
Score = Salat 25% + Gym 20% + Supplements 15% + Steps 15% + Running 10% + Savings 15%
```

---

## Using in Microsoft Excel

1. Open `Personal_Life_Tracker.xlsx` directly in Excel 365.
2. Checkbox cells (prayers, gym, supplements, running, tasks) are **native click-to-toggle** — tap once to check (green), tap again to uncheck (red).
3. Enter step counts, distances, and expenses in the number columns.
4. Check the **Dashboard** tab for live scores.

---

## Using in Google Sheets (one-time setup)

### Step 1 — Upload and convert
1. Go to [drive.google.com](https://drive.google.com)
2. Drag and drop `Personal_Life_Tracker.xlsx` onto Drive
3. Right-click the uploaded file → **Open with → Google Sheets**
   *(This converts it to a native Sheets file — do NOT just open as xlsx viewer)*

### Step 2 — Add tap-to-toggle checkboxes (run once, ~30 seconds)
1. In the spreadsheet: click **Extensions → Apps Script**
2. Delete all the default code in the editor
3. Open [`add_gsheets_checkboxes.gs`](add_gsheets_checkboxes.gs) from this repo, copy all the code, paste it into the editor
4. Click the **▶ Run** button
5. If a permissions dialog appears: click **Review permissions → Advanced → Go to (unsafe) → Allow**
6. Click **▶ Run** again — a popup confirms checkboxes are active
7. Close the Apps Script tab

### Step 3 — Done
Every checkbox cell in Salat, Gym, Supplements, Running, and Task Tracker is now a **single-tap toggle**:
- **Red** = unchecked / not done
- **Green** = checked / complete

> The script runs once — checkboxes are saved permanently in the file.

### Google Sheets Tips
- Pin header rows: **View → Freeze → Up to row 3** on each sheet
- Change any goal in **Goals & Settings** — all sheets update automatically
- The **Monthly Summary** chart updates as you fill data

---

## Gym Schedule (auto-filled)

| Day | Type |
|---|---|
| Monday | Grind Day |
| Tuesday | Grind Day |
| Wednesday | Grind Day |
| Thursday | **Rest Day** |
| Friday | Grind Day |
| Saturday | Grind Day |
| Sunday | **Rest Day** |

---

## Regenerate from source

```bash
pip install openpyxl
python generate_tracker.py
# → Personal_Life_Tracker.xlsx
```

## File Structure

```
habit-tracker/
├── generate_tracker.py            # Python script to regenerate the workbook
├── Personal_Life_Tracker.xlsx     # Ready-to-use Excel workbook (Excel 365 + Google Sheets)
├── add_gsheets_checkboxes.gs      # Paste into Apps Script for Google Sheets tap-to-toggle
└── README.md
```
