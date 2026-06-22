# Personal Life Tracker 2026

A complete personal habit & finance tracking workbook for the full year 2026, built with Python + openpyxl.

## Features

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

## Overall Life Score Formula

```
Score = Salat 25% + Gym 20% + Supplements 15% + Steps 15% + Running 10% + Savings 15%
```

## How to generate the workbook

```bash
pip install openpyxl
python generate_tracker.py
# → Personal_Life_Tracker.xlsx
```

## Daily Usage

1. **Open the workbook** in Excel or upload to Google Sheets.
2. **Go to today's sheet** (Salat, Gym, etc.) and find today's row — it is pre-dated for the whole year.
3. **Click the cell** in prayer/supplement/workout columns and pick `TRUE` or `FALSE` from the dropdown.
4. **Enter numbers** (Steps, Distance, Expenses) directly.
5. **Check the Dashboard** anytime to see your live scores and trends.

### Tips
- Change goals anytime in **Goals & Settings** — all sheets update automatically.
- The **Monthly Summary** chart updates as you fill in data.
- Works in both **Microsoft Excel** and **Google Sheets** (upload the .xlsx file to Google Drive and open with Sheets).

## Google Sheets — One-Time Checkbox Setup

After uploading the `.xlsx` to Google Drive, open it **as Google Sheets** (right-click → Open with → Google Sheets).  
Cells will show as empty colored boxes (red = unchecked, green = checked) — no text visible.

**To enable tap-to-toggle checkboxes (run once):**

1. In the spreadsheet click **Extensions → Apps Script**
2. Delete all code in the editor, paste the contents of [`add_gsheets_checkboxes.gs`](add_gsheets_checkboxes.gs)
3. Click **Run** → grant permissions → click **Run** again
4. A dialog confirms checkboxes are added — close it

After that, every red/green cell in Salat, Gym, Supplements, and Running toggles on a single tap.

**Tips:**
- Pin rows with **View → Freeze → Up to row 3** on each sheet.
- The script only needs to run once — checkboxes persist in the file.

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

## File Structure

```
habit-tracker/
├── generate_tracker.py            # Python script to regenerate the workbook
├── Personal_Life_Tracker.xlsx     # The ready-to-use Excel workbook
├── add_gsheets_checkboxes.gs      # Google Apps Script — run once for tap-to-toggle
└── README.md
```
