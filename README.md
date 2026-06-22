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

## Google Sheets Tips

When uploading to Google Sheets:
- TRUE/FALSE dropdowns work natively.
- For real checkboxes: select the prayer/supplement columns → **Insert → Checkbox** — Sheets maps TRUE/FALSE automatically.
- Pin rows with **View → Freeze → Up to row 3** on each sheet.
- Use **Format → Conditional formatting** to add mobile-friendly color rules.

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
├── generate_tracker.py        # Python script to regenerate the workbook
├── Personal_Life_Tracker.xlsx # The ready-to-use Excel workbook
└── README.md
```
