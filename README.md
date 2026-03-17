# Offering Project

## Overview

This project extracts and consolidates donation data from email confirmations across three payment platforms — Zelle, CashApp, and PayPal — and exports the structured results to a formatted, multi-tab Excel workbook for record-keeping and analysis.

Emails are exported from a mail client as a CSV file and processed through a two-script pipeline: extraction/transformation, then Excel export.

> **Privacy Notice:** Input data (email confirmations) and output data (exported Excel file) are omitted from this repository for privacy purposes.

---

## Pipeline

```
Raw_Data/
└── March_03_2026.csv          ← CSV export of donation emails
        │
        ▼
extract_transform.py           ← Parse & clean each platform's emails
        │
        ▼
export.py                      ← Write 3-tab formatted Excel workbook
        │
        ▼
Output/offerings_export.xlsx   ← Final output
```

---

## Scripts

### `extract_transform.py` — Extraction & Transformation

Reads the raw email CSV (`Raw_Data/March_03_2026.csv`) with columns: `Title`, `Payment_Type`, `Receiver_email`, `Date`, `Star`, `Reference_num`, `Email_body`.

Filters down to three payment senders and processes each separately:

**Zelle** (sender: `pncalerts@pnc.com`)
- Keeps emails with subject containing `"sent you a Zelle® payment"`
- Extracts via regex: `full_name`, `date`, `time`, `amount`, `note`
- Applies `time_extract()` to compute a `realized_date`:
  - If the payment was received on a business day at or before 9:59 PM → same day
  - Otherwise (weekend, US federal holiday, or after cutoff) → next business day
- Uses `USFederalHolidayCalendar` and `CustomBusinessDay` from pandas for holiday-aware logic

**PayPal** (sender: `service@paypal.com`)
- Keeps emails with subject containing `"You've got money"`
- Extracts: `full_name`, `Date`, `Time`, `amount_received`, `Fee`, `Total`, `Transaction_id`

**CashApp** (sender: `cash@square.com`)
- Keeps emails with subject containing `"Payment received"`
- Extracts: `full_name`, `Date`, `Time`, `amount_received`, `Note`, `Transaction_id`

---

### `export.py` — Excel Export

Takes the three cleaned DataFrames and writes them to a single 3-tab Excel workbook using openpyxl.

| Tab | Color | Columns |
|-----|-------|---------|
| Zelle | Purple (`#4A235A`) | Date, Amount, Sender Name, Memo |
| CashApp | Green (`#00C244`) | Date, Amount, Sender Name, Cashtag, Transaction ID, Note |
| PayPal | Blue (`#003087`) | Date, Gross Amount, Fee, Net Amount, Sender Name, Sender Email, Transaction ID, Description, Status |

**Formatting applied to every tab:**
- Bold white headers on platform-branded background
- Alternating row shading for readability
- Currency columns formatted as `$#,##0.00` with a live `=SUM()` totals row
- Date columns center-aligned
- Auto-fitted column widths (max 40 characters)
- Frozen header row

---

## Project Structure

```
Offering_Project/
├── Raw_Data/
│   └── March_03_2026.csv      # Email CSV export (omitted for privacy)
├── Output/
│   └── offerings_export.xlsx  # Generated workbook (omitted for privacy)
├── extract_transform.py       # Email parsing & transformation
├── export.py                  # 3-tab Excel export
├── analysis.py                # Prototype / scratch script
├── .gitignore
└── README.md
```

---

## How to Run

1. Install dependencies:
   ```bash
   pip install pandas openpyxl
   ```

2. Place your email CSV export in `Raw_Data/` and update the filename reference in `extract_transform.py` if needed.

3. Run extraction and transformation:
   ```bash
   python extract_transform.py
   ```

4. Run the export:
   ```bash
   python export.py
   ```

   The formatted Excel file will be saved to `Output/offerings_export.xlsx`.

---

## Tools and Technologies

- Python 3.8+
- pandas — data loading, filtering, and transformation
- `re` — regex-based field extraction from email body text
- `pandas.tseries` — US federal holiday calendar and business day offset logic
- openpyxl — Excel workbook creation and formatting
