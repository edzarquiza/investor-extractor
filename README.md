# Capspace Statement Extractor

**Built by [Dexterous](https://dexterous.com.au)**

A web-based automation tool that replaces the manual Excel macro workflow for Capspace monthly reconciliations. Upload raw statement files and download clean, formatted Excel outputs in seconds — no Excel, no macros, no manual copy-paste.

---

## What it does

The app has two tools in one, switchable via tabs:

### 📋 Unit Register
Extracts investor data from monthly **Investor Statement** files (CPDF, DLOT, CDLOT2).

**Input:** One or more raw `Investor_Statements_*.xlsx` files
**Output:** `Combined_Statement_Extracted.xlsx` with columns:

| Entity | Investor | Entity \| Investor | Balance |
|--------|----------|--------------------|---------|
| CPDF | Skoufa Pty Ltd ATF SRA Skoufis Super Fund | CPDF \| Skoufa Pty Ltd... | 464,292.95 |

- Supports multiple files at once — all investors combined into one sheet
- Fund code (CPDF / DLOT / CDLOT2) auto-detected from filename
- Investor name corrections applied automatically (typos, spelling variants)

---

### 🏦 Loan Register
Extracts borrower data from the monthly **All Statements Capspace Loans** file.

**Input:** One `All_Statements_Capspace_Loans_*.xlsx` file
**Output:** `Loan_Register_Extracted.xlsx` with columns:

| Entity | Borrower | Statement Balance | Interest for the month | Reserve Balance |
|--------|----------|-------------------|------------------------|-----------------|
| CPDF | 659 - Loan - Crowdy Bay Trust | 2,661,000.00 | 25,151.92 | - |

- Statement month is **auto-detected** from transaction dates in the file — no manual selection needed
- Borrower names mapped to standardised loan codes via built-in Master map
- Handles multi-tranche loans (same borrower appearing multiple times) as separate rows

---

## How to use

1. Open the app URL in any browser
2. Click the relevant tab — **Unit Register** or **Loan Register**
3. Upload your file(s) by dragging and dropping or clicking Browse
4. Click **Extract**
5. Review the summary stats and preview table
6. Click **Download** to save the output Excel file

> ⚠️ Always upload the **raw monthly statement file** — not a previously extracted output file.

---

## File naming conventions

The app detects the fund automatically from the filename. Make sure the filename contains one of these codes:

| Fund Code | Example Filename |
|-----------|-----------------|
| `CPDF` | `Investor_Statements_CPDF_November_2025.xlsx` |
| `DLOT` | `Investor_Statements_-_DLOT_-_September_2025.xlsx` |
| `CDLOT2` | `Investor_Statements_CDLOT2_February_2026.xlsx` |

For the Loan Register, any filename is accepted — the tool reads borrower data from the file content directly.

---

## Deployment

This app is built with [Streamlit](https://streamlit.io) and hosted on Streamlit Cloud.

### Files in this repo

```
app.py              # Main application
requirements.txt    # Python dependencies
README.md           # This file
```

### Requirements

```
streamlit
pandas
openpyxl
xlsxwriter
```

### How to update

1. Make changes to `app.py` locally or edit directly on GitHub
2. Commit the changes to the `main` branch
3. Streamlit Cloud detects the change and redeploys automatically (takes ~1–2 minutes)
4. Refresh the app URL to see the updated version

### How to deploy from scratch

1. Fork or clone this repo to your GitHub account
2. Log in to [share.streamlit.io](https://share.streamlit.io)
3. Click **New app**
4. Select your GitHub repo, branch (`main`), and set main file to `app.py`
5. Click **Deploy**

---

## Updating the Master maps

The app has two hardcoded name correction maps:

- **Unit Register** — corrects investor name typos (e.g. "Supperannuation" → "Superannuation")
- **Loan Register** — maps raw borrower names to standardised loan codes (e.g. "Hayden Consulting Group Pty Ltd" → "709 - Loan - Hayden Consulting Pty Ltd")

To add or update an entry, edit the relevant dictionary in `app.py`:

```python
# Unit Register corrections — in UNIT_MASTER dict
"Wrong Name Here": "Correct Name Here",

# Loan Register mappings — in LOAN_MASTER dict
"Raw Borrower Name from Statement": ("CPDF", "000 - Loan - Short Name"),
```

Then commit `app.py` to GitHub and the app will redeploy with the updated map.

---

## Adding a new fund

To support a new fund code in the Unit Register:

1. Open `app.py`
2. Find the `UNIT_FUND_CODES` list near the top
3. Add the new code — **place longer codes before shorter ones** to avoid partial matches:

```python
UNIT_FUND_CODES = ["CDLOT2", "CPDF", "DLOT", "CDLOT", "NEWFUND"]
```

4. Commit and the app redeploys automatically.

---

## Notes

- The app runs entirely in the browser — uploaded files are not stored anywhere
- All processing happens in-memory; nothing is saved to a server
- Output filenames are based on the input filename (Unit Register) or a fixed name (Loan Register)
- The Loan Register interest figure represents the **last positive interest payment** in the most recent payment due month found in the file
- Loans with no payment in the current month will show interest as `0` (e.g. a loan that paid early and skipped a month)

---

*For issues or feature requests, contact the Dexterous team.*
