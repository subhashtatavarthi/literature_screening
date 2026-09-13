# Literature Screening & Retrieval Pipeline

## What it does

1. Reads `UI`, `Title`, and `DOI` from an input Excel workbook.
2. Creates/replaces the `Titles for Screening` sheet.
3. Highlights exact duplicate DOI/title records.
4. Resolves DOI metadata with Crossref.
5. Checks legitimate open-access locations through Unpaywall.
6. Downloads available PDFs and names them `<UI>.pdf`.
7. Produces:
   - `Screening_Output.xlsx`
   - `Download_Log.xlsx`
   - `Purchase_List.xlsx`
   - `Literature_PDFs.zip`

## Setup

Python 3.10+ recommended.

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

### Windows

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

## Input

Place your Excel file at:

`input/literature.xlsx`

It must contain these columns:

- `UI`
- `Title`
- `DOI`

If your workbook uses different column names, edit the three column constants near the top of `literature_pipeline.py`.

## Run

Unpaywall requires an email parameter. Set it before running.

macOS / Linux:

```bash
export UNPAYWALL_EMAIL="your.email@example.com"
python literature_pipeline.py
```

Windows PowerShell:

```powershell
$env:UNPAYWALL_EMAIL="your.email@example.com"
python literature_pipeline.py
```

## Output statuses

- `DOWNLOADED`
- `PURCHASE_REQUIRED`
- `OA_DOWNLOAD_FAILED`
- `MANUAL_REVIEW`
- `ERROR`

## Important note about purchase prices

The pipeline does not invent article prices. Publisher prices may depend on region,
institution, authentication, or interactive checkout. The purchase workbook records
the publisher/DOI purchase page and marks the cost as `Price not publicly available`
when no verified price has been obtained.

## API / access behavior

The tool uses Crossref for DOI metadata and Unpaywall for legitimate open-access
locations. It does not intentionally bypass publisher paywalls or access controls.
