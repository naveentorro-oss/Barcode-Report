# Barcode-Report

This project downloads an Excel report from a website and builds a pivot summary in a new Excel workbook.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

### Streamlit app

Install the requirements, then start the app:

```bash
streamlit run app.py
```

Upload a CSV or Excel report, choose the fields and Actioned By values to include, and select **Generate report**. Use **Download Excel report** to download the completed workbook directly from the app. For barcode validation reports, `Time Frame` is derived from `Time Taken From Last Action` when it is not already present, and the summary retains the reference workbook layout when the matching fields are selected.

From a local Excel file:

```bash
python barcode_report.py \
  --source-file "source_report.xlsx" \
  --output "pivot_report.xlsx" \
  --index "Barcode" \
  --columns "Category" \
  --values "Qty" \
  --aggfunc "sum"
```

From a public website URL:

```bash
python barcode_report.py \
  --source-url "https://example.com/report.xlsx" \
  --output "pivot_report.xlsx" \
  --index "Barcode" \
  --columns "Category" \
  --values "Qty" \
  --aggfunc "sum"
```

The script saves a sheet named `Pivot` inside the output workbook. The sheet contains the summary table that Excel can open and manipulate.

## Notes

- This creates a pivot-style summary table in a new Excel file.
- Works best with a direct Excel download URL from the report source.
- If you want a native Excel PivotTable object created by Excel itself, use a Windows environment with Excel/pywin32 automation.
