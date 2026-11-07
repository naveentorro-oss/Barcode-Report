# Barcode-Report

Download a report from a website or upload a CSV/Excel file, then create a pivot-style summary workbook.

## Run locally

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
streamlit run app.py
```

Upload a CSV or Excel report, choose the row, column, value, and aggregation fields, then select **Generate report**. Use **Download Excel report** to save the generated workbook. If the file contains an `Actioned By` column, you can optionally filter its values.

## Deploy on Streamlit Community Cloud

1. Push this repository to GitHub.
2. In [Streamlit Community Cloud](https://share.streamlit.io/), create a new app and select this repository and branch.
3. Set the **Main file path** to `app.py` and deploy. Cloud installs the packages listed in `requirements.txt`.

## Command-line usage

From a local file:

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

This produces a pivot-style summary table, not a native Excel PivotTable object.
