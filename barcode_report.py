from __future__ import annotations

import argparse
import re
from copy import copy
from pathlib import Path
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter


def extract_excel_download_url(page_url: str, html: str) -> str | None:
    """Find the likely Excel export URL from a report page."""
    matches = re.findall(r'''(?:href|src)=["']([^"']+)["']''', html, flags=re.IGNORECASE)
    for candidate in matches:
        normalized = candidate.strip()
        if not normalized or normalized.startswith(("javascript:", "mailto:", "#")):
            continue
        full_url = urljoin(page_url, normalized)
        lowered = full_url.lower()
        if any(token in lowered for token in (".xlsx", ".xls", ".xlsm", ".csv", "download", "export")):
            return full_url
    return None


def _save_downloaded_file(response: requests.Response, destination_dir: str | Path, explicit_url: str | None = None) -> Path:
    target_dir = Path(destination_dir)
    target_dir.mkdir(parents=True, exist_ok=True)

    url = explicit_url or response.url
    file_name = Path(urlparse(url).path).name or "downloaded_report.xlsx"
    if not file_name.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv")):
        file_name = f"{file_name}.xlsx"

    local_path = target_dir / file_name
    local_path.write_bytes(response.content)
    return local_path


def _seconds_from_time_value(value: object) -> int:
    if pd.isna(value):
        return 0
    if isinstance(value, str):
        value = value.strip()
        if not value:
            return 0
        parts = [int(part) for part in value.split(":")]
        if len(parts) == 2:
            return parts[0] * 60 + parts[1]
        if len(parts) == 3:
            return parts[0] * 3600 + parts[1] * 60 + parts[2]
    return int(value)


def _derive_time_frame_column(df: pd.DataFrame) -> pd.DataFrame:
    if "Time Frame" in df.columns:
        return df

    time_column = None
    for candidate in df.columns:
        lowered = str(candidate).lower()
        if "time taken" in lowered or "timeframe" in lowered or "time frame" in lowered:
            time_column = candidate
            break

    if time_column is None:
        return df

    df = df.copy()
    df["Time Frame"] = df[time_column].apply(
        lambda value: "15 to 30" if 15 <= _seconds_from_time_value(value) <= 30 else "Less than 15"
    )
    return df


def load_excel_data(source_file: str | Path) -> pd.DataFrame:
    """Read an existing spreadsheet file and return its rows as a DataFrame."""
    source = Path(source_file)
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")

    if source.suffix.lower() == ".csv":
        df = pd.read_csv(source)
    else:
        df = pd.read_excel(source)

    return _derive_time_frame_column(df)


def download_excel_report(url: str, destination_dir: str | Path) -> Path:
    """Download an Excel report from a website and save it locally."""
    if not url.startswith(("http://", "https://")):
        raise ValueError("The source URL must start with http:// or https://")

    session = requests.Session()
    response = session.get(url, timeout=60, allow_redirects=True)
    response.raise_for_status()

    content_type = response.headers.get("Content-Type", "").lower()
    if (
        "excel" in content_type
        or "csv" in content_type
        or url.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv"))
        or response.url.lower().endswith((".xlsx", ".xls", ".xlsm", ".csv"))
        or response.content.startswith(b"PK")
    ):
        return _save_downloaded_file(response, destination_dir, response.url)

    download_url = extract_excel_download_url(response.url, response.text)
    if not download_url:
        raise ValueError(
            "Could not find an Excel export/download link on the provided website page. "
            "Use a direct .xlsx/.xls/.csv URL or a report page that exposes an export link."
        )

    export_response = session.get(download_url, timeout=60, allow_redirects=True)
    export_response.raise_for_status()
    return _save_downloaded_file(export_response, destination_dir, download_url)


def _write_reference_summary_sheet(df: pd.DataFrame, work_book: Workbook) -> None:
    ws_data = work_book.active
    ws_data.title = "data"
    ws_data.append(list(df.columns))
    for row in df.itertuples(index=False, name=None):
        ws_data.append(row)

    ws_summary = work_book.create_sheet("Sheet1")
    ws_summary["A1"] = "Actioned By"
    ws_summary["B1"] = "(Multiple Items)"

    allowed_actioned_by = [
        str(value).strip()
        for value in sorted({str(v).strip() for v in df["Actioned By"].dropna().astype(str)})
        if str(value).strip()
    ]
    if not allowed_actioned_by:
        allowed_actioned_by = ["nr3", "mhussain3"]

    actioned_by_validation = DataValidation(
        type="list",
        formula1=f'"{",".join(allowed_actioned_by)}"',
        allow_blank=False,
    )
    actioned_by_validation.prompt = "Select an Actioned By value"
    actioned_by_validation.error = "Please choose one of the allowed Actioned By values."
    actioned_by_validation.add("B1")
    ws_summary.add_data_validation(actioned_by_validation)

    ws_summary["A3"] = "Count of Ticket Number"
    ws_summary["B3"] = "Time Frame"
    ws_summary["A4"] = "Ticket Status"
    ws_summary["B4"] = "15 to 30"
    ws_summary["C4"] = "Less than 15"
    ws_summary["D4"] = "Grand Total"

    pivot = pd.pivot_table(
        df,
        index="Ticket Status",
        columns="Time Frame",
        values="Ticket Number",
        aggfunc="count",
        fill_value=0,
    )

    row_num = 5
    for status in pivot.index:
        status_row = pivot.loc[status]
        val_15_30 = int(status_row.get("15 to 30", 0))
        val_lt_15 = int(status_row.get("Less than 15", 0))
        grand_total = val_15_30 + val_lt_15
        ws_summary[f"A{row_num}"] = status
        ws_summary[f"B{row_num}"] = val_15_30
        ws_summary[f"C{row_num}"] = val_lt_15
        ws_summary[f"D{row_num}"] = grand_total
        row_num += 1

    grand_15_30 = int(pivot["15 to 30"].sum()) if "15 to 30" in pivot.columns else 0
    grand_lt_15 = int(pivot["Less than 15"].sum()) if "Less than 15" in pivot.columns else 0
    ws_summary[f"A{row_num}"] = "Grand Total"
    ws_summary[f"B{row_num}"] = grand_15_30
    ws_summary[f"C{row_num}"] = grand_lt_15
    ws_summary[f"D{row_num}"] = grand_15_30 + grand_lt_15

    ws_summary["D10"] = "TAT Achived"
    ws_summary["E10"] = round((grand_lt_15 / len(df)) * 100, 1) if len(df) else 0.0

    for row in ws_summary.iter_rows(min_row=1, max_row=ws_summary.max_row, min_col=1, max_col=ws_summary.max_column):
        for cell in row:
            if cell.row == 1 or cell.row == 3 or cell.row == 4:
                font = copy(cell.font)
                font.bold = True
                cell.font = font

    for column_cells in ws_summary.columns:
        max_length = max(len(str(cell.value)) if cell.value is not None else 0 for cell in column_cells)
        ws_summary.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max_length + 2, 28)


def build_pivot_table_sheet(
    source_path: str | Path,
    output_path: str | Path,
    index: str,
    columns: str,
    values: str,
    aggfunc: str = "sum",
    sheet_name: str = "Pivot",
    actioned_by_filter: list[str] | None = None,
) -> Path:
    """Create a pivot table summary sheet in a new Excel workbook."""
    source_file = Path(source_path)
    output_file = Path(output_path)

    if not source_file.exists():
        raise FileNotFoundError(f"Source file not found: {source_file}")

    df = load_excel_data(source_file)

    if actioned_by_filter:
        filter_values = {str(v).strip().lower() for v in actioned_by_filter}
        if "Actioned By" in df.columns:
            df = df[df["Actioned By"].fillna("").astype(str).str.strip().str.lower().isin(filter_values)].copy()

    required_fields = {index, columns, values}
    missing = sorted(required_fields - set(df.columns))
    if missing:
        raise ValueError(f"Missing required columns: {', '.join(missing)}")

    pivot_df = pd.pivot_table(
        df,
        index=index,
        columns=columns,
        values=values,
        aggfunc=aggfunc,
        fill_value=0,
    ).reset_index()

    output_file.parent.mkdir(parents=True, exist_ok=True)

    if actioned_by_filter and "Actioned By" in df.columns and "Time Frame" in df.columns and index == "Ticket Status":
        wb = Workbook()
        _write_reference_summary_sheet(df, wb)
        wb.save(output_file)
        return output_file

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Data", index=False)
        pivot_df.to_excel(writer, sheet_name=sheet_name, index=False)

    return output_file


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Download an Excel report from a website and create a pivot summary in a new Excel workbook."
    )
    parser.add_argument("--source-url", help="Public URL of the source Excel file to download.")
    parser.add_argument("--source-file", help="Local path to the source Excel file if already downloaded.")
    parser.add_argument("--download-dir", default="downloads", help="Directory to save downloaded files when using --source-url.")
    parser.add_argument("--output", required=True, help="Output Excel file for the pivot summary.")
    parser.add_argument("--index", required=True, help="Column used as the row index in the pivot table.")
    parser.add_argument("--columns", required=True, help="Column used as the pivot columns.")
    parser.add_argument("--values", required=True, help="Numeric column to aggregate.")
    parser.add_argument("--aggfunc", default="sum", help="Aggregation function for the pivot table (default: sum).")
    parser.add_argument("--sheet-name", default="Pivot", help="Name of the worksheet to store the pivot summary.")
    parser.add_argument(
        "--actioned-by-filter",
        nargs="*",
        default=["nr3", "mhussain3"],
        help="Optional list of Actioned By values to keep in the report. Default: nr3 mhussain3",
    )
    return parser


def main() -> None:
    parser = _build_parser()
    args = parser.parse_args()

    if bool(args.source_url) == bool(args.source_file):
        parser.error("Use exactly one of --source-url or --source-file.")

    source_file = args.source_file
    if args.source_url:
        source_file = download_excel_report(args.source_url, args.download_dir)

    output_file = build_pivot_table_sheet(
        source_path=source_file,
        output_path=args.output,
        index=args.index,
        columns=args.columns,
        values=args.values,
        aggfunc=args.aggfunc,
        sheet_name=args.sheet_name,
        actioned_by_filter=args.actioned_by_filter,
    )

    print(f"Pivot summary saved to: {output_file}")


if __name__ == "__main__":
    main()
