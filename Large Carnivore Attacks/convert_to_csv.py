"""Convert carnivore_attacks_enriched.xlsx into one CSV per sheet, plus a codebook.

Run from the repository root (needs pandas and openpyxl, both in requirements.txt):

    python "Large Carnivore Attacks/convert_to_csv.py"

The CSVs are written to "Large Carnivore Attacks/csv/". The .xlsx files are only read, never changed.

What this does beyond a plain export:
- `date`: Excel stores dates as day counts (41524); they are written as 2013-09-07.
- `time`: 112 cells hold a bare hour (17 meaning 17:00) but carry a date format, so a plain
  export would write "1900-01-17 00:00:00". They are written as "17:00:00", like the rest
  of the column.
- Columns holding only whole numbers are written as 1, 2 instead of 1.0, 2.0.
- codebook.csv: the original authors explain their coded columns (1 = male, 2 = female, ...)
  only in cell comments of "original data/pbio.3001946.s005.xlsx". Those comments are saved here.
Text is otherwise written exactly as it appears in the workbook (nothing is trimmed or renamed).
"""
import re
from datetime import datetime
from pathlib import Path

import openpyxl
import pandas as pd

HERE = Path(__file__).resolve().parent
ENRICHED = HERE / "carnivore_attacks_enriched.xlsx"
ORIGINAL = HERE / "original data" / "pbio.3001946.s005.xlsx"
OUT = HERE / "csv"

SHEETS = {
    "Enriched incidents": "carnivore_attacks.csv",
    "Data dictionary": "data_dictionary.csv",
    "S2 Figure 2 source": "figure2_source.csv",
}


def fix_time(value):
    """Turn a bare hour that Excel read as a date (17 -> 1900-01-17) back into '17:00:00'."""
    if isinstance(value, datetime):
        if value.year != 1900 or value.month != 1 or value.hour or value.minute or value.second:
            raise ValueError(f"Unexpected value in the time column: {value!r}")
        return f"{value.day:02d}:00:00"
    return value


def whole_numbers_as_int(df):
    """Write float columns that only hold whole numbers as integers (1 instead of 1.0)."""
    for col in df.columns:
        values = df[col].dropna()
        if pd.api.types.is_float_dtype(df[col]) and len(values) and (values % 1 == 0).all():
            df[col] = df[col].astype("Int64")
    return df


def convert_sheets():
    sheets = pd.read_excel(ENRICHED, sheet_name=None, keep_default_na=False, na_values=[""])
    if set(sheets) != set(SHEETS):
        raise ValueError(f"Expected sheets {sorted(SHEETS)}, found {sorted(sheets)}")

    incidents = sheets["Enriched incidents"]
    incidents["date"] = pd.to_datetime(incidents["date"]).dt.strftime("%Y-%m-%d")
    incidents["time"] = incidents["time"].map(fix_time)

    for sheet, filename in SHEETS.items():
        df = whole_numbers_as_int(sheets[sheet])
        df.to_csv(OUT / filename, index=False, encoding="utf-8")
        print(f"{filename}: {len(df)} rows x {len(df.columns)} columns")


def export_codebook():
    """Save the cell comments of the original authors' file as codebook.csv."""
    ws = openpyxl.load_workbook(ORIGINAL).worksheets[0]
    headers = [cell.value for cell in ws[1]]
    rows = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.comment is None:
                continue
            text = re.sub(r"^\s*(Giulia|Author):\s*", "", cell.comment.text)
            text = "; ".join(line.strip() for line in text.splitlines() if line.strip())
            rows.append({
                "field": headers[cell.column - 1],
                # Comments in the header row describe the whole column; others describe one incident.
                "case_id": None if cell.row == 1 else ws.cell(cell.row, 1).value,
                "note": text,
            })
    codebook = pd.DataFrame(rows, columns=["field", "case_id", "note"])
    codebook["case_id"] = codebook["case_id"].astype("Int64")
    codebook.to_csv(OUT / "codebook.csv", index=False, encoding="utf-8")
    print(f"codebook.csv: {len(codebook)} notes")


if __name__ == "__main__":
    OUT.mkdir(exist_ok=True)
    convert_sheets()
    export_codebook()
