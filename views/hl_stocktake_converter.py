import csv
from datetime import datetime
from io import BytesIO, StringIO

import openpyxl
import streamlit as st


def _rows_from_xlsx(file_bytes: bytes) -> list[tuple]:
    wb = openpyxl.load_workbook(BytesIO(file_bytes), data_only=True, read_only=True)
    ws = wb.active
    rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i == 0 and row[0] is not None and str(row[0]).strip().lower() in {"barcode", "code", "sku"}:
            continue
        rows.append((row[0], row[1]))
    wb.close()
    return rows


def _rows_from_csv(file_bytes: bytes) -> list[tuple]:
    text = file_bytes.decode("utf-8-sig")
    reader = csv.DictReader(StringIO(text))
    fieldnames = [f.strip().lower() for f in (reader.fieldnames or [])]

    # Map column names: barcode = codigo/barcode/code/sku, count = cantidad/count/qty
    barcode_col = next((f for f in fieldnames if f in {"codigo", "barcode", "code", "sku"}), None)
    count_col = next((f for f in fieldnames if f in {"cantidad", "count", "qty", "quantity"}), None)

    if barcode_col is None or count_col is None:
        raise ValueError(
            f"Could not find barcode/count columns. Found: {fieldnames}"
        )

    rows = []
    for row in reader:
        norm = {k.strip().lower(): v for k, v in row.items()}
        rows.append((norm.get(barcode_col), norm.get(count_col)))
    return rows


def _parse_barcode(raw) -> tuple[str, bool]:
    """Return (barcode_str, had_scientific_notation).

    Only converts through float when scientific notation is detected, so that
    leading zeros in plain strings (e.g. '082184045367') are preserved.
    """
    s = str(raw).strip()
    is_sci = "e" in s.lower()
    if is_sci:
        try:
            return str(int(float(s))), True
        except (ValueError, OverflowError):
            return s, True
    return s, False


def _build_dat(all_rows: list[tuple], dt: datetime) -> tuple[bytes, int, int, bool]:
    """Returns (dat_bytes, written, skipped, had_scientific_notation)."""
    dt_str = dt.strftime("%d%m%y%H%M")
    lines = []
    written = 0
    skipped = 0
    had_sci = False

    for barcode_raw, count_raw in all_rows:
        if barcode_raw is None or str(barcode_raw).strip() == "":
            skipped += 1
            continue
        if count_raw is None or str(count_raw).strip() == "":
            skipped += 1
            continue

        barcode, is_sci = _parse_barcode(barcode_raw)
        if is_sci:
            had_sci = True
        count = str(count_raw).strip()

        line = f"{barcode:<16}{count:<9}{dt_str}"
        lines.append(line)
        written += 1

    content = "\r\n".join(lines) + "\r\n" if lines else ""
    return content.encode("ascii"), written, skipped, had_sci


def render():
    st.header("📤 Stocktake → H&L DAT")
    st.caption("Converts one or more stocktake files (.xlsx or .csv) into a single .dat file for H&L import.")

    uploaded_files = st.file_uploader(
        "Upload stocktake files (.xlsx or .csv)",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
        key="hl_uploader",
    )

    if not uploaded_files:
        st.info("Upload one or more .xlsx / .csv files to continue.")
        return

    dt = datetime.now()
    all_rows: list[tuple] = []
    file_errors: list[str] = []

    for f in uploaded_files:
        file_bytes = f.read()
        try:
            if f.name.lower().endswith(".csv"):
                rows = _rows_from_csv(file_bytes)
            else:
                rows = _rows_from_xlsx(file_bytes)
            all_rows.extend(rows)
        except Exception as e:
            file_errors.append(f"**{f.name}**: {e}")

    if file_errors:
        for err in file_errors:
            st.error(err)
        if not all_rows:
            return

    try:
        dat_bytes, written, skipped, had_sci = _build_dat(all_rows, dt)
    except Exception as e:
        st.error(f"Error building DAT: {e}")
        return

    if had_sci:
        st.warning(
            "⚠️ One or more barcodes were in scientific notation (e.g. `9.34101E+12`). "
            "This means the CSV was saved from Excel with insufficient precision and the last digits of those barcodes are lost. "
            "**Fix:** open the original file in Excel, format the barcode column as *Text* or *Number (0 decimals)*, then export to CSV again."
        )

    base_name = uploaded_files[0].name.rsplit(".", 1)[0] if len(uploaded_files) == 1 else "stocktake"
    out_name = base_name + ".dat"

    st.success(f"Conversion successful: **{written}** rows written, **{skipped}** rows skipped.")

    st.download_button(
        label=f"⬇️ Download {out_name}",
        data=dat_bytes,
        file_name=out_name,
        mime="text/plain",
    )

    with st.expander("Preview (first 10 lines)"):
        preview_lines = dat_bytes.decode("ascii").replace("\r\n", "\n").splitlines()[:10]
        st.code("\n".join(preview_lines), language="text")
