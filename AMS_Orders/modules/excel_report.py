"""
Excel Report Engine — reads pivot tables from the AOMOSO engine workbook
and appends summary rows to destination tables.

Migrated from Excel/aomoXL.py into the AMS_Orders module architecture.
"""

import gc
import glob
import os
import shutil
import time
import win32com.client as win32
from datetime import datetime

from logger import logger
from helpers import com_context, subtract_one_business_day


# ---------------------------------------------------------------------------
# File discovery and backup
# ---------------------------------------------------------------------------

ENGINE_PATTERNS = ('*AO MO SO CHECKER*.xlsm', '*AO MO SO CHECKER*.xlsx')


def _find_engine_file():
    """Locate the engine workbook in the current working directory.

    Uses glob patterns, preferring .xlsm over .xlsx.  When multiple files
    match, the most-recently-modified one is returned.
    Raises FileNotFoundError if nothing matches.
    """
    cwd = os.getcwd()
    for pattern in ENGINE_PATTERNS:
        matches = glob.glob(os.path.join(cwd, pattern))
        if matches:
            return max(matches, key=os.path.getmtime)
    raise FileNotFoundError(
        f"No engine workbook ({', '.join(ENGINE_PATTERNS)}) found in {cwd}"
    )


def _backup_engine_file(engine_path):
    """Copy the engine workbook to a Backup/ subfolder with a date stamp.

    The date stamp uses the previous business day (matching the original
    aomoXL.py convention).  The Backup/ folder is created if it doesn't exist.
    """
    backup_folder = os.path.join(os.path.dirname(engine_path), "Backup")
    os.makedirs(backup_folder, exist_ok=True)

    prev_day = subtract_one_business_day(datetime.now())
    archive_date = prev_day.strftime("%m%d%Y")

    name, ext = os.path.splitext(os.path.basename(engine_path))
    new_name = f"{name}_{archive_date}{ext}"
    dest = os.path.join(backup_folder, new_name)

    shutil.copy(engine_path, dest)
    logger.info(f"Engine file backed up as {new_name} to {backup_folder}")


# ---------------------------------------------------------------------------
# COM helpers
# ---------------------------------------------------------------------------

def _release_com_object(obj):
    """Release a COM object and force garbage collection."""
    if obj is None:
        return
    try:
        del obj
        gc.collect()
    except Exception:
        pass


def _wait_for_calculations(excel, max_wait=60):
    """Wait for Excel async calculations to finish (exponential backoff).

    Returns True if calculations completed within *max_wait* seconds.
    """
    start = time.time()
    interval = 0.5

    while excel.CalculationState != 0:          # 0 = xlDone
        if time.time() - start > max_wait:
            logger.warning("Calculation timeout — proceeding anyway.")
            return False
        time.sleep(interval)
        interval = min(interval * 1.5, 3)       # cap at 3 s

    excel.CalculateUntilAsyncQueriesDone()
    return True


# ---------------------------------------------------------------------------
# Pattern functions
# ---------------------------------------------------------------------------

def _copy_data_body_range(workbook, source_sheet, op):
    """Pattern A: copy all DataBodyRange rows into a destination table.

    Reads the entire DataBodyRange in one COM call (batch), then writes
    each value (excluding the last grand-total column) into the destination
    table at the column position returned by ``op["col_offset"]``.
    """
    pivot = source_sheet.PivotTables(op["pivot"])
    data_range = pivot.DataBodyRange

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    next_row = dest_table.ListRows.Count + 1
    dest_table.ListRows.Add()
    new_row = dest_table.ListRows(next_row)

    if op.get("date_col"):
        col, date_val = op["date_col"]
        new_row.Range.Cells(1, col).Value = date_val

    # Batch-read all values in one COM call (returns 2D tuple)
    values = data_range.Value
    if not isinstance(values, tuple):
        values = ((values,),)

    col_offset = op["col_offset"]
    num_cols = data_range.Columns.Count
    for i, row_data in enumerate(values, 1):
        for j in range(1, num_cols):   # skip last column
            new_row.Range.Cells(i, col_offset(j)).Value = row_data[j - 1]

    logger.info(f"Pattern A complete: {op['name']}")


def _search_row_copy_columns(workbook, source_sheet, op):
    """Pattern B: search TableRange2 for a keyword row, copy columns 2+ to dest.

    Batch-reads the entire TableRange2 in one COM call, searches for the
    keyword row, then copies columns 2..N-1 into the destination table.
    """
    pivot = source_sheet.PivotTables(op["pivot"])
    full_range = pivot.TableRange2

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    next_row = dest_table.ListRows.Count + 1
    dest_table.ListRows.Add()
    new_row = dest_table.ListRows(next_row)

    if op.get("date_col"):
        col, date_val = op["date_col"]
        new_row.Range.Cells(1, col).Value = date_val

    # Batch-read entire range once
    values = full_range.Value
    num_cols = full_range.Columns.Count

    row_start = op.get("row_start", 1)
    target_row = None
    for i in range(row_start - 1, len(values)):
        cell_val = values[i][0]
        if cell_val and op["keyword"] in str(cell_val):
            target_row = values[i]
            break

    if target_row is None:
        logger.warning(f"Keyword '{op['keyword']}' not found for {op['name']}")
        return

    col_offset = op["col_offset"]
    for j in range(2, num_cols):            # skip last column
        new_row.Range.Cells(1, col_offset(j)).Value = target_row[j - 1]

    logger.info(f"Pattern B complete: {op['name']}")


def _previous_full_day_lookup(workbook, source_sheet, op):
    """Pattern C: find first empty cell in dest column, write last pivot col value.

    1. Locate ``op["dest_col_header"]`` in the destination table's header row.
    2. Walk down that column to find the first empty cell.
    3. Search the pivot's TableRange2 for ``op["keyword"]``.
    4. Write the value from the *last* column of that row into the empty cell.
    """
    pivot = source_sheet.PivotTables(op["pivot"])
    full_range = pivot.TableRange2

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    # Batch-read header row to find destination column
    header_row = dest_table.HeaderRowRange
    header_vals = header_row.Value
    if header_vals and isinstance(header_vals[0], tuple):
        header_vals = header_vals[0]
    dest_col = None
    for idx, val in enumerate(header_vals, 1):
        if val == op["dest_col_header"]:
            dest_col = idx
            break

    if dest_col is None:
        logger.warning(f"Column '{op['dest_col_header']}' not found for {op['name']}")
        return

    # Batch-read data body column to find first empty row
    data_body = dest_table.DataBodyRange
    body_vals = data_body.Value
    target_row = None
    for row_idx, row_data in enumerate(body_vals):
        if not row_data[dest_col - 1]:
            target_row = row_idx + 1 + data_body.Row - 1
            break
    if target_row is None:
        target_row = data_body.Rows.Count + data_body.Row

    # Batch-read pivot range to find keyword row
    values = full_range.Value
    num_cols = len(values[0]) if values else 0
    target_data = None
    for row_data in values:
        cell_val = row_data[0]
        if cell_val and op["keyword"] in str(cell_val):
            target_data = row_data
            break

    if target_data is None:
        logger.warning(f"Keyword '{op['keyword']}' not found for {op['name']}")
        return

    last_val = target_data[num_cols - 1] if num_cols else None
    if last_val is not None:
        data_body.Cells(target_row - data_body.Row + 1, dest_col).Value = last_val
        logger.info(f"Pattern C complete: {op['name']}")
    else:
        logger.warning(f"No data in last column for {op['name']}")


def _single_cell_extraction(workbook, source_sheet, op):
    """Pattern D: extract one cell value from a pivot row and write to dest.

    Searches column 1 of TableRange2 for ``op["keyword"]``, reads the value
    at ``op["extract_col"]``, and writes it into a new row's first cell in
    the destination table.
    """
    pivot = source_sheet.PivotTables(op["pivot"])
    full_range = pivot.TableRange2

    # Batch-read all values
    values = full_range.Value
    row_start = op.get("row_start", 2)
    target_row = None
    for i in range(row_start - 1, len(values)):
        cell_val = values[i][0]
        if cell_val and op["keyword"] in str(cell_val):
            target_row = values[i]
            break

    if target_row is None:
        logger.warning(f"Keyword '{op['keyword']}' not found for {op['name']}")
        return

    value = target_row[op["extract_col"] - 1]

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    next_row = dest_table.ListRows.Count + 1
    dest_table.ListRows.Add()
    dest_table.DataBodyRange.Rows(next_row).Cells(1).Value = value

    logger.info(f"Pattern D complete: {op['name']}")


def _search_row_with_blank_check(workbook, source_sheet, op):
    """Pattern E: like B but includes the last column and replaces blanks with 0.

    Batch-reads the pivot range, searches for the keyword row, checks if
    all data values are blank/zero, and copies accordingly.
    """
    pivot = source_sheet.PivotTables(op["pivot"])
    full_range = pivot.TableRange2

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    next_row = dest_table.ListRows.Count + 1
    dest_table.ListRows.Add()
    new_row = dest_table.ListRows(next_row)

    if op.get("date_col"):
        col, date_val = op["date_col"]
        new_row.Range.Cells(1, col).Value = date_val

    # Batch-read entire range once
    values = full_range.Value
    num_cols = len(values[0]) if values else 0

    row_start = op.get("row_start", 2)
    target_row = None
    for i in range(row_start - 1, len(values)):
        cell_val = values[i][0]
        if cell_val and op["keyword"] in str(cell_val):
            target_row = values[i]
            break

    if target_row is None:
        logger.warning(f"Keyword '{op['keyword']}' not found for {op['name']}")
        return

    # Check if entire row is blank / zero (columns 2+ in batch data)
    is_blank = all(v in (None, "", 0) for v in target_row[1:])

    col_offset = op["col_offset"]
    for j in range(2, num_cols + 1):       # includes last col
        new_val = 0 if is_blank else target_row[j - 1]
        new_row.Range.Cells(1, col_offset(j)).Value = new_val

    logger.info(f"Pattern E complete: {op['name']}")


def _sheet_range_copy(workbook, op):
    """Pattern F: copy a direct cell range (not a pivot table) to a dest table.

    Reads a rectangular range defined by ``op["start_row"]``, ``start_col``,
    ``end_row``, ``end_col`` from ``op["source_sheet"]`` and writes each cell
    into the destination table at the column returned by ``op["col_offset"]``.
    """
    src_sheet = workbook.Sheets(op["source_sheet"])
    start_cell = src_sheet.Cells(op["start_row"], op["start_col"])
    end_cell = src_sheet.Cells(op["end_row"], op["end_col"])
    source_range = src_sheet.Range(start_cell, end_cell)

    dest_sheet = workbook.Sheets(op["dest_sheet"])
    dest_table = dest_sheet.ListObjects(op["dest_table"])

    next_row = dest_table.ListRows.Count + 1
    dest_table.ListRows.Add()
    new_row = dest_table.ListRows(next_row).Range

    if op.get("date_col"):
        col, date_val = op["date_col"]
        new_row.Cells(1, col).Value = date_val

    # Batch-read source range in one COM call
    values = source_range.Value
    # Flatten: may be a 2D tuple for a single row
    if isinstance(values, tuple) and values and isinstance(values[0], tuple):
        flat_values = values[0]
    elif isinstance(values, tuple):
        flat_values = values
    else:
        flat_values = (values,)

    col_offset = op["col_offset"]
    for i, val in enumerate(flat_values, start=1):
        new_row.Cells(1, col_offset(i)).Value = val

    logger.info(f"Pattern F complete: {op['name']}")


_DISPATCH = {
    "A": _copy_data_body_range,
    "B": _search_row_copy_columns,
    "C": _previous_full_day_lookup,
    "D": _single_cell_extraction,
    "E": _search_row_with_blank_check,
    "F": _sheet_range_copy,
}


# ---------------------------------------------------------------------------
# Operations list
# ---------------------------------------------------------------------------

def _build_operations(today_str, prev_bday_str):
    """Return the ordered list of all operation dicts."""
    return [
        # ── MO YR SUMMARY ────────────────────────────────────────────
        {   # Op 1 — Incomplete, Inventory > 0
            "name": "Incomplete Inventory > 0",
            "pattern": "A",
            "pivot": "PivotTable5",
            "dest_sheet": "MO YR SUMMARY",
            "dest_table": "YR_INCOMP",
            "date_col": (2, today_str),
            "col_offset": lambda j: j + 2,
        },
        {   # Op 2 — No Inventory, Inventory = 0
            "name": "No Inventory = 0",
            "pattern": "A",
            "pivot": "PivotTable7",
            "dest_sheet": "MO YR SUMMARY",
            "dest_table": "YR_NOINV",
            "col_offset": lambda j: j,
        },
        {   # Op 3 — Total MO Created
            "name": "Total MO Created",
            "pattern": "A",
            "pivot": "PivotTable4",
            "dest_sheet": "MO YR SUMMARY",
            "dest_table": "MB51_submit18",
            "col_offset": lambda j: j,
        },
        {   # Op 4 — Daily Reservation Items Submitted
            "name": "Daily Reservation Submitted",
            "pattern": "B",
            "pivot": "PivotTable11",
            "keyword": "CURRENT UNIL 6 PM",
            "dest_sheet": "MO YR SUMMARY",
            "dest_table": "MB51_submit",
            "col_offset": lambda j: j - 1,
        },
        {   # Op 5 — Previous Full Day Submitted
            "name": "Prev Full Day Submitted",
            "pattern": "C",
            "pivot": "PivotTable11",
            "keyword": "PREVIOUS FULL DAY",
            "dest_sheet": "MO YR SUMMARY",
            "dest_table": "MB51_submit",
            "dest_col_header": "Full Day SUBMIT",
        },
        # ── DN AO YR SUMMARY ─────────────────────────────────────────
        {   # Op 6 — DN AO Inventory Available
            "name": "DN AO Inventory Available",
            "pattern": "B",
            "pivot": "PivotTable3",
            "keyword": "Available",
            "dest_sheet": "DN AO YR SUMMARY",
            "dest_table": "AO_INV_AVAIL",
            "date_col": (2, today_str),
            "col_offset": lambda j: j + 1,
        },
        {   # Op 7 — DN AO No Inventory
            "name": "DN AO No Inventory",
            "pattern": "B",
            "pivot": "PivotTable3",
            "keyword": "No Inventory",
            "dest_sheet": "DN AO YR SUMMARY",
            "dest_table": "AO_NO_INV",
            "col_offset": lambda j: j - 1,
        },
        {   # Op 8 — DN AO Partial Inventory
            "name": "DN AO Partial Inventory",
            "pattern": "B",
            "pivot": "PivotTable3",
            "keyword": "Partial",
            "dest_sheet": "DN AO YR SUMMARY",
            "dest_table": "AO_PART_INV",
            "col_offset": lambda j: j - 1,
        },
        {   # Op 9 — Daily DN AO Submitted
            "name": "Daily DN AO Submitted",
            "pattern": "B",
            "pivot": "PivotTable1",
            "keyword": "CURRENT UNIL 6 PM",
            "dest_sheet": "DN AO YR SUMMARY",
            "dest_table": "Table16",
            "col_offset": lambda j: j - 1,
        },
        {   # Op 10 — Previous Full Day AO Submitted
            "name": "Prev Full Day AO Submitted",
            "pattern": "C",
            "pivot": "PivotTable1",
            "keyword": "PREVIOUS FULL DAY",
            "dest_sheet": "DN AO YR SUMMARY",
            "dest_table": "Table16",
            "dest_col_header": "FULL DAY SUBMIT",
        },
        # ── SO YR COMP ───────────────────────────────────────────────
        {   # Op 11 — Assembly Completed
            "name": "SO YR COMP - Assembly Completed",
            "pattern": "B",
            "pivot": "PivotTable6",
            "keyword": "ASSEMBLY COMPLETED",
            "dest_sheet": "SO YR COMP",
            "dest_table": "Table9",
            "date_col": (2, prev_bday_str),
            "col_offset": lambda j: j + 1,
            "row_start": 2,
        },
        {   # Op 12 — eStore
            "name": "SO YR COMP - eStore",
            "pattern": "D",
            "pivot": "PivotTable6",
            "keyword": "eStore",
            "dest_sheet": "SO YR COMP",
            "dest_table": "Table11",
            "extract_col": 6,
            "row_start": 2,
        },
        {   # Op 13 — HUB ORDER
            "name": "SO YR COMP - HUB ORDER",
            "pattern": "B",
            "pivot": "PivotTable6",
            "keyword": "HUB ORDER",
            "dest_sheet": "SO YR COMP",
            "dest_table": "Table15",
            "col_offset": lambda j: j - 1,
            "row_start": 2,
        },
        {   # Op 14 — REGULAR
            "name": "SO YR COMP - REGULAR",
            "pattern": "B",
            "pivot": "PivotTable6",
            "keyword": "REGULAR",
            "dest_sheet": "SO YR COMP",
            "dest_table": "Table19",
            "col_offset": lambda j: j - 1,
            "row_start": 2,
        },
        # ── SO YR INCMP ──────────────────────────────────────────────
        {   # Op 15 — Assembly Completed (incomplete)
            "name": "SO YR INCMP - Assembly Completed",
            "pattern": "E",
            "pivot": "PivotTable8",
            "keyword": "ASSEMBLY COMPLETED",
            "dest_sheet": "SO YR INCMP",
            "dest_table": "Table21",
            "date_col": (2, today_str),
            "col_offset": lambda j: j + 1,
            "row_start": 2,
        },
        {   # Op 16 — eStore (incomplete)
            "name": "SO YR INCMP - eStore",
            "pattern": "D",
            "pivot": "PivotTable8",
            "keyword": "eStore",
            "dest_sheet": "SO YR INCMP",
            "dest_table": "Table2326",
            "extract_col": 7,
            "row_start": 2,
        },
        {   # Op 17 — HUB ORDER (incomplete)
            "name": "SO YR INCMP - HUB ORDER",
            "pattern": "E",
            "pivot": "PivotTable8",
            "keyword": "HUB ORDER",
            "dest_sheet": "SO YR INCMP",
            "dest_table": "Table27",
            "col_offset": lambda j: j - 1,
            "row_start": 2,
        },
        {   # Op 18 — REGULAR (incomplete)
            "name": "SO YR INCMP - REGULAR",
            "pattern": "E",
            "pivot": "PivotTable8",
            "keyword": "REGULAR",
            "dest_sheet": "SO YR INCMP",
            "dest_table": "Table28",
            "col_offset": lambda j: j - 1,
            "row_start": 2,
        },
        # ── MO % ─────────────────────────────────────────────────────
        {   # Op 19 — MO % sheet range copy (C4:DB4 → Table18)
            "name": "MO % Sheet Copy",
            "pattern": "F",
            "source_sheet": "MO %",
            "start_row": 4, "start_col": 3,
            "end_row": 4, "end_col": 106,
            "dest_sheet": "MO %",
            "dest_table": "Table18",
            "date_col": (2, today_str),
            "col_offset": lambda i: i + 11,
        },
    ]


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def main(progress_callback=None):
    """Run the full Excel report pipeline.

    Parameters
    ----------
    progress_callback : callable, optional
        If provided, called as ``progress_callback(percent, stage)`` where
        *percent* is 0-100 and *stage* is a short description string.
    """
    def emit(pct, stage):
        logger.info(f"[{pct}%] {stage}")
        if progress_callback:
            progress_callback(pct, stage)

    # --- Step 1: Discover and back up engine file ---
    engine_path = _find_engine_file()
    emit(0, f"Engine file found: {os.path.basename(engine_path)}")

    emit(2, "Backing up engine file...")
    _backup_engine_file(engine_path)

    # --- Step 2: Open workbook and set dates ---
    with com_context():
        excel = None
        workbook = None
        try:
            emit(5, "Opening engine workbook...")
            excel = win32.Dispatch('Excel.Application')
            excel.Visible = True
            workbook = excel.Workbooks.Open(engine_path)

            today = datetime.now()
            prev_bday = subtract_one_business_day(today)

            # UTILITY sheet cells use MM/DD/YYYY
            utility = workbook.Sheets('UTILITY')
            utility.Range('F3').Value = today.strftime("%m/%d/%Y")
            utility.Range('E3').Value = prev_bday.strftime("%m/%d/%Y")

            # Table date columns use YYYY-MM-DD
            today_str = today.strftime("%Y-%m-%d")
            prev_bday_str = prev_bday.strftime("%Y-%m-%d")

            # --- Step 3: Refresh workbook queries (3 cycles to match original) ---
            emit(10, "Refreshing workbook data...")
            for _ in range(3):
                workbook.RefreshAll()
                _wait_for_calculations(excel, max_wait=120)
            emit(20, "Refresh complete.")

            excel.Visible = False

            # --- Step 4: Execute operations ---
            source_sheet = workbook.Sheets('UTILITY')
            operations = _build_operations(today_str, prev_bday_str)
            total = len(operations) or 1

            for idx, op in enumerate(operations, 1):
                pattern = op["pattern"]
                handler = _DISPATCH.get(pattern)
                try:
                    if handler is None:
                        logger.error(f"Unknown pattern '{pattern}' for {op['name']}")
                        continue
                    if pattern == "F":
                        handler(workbook, op)
                    else:
                        handler(workbook, source_sheet, op)
                except Exception as e:
                    logger.error(f"Operation '{op['name']}' failed: {e}")

                pct = 20 + int((idx / total) * 70)     # 20-90% range
                emit(pct, op['name'])

            # --- Step 5: Final refresh and save ---
            emit(92, "Final refresh...")
            workbook.RefreshAll()
            _wait_for_calculations(excel, max_wait=60)

            workbook.Save()
            excel.Visible = True
            emit(100, "Excel report complete!")

        except Exception as e:
            logger.error(f"Excel report failed: {e}")
            raise
        finally:
            try:
                if workbook is not None:
                    workbook.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                if excel is not None:
                    excel.Quit()
            except Exception:
                pass
            _release_com_object(workbook)
            _release_com_object(excel)
