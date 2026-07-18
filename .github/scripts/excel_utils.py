#!/usr/bin/env python3
"""
Excel read/write utilities for PyTorch NPU test case management.

Provides:
  1. read_excel()           - Read test case list from Excel, return list of rows
  2. write_excel()          - Write test case data to Excel
  3. discover_from_excel()  - Compatible replacement for discover_test_files.discover_test_files()
  4. generate_sample_excel() - Scan test_*.py files and generate a template Excel

The Excel is the single source of truth for which test files to run.
Each row = one test file, with columns: file_path, test_type, enabled.

Usage as CLI:
    # Read an Excel and print file list for a given type
    python excel_utils.py read --excel test_upstream/case_list.xlsx --type distributed
    python excel_utils.py read --excel test_upstream/case_list.xlsx --type regular

    # Generate a sample Excel from a test directory (first-time setup)
    python excel_utils.py generate --test-dir /path/to/pytorch/test \
        --output test_upstream/case_list.xlsx
"""

import argparse
import json
import re
import sys
from collections import Counter
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# openpyxl is a pure-Python library with no C extensions.
# In CI, install with: pip install openpyxl
# ---------------------------------------------------------------------------
try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.worksheet import Worksheet

    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ---------------------------------------------------------------------------
# Default column mapping
# ---------------------------------------------------------------------------
COL_FILE_PATH = "file_path"
COL_TEST_TYPE = "test_type"
COL_ENABLED = "enabled"
COL_NOTES = "notes"

DEFAULT_COLUMNS = [COL_FILE_PATH, COL_TEST_TYPE, COL_ENABLED, COL_NOTES]

# Header styling
HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
HEADER_ALIGNMENT = Alignment(horizontal="center", vertical="center")
BODY_FONT = Font(name="Consolas", size=10)
THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)

# ---------------------------------------------------------------------------
# Core read / write
# ---------------------------------------------------------------------------


def read_excel(
    file_path: str,
    sheet_name: Optional[str] = None,
) -> List[Dict[str, str]]:
    """
    Read an Excel file and return rows as list of dicts.

    Auto-detects header row (first row). Returns empty list if file not found
    or sheet is empty.

    Args:
        file_path: Path to .xlsx file
        sheet_name: Sheet name or index (default: first sheet)

    Returns:
        List of dicts, each dict = one row with column-name keys

    Raises:
        ImportError: if openpyxl is not installed
        FileNotFoundError: if file_path does not exist
    """
    _ensure_openpyxl()

    wb = load_workbook(file_path, read_only=True, data_only=True)
    try:
        if sheet_name is None:
            ws = wb[wb.sheetnames[0]]
        elif isinstance(sheet_name, int):
            ws = wb[wb.sheetnames[sheet_name]]
        else:
            ws = wb[sheet_name]

        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return []

        # First row = header
        headers = [str(h).strip() if h else "" for h in rows[0]]
        data_rows = rows[1:]

        result = []
        for row in data_rows:
            entry: Dict[str, str] = {}
            for i, header in enumerate(headers):
                if not header:
                    continue
                value = row[i] if i < len(row) else ""
                entry[header] = str(value).strip() if value is not None else ""
            if any(v for v in entry.values()):
                result.append(entry)
        return result
    finally:
        wb.close()


def write_excel(
    file_path: str,
    data: List[Dict[str, Any]],
    sheet_name: str = "Sheet1",
    *,
    columns: Optional[List[str]] = None,
    auto_width: bool = True,
) -> None:
    """
    Write data to an Excel file.

    Args:
        file_path: Output .xlsx path (will be overwritten)
        data: List of dicts to write
        sheet_name: Sheet name
        columns: Ordered list of column keys (default: DEFAULT_COLUMNS)
        auto_width: Auto-fit column widths
    """
    _ensure_openpyxl()

    if columns is None:
        columns = DEFAULT_COLUMNS

    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    # --- Header row ---
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT
        cell.border = THIN_BORDER

    # --- Data rows ---
    for row_idx, entry in enumerate(data, 2):
        for col_idx, col_name in enumerate(columns, 1):
            value = entry.get(col_name, "")
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER

    # --- Column widths ---
    if auto_width:
        _auto_fit_columns(ws, columns, data)

    # Freeze header row
    ws.freeze_panes = "A2"

    # Auto-filter
    if data:
        last_col_letter = get_column_letter(len(columns))
        ws.auto_filter.ref = f"A1:{last_col_letter}{len(data) + 1}"

    wb.save(file_path)


# ---------------------------------------------------------------------------
# Discover function (compatible with discover_test_files.discover_test_files)
# ---------------------------------------------------------------------------


def discover_from_excel(
    excel_path: str,
    test_type: str,
    *,
    sheets: Optional[List[str]] = None,
) -> Tuple[List[str], Dict[str, Any]]:
    """
    Read test files from Excel, return format compatible with
    discover_test_files.discover_test_files().

    Supports two Excel formats:

    1. **Simple format**: single sheet with columns file_path / test_type / enabled.
       Used when the Excel has a "file_path" column.

    2. **Tracker format**: multi-sheet tracking workbook with columns:
       Classification | Specialization | File(N) | Status | Priority | ...
       Reads all data sheets, finds the "File(…)" column, and filters by
       Status containing "Done".

    Args:
        excel_path: Path to the .xlsx file
        test_type: "distributed" or "regular"
        sheets: Optional list of sheet names to scan. If None, auto-detects
                sheets with a "File(…)" column, falling back to all sheets.

    Returns:
        Tuple of (selected_file_paths, metadata_dict)
        - selected_file_paths: sorted list of file paths
        - metadata_dict: discovery metadata dict
    """
    _ensure_openpyxl()

    wb = load_workbook(excel_path, read_only=True, data_only=True)

    all_selected: List[str] = []
    total_files = 0
    total_rows = 0
    sheets_scanned: List[str] = []

    try:
        for sname in wb.sheetnames:
            # Skip README / backup sheets
            if sname.lower().startswith("readme") or sname.lower().endswith("(bak)"):
                continue
            ws = wb[sname]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = [str(h).strip() if h else "" for h in rows[0]]
            data_rows = rows[1:]

            # Detect format: simple vs tracker
            has_file_col = any(h == COL_FILE_PATH for h in headers)
            has_file_N_col = _find_file_N_column(headers)

            if has_file_col:
                # ---- Simple format ----
                file_idx = headers.index(COL_FILE_PATH)
                type_idx = headers.index(COL_TEST_TYPE) if COL_TEST_TYPE in headers else -1
                enabled_idx = headers.index(COL_ENABLED) if COL_ENABLED in headers else -1

                for row in data_rows:
                    total_rows += 1
                    fp = _cell_str(row, file_idx)
                    if not fp:
                        continue
                    total_files += 1

                    # Enabled check
                    enabled_val = _cell_str(row, enabled_idx) if enabled_idx >= 0 else "yes"
                    if enabled_val.strip().lower() not in ("yes", "true", "1", "y", ""):
                        continue

                    # Type classification
                    if type_idx >= 0:
                        t = _cell_str(row, type_idx).strip().lower()
                    else:
                        t = "distributed" if fp.startswith("test/distributed/") else "regular"

                    if t == test_type:
                        all_selected.append(fp)

            elif has_file_N_col is not None:
                # ---- Tracker format ----
                file_idx = has_file_N_col
                status_idx = _find_status_column(headers)

                for row in data_rows:
                    total_rows += 1
                    fp = _cell_str(row, file_idx)
                    if not fp:
                        continue
                    total_files += 1

                    # Status check: must contain "Done"
                    status_val = _cell_str(row, status_idx) if status_idx >= 0 else ""
                    if "Done" not in status_val and "done" not in status_val.lower():
                        continue

                    # Classify by path
                    t = "distributed" if fp.startswith("test/distributed/") else "regular"
                    if t == test_type:
                        all_selected.append(fp)
            else:
                # No recognizable columns — skip
                continue

            sheets_scanned.append(sname)
    finally:
        wb.close()

    all_selected = sorted(set(all_selected))

    metadata = {
        "source": "excel",
        "excel_path": excel_path,
        "test_type": test_type,
        "sheets_scanned": sheets_scanned,
        "total_rows": total_rows,
        "total_files_in_excel": total_files,
        "type_selected": len(all_selected),
        "type_excluded": total_files - len(all_selected),
        "total_files": total_files,
        "whitelist_entries": 0,
        "blacklist_entries": 0,
        "rules_selected": len(all_selected),
        "rules_excluded": total_files - len(all_selected),
        "case_paths_config": excel_path,
    }
    return all_selected, metadata


# ---------------------------------------------------------------------------
# Generate sample Excel from test directory
# ---------------------------------------------------------------------------


def generate_sample_excel(
    test_dir: str,
    output_path: str,
    *,
    existing_excel: Optional[str] = None,
) -> None:
    """
    Scan test_*.py files under test_dir and generate a template Excel.

    If existing_excel is provided, preserve existing enabled/notes values
    and only add new files.

    Args:
        test_dir: Path to PyTorch test/ directory
        output_path: Output .xlsx path
        existing_excel: Optional path to existing Excel to merge with
    """
    test_dir_p = Path(test_dir).resolve()
    if not test_dir_p.is_dir():
        raise FileNotFoundError(f"Test directory not found: {test_dir_p}")

    # Scan all test_*.py files
    raw_files: List[str] = []
    for tf in sorted(test_dir_p.rglob("test_*.py")):
        rel = tf.relative_to(test_dir_p).as_posix()
        raw_files.append(f"test/{rel}")

    print(f"Scanned {len(raw_files)} test_*.py files under {test_dir_p}")

    # Load existing Excel state if provided
    existing_state: Dict[str, Dict[str, str]] = {}
    if existing_excel and Path(existing_excel).exists():
        old_rows = read_excel(existing_excel)
        for r in old_rows:
            fp = r.get(COL_FILE_PATH, "").strip()
            if fp:
                existing_state[fp] = r
        print(f"Loaded {len(existing_state)} entries from existing Excel: {existing_excel}")

    # Build data rows
    data = []
    for fp in raw_files:
        if fp in existing_state:
            data.append(existing_state[fp])
        else:
            # Auto-classify by path
            t = "distributed" if fp.startswith("test/distributed/") else "regular"
            data.append({
                COL_FILE_PATH: fp,
                COL_TEST_TYPE: t,
                COL_ENABLED: "yes",
                COL_NOTES: "",
            })

    write_excel(output_path, data)
    print(f"Generated Excel: {output_path}")
    print(f"  Total files: {len(data)}")
    distributed = sum(1 for d in data if d.get(COL_TEST_TYPE) == "distributed")
    regular = sum(1 for d in data if d.get(COL_TEST_TYPE) == "regular")
    print(f"  Distributed: {distributed}, Regular: {regular}")
    enabled = sum(1 for d in data if d.get(COL_ENABLED, "yes").strip().lower() in ("yes", "true", "1", "y", ""))
    disabled = len(data) - enabled
    if disabled:
        print(f"  Enabled: {enabled}, Disabled: {disabled}")


# ---------------------------------------------------------------------------
# Whitelist generation from Excel
# ---------------------------------------------------------------------------


def get_whitelist_from_excel(excel_path: str) -> List[str]:
    """
    Extract all Status=Done file paths from the tracker Excel.

    Walks all data sheets (skip README/Bak), finds the File(N) column,
    and returns file paths for rows where Status contains "Done".

    Returns sorted deduplicated list.
    """
    _ensure_openpyxl()

    wb = load_workbook(excel_path, read_only=True, data_only=True)
    all_files: List[str] = []

    try:
        for sname in wb.sheetnames:
            if sname.lower().startswith("readme") or sname.lower().endswith("(bak)"):
                continue
            ws = wb[sname]
            rows = list(ws.iter_rows(values_only=True))
            if not rows:
                continue

            headers = [str(h).strip() if h else "" for h in rows[0]]
            file_idx = _find_file_N_column(headers)
            if file_idx is None:
                continue
            status_idx = _find_status_column(headers)

            for row in rows[1:]:
                fp = _cell_str(row, file_idx)
                if not fp:
                    continue
                st = _cell_str(row, status_idx) if status_idx >= 0 else ""
                if "Done" not in st and "done" not in st.lower():
                    continue
                all_files.append(fp)
    finally:
        wb.close()

    return sorted(set(all_files))


def update_case_paths_from_excel(
    excel_path: str,
    yaml_path: str,
    output_path: str,
) -> None:
    """
    Generate a new case_paths_ci.yml with whitelist from Excel, preserving blacklist.

    Args:
        excel_path: Path to Test Class Refactoring Tracker.xlsx
        yaml_path: Path to original case_paths_ci.yml (for blacklist)
        output_path: Path to write the updated case_paths_ci.yml

    The original blacklist is preserved, while the whitelist is replaced
    with files from the Excel that have Status=Done.
    """
    whitelist_files = get_whitelist_from_excel(excel_path)
    if not whitelist_files:
        print("WARNING: No Status=Done files found in Excel, keeping original yaml unchanged")
        # Copy original to output
        if yaml_path != output_path:
            import shutil
            shutil.copy2(yaml_path, output_path)
        return

    # Read original yaml to extract blacklist
    blacklist: List[str] = []
    yaml_path_obj = Path(yaml_path)
    if yaml_path_obj.exists():
        raw = yaml_path_obj.read_text(encoding="utf-8")
        in_blacklist = False
        for line in raw.splitlines():
            stripped = line.strip()
            if stripped.startswith("blacklist:"):
                in_blacklist = True
                continue
            if in_blacklist and stripped.startswith("- "):
                value = stripped[2:].strip().strip("\"'")
                if value:
                    blacklist.append(value)
            elif in_blacklist and not stripped.startswith("- ") and stripped:
                # End of blacklist (reached next top-level key)
                in_blacklist = False

    # Build new yaml
    lines = []
    lines.append("whitelist:")
    for fp in whitelist_files:
        lines.append(f"  - {fp}")
    lines.append("blacklist:")
    if blacklist:
        for item in blacklist:
            lines.append(f"  - {item}")
    else:
        lines.append("  # (preserved from original)")

    Path(output_path).write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"Updated case_paths_ci.yml: {len(whitelist_files)} whitelist entries, {len(blacklist)} blacklist entries → {output_path}")


# ---------------------------------------------------------------------------
# Result Excel generation
# ---------------------------------------------------------------------------

# Error type classification patterns: (label, regex)
_SEG_PATTERNS = [
    ("SIGSEGV", r"SIGSEGV|segmentation\s*fault|segfault|signal\s*11"),
    ("SIGABRT", r"SIGABRT|aborted|abort|signal\s*6"),
    ("SIGBUS", r"SIGBUS|bus\s*error|signal\s*(7|10)"),
    ("SIGFPE", r"SIGFPE|floating\s*point\s*exception|signal\s*8"),
    ("SIGILL", r"SIGILL|illegal\s*instruction|signal\s*4"),
    ("OOM", r"out\s*of\s*memory|MemoryError|cannot\s*allocate\s*memory"),
    ("TIMEOUT", r"timeout|timed\s*out|Timeout"),
    ("IMPORT_ERROR", r"ImportError|ModuleNotFoundError|No\s*module\s*named"),
    ("ASSERTION_ERROR", r"AssertionError"),
    ("RUNTIME_ERROR", r"RuntimeError"),
    ("NPU_ERROR", r"npu|ascend|cann|hcc[pl]|NPU|ACL"),
    ("UNKNOWN", r""),  # fallback
]


def _classify_error(message: str) -> str:
    """Classify an error message string into SEG error types.

    Returns the first matching pattern label, or empty string if no error."""
    if not message or not message.strip():
        return ""
    for label, pattern in _SEG_PATTERNS:
        if pattern and re.search(pattern, message, re.IGNORECASE):
            return label
    return "OTHER"


def generate_result_excel(
    tracker_excel: str,
    output_path: str,
    cases_results_dir: str,
) -> None:
    """
    Generate a result Excel merging tracker structure with test execution results.

    For each file in the tracker Excel (Status=Done), appends:
      Col D: All collected nodeids
      Col E: Execution result summary (P:passed F:failed E:errors T:timeout)
      Col F: Error messages from failed cases (truncated)
      Col G: SEG error type classification

    Args:
        tracker_excel: Path to the original 'Test Class Refactoring Tracker.xlsx'
        output_path: Path to write the result Excel
        cases_results_dir: Directory containing *_cases_results_by_file.jsonl
            (generated by generate_npu_full_test_report.py)
    """
    _ensure_openpyxl()

    # ---- Step 1: Load case results keyed by file path ----
    file_results: Dict[str, Dict[str, Any]] = {}  # file_path -> aggregated results
    cases_dir = Path(cases_results_dir)
    if not cases_dir.is_dir():
        print(f"WARNING: cases_results_dir not found: {cases_results_dir}")
        cases_dir = None

    if cases_dir:
        for jsonl_path in sorted(cases_dir.glob("*_cases_results_by_file.jsonl")):
            print(f"Loading results: {jsonl_path}")
            with open(jsonl_path, "r", encoding="utf-8") as f:
                for line in f:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        obj = json.loads(line)
                    except json.JSONDecodeError:
                        continue
                    fp = obj.get("file_path", "")
                    if not fp:
                        continue
                    # Normalize path
                    if fp.startswith("test/"):
                        fp_key = fp
                    else:
                        fp_key = f"test/{fp}"
                    file_results[fp_key] = {
                        "case_count": obj.get("case_count", 0),
                        "cases": obj.get("cases", []),
                    }

    # ---- Step 2: Walk the tracker Excel sheet-by-sheet ----
    wb_src = load_workbook(tracker_excel, read_only=False, data_only=True)

    # Prepare output workbook (will copy structure + add columns)
    wb_out = Workbook()
    wb_out.remove(wb_out.active)  # remove default sheet

    # Results columns to append after original columns
    RESULT_COLUMNS = [
        "NodeIDs",
        "Result",
        "Error Logs",
        "SEG Type",
    ]

    # Define status pass/fail colors
    PASS_FILL = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # green
    FAIL_FILL = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # red
    WARN_FILL = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")  # yellow
    BODY_FONT_SMALL = Font(name="Consolas", size=9)

    total_files_matched = 0
    total_files_not_found = 0

    try:
        for sname in wb_src.sheetnames:
            # Skip README / backup sheets
            if sname.lower().startswith("readme") or sname.lower().endswith("(bak)"):
                continue

            ws_src = wb_src[sname]
            rows_src = list(ws_src.iter_rows(values_only=False))  # get cell objects for formatting
            if not rows_src:
                continue

            headers_src = [str(c.value).strip() if c.value else "" for c in rows_src[0]]
            file_idx = _find_file_N_column(headers_src)
            if file_idx is None:
                continue

            status_idx = _find_status_column(headers_src)
            if status_idx < 0:
                status_idx = headers_src.index("Status") if "Status" in headers_src else -1

            # Find existing column count
            num_orig_cols = len(headers_src)
            num_result_cols = len(RESULT_COLUMNS)

            # Create output sheet
            ws_out = wb_out.create_sheet(title=sname[:31])  # Excel sheet name limit

            # ---- Copy header row + add result columns ----
            for col_idx in range(1, num_orig_cols + 1):
                src_cell = rows_src[0][col_idx - 1]
                out_cell = ws_out.cell(row=1, column=col_idx)
                out_cell.value = src_cell.value
                out_cell.font = HEADER_FONT
                out_cell.fill = HEADER_FILL
                out_cell.alignment = HEADER_ALIGNMENT
                out_cell.border = THIN_BORDER

            for ri, label in enumerate(RESULT_COLUMNS):
                out_cell = ws_out.cell(row=1, column=num_orig_cols + 1 + ri)
                out_cell.value = label
                out_cell.font = HEADER_FONT
                out_cell.fill = HEADER_FILL
                out_cell.alignment = HEADER_ALIGNMENT
                out_cell.border = THIN_BORDER

            # ---- Track fill-down values for Classification / Specialization ----
            last_classification = ""
            last_specialization = ""

            # ---- Process data rows ----
            out_row = 2
            for src_row_idx in range(1, len(rows_src)):
                src_row = rows_src[src_row_idx]

                # Extract file path
                fp = _cell_str_by_ref(src_row, file_idx)
                if not fp:
                    continue  # skip rows without a file path

                # Check Status
                st = _cell_str_by_ref(src_row, status_idx) if status_idx >= 0 else ""
                if "Done" not in st and "done" not in st.lower():
                    continue  # only Done files

                # Update fill-down values
                class_val = _cell_str_by_ref(src_row, 0)
                if class_val:
                    last_classification = class_val
                spec_val = _cell_str_by_ref(src_row, 1)
                if spec_val:
                    last_specialization = spec_val

                # Copy original cells to output
                for col_idx in range(1, num_orig_cols + 1):
                    src_cell = src_row[col_idx - 1]
                    out_cell = ws_out.cell(row=out_row, column=col_idx)
                    out_cell.value = src_cell.value
                    out_cell.font = BODY_FONT_SMALL
                    out_cell.border = THIN_BORDER

                # Fill down Classification if empty
                if not _cell_str_by_ref(src_row, 0):
                    ws_out.cell(row=out_row, column=1).value = last_classification
                if not _cell_str_by_ref(src_row, 1):
                    ws_out.cell(row=out_row, column=2).value = last_specialization

                # ---- Query test results for this file ----
                # Normalize path for lookup
                fp_norm = fp if fp.startswith("test/") else f"test/{fp}"
                fr = file_results.get(fp_norm, None)

                # Col D: NodeIDs
                if fr and fr["cases"]:
                    nodeids = []
                    for c in fr["cases"]:
                        nid = c.get("nodeid", "")
                        if nid:
                            # Extract just the test method part (after ::)
                            parts = nid.split("::")
                            if len(parts) >= 3:
                                nodeids.append(f"{parts[-2]}::{parts[-1]}")
                            elif len(parts) == 2:
                                nodeids.append(parts[-1])
                            else:
                                nodeids.append(nid)
                    nodeid_text = "\n".join(nodeids)
                else:
                    nodeid_text = ""
                ws_out.cell(row=out_row, column=num_orig_cols + 1).value = nodeid_text
                ws_out.cell(row=out_row, column=num_orig_cols + 1).font = BODY_FONT_SMALL

                # Col E: Result summary
                if fr and fr["cases"]:
                    statuses = Counter(c.get("status", "unknown") for c in fr["cases"])
                    passed = statuses.get("passed", 0)
                    failed = statuses.get("failed", 0)
                    errors = statuses.get("errors", 0)
                    timeout = statuses.get("timeout", 0)
                    skipped = statuses.get("skipped", 0)
                    not_executed = statuses.get("not_executed", 0)
                    total = len(fr["cases"])

                    result_text = f"P:{passed} F:{failed} E:{errors} T:{timeout} S:{skipped} / {total}"
                    # Color cell based on result
                    result_cell = ws_out.cell(row=out_row, column=num_orig_cols + 2)
                    result_cell.value = result_text
                    if failed + errors + timeout + not_executed > 0:
                        result_cell.fill = FAIL_FILL
                    elif skipped > 0:
                        result_cell.fill = WARN_FILL
                    else:
                        result_cell.fill = PASS_FILL
                else:
                    ws_out.cell(row=out_row, column=num_orig_cols + 2).value = "NO RESULT"
                    ws_out.cell(row=out_row, column=num_orig_cols + 2).fill = WARN_FILL
                    total_files_not_found += 1
                    total_files_matched -= 1  # will be incremented back below

                ws_out.cell(row=out_row, column=num_orig_cols + 2).font = BODY_FONT_SMALL
                total_files_matched += 1

                # Col F: Error logs
                error_messages = []
                if fr and fr["cases"]:
                    for c in fr["cases"]:
                        msg = c.get("message", "")
                        if msg and c.get("status", "") not in ("passed", "skipped", "not_executed"):
                            # Truncate each message to avoid huge cells
                            error_messages.append(msg[:2000])
                error_text = "\n---\n".join(error_messages)
                if len(error_text) > 32767:  # Excel cell limit
                    error_text = error_text[:32767]
                ws_out.cell(row=out_row, column=num_orig_cols + 3).value = error_text
                ws_out.cell(row=out_row, column=num_orig_cols + 3).font = BODY_FONT_SMALL

                # Col G: SEG error type
                all_errors = "\n".join(error_messages) if error_messages else ""
                seg_types = set()
                if all_errors:
                    for line in all_errors.split("\n"):
                        seg_type = _classify_error(line)
                        if seg_type:
                            seg_types.add(seg_type)
                seg_text = ", ".join(sorted(seg_types))
                ws_out.cell(row=out_row, column=num_orig_cols + 4).value = seg_text
                ws_out.cell(row=out_row, column=num_orig_cols + 4).font = BODY_FONT_SMALL

                out_row += 1

            # ---- Auto-fit columns ----
            for col_idx in range(1, num_orig_cols + num_result_cols + 1):
                col_letter = get_column_letter(col_idx)
                max_width = 8
                for row in ws_out.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
                    cell_val = str(row[0]) if row[0] else ""
                    # For nodeids column, cap at 60; for error logs, cap at 80
                    if col_idx == num_orig_cols + 1:  # NodeIDs
                        max_width = min(max(max_width, max(len(l) for l in cell_val.split("\n")[:5] if l) + 2), 60)
                    elif col_idx == num_orig_cols + 3:  # Error Logs
                        max_width = min(max(max_width, 80), 80)
                    else:
                        max_width = min(max(max_width, len(cell_val) + 2), 60)
                ws_out.column_dimensions[col_letter].width = max_width

            # Freeze header row
            ws_out.freeze_panes = "A2"

            # Auto-filter on all columns
            last_col_letter = get_column_letter(num_orig_cols + num_result_cols)
            ws_out.auto_filter.ref = f"A1:{last_col_letter}{out_row - 1}"

            print(f"  [{sname}] {out_row - 2} files written")
    finally:
        wb_src.close()

    wb_out.save(output_path)
    print(f"Result Excel saved: {output_path}")
    print(f"  Files matched: {total_files_matched}")
    print(f"  Files not found: {total_files_not_found}")


def _ensure_openpyxl() -> None:
    """Raise ImportError with install hint if openpyxl is missing."""
    if not HAS_OPENPYXL:
        raise ImportError(
            "openpyxl is required for Excel operations.\n"
            "Install with: pip install openpyxl"
        )


def _cell_str(row: tuple, idx: int) -> str:
    """Safely extract a string value from a row tuple at given index."""
    if idx < 0 or idx >= len(row):
        return ""
    v = row[idx]
    return str(v).strip() if v is not None else ""


def _cell_str_by_ref(row: tuple, idx: int) -> str:
    """Safely extract a string value from a row of Cell objects at given index."""
    if idx < 0 or idx >= len(row):
        return ""
    cell = row[idx]
    if hasattr(cell, "value"):
        v = cell.value
        return str(v).strip() if v is not None else ""
    return str(cell).strip() if cell is not None else ""


def _find_file_N_column(headers: List[str]) -> Optional[int]:
    """
    Find the 'File(N)' column in tracker-format headers.
    Matches patterns like 'File(53)', 'File(70)', 'File(285)', or plain 'File'.
    """
    for i, h in enumerate(headers):
        if not h:
            continue
        if h == "File":
            return i
        if h.startswith("File(") and h.endswith(")"):
            return i
    return None


def _find_status_column(headers: List[str]) -> int:
    """Find the 'Status' column index, or -1 if not found."""
    for i, h in enumerate(headers):
        if h and h.strip().lower() == "status":
            return i
    return -1


def _auto_fit_columns(ws: Worksheet, columns: List[str], data: List[Dict[str, Any]]) -> None:
    """Set column widths based on content length."""
    for col_idx, col_name in enumerate(columns, 1):
        max_len = len(str(col_name))
        for row in data:
            val = str(row.get(col_name, ""))
            max_len = max(max_len, len(val))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 80)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def _cmd_read(args: argparse.Namespace) -> None:
    """CLI: read Excel and output file list."""
    files, meta = discover_from_excel(args.excel, test_type=args.type)
    if args.format == "json":
        print(json.dumps({"files": files, "metadata": meta}, indent=2))
    else:
        for f in files:
            print(f)
    if args.verbose:
        print(f"\n--- metadata ---", file=sys.stderr)
        for k, v in meta.items():
            print(f"  {k}: {v}", file=sys.stderr)


def _cmd_generate(args: argparse.Namespace) -> None:
    """CLI: generate sample Excel."""
    generate_sample_excel(
        test_dir=args.test_dir,
        output_path=args.output,
        existing_excel=getattr(args, "merge", None),
    )


def _cmd_list_sheets(args: argparse.Namespace) -> None:
    """CLI: list sheet names in an Excel file."""
    _ensure_openpyxl()
    wb = load_workbook(args.excel, read_only=True)
    for name in wb.sheetnames:
        print(name)
    wb.close()


def _cmd_summary(args: argparse.Namespace) -> None:
    """CLI: print summary stats for an Excel file."""
    _ensure_openpyxl()
    rows = read_excel(args.excel)
    by_type: Dict[str, int] = {}
    enabled_count = 0
    disabled_count = 0
    for r in rows:
        t = r.get(COL_TEST_TYPE, "regular").strip().lower()
        by_type[t] = by_type.get(t, 0) + 1
        if r.get(COL_ENABLED, "yes").strip().lower() in ("yes", "true", "1", "y", ""):
            enabled_count += 1
        else:
            disabled_count += 1
    print(f"File: {args.excel}")
    print(f"Total rows: {len(rows)}")
    print(f"Enabled: {enabled_count}, Disabled: {disabled_count}")
    for t, cnt in sorted(by_type.items()):
        print(f"  {t}: {cnt}")


def _cmd_report(args: argparse.Namespace) -> None:
    """CLI: generate result Excel from tracker + test results."""
    generate_result_excel(
        tracker_excel=args.tracker,
        output_path=args.output,
        cases_results_dir=args.results_dir,
    )


def _cmd_update_yaml(args: argparse.Namespace) -> None:
    """CLI: update case_paths_ci.yml whitelist from Excel."""
    update_case_paths_from_excel(
        excel_path=args.excel,
        yaml_path=args.yaml,
        output_path=args.output or args.yaml,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Excel test case management")
    sub = parser.add_subparsers(dest="command", help="Sub-command")

    # --- read ---
    p_read = sub.add_parser("read", help="Read test files from Excel")
    p_read.add_argument("--excel", required=True, help="Path to .xlsx file")
    p_read.add_argument("--type", choices=["distributed", "regular"], default="regular")
    p_read.add_argument("--format", choices=["text", "json"], default="text")
    p_read.add_argument("--verbose", "-v", action="store_true")
    p_read.set_defaults(func=_cmd_read)

    # --- generate ---
    p_gen = sub.add_parser("generate", help="Generate template Excel from test dir")
    p_gen.add_argument("--test-dir", required=True, help="Path to PyTorch test/ directory")
    p_gen.add_argument("--output", required=True, help="Output .xlsx path")
    p_gen.add_argument("--merge", help="Existing Excel to merge with (preserve enabled/notes)")
    p_gen.set_defaults(func=_cmd_generate)

    # --- sheets ---
    p_sheets = sub.add_parser("sheets", help="List sheet names")
    p_sheets.add_argument("--excel", required=True)
    p_sheets.set_defaults(func=_cmd_list_sheets)

    # --- summary ---
    p_summary = sub.add_parser("summary", help="Print summary stats")
    p_summary.add_argument("--excel", required=True)
    p_summary.set_defaults(func=_cmd_summary)

    # --- report ---
    p_report = sub.add_parser("report", help="Generate result Excel from tracker + test results")
    p_report.add_argument("--tracker", required=True, help="Path to Test Class Refactoring Tracker.xlsx")
    p_report.add_argument("--output", required=True, help="Output .xlsx path")
    p_report.add_argument("--results-dir", required=True, help="Directory with *_cases_results_by_file.jsonl")
    p_report.set_defaults(func=_cmd_report)

    # --- update-yaml ---
    p_uy = sub.add_parser("update-yaml", help="Update case_paths_ci.yml whitelist from Excel")
    p_uy.add_argument("--excel", required=True, help="Path to Test Class Refactoring Tracker.xlsx")
    p_uy.add_argument("--yaml", required=True, help="Path to original case_paths_ci.yml")
    p_uy.add_argument("--output", help="Output path (default: overwrite --yaml)")
    p_uy.set_defaults(func=_cmd_update_yaml)

    args = parser.parse_args()
    if args.command is None:
        parser.print_help()
        sys.exit(1)
    args.func(args)


if __name__ == "__main__":
    main()
