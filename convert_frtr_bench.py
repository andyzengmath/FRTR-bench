"""
convert_frtr_bench.py

Converts FRTR-bench dataset (30 IRM-encrypted Excel workbooks) into the format
expected by the old_prod_simulation pipeline:
  - Unencrypted XLSX files with data sheets only (no Readme/Questions sheets)
  - A JSONL file with one question per line matching run_data_set.py schema

Requires: Windows + Excel installed + IRM access + pywin32
Usage:
  python convert_frtr_bench.py --src_dir "path/to/FRTR-bench" --dst_dir "path/to/frtr_data"
"""

import argparse
import json
import os
import re
import sys
import time

import win32com.client
import pythoncom

# COM retry delay (seconds) — gives Excel time to recover between operations
COM_RETRY_DELAY = 3
COM_MAX_RETRIES = 5
WORKBOOK_DELAY = 8  # seconds between workbooks


# ---------------------------------------------------------------------------
# Constants & mappings
# ---------------------------------------------------------------------------

IMAGE_REASONING_TYPES = {
    "Chart reading", "Trend/Visual", "Trend reading",
    "TrendReasoning", "ExtractionFromImage", "Visual Min",
}

OPERATION_MAP = {
    "Aggregation": "aggregate",
    "Computation": "aggregate",
    "Cross-sheet SUMIFS": "aggregate",
    "Counting/Filter": "aggregate",
    "Comparison": "lookup",
    # Extended types found in frtr_0001, frtr_0002, frtr_00027
    "Aggregation/Lookup": "aggregate",
    "Lookup/Computation": "lookup",
    "Lookup/Count": "lookup",
    "Averaging": "aggregate",
    "FormulaReasoning": "aggregate",
    "Threshold/Logic": "aggregate",
    "Percentage + Filter": "aggregate",
    "AVERAGEIFS": "aggregate",
    "SUMPRODUCT Count": "aggregate",
    "Spread/Variance": "aggregate",
}

DOMAIN_MAP = {
    "quarterly-summary": "finance",
    "ops-inventory": "supply_chain",
    "consolidation": "finance",
    "payroll-summary": "finance",
    "ap-vendor-spend": "finance",
    "manufacturing-ops": "manufacturing",
    "ecommerce-metrics": "retail",
    "cashflow-consolidation": "finance",
    "healthcare-claims": "healthcare",
    "energy-production": "energy",
    "gov-grants-tracking": "government",
    "nonprofit-donations": "nonprofit",
    "education-outcomes": "education",
    "retail-pos": "retail",
    "telecom-churn": "telecom",
    "airline-ops": "transportation",
    "realestate-portfolio": "real_estate",
    "hospitality-bookings": "hospitality",
    "agriculture-yield": "agriculture",
    "pharma-trials": "pharma",
    "insurance-claims": "insurance",
    "logistics-routing": "supply_chain",
    "media-streaming": "media",
    "cybersecurity-incidents": "cybersecurity",
    "construction-controls": "construction",
    "consolidation-multi-subsidiary": "finance",
    "supply-chain-regional": "supply_chain",
    "payroll-analysis": "finance",
    "customer-analytics": "retail",
    "capex-project-portfolio": "finance",
}

FORMULA_PREFIXES = (
    "SUM(", "SUMIF(", "SUMIFS(", "COUNTIF(", "COUNTIFS(",
    "XLOOKUP(", "VLOOKUP(", "HLOOKUP(", "MAX(", "MIN(",
    "AVERAGE(", "AVERAGEIF(", "AVERAGEIFS(", "INDEX(", "MATCH(",
    "ROWS(", "COLUMNS(", "COUNTA(", "COUNT(", "ABS(",
    "IF(", "IFS(", "FILTER(", "UNIQUE(", "SORT(",
    "LET(", "MAXIFS(", "MINIFS(",
)

# Regex for bare cell reference like "Sheet!B4" or "Sheet!B4:C10"
BARE_CELL_REF_RE = re.compile(
    r"^([A-Za-z_][\w\s\-]*)!\s*([A-Z]+\d+(?::[A-Z]+\d+)?)\s*$"
)


# ---------------------------------------------------------------------------
# Helper functions
# ---------------------------------------------------------------------------

def com_retry(func, *args, retries=COM_MAX_RETRIES, delay=COM_RETRY_DELAY):
    """Retry a COM call if Excel rejects it (RPC_E_CALL_REJECTED)."""
    for attempt in range(retries):
        try:
            return func(*args)
        except pythoncom.com_error as e:
            # -2147418111 = RPC_E_CALL_REJECTED (callee busy)
            if e.args[0] == -2147418111 and attempt < retries - 1:
                print(f"    COM busy, retrying in {delay}s (attempt {attempt + 1}/{retries})...")
                time.sleep(delay)
                delay *= 2  # exponential backoff
            else:
                raise
    raise RuntimeError("COM retries exhausted")


def parse_filename(fname):
    """Extract file_id and domain_slug from filename like frtr_0001_quarterly-summary.xlsx."""
    match = re.match(r"frtr_0*(\d+)_(.+)\.xlsx$", fname)
    if not match:
        return None, None
    file_id = match.group(1).zfill(4)
    domain_slug = match.group(2)
    return file_id, domain_slug


def extract_sheet_refs(provenance_text):
    """Extract distinct sheet names referenced in the Provenance column.

    Looks for patterns like SheetName!CellRef or SheetName!Range.
    """
    if not provenance_text:
        return set()
    pattern = r"([A-Za-z_][\w\s\-]*)!"
    matches = re.findall(pattern, str(provenance_text))
    return {m.strip() for m in matches}


def determine_scope(provenance_text):
    """Return 'multiple_sheets' or 'single_sheet' based on provenance."""
    refs = extract_sheet_refs(provenance_text)
    if len(refs) >= 2:
        return "multiple_sheets"
    return "single_sheet"


def map_question_type(reasoning_type, provenance_text):
    """Map FRTR ReasoningType + Provenance -> pipeline question_type."""
    operation = OPERATION_MAP.get(reasoning_type, "aggregate")
    scope = determine_scope(provenance_text)
    return f"{operation}_{scope}"


def infer_output_type(answer_value):
    """Infer output_type from the resolved answer value."""
    val_str = str(answer_value).replace(",", "").replace("$", "").replace("%", "").strip()
    try:
        float(val_str)
        return "number"
    except (ValueError, TypeError):
        return "text"


def eval_formula_in_temp_cell(wb, formula_str):
    """Evaluate a formula by writing to a temp cell and reading the result.

    Always cleans up the temp cell, even on error.
    """
    ws = wb.Sheets(1)
    try:
        ws.Range("ZZ1").Formula = formula_str
        time.sleep(0.5)  # give Excel time to calculate
        val = ws.Range("ZZ1").Value
        return val
    finally:
        try:
            ws.Range("ZZ1").ClearContents()
        except Exception:
            pass


def _resolve_cell_ref(wb, sheet_name, cell_ref):
    """Read a single cell value from a sheet."""
    sheet_name = sheet_name.strip().strip("'")
    cell_ref = cell_ref.strip()
    ws = wb.Sheets(sheet_name)
    val = ws.Range(cell_ref).Value
    return format_value(val)


def resolve_answer(wb, answer_text):
    """Resolve an answer string to a plain-text value using COM automation.

    Returns (resolved_value: str, resolution_failed: bool).
    """
    if not answer_text or not str(answer_text).strip():
        return str(answer_text), True

    answer_text = str(answer_text).strip()

    # If answer contains " OR ", try the first option only
    if " OR " in answer_text:
        first_option = answer_text.split(" OR ")[0].strip()
        val, failed = resolve_answer(wb, first_option)
        if not failed:
            return val, False
        # Fall through to try full text

    try:
        # Pattern 1: "See Sheet!Cell" (possibly with trailing text)
        see_match = re.match(r"^See\s+([A-Za-z_][\w\s\-]*)!\s*([A-Z]+\d+)", answer_text, re.IGNORECASE)
        if see_match:
            sheet_name = see_match.group(1)
            cell_ref = see_match.group(2)
            return _resolve_cell_ref(wb, sheet_name, cell_ref), False

        # Pattern 2: Bare cell reference "Sheet!Cell" (exact match)
        bare_match = BARE_CELL_REF_RE.match(answer_text)
        if bare_match:
            sheet_name = bare_match.group(1)
            cell_ref = bare_match.group(2)
            return _resolve_cell_ref(wb, sheet_name, cell_ref), False

        # Pattern 3: "Max of Sheet!Range" / "Min of Sheet!Range"
        maxmin_match = re.match(r"^(Max|Min)\s+of\s+(.+)$", answer_text, re.IGNORECASE)
        if maxmin_match:
            func = maxmin_match.group(1).upper()
            range_ref = maxmin_match.group(2).strip()
            val = eval_formula_in_temp_cell(wb, f"={func}({range_ref})")
            return format_value(val), False

        # Pattern 4: Formula — starts with known function name
        upper_text = answer_text.upper()
        if any(upper_text.startswith(p) for p in FORMULA_PREFIXES):
            val = eval_formula_in_temp_cell(wb, "=" + answer_text)
            return format_value(val), False

        # Pattern 5: Plain text — return as-is
        return answer_text, False

    except Exception as e:
        print(f"    WARNING: Resolution failed for \"{answer_text[:60]}\": {e}")
        return answer_text, True


def format_value(val):
    """Format a COM cell value to a clean string."""
    if val is None:
        return ""
    if isinstance(val, float):
        if val == int(val):
            return str(int(val))
        return str(val)
    return str(val).strip()


def read_questions_sheet(wb):
    """Read the Questions sheet and return a list of question dicts.

    Each dict has keys: question, reasoning_type, answer_raw, provenance, difficulty.
    Returns empty list if no Questions sheet found.
    """
    questions_sheet = None
    for i in range(1, wb.Sheets.Count + 1):
        sheet_name = wb.Sheets(i).Name.lower().strip()
        if sheet_name == "questions" or sheet_name.startswith("question"):
            questions_sheet = wb.Sheets(i)
            break

    if questions_sheet is None:
        # Try the last sheet as fallback (Questions is always the last sheet)
        last_sheet = wb.Sheets(wb.Sheets.Count)
        first_cell = str(last_sheet.Cells(1, 1).Value or "").lower().strip()
        if "question" in first_cell:
            questions_sheet = last_sheet

    if questions_sheet is None:
        return []

    questions = []
    used_range = questions_sheet.UsedRange
    row_count = used_range.Rows.Count

    # Row 1 is the header; data starts at row 2
    for row_idx in range(2, row_count + 1):
        question_text = questions_sheet.Cells(row_idx, 1).Value
        if not question_text or not str(question_text).strip():
            continue

        reasoning_type = str(questions_sheet.Cells(row_idx, 2).Value or "").strip()
        answer_raw = str(questions_sheet.Cells(row_idx, 3).Value or "").strip()
        provenance = str(questions_sheet.Cells(row_idx, 4).Value or "").strip()
        difficulty = str(questions_sheet.Cells(row_idx, 5).Value or "").strip()

        questions.append({
            "question": str(question_text).strip(),
            "reasoning_type": reasoning_type,
            "answer_raw": answer_raw,
            "provenance": provenance,
            "difficulty": difficulty,
        })

    return questions


def should_strip_row1(ws):
    """Determine if row 1 of a data sheet is a description note (not a real header).

    Returns (should_strip: bool, reason: str).
    """
    used_range = ws.UsedRange
    col_count = used_range.Columns.Count

    row1_values = []
    row2_values = []
    for col_idx in range(1, col_count + 1):
        v1 = ws.Cells(1, col_idx).Value
        v2 = ws.Cells(2, col_idx).Value
        if v1 is not None and str(v1).strip():
            row1_values.append(str(v1).strip())
        if v2 is not None and str(v2).strip():
            row2_values.append(str(v2).strip())

    # If row 2 is empty, row 1 is likely the real header
    if len(row2_values) == 0:
        return False, f"row2 is empty; row1 has {len(row1_values)} values -> keeping row 1"

    # Heuristic: description note has few cells, real headers have many
    if len(row1_values) <= 2 and len(row2_values) >= 3:
        # Extra check: row 1 cells are long text (description-like)
        avg_len_row1 = sum(len(v) for v in row1_values) / max(len(row1_values), 1)
        avg_len_row2 = sum(len(v) for v in row2_values) / max(len(row2_values), 1)
        if avg_len_row1 > avg_len_row2:
            return True, (
                f"row1 has {len(row1_values)} cell(s) (avg len {avg_len_row1:.0f}), "
                f"row2 has {len(row2_values)} cell(s) (avg len {avg_len_row2:.0f}) -> stripping row 1"
            )

    # If row 1 and row 2 have similar column counts, row 1 is likely a real header
    if len(row1_values) >= 3 and abs(len(row1_values) - len(row2_values)) <= 2:
        return False, (
            f"row1 has {len(row1_values)} values, row2 has {len(row2_values)} values "
            f"-> both look like headers, keeping row 1"
        )

    return False, (
        f"row1 has {len(row1_values)} values, row2 has {len(row2_values)} values -> keeping row 1 (default)"
    )


def clean_workbook(wb):
    """Remove Readme and Questions sheets, strip row 1 description notes from data sheets.

    Returns list of log messages.
    """
    logs = []

    # Collect sheet names to delete (iterate by name to avoid index shifting)
    sheets_to_delete = []
    for i in range(1, wb.Sheets.Count + 1):
        name = wb.Sheets(i).Name.lower().strip()
        if name in ("readme", "questions"):
            sheets_to_delete.append(wb.Sheets(i).Name)

    for name in sheets_to_delete:
        wb.Sheets(name).Delete()
        logs.append(f"  Deleted sheet: \"{name}\"")

    # Check row 1 on remaining data sheets
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        sheet_name = ws.Name
        strip, reason = should_strip_row1(ws)
        if strip:
            ws.Rows(1).Delete()
            logs.append(f"  Sheet \"{sheet_name}\": {reason}")
        else:
            logs.append(f"  Sheet \"{sheet_name}\": {reason}")

    return logs


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def process_workbook(excel_app, src_path, dst_dir, file_id, domain_slug, fname):
    """Process a single workbook. Returns (records, stats, logs).

    Single-open approach to protect the original file:
      1. Open original READ-ONLY
      2. Extract questions and resolve answers (no modifications)
      3. SaveAs to destination as unencrypted XLSX (workbook now points to dest)
      4. Clean the workbook (modifications only affect destination)
      5. Save and close
    """
    records = []
    stats = {"extracted": 0, "skipped_image": 0, "resolution_failed": 0}
    logs = []
    abs_src = os.path.abspath(src_path)
    dst_path = os.path.join(os.path.abspath(dst_dir), fname)

    wb = com_retry(
        excel_app.Workbooks.Open,
        abs_src,
        False,   # UpdateLinks
        True,    # ReadOnly — protects original from any write-back
    )

    try:
        # Step 1: Read questions (no modifications to workbook)
        questions = read_questions_sheet(wb)
        logs.append(f"  Found {len(questions)} question(s) in Questions sheet")

        # Step 2: Resolve answers (uses temp cell ZZ1 but workbook is read-only
        #         so changes can't propagate back to original)
        for q_idx, q in enumerate(questions, start=1):
            q_id = f"frtr_{file_id}_q{str(q_idx).zfill(2)}"

            if q["reasoning_type"] in IMAGE_REASONING_TYPES:
                logs.append(
                    f"  {q_id}: SKIPPED ({q['reasoning_type']}) "
                    f"\"{q['question'][:60]}...\""
                )
                stats["skipped_image"] += 1
                continue

            resolved_answer, failed = resolve_answer(wb, q["answer_raw"])
            if failed:
                stats["resolution_failed"] += 1

            question_type = map_question_type(q["reasoning_type"], q["provenance"])
            output_type = infer_output_type(resolved_answer)
            domain = DOMAIN_MAP.get(domain_slug, domain_slug)

            record = {
                "id": q_id,
                "question": q["question"],
                "reference": resolved_answer,
                "answer_source_doc": [fname],
                "input_type": ["xlsx"],
                "question_type": question_type,
                "output_type": output_type,
                "domain": domain,
                "reference_tag": {
                    "human_verified": "TRUE",
                    "AI_verified": "FALSE",
                },
                "reference_image_or_file": "",
                "archived": {
                    "question_type_old": q["reasoning_type"],
                    "reference_location": q["answer_raw"],
                    "note": f"Difficulty: {q['difficulty']}; Provenance: {q['provenance']}",
                },
                "question_generation_method": "benchmark",
                "c1_bot_name": "FRTR_bench",
                "original_question_id": q_idx,
                "query_source": "FRTR-bench",
            }

            records.append(record)
            stats["extracted"] += 1

            status = "FAILED" if failed else output_type
            logs.append(
                f"  {q_id}: \"{q['question'][:50]}...\" -> "
                f"reference={resolved_answer[:40]} ({status})"
            )

        # Step 3: Remove IRM protection and SaveAs to destination
        try:
            wb.Permission.Enabled = False
            time.sleep(2)  # let Excel settle after IRM removal
            logs.append(f"  IRM protection removed")
        except Exception as e:
            logs.append(f"  IRM removal skipped: {e}")
        wb.SaveAs(dst_path, FileFormat=51)  # 51 = xlOpenXMLWorkbook
        time.sleep(1)  # let SaveAs complete
        logs.append(f"  SaveAs: {dst_path}")

        # Step 4: Clean workbook (now pointing to destination, not original)
        clean_logs = clean_workbook(wb)
        logs.extend(clean_logs)

        # Step 5: Save cleaned version (re-SaveAs to ensure ZIP/XLSX format)
        wb.SaveAs(dst_path, FileFormat=51)
        logs.append(f"  Saved (cleaned)")

    finally:
        wb.Close(SaveChanges=False)

    return records, stats, logs


def main():
    parser = argparse.ArgumentParser(
        description="Convert FRTR-bench dataset to old_prod_simulation pipeline format."
    )
    parser.add_argument(
        "--src_dir",
        required=True,
        help="Path to FRTR-bench directory containing the 30 Excel workbooks.",
    )
    parser.add_argument(
        "--dst_dir",
        required=True,
        help="Output directory for cleaned XLSX files and frtr_bench.jsonl.",
    )
    args = parser.parse_args()

    src_dir = args.src_dir
    dst_dir = args.dst_dir

    if not os.path.isdir(src_dir):
        print(f"ERROR: Source directory not found: {src_dir}")
        sys.exit(1)

    os.makedirs(dst_dir, exist_ok=True)

    # Discover workbook files
    workbook_files = sorted(
        f for f in os.listdir(src_dir)
        if f.startswith("frtr_") and f.endswith(".xlsx")
    )

    if not workbook_files:
        print(f"ERROR: No frtr_*.xlsx files found in {src_dir}")
        sys.exit(1)

    print(f"Found {len(workbook_files)} workbook(s) in {src_dir}")
    print(f"Output directory: {dst_dir}")
    print()

    # Initialize Excel COM
    pythoncom.CoInitialize()

    def start_excel():
        app = win32com.client.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        app.Interactive = False
        app.AskToUpdateLinks = False
        return app

    def kill_excel(app):
        try:
            app.Quit()
        except Exception:
            pass
        # Force-kill any lingering Excel processes
        import subprocess
        subprocess.run(
            ["taskkill", "/F", "/IM", "EXCEL.EXE"],
            capture_output=True,
        )
        time.sleep(8)

    excel = start_excel()

    all_records = []
    total_stats = {"extracted": 0, "skipped_image": 0, "resolution_failed": 0}
    consecutive_failures = 0

    try:
        for idx, fname in enumerate(workbook_files, start=1):
            file_id, domain_slug = parse_filename(fname)
            if file_id is None:
                print(f"[{idx}/{len(workbook_files)}] SKIPPING {fname} (cannot parse filename)")
                continue

            print(f"[{idx}/{len(workbook_files)}] Processing {fname}...")
            src_path = os.path.join(src_dir, fname)
            dst_check = os.path.join(os.path.abspath(dst_dir), fname)

            # Skip if output already exists and is a valid ZIP (XLSX)
            if os.path.exists(dst_check):
                try:
                    with open(dst_check, "rb") as fcheck:
                        if fcheck.read(2) == b"PK":
                            print(f"  SKIPPED (already converted)")
                            print()
                            continue
                except Exception:
                    pass

            # Fresh Excel instance per file for maximum COM stability
            kill_excel(excel)
            excel = start_excel()

            try:
                records, stats, logs = process_workbook(
                    excel, src_path, dst_dir, file_id, domain_slug, fname
                )
                all_records.extend(records)
                for key in total_stats:
                    total_stats[key] += stats[key]
                for log_line in logs:
                    print(log_line.encode("ascii", errors="replace").decode("ascii"))
            except Exception as e:
                print(f"  ERROR processing {fname}: {e}")

            print()

    finally:
        kill_excel(excel)
        pythoncom.CoUninitialize()

    # Write JSONL — merge with existing records from previous runs
    jsonl_path = os.path.join(dst_dir, "frtr_bench.jsonl")
    existing_records = []
    existing_ids = set()
    if os.path.exists(jsonl_path):
        with open(jsonl_path, "r", encoding="utf-8") as f:
            for line in f:
                r = json.loads(line.strip())
                existing_records.append(r)
                existing_ids.add(r["id"])
    # Add only new records (avoid duplicates)
    for r in all_records:
        if r["id"] not in existing_ids:
            existing_records.append(r)
    with open(jsonl_path, "w", encoding="utf-8") as f:
        for record in existing_records:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    print(f"  JSONL total: {len(existing_records)} ({len(existing_records) - len(existing_ids)} new)")

    print("=" * 60)
    print(f"Done. {len(workbook_files)} workbooks processed.")
    print(f"  Questions extracted: {total_stats['extracted']}")
    print(f"  Questions skipped (image): {total_stats['skipped_image']}")
    print(f"  Resolution failures: {total_stats['resolution_failed']}")
    print(f"  JSONL written to: {jsonl_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
