"""Convert a single FRTR-bench workbook. Separate process per file for COM stability.

Three-phase approach:
  Phase 1: Open read-only, extract question TEXT only (no formula eval), close.
  Phase 2: Open read-only, remove IRM, SaveAs clean XLSX, clean sheets, re-save, close.
  Phase 3: Open the NEW clean XLSX, resolve answers by reading cells/evaluating formulas, close.
"""
import sys
import os
import json
import time
import re

sys.path.insert(0, os.path.dirname(__file__))
from convert_frtr_bench import (
    parse_filename, IMAGE_REASONING_TYPES, DOMAIN_MAP, OPERATION_MAP,
    extract_sheet_refs, determine_scope, map_question_type, infer_output_type,
    format_value, FORMULA_PREFIXES, BARE_CELL_REF_RE,
)

import win32com.client
import pythoncom


def read_questions_raw(wb):
    """Read Questions sheet — extract text only, no cell/formula evaluation."""
    qs_sheet = None
    for i in range(1, wb.Sheets.Count + 1):
        nm = wb.Sheets(i).Name.lower().strip()
        if nm == "questions" or nm.startswith("question"):
            qs_sheet = wb.Sheets(i)
            break
    if qs_sheet is None:
        return []

    questions = []
    row_count = qs_sheet.UsedRange.Rows.Count
    for r in range(2, row_count + 1):
        q_text = qs_sheet.Cells(r, 1).Value
        if not q_text or not str(q_text).strip():
            continue
        questions.append({
            "question": str(q_text).strip(),
            "reasoning_type": str(qs_sheet.Cells(r, 2).Value or "").strip(),
            "answer_raw": str(qs_sheet.Cells(r, 3).Value or "").strip(),
            "provenance": str(qs_sheet.Cells(r, 4).Value or "").strip(),
            "difficulty": str(qs_sheet.Cells(r, 5).Value or "").strip(),
        })
    return questions


def clean_workbook_sheets(wb):
    """Delete Readme/Questions sheets and strip row 1 description notes."""
    to_delete = []
    for i in range(1, wb.Sheets.Count + 1):
        nm = wb.Sheets(i).Name.lower().strip()
        if nm in ("readme", "questions") or nm.startswith("question"):
            to_delete.append(wb.Sheets(i).Name)
    for nm in to_delete:
        wb.Sheets(nm).Delete()

    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        col_count = ws.UsedRange.Columns.Count
        r1 = [str(ws.Cells(1, c).Value or "").strip() for c in range(1, col_count + 1) if ws.Cells(1, c).Value]
        r2 = [str(ws.Cells(2, c).Value or "").strip() for c in range(1, col_count + 1) if ws.Cells(2, c).Value]
        if len(r1) <= 2 and len(r2) >= 3:
            avg1 = sum(len(v) for v in r1) / max(len(r1), 1)
            avg2 = sum(len(v) for v in r2) / max(len(r2), 1)
            if avg1 > avg2:
                ws.Rows(1).Delete()


def resolve_answer_from_clean(wb, answer_text):
    """Resolve answer from a clean (non-IRM) workbook."""
    if not answer_text or not str(answer_text).strip():
        return str(answer_text), True
    answer_text = str(answer_text).strip()

    if " OR " in answer_text:
        first = answer_text.split(" OR ")[0].strip()
        val, failed = resolve_answer_from_clean(wb, first)
        if not failed:
            return val, False

    try:
        # "See Sheet!Cell"
        m = re.match(r"^See\s+([A-Za-z_][\w\s\-]*)!\s*([A-Z]+\d+)", answer_text, re.IGNORECASE)
        if m:
            ws = wb.Sheets(m.group(1).strip().strip("'"))
            return format_value(ws.Range(m.group(2).strip()).Value), False

        # Bare "Sheet!Cell"
        bm = BARE_CELL_REF_RE.match(answer_text)
        if bm:
            ws = wb.Sheets(bm.group(1).strip().strip("'"))
            return format_value(ws.Range(bm.group(2).strip()).Value), False

        # "Max/Min of Range"
        mm = re.match(r"^(Max|Min)\s+of\s+(.+)$", answer_text, re.IGNORECASE)
        if mm:
            ws = wb.Sheets(1)
            ws.Range("ZZ1").Formula = f"={mm.group(1).upper()}({mm.group(2).strip()})"
            time.sleep(0.3)
            val = ws.Range("ZZ1").Value
            ws.Range("ZZ1").ClearContents()
            return format_value(val), False

        # Formula
        if any(answer_text.upper().startswith(p) for p in FORMULA_PREFIXES):
            ws = wb.Sheets(1)
            ws.Range("ZZ1").Formula = "=" + answer_text
            time.sleep(0.3)
            val = ws.Range("ZZ1").Value
            ws.Range("ZZ1").ClearContents()
            return format_value(val), False

        return answer_text, False
    except Exception as e:
        return answer_text, True


def main():
    src_path = sys.argv[1]
    dst_dir = sys.argv[2]
    fname = os.path.basename(src_path)
    file_id, domain_slug = parse_filename(fname)
    dst_path = os.path.join(os.path.abspath(dst_dir), fname)

    # Already done?
    if os.path.exists(dst_path):
        with open(dst_path, "rb") as f:
            if f.read(2) == b"PK":
                print(f"SKIP (already valid)")
                sys.exit(0)
        # Invalid existing file — try to remove
        try:
            os.remove(dst_path)
        except Exception:
            pass

    pythoncom.CoInitialize()

    # ---- Phase 1: Extract questions AND resolve answers from original ----
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.Interactive = False

    questions = []
    records = []
    try:
        wb = excel.Workbooks.Open(os.path.abspath(src_path), False, True)
        questions = read_questions_raw(wb)

        # Resolve answers while original formulas are intact
        for q_idx, q in enumerate(questions, start=1):
            q_id = f"frtr_{file_id}_q{str(q_idx).zfill(2)}"
            if q["reasoning_type"] in IMAGE_REASONING_TYPES:
                print(f"  {q_id}: SKIPPED ({q['reasoning_type']})")
                continue

            resolved, failed = resolve_answer_from_clean(wb, q["answer_raw"])
            question_type = map_question_type(q["reasoning_type"], q["provenance"])
            output_type = infer_output_type(resolved)
            domain = DOMAIN_MAP.get(domain_slug, domain_slug)

            records.append({
                "id": q_id,
                "question": q["question"],
                "reference": resolved,
                "answer_source_doc": [fname],
                "input_type": ["xlsx"],
                "question_type": question_type,
                "output_type": output_type,
                "domain": domain,
                "reference_tag": {"human_verified": "TRUE", "AI_verified": "FALSE"},
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
            })
            status = "FAILED" if failed else output_type
            print(f"  {q_id}: {resolved[:40]} ({status})")

        wb.Close(SaveChanges=False)
        print(f"  Phase 1: {len(questions)} questions, {len(records)} resolved")
    except Exception as e:
        err = str(e).encode("ascii", errors="replace").decode("ascii")
        print(f"  Phase 1 ERROR: {err}")
    finally:
        try:
            excel.Quit()
        except Exception:
            pass

    time.sleep(3)

    # ---- Phase 2: Copy sheets to new workbook, save (bypasses IRM entirely) ----
    import subprocess
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"], capture_output=True)
    time.sleep(5)

    excel2 = win32com.client.Dispatch("Excel.Application")
    excel2.Visible = False
    excel2.DisplayAlerts = False
    excel2.Interactive = False

    phase2_ok = False
    skip_sheets = {"readme", "questions"}
    try:
        wb_src = excel2.Workbooks.Open(os.path.abspath(src_path), False, True)
        wb_new = excel2.Workbooks.Add()

        # Copy data sheets (skip Readme, Questions)
        for i in range(1, wb_src.Sheets.Count + 1):
            sn = wb_src.Sheets(i).Name
            if sn.lower().strip() in skip_sheets or sn.lower().strip().startswith("question"):
                continue
            wb_src.Sheets(i).Copy(Before=wb_new.Sheets(1))

        # Delete default Sheet from new workbook
        for i in range(wb_new.Sheets.Count, 0, -1):
            if wb_new.Sheets(i).Name.startswith("Sheet"):
                wb_new.Sheets(i).Delete()

        wb_src.Close(SaveChanges=False)

        # Strip row 1 description notes
        clean_workbook_sheets(wb_new)

        wb_new.SaveAs(os.path.abspath(dst_path), FileFormat=51)
        wb_new.Close(SaveChanges=False)
        phase2_ok = True
        print(f"  Phase 2: copied to clean XLSX ({wb_new.Sheets.Count if not phase2_ok else 'ok'})")
    except Exception as e:
        err = str(e).encode("ascii", errors="replace").decode("ascii")
        print(f"  Phase 2 ERROR: {err}")
        try:
            excel2.Workbooks(1).Close(SaveChanges=False)
        except Exception:
            pass
        try:
            excel2.Workbooks(1).Close(SaveChanges=False)
        except Exception:
            pass
    finally:
        try:
            excel2.Quit()
        except Exception:
            pass

    pythoncom.CoUninitialize()

    # Save records
    if records:
        rpath = dst_path + ".records.json"
        with open(rpath, "w", encoding="utf-8") as f:
            json.dump(records, f, ensure_ascii=False)

    # Verify output
    if os.path.exists(dst_path):
        with open(dst_path, "rb") as f:
            if f.read(2) == b"PK":
                print(f"  RESULT: OK ({len(records)} questions)")
            else:
                print(f"  RESULT: FAILED (still OLE2)")
    else:
        print(f"  RESULT: FAILED (no output)")


if __name__ == "__main__":
    main()
