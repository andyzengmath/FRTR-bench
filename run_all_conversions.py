"""Wrapper that converts all FRTR-bench files by spawning a separate process per file."""
import subprocess
import os
import sys
import json
import time
import glob

PYTHON = r"C:\Users\andyzeng\AppData\Local\anaconda3\python.exe"
SCRIPT = os.path.join(os.path.dirname(__file__), "convert_one_file.py")

SRC_DIR = sys.argv[1] if len(sys.argv) > 1 else os.path.dirname(__file__)
DST_DIR = sys.argv[2] if len(sys.argv) > 2 else r"C:\temp\frtr_data"

os.makedirs(DST_DIR, exist_ok=True)

workbooks = sorted(
    f for f in os.listdir(SRC_DIR)
    if f.startswith("frtr_") and f.endswith(".xlsx")
)

print(f"Found {len(workbooks)} workbooks")
print(f"Output: {DST_DIR}")
print()

all_records = []
success = 0
skipped = 0
failed = 0

for idx, fname in enumerate(workbooks, 1):
    src = os.path.join(SRC_DIR, fname)
    dst = os.path.join(DST_DIR, fname)

    # Check if already done
    if os.path.exists(dst):
        with open(dst, "rb") as f:
            if f.read(2) == b"PK":
                print(f"[{idx}/{len(workbooks)}] {fname} - already valid, skipping")
                skipped += 1
                # Load existing records
                rpath = dst + ".records.json"
                if os.path.exists(rpath):
                    with open(rpath, "r", encoding="utf-8") as rf:
                        all_records.extend(json.load(rf))
                continue

    print(f"[{idx}/{len(workbooks)}] {fname}...")

    # Kill any lingering Excel
    subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                   capture_output=True, creationflags=0x08000000)
    time.sleep(5)

    # Run conversion in a separate process
    try:
        result = subprocess.run(
            [PYTHON, SCRIPT, src, DST_DIR],
            capture_output=True, text=True, timeout=600,
            encoding="utf-8", errors="replace",
        )
        stdout = result.stdout.strip() if result.stdout else "(no output)"
        print(stdout.encode("ascii", errors="replace").decode("ascii"))
        if result.stderr:
            for line in result.stderr.strip().split("\n")[:3]:
                err_line = line.encode("ascii", errors="replace").decode("ascii")
                print(f"  STDERR: {err_line}")
    except subprocess.TimeoutExpired:
        print("  TIMEOUT (600s)")
        subprocess.run(["taskkill", "/F", "/IM", "EXCEL.EXE"],
                       capture_output=True, creationflags=0x08000000)
        time.sleep(5)

    # Check if output is valid
    if os.path.exists(dst):
        with open(dst, "rb") as f:
            if f.read(2) == b"PK":
                success += 1
                rpath = dst + ".records.json"
                if os.path.exists(rpath):
                    with open(rpath, "r", encoding="utf-8") as rf:
                        all_records.extend(json.load(rf))
            else:
                print(f"  OUTPUT IS STILL OLE2 - removing")
                os.remove(dst)
                failed += 1
    else:
        failed += 1

    print()

# Write combined JSONL
jsonl_path = os.path.join(DST_DIR, "frtr_bench.jsonl")
seen_ids = set()
unique_records = []
for r in all_records:
    if r["id"] not in seen_ids:
        unique_records.append(r)
        seen_ids.add(r["id"])

with open(jsonl_path, "w", encoding="utf-8") as f:
    for r in unique_records:
        f.write(json.dumps(r, ensure_ascii=False) + "\n")

# Write runtime config
cfg_path = os.path.join(DST_DIR, "frtr_runtime_config.json")
if not os.path.exists(cfg_path):
    with open(cfg_path, "w") as f:
        json.dump({
            "max_length": 60000,
            "xlsx_meta_analysis_prompt": "excel_processor_v2_explanation.md",
            "use_xlsx_processor_in_main_prompt": False,
            "token_budget": 64000,
            "retry": 1,
            "enable_early_stop": True,
            "excel_workflow": "excel2csv",
            "large_dataset": False,
        }, f, indent=2)

print("=" * 60)
print(f"Done. {success} converted, {skipped} skipped, {failed} failed")
print(f"JSONL: {len(unique_records)} questions -> {jsonl_path}")
print("=" * 60)
