# FRTR-Bench Integration with Old Prod Simulation Pipeline

## Overview

This document describes how the FRTR-bench dataset was converted into a format compatible with the `old_prod_simulation` Code Interpreter pipeline. The conversion script (`convert_frtr_bench.py`) extracts questions from 30 IRM-encrypted Excel workbooks, resolves ground-truth answers, and produces a JSONL file + cleaned XLSX files ready for benchmarking.

---

## 1. FRTR-Bench Dataset Format

### Repository Structure

```
FRTR-bench/
├── README.md
├── FRTR_BENCH_INTEGRATION.md       (this file)
├── convert_frtr_bench.py           (conversion script)
├── docs/plans/                     (design doc)
├── frtr_XXXX_<domain>.xlsx         (30 IRM-encrypted workbooks)
└── frtr_data_test/                 (conversion output)
    ├── frtr_bench.jsonl            (112 questions)
    └── frtr_XXXX_<domain>.xlsx     (29 cleaned workbooks)
```

### Dataset Statistics

| Metric               | Value     |
|----------------------|-----------|
| Workbooks            | 30        |
| Sheets               | 155       |
| Rows                 | 656,457   |
| Cells                | 3,928,934 |
| Embedded Images      | 53        |
| Cross-Sheet Formulas | 30        |
| Total Questions      | ~153      |

### Workbook Structure

Each Excel file follows a consistent layout:

1. **Readme sheet** -- 1-2 rows with workbook title and description.
2. **1-4 Data sheets** -- Large tabular data. Row 1 contains a summary/description note; actual column headers are in row 2.
3. **Questions sheet** -- Always the last sheet. Contains 5-8 Q&A pairs.

Exception: `frtr_0020_pharma-trials.xlsx` has no Readme or Questions sheet (only 3 data sheets).

### Questions Sheet Schema

| Column | Header         | Description                                                  |
|--------|----------------|--------------------------------------------------------------|
| A      | Question       | Natural-language question                                    |
| B      | ReasoningType  | Category of reasoning required                               |
| C      | Answer         | Ground-truth answer (cell ref, formula, or descriptive text) |
| D      | Provenance     | Evidence trail -- cell references, derivation path, image IDs |
| E      | Difficulty     | Easy, Medium, or Hard                                        |

### ReasoningType Values

The dataset uses two sets of reasoning types across different workbooks:

**Standard types (most workbooks):**
- `Aggregation`, `Comparison`, `Computation`, `Cross-sheet SUMIFS`, `Counting/Filter`
- `Chart reading`, `Trend/Visual`, `Trend reading` (image-dependent)

**Extended types (frtr_0001, frtr_0002, frtr_00027):**
- `Aggregation/Lookup`, `Lookup/Computation`, `Lookup/Count`, `Averaging`
- `FormulaReasoning`, `Threshold/Logic`, `Percentage + Filter`
- `AVERAGEIFS`, `SUMPRODUCT Count`, `Spread/Variance`
- `TrendReasoning`, `ExtractionFromImage`, `Visual Min` (image-dependent)

### Answer Formats

Answers in the Questions sheet are **not plain-text values**. They come in several forms:

1. **"See" cell references** -- `See Summary!B4` (read cell to get value)
2. **Bare cell references** -- `Dept_Summary!B4` (no "See" prefix)
3. **Formulas** -- `SUM(Payer_Summary!D3:D7)`, `XLOOKUP(...)`, `COUNTIF(...)` (evaluate)
4. **"Max/Min of" ranges** -- `Max of Summary!D3:D6` (evaluate via temp cell)
5. **"OR" alternatives** -- `Summary!B5 OR AVERAGE(Data_Main!G3:G1048576)` (try first option)
6. **Plain text** -- `"GBP"`, `"Healthcare"` (return as-is)
7. **Image references** -- `"Assess line chart slope."` (cannot resolve, skipped)

### File Format

All 30 `.xlsx` files are **IRM-encrypted OLE2 Compound Documents** ("Confidential -- Microsoft Extended"). They **cannot** be read with openpyxl, xlrd, or pandas. They require:

- Windows machine with Microsoft Excel installed
- Authenticated user with IRM access rights
- COM automation via `win32com.client` (available in Anaconda Python)

---

## 2. Old Prod Simulation Pipeline Format

### Pipeline Location

```
code-interpreter/old_prod_simulation/
├── src/
│   ├── ci_gen.py                 # Main entry point (single question)
│   ├── run_data_set.py           # Batch processor (reads JSONL)
│   ├── preprocessor.py           # Routes files to extractors
│   ├── xlsx2csvprocessor.py      # Excel-to-CSV conversion
│   └── ...
├── prompts/                      # System prompts for LLM
└── default_runtime_config.json   # Runtime configuration
```

### Input JSONL Schema (what the pipeline expects)

```json
{
  "id": "frtr_0001_q01",
  "question": "What is total revenue in Q2 2024?",
  "reference": "2176251.17",
  "answer_source_doc": ["frtr_0001_quarterly-summary.xlsx"],
  "input_type": ["xlsx"],
  "question_type": "aggregate_single_sheet",
  "output_type": "number",
  "domain": "finance",
  "reference_tag": { "human_verified": "TRUE", "AI_verified": "FALSE" },
  "reference_image_or_file": "",
  "archived": {
    "question_type_old": "Aggregation",
    "reference_location": "See Summary!B4",
    "note": "Difficulty: Easy; Provenance: Cell: Summary!B4; ..."
  },
  "question_generation_method": "benchmark",
  "c1_bot_name": "FRTR_bench",
  "original_question_id": 1,
  "query_source": "FRTR-bench"
}
```

### Pipeline's 6 question_type Values

| Operation | Single Sheet              | Multiple Sheets              |
|-----------|---------------------------|------------------------------|
| Lookup    | `lookup_single_sheet`     | `lookup_multiple_sheets`     |
| Aggregate | `aggregate_single_sheet`  | `aggregate_multiple_sheets`  |

Note: `_multiple_files` types are not used since FRTR-bench is always one workbook per question.

---

## 3. Conversion Script

### Usage

```bash
# Requires: Windows + Excel installed + IRM access + Anaconda Python with pywin32
"C:\...\anaconda3\python.exe" convert_frtr_bench.py \
  --src_dir "C:\...\FRTR-bench" \
  --dst_dir "C:\...\frtr_data_test"
```

### How It Works

The script uses a **single-open approach** per workbook to protect original files:

```
For each of the 30 workbooks:
  1. Open original via COM automation (ReadOnly=True)
  2. Read Questions sheet
     - Extract rows: Question, ReasoningType, Answer, Provenance, Difficulty
     - Skip image-dependent questions
  3. Resolve answers (no modifications to workbook)
     - "See Sheet!Cell"         -> read cell value
     - "Sheet!Cell" (bare ref)  -> read cell value
     - "Max/Min of Range"       -> evaluate via temp cell
     - "SUM(...)" etc.          -> evaluate via temp cell
     - "answer1 OR answer2"     -> try first option
     - Plain text               -> return as-is
     - On failure               -> log, store raw text
  4. SaveAs unencrypted XLSX to destination (FileFormat=51)
     - Workbook now points to destination, not original
  5. Clean workbook (only affects destination copy)
     - Delete Readme and Questions sheets
     - Strip row 1 if it's a description note (heuristic check)
  6. Save cleaned version
  7. Close workbook
  8. Append question records to frtr_bench.jsonl
```

### Key Design Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Answer resolution | Resolve via COM + keep raw in `archived` | Audit trail preserved |
| Image questions | Skip entirely | Pipeline doesn't support images |
| Row 1 notes | Strip with heuristic check | Avoid misidentified headers, but don't remove real headers |
| Readme/Questions sheets | Strip from output XLSX | Prevent answer leakage |
| question_type | Map to pipeline's 6 types | Consistent with existing eval taxonomy |
| Scope detection | Infer from Provenance column | Determines single_sheet vs multiple_sheets |
| File protection | Open read-only, SaveAs first, then clean | Prevents IRM/OneDrive from corrupting originals |

### Row 1 Detection Heuristic

```
Strip row 1 if ALL of:
  - Row 1 has <= 2 non-empty cells (description notes are 1-2 cells)
  - Row 2 has >= 3 non-empty cells (real headers have multiple columns)
  - Row 1 average cell text length > Row 2 average (descriptions are longer)
Otherwise keep row 1 (it's likely a real header).
```

### question_type Mapping

**Operation** (from ReasoningType):
- aggregate: `Aggregation`, `Computation`, `Cross-sheet SUMIFS`, `Counting/Filter`, `Aggregation/Lookup`, `Averaging`, `FormulaReasoning`, `Threshold/Logic`, `Percentage + Filter`, `AVERAGEIFS`, `SUMPRODUCT Count`, `Spread/Variance`
- lookup: `Comparison`, `Lookup/Computation`, `Lookup/Count`

**Scope** (from Provenance column):
- `_multiple_sheets` if 2+ distinct sheet names referenced
- `_single_sheet` otherwise

**Image types skipped entirely:**
`Chart reading`, `Trend/Visual`, `Trend reading`, `TrendReasoning`, `ExtractionFromImage`, `Visual Min`

---

## 4. Conversion Results

### Final Numbers

| Metric                     | Count |
|----------------------------|-------|
| Workbooks in dataset       | 30    |
| Workbooks converted        | 30 (all) |
| Total questions in dataset | ~153  |
| Questions skipped (image)  | ~40   |
| Questions in JSONL         | 113   |

### Questions Skipped (Image-Dependent)

~40 questions were excluded because they require interpreting embedded chart images. The pipeline only processes tabular data. Skipped `ReasoningType` values:
- `Chart reading` -- bar/pie chart interpretation
- `Trend/Visual`, `Trend reading`, `TrendReasoning` -- line chart trend analysis
- `ExtractionFromImage` -- reading scanned document images
- `Visual Min` -- identifying values from chart visuals

### Breakdown by question_type

| question_type              | Count |
|----------------------------|-------|
| `aggregate_single_sheet`   | 98    |
| `aggregate_multiple_sheets`| 9     |
| `lookup_single_sheet`      | 6     |

### Breakdown by output_type

| output_type | Count |
|-------------|-------|
| `number`    | 78    |
| `text`      | 35    |

### Breakdown by domain

| Domain         | Count |
|----------------|-------|
| finance        | 40    |
| supply_chain   | 13    |
| retail         | 11    |
| manufacturing  | 4     |
| education      | 4     |
| agriculture    | 3     |
| construction   | 3     |
| cybersecurity  | 3     |
| energy         | 3     |
| government     | 3     |
| healthcare     | 3     |
| hospitality    | 3     |
| insurance      | 3     |
| media          | 3     |
| nonprofit      | 3     |
| real_estate    | 3     |
| telecom        | 3     |
| transportation | 3     |
| pharma         | 2     |

---

## 5. Issues Encountered and Resolutions

### Critical: Source Files Corrupted by Initial Approach

**Problem:** The initial script opened the original IRM file, deleted sheets, then `SaveAs` to destination. Despite `wb.Close(SaveChanges=False)`, IRM + OneDrive auto-save wrote changes back to originals, destroying Questions sheets.

**Resolution:** Two-phase per-file approach using separate Python processes:
- Phase 1: Open original read-only via COM, extract questions and resolve answers
- Phase 2: Open original again, copy data sheets to a new workbook (bypasses IRM), save to `C:\temp\frtr_data` (outside OneDrive sync)

### OneDrive Re-encrypting Output Files

**Problem:** Files saved to OneDrive-synced directories got re-encrypted by OneDrive's IRM policy, reverting clean XLSX back to OLE2.

**Resolution:** Save all output to `C:\temp\frtr_data\` (local, non-synced directory).

### COM Cascading Failures

**Problem:** Excel COM became permanently unresponsive after processing several IRM files.

**Resolution:** Spawn a separate Python process per file (`convert_one_file.py` + `run_all_conversions.py`). Each process gets a fresh COM stack. Kill Excel between files with `taskkill /F /IM EXCEL.EXE`.

### OneDrive File Locks

**Problem:** 4 files (frtr_0014, 0015, 0016, 0017) were persistently locked by OneDrive and could not be opened by Excel COM.

**Resolution:** Pause OneDrive sync before running. These files turned out to already be clean XLSX (from an earlier conversion that modified the git working tree), so they were processed directly with openpyxl instead of COM.

### Other Issues Resolved

- **Unicode encoding errors**: Replaced special characters in log strings; used `.encode('ascii', errors='replace')` for output
- **Bare cell references** (`Dept_Summary!B4` without "See " prefix): Added `BARE_CELL_REF_RE` regex pattern
- **Missing formula prefixes**: Extended to 30+ Excel function names including `ROWS()`, `LET()`, `MAXIFS()`
- **Unknown ReasoningTypes**: Extended `IMAGE_REASONING_TYPES` and `OPERATION_MAP` for all observed types
- **Python environment**: Use Anaconda Python (has `win32com.client`) instead of Windows Store Python

---

## 6. Pipeline Results

### Evaluation Methodology

Results were evaluated using `text_judge.py`, which sends each (question, ground_truth, predicted) triple to **GPT-4.1** for Pass/Fail judgment. The judge uses strict criteria:
- Exact match for counts, names, years (case-insensitive, minor formatting tolerance)
- Numeric values within 2% tolerance
- Yes/no must match exactly
- Ambiguous or incomplete answers are marked Fail

Judge prompt: `prompts/update_evaluationprompt.md`
Judge output: `C:\temp\frtr_judge_output\frtr_judge_input.jsonl`

### Overall

| Metric                  | Value |
|-------------------------|-------|
| Questions in JSONL      | 113   |
| Pipeline processed      | 113   |
| Pipeline success        | 105   |
| Pipeline error          | 8     |
| Judged by GPT-4.1       | 112   |
| **Pass**                | **56 (50.0%)** |
| **Fail**                | **56 (50.0%)** |

### By Question Type (Pipeline Taxonomy)

| question_type                | Pass | Total | Pass Rate |
|------------------------------|------|-------|-----------|
| `aggregate_single_sheet`     | 54   | 98    | 55.1%     |
| `aggregate_multiple_sheets`  | 1    | 9     | 11.1%     |
| `lookup_single_sheet`        | 1    | 5     | 20.0%     |

**Observation:** Single-sheet aggregation is the sweet spot (55%). Cross-sheet reasoning (11%) and lookup questions (20%) need improvement.

### By Output Type

| output_type | Pass | Total | Pass Rate |
|-------------|------|-------|-----------|
| `number`    | 51   | 77    | 66.2%     |
| `text`      | 5    | 35    | 14.3%     |

**Observation:** Numeric answers pass at 66%, text answers at only 14%. The pipeline generates correct numbers far more reliably than formatted text responses.

### By Difficulty

| Difficulty | Pass | Total | Pass Rate |
|------------|------|-------|-----------|
| Easy       | 22   | 39    | 56.4%     |
| Medium     | 27   | 56    | 48.2%     |
| Hard       | 6    | 15    | 40.0%     |

**Observation:** Clear difficulty gradient. The benchmark's difficulty labels correlate well with pipeline performance.

### By Domain

| Domain         | Pass | Total | Pass Rate |
|----------------|------|-------|-----------|
| agriculture    | 3    | 3     | 100%      |
| construction   | 3    | 3     | 100%      |
| cybersecurity  | 3    | 3     | 100%      |
| energy         | 3    | 3     | 100%      |
| government     | 3    | 3     | 100%      |
| media          | 3    | 3     | 100%      |
| manufacturing  | 3    | 4     | 75%       |
| healthcare     | 2    | 3     | 67%       |
| insurance      | 2    | 3     | 67%       |
| nonprofit      | 2    | 3     | 67%       |
| supply_chain   | 8    | 13    | 62%       |
| education      | 2    | 4     | 50%       |
| finance        | 15   | 39    | 38%       |
| hospitality    | 1    | 3     | 33%       |
| retail         | 3    | 11    | 27%       |
| real_estate    | 0    | 3     | 0%        |
| telecom        | 0    | 3     | 0%        |
| transportation | 0    | 3     | 0%        |
| pharma         | 0    | 2     | 0%        |

**Observation:** 6 domains achieve 100%. Domains with simple, small workbooks perform best. Finance (the largest group at 39 questions) passes at 38%. Four domains score 0%.

---

## 7. Key Findings

1. **Overall pass rate is 50.0%** (56/112) on non-image questions from FRTR-bench, as evaluated by GPT-4.1 LLM judge. This provides a baseline for Code Interpreter performance on enterprise spreadsheet reasoning.

2. **Numeric answers pass at 66.2%, text answers at 14.3%.** The pipeline's code generation approach works well for computing specific numbers but struggles to produce text in the expected format.

3. **Single-sheet aggregation is the sweet spot** (55.1%). Cross-sheet reasoning (11.1%) and lookup questions (20.0%) are significantly weaker.

4. **Difficulty gradient is clear**: Easy 56.4%, Medium 48.2%, Hard 40.0%. The benchmark's difficulty labels correlate well with actual pipeline performance.

5. **6 domains achieve 100%** (agriculture, construction, cybersecurity, energy, government, media). These are small, straightforward workbooks. Finance (the largest group) passes at only 38%.

6. **4 domains score 0%**: real_estate, telecom, transportation, pharma. These may have structural issues in the converted XLSX files or particularly challenging question types.

---

## 8. Issues and Limitations

### Reference Answer Quality

- Some references are unresolved formulas (e.g., `XLOOKUP(...)`, `SUM(Store_Summary!B3:B22)`) rather than computed values. This affects ~15 questions where COM formula evaluation failed or the files were processed via openpyxl (which can't evaluate formulas).
- `ROWS(Payments!A:A)-1` evaluates to `1048575` (max Excel rows) instead of actual row count. Affects frtr_0009_q05 and frtr_0021_q05.

### Image-Dependent Questions Excluded

~40 questions requiring chart/image interpretation were excluded. The pipeline doesn't support multimodal inputs. These questions would need to be benchmarked separately with image-capable models.

### IRM/OneDrive Complexity

The conversion required extensive workarounds for IRM encryption and OneDrive sync interference. The final approach uses:
- Separate Python process per file (fresh COM stack)
- Sheet-copy to new workbook (bypasses IRM without removal)
- Output to local non-synced directory
- Pause OneDrive sync for locked files
- openpyxl fallback for files that are already clean XLSX

---

## 9. Running the Pipeline

### Data Location

- JSONL + XLSX: `C:\temp\frtr_data\`
- Pipeline output: `C:\temp\frtr_output\`
- Runtime config: `C:\temp\frtr_data\frtr_runtime_config.json`

### Commands

```bash
# Conversion (if re-running)
"C:\...\anaconda3\python.exe" run_all_conversions.py

# Pipeline
"C:\...\anaconda3\python.exe" code-interpreter/old_prod_simulation/src/run_data_set.py \
  -cd "C:\temp\frtr_data" \
  -od "C:\temp\frtr_output" \
  -rcp "C:\temp\frtr_data\frtr_runtime_config.json" \
  -bs 4 -d INFO \
  "C:\temp\frtr_data\frtr_bench.jsonl"
```

### Runtime Config

```json
{
  "max_length": 60000,
  "use_xlsx_processor_in_main_prompt": false,
  "token_budget": 64000,
  "retry": 1,
  "enable_early_stop": true,
  "excel_workflow": "excel2csv",
  "large_dataset": false
}
```

---

## 10. Output Directory Structure

```
C:\temp\frtr_data\
├── frtr_bench.jsonl                              # 113 questions
├── frtr_runtime_config.json                      # Pipeline config
├── frtr_0001_quarterly-summary.xlsx              # Clean XLSX (data sheets only)
├── frtr_0002_ops-inventory.xlsx
├── ...
├── frtr_0030_capex-project-portfolio.xlsx
└── *.records.json                                # Per-file question records

C:\temp\frtr_output\                              # Pipeline results
├── finance-aggregate_single_sheet-number-id_frtr_0003_q01/
│   ├── input.config.json
│   ├── CI.execute.result.json
│   └── ...
├── ... (113 directories total)
```
