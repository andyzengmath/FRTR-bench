# FRTR-Bench Conversion Script Design

**Date:** 2026-02-10
**Goal:** Convert FRTR-bench dataset into the format expected by `old_prod_simulation` pipeline so we can benchmark Code Interpreter against 157 enterprise spreadsheet questions.

---

## Deliverable

A single Python script (`convert_frtr_bench.py`) that:
- Decrypts 30 IRM-encrypted Excel workbooks
- Extracts questions and resolves answers to plain-text values
- Cleans workbooks (removes non-data sheets, fixes row 1 headers)
- Generates a JSONL file compatible with `run_data_set.py`

## Design Decisions

| Decision | Choice | Rationale |
|----------|--------|-----------|
| Answer resolution | Resolve via COM + keep raw in `archived` | Audit trail preserved |
| Image questions | Skip entirely | Pipeline doesn't support images; cleaner dataset |
| Row 1 description notes | Strip with heuristic check | Avoid misidentified headers, but don't remove real headers |
| Readme/Questions sheets | Strip from output XLSX | Prevent answer leakage and noisy metadata |
| question_type mapping | Map to pipeline's 6 types (aggregate/lookup x single/multi sheet) | Consistent with existing eval taxonomy |
| Scope detection | Infer from Provenance column | Determines single_sheet vs multiple_sheets |

---

## Script Flow

```
convert_frtr_bench.py --src_dir <FRTR-bench> --dst_dir <frtr_data>

For each of the 30 workbooks:
  1. Open via win32com COM automation (handles IRM)
  2. Read Questions sheet
     - Extract rows: Question, ReasoningType, Answer, Provenance, Difficulty
     - Skip image questions (ReasoningType in: Chart reading, Trend/Visual, Trend reading)
  3. Resolve answers for remaining questions
     - Cell reference ("See Sheet!Cell") → read cell value
     - Formula ("SUM(...)") → evaluate via temp cell
     - Plain text → return as-is
     - On failure → log, store raw text, flag resolution_failed
  4. Delete Readme and Questions sheets from workbook
  5. For each data sheet, check row 1:
     - If description note → delete row 1 so headers become row 1
     - If real header → leave as-is
  6. SaveAs unencrypted XLSX (FileFormat=51) to dst_dir
  7. Append question records to frtr_bench.jsonl

Final: print summary (total extracted, skipped, failed)
```

---

## Row 1 Detection Heuristic

```
For each data sheet:
  row1_values = non-empty cells in row 1
  row2_values = non-empty cells in row 2

  Strip row 1 if ALL of:
    - len(row1_values) <= 2    (description notes are typically 1-2 cells)
    - len(row2_values) >= 3    (real headers have multiple columns)
    - row 2 values look like headers (short strings, no long sentences)

  Log every decision with sheet name and detected values.
```

---

## Answer Resolution

| Answer Pattern | Detection | Resolution |
|----------------|-----------|------------|
| `See Summary!B4` | Starts with "See " | `wb.Sheets("Summary").Range("B4").Value` |
| `SUM(Payer_Summary!D3:D7)` | Starts with known function name | Write `=formula` to temp cell ZZ1, read value, clear |
| `Max of Summary!D3:D6` | Starts with "Max of " / "Min of " | Parse range, evaluate via temp cell |
| `"May/2024-05"` | No pattern match | Return as-is (plain text) |
| Any resolution error | Exception caught | Log error, store raw text, set `resolution_failed: true` in archived |

---

## question_type Mapping

### Pipeline's Existing Types

| Operation | Single Sheet | Multiple Sheets |
|-----------|-------------|-----------------|
| Lookup | `lookup_single_sheet` | `lookup_multiple_sheets` |
| Aggregate | `aggregate_single_sheet` | `aggregate_multiple_sheets` |

Note: `_multiple_files` types are not used since FRTR-bench is always one workbook per question.

### Mapping Rules

**Operation** (from ReasoningType):
- `Aggregation` → aggregate
- `Computation` → aggregate
- `Cross-sheet SUMIFS` → aggregate
- `Counting/Filter` → aggregate
- `Comparison` → lookup

**Scope** (from Provenance column):
- References 2+ distinct sheet names → `_multiple_sheets`
- References 1 sheet or unclear → `_single_sheet`

**Combined**: `{operation}_{scope}`

Original `ReasoningType` preserved in `archived.question_type_old`.

---

## output_type Inference

```python
def infer_output_type(answer_value):
    try:
        float(str(answer_value).replace(",", "").replace("$", ""))
        return "number"
    except ValueError:
        return "text"
```

---

## JSONL Record Format

```json
{
  "id": "frtr_0001_q01",
  "question": "What is total revenue in Q2 2024?",
  "reference": "1234567.89",
  "answer_source_doc": ["frtr_0001_quarterly-summary.xlsx"],
  "input_type": ["xlsx"],
  "question_type": "aggregate_single_sheet",
  "output_type": "number",
  "domain": "finance",
  "reference_tag": {
    "human_verified": "TRUE",
    "AI_verified": "FALSE"
  },
  "reference_image_or_file": "",
  "archived": {
    "question_type_old": "Aggregation",
    "reference_location": "See Summary!B4",
    "note": "Difficulty: Easy; Provenance: Cell: Summary!B4; Derived from Revenue!F3:F602"
  },
  "question_generation_method": "benchmark",
  "c1_bot_name": "FRTR_bench",
  "original_question_id": 1,
  "query_source": "FRTR-bench"
}
```

### Field Derivations

| JSONL Field | Source |
|-------------|--------|
| `id` | `frtr_{file_id}_q{question_number}` (zero-padded) |
| `question` | Questions sheet column A |
| `reference` | Resolved answer value (plain text) |
| `answer_source_doc` | `["{filename}.xlsx"]` (the cleaned output file) |
| `input_type` | `["xlsx"]` (always) |
| `question_type` | Mapped from ReasoningType + Provenance (see above) |
| `output_type` | Inferred from resolved answer (`number` or `text`) |
| `domain` | Mapped from filename slug via DOMAIN_MAP |
| `archived.question_type_old` | Raw ReasoningType value |
| `archived.reference_location` | Raw Answer column (cell ref/formula) |
| `archived.note` | `"Difficulty: {difficulty}; Provenance: {provenance}"` |

---

## Domain Mapping

| Filename Slug | Domain |
|---------------|--------|
| quarterly-summary | finance |
| ops-inventory | supply_chain |
| consolidation | finance |
| payroll-summary | finance |
| ap-vendor-spend | finance |
| manufacturing-ops | manufacturing |
| ecommerce-metrics | retail |
| cashflow-consolidation | finance |
| healthcare-claims | healthcare |
| energy-production | energy |
| gov-grants-tracking | government |
| nonprofit-donations | nonprofit |
| education-outcomes | education |
| retail-pos | retail |
| telecom-churn | telecom |
| airline-ops | transportation |
| realestate-portfolio | real_estate |
| hospitality-bookings | hospitality |
| agriculture-yield | agriculture |
| pharma-trials | pharma |
| insurance-claims | insurance |
| logistics-routing | supply_chain |
| media-streaming | media |
| cybersecurity-incidents | cybersecurity |
| construction-controls | construction |
| consolidation-multi-subsidiary | finance |
| supply-chain-regional | supply_chain |
| payroll-analysis | finance |
| customer-analytics | retail |
| capex-project-portfolio | finance |

---

## CLI Interface

```bash
python convert_frtr_bench.py \
  --src_dir "C:\...\FRTR-bench" \
  --dst_dir "C:\...\frtr_data"
```

## Logging

- Progress per workbook: `[1/30] Processing frtr_0001_quarterly-summary.xlsx...`
- Per question: `  Q1: "What is total revenue..." → reference=1234567.89 (number)`
- Skipped questions: `  Q5: SKIPPED (Chart reading) "Which quarter shows..."`
- Row 1 decisions: `  Sheet "Revenue": row1 has 1 cell, row2 has 8 cells → stripping row 1`
- Resolution failures: `  Q3: RESOLUTION FAILED for "XLOOKUP(...)" → storing raw text`
- Summary: `Done. 30 workbooks, 132 questions extracted, 25 skipped (image), 0 failed.`

---

## Output Directory Structure

```
frtr_data/
├── frtr_bench.jsonl
├── frtr_0001_quarterly-summary.xlsx   (unencrypted, data-only, headers in row 1)
├── frtr_0002_ops-inventory.xlsx
├── ...
└── frtr_0030_capex-project-portfolio.xlsx
```

## Running the Pipeline After Conversion

```bash
python src/run_data_set.py \
  --input_file frtr_data/frtr_bench.jsonl \
  --data_dir frtr_data/ \
  --output_dir frtr_output/ \
  --runtime_config frtr_runtime_config.json \
  --batch_size 4
```

Recommended runtime config overrides for FRTR-bench:
- `excel_workflow: "excel2csv"`
- `enable_early_stop: true`
- `token_budget: 64000` (increased for large workbooks)
- `retry: 1`
