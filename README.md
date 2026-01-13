# FRTR-Bench

**FRTR-Bench** is a large-scale, multimodal benchmark for **enterprise spreadsheet reasoning**, designed to evaluate retrieval-augmented and multimodal LLM systems on realistic Excel workloads.

Unlike prior spreadsheet benchmarks that focus on single-sheet or text-only tables, FRTR-Bench captures the scale, structure, and modality of real-world enterprise workbooks.

---

## Overview

FRTR-Bench consists of **30 enterprise-grade Excel workbooks** spanning finance, supply chain, healthcare, energy, government, and education domains. The benchmark emphasizes three core challenges:

* **Scale**: Hundreds of thousands of rows per workbook, totaling ~4M cells
* **Cross-sheet reasoning**: Formulas and logic spanning multiple worksheets
* **Multimodality**: Embedded images such as charts, dashboards, and scanned receipts

Each workbook includes natural-language questions that require lookup, aggregation, formula reasoning, and multimodal interpretation.

---

## Dataset Statistics

| Metric               | Count     |
| -------------------- | --------- |
| Workbooks            | 30        |
| Sheets               | 155       |
| Rows                 | 656,457   |
| Cells                | 3,928,934 |
| Embedded Images      | 53        |
| Cross-Sheet Formulas | 30        |
| Questions            | 157       |

---

## Workbook Structure

Each Excel file typically contains:

* **Metadata sheet** (schema, descriptions)
* **1–5 data sheets** with large tabular content
* **Embedded images** (PNG charts, receipts, dashboards)
* **Questions sheet** with:

  * Natural-language queries
  * Ground-truth answers
  * Explicit provenance (cell references, formulas, or image IDs)

Questions span difficulty levels (*Easy, Medium, Hard*) based on workbook size and reasoning complexity.

---

## Tasks Evaluated

FRTR-Bench supports evaluation of:

* Spreadsheet question answering
* Cross-sheet numerical reasoning
* Formula identification and synthesis
* Multimodal reasoning over tables + images
* Retrieval-augmented spreadsheet understanding

The benchmark is **model-agnostic** and can be used with retrieval-based, compression-based, or long-context approaches.

---

## Usage

FRTR-Bench is provided as raw Excel workbooks (`.xlsx`) with accompanying question annotations.
No execution engine or evaluation script is enforced, allowing flexibility in how models retrieve, reason, and answer.

Researchers are encouraged to report:

* Answer accuracy
* Token usage
* Latency
* Evidence/provenance correctness

---

## Related Work

FRTR-Bench is introduced alongside the **FRTR (From Rows to Reasoning)** framework in:

> *From Rows to Reasoning: A Retrieval-Augmented Multimodal Framework for Spreadsheet Understanding*

The benchmark is designed to complement (and extend beyond) existing datasets such as SpreadsheetLLM and SpreadsheetBench.

---

## Citation

If you use FRTR-Bench in your research, please cite:

```bibtex
@article{gulati2025frtr,
  title={From Rows to Reasoning: A Retrieval-Augmented Multimodal Framework for Spreadsheet Understanding},
  author={Gulati, Anmol and Sen, Sahil and Sarguroh, Waqar and Paul, Kevin},
  year={2025}
}
```

