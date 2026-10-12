# `xlsx2json.py` — Excel → JSON converter for patent-claim error annotations

This tool converts the annotated Excel (`.xlsx`) files used in the Patent Claims
Translation/Evaluation Task into a single JSON file per workbook. Each JSON holds
the source text, the translated text, the quality score, and the error
annotations (spans, categories, severities), in a format that is easy to process
programmatically.

The converter auto-detects the translation direction (en-ja / ja-en) from the
text, and auto-detects how the annotation highlights are laid out, so the same
command works for every file in the task.

## Requirements

- Python 3.8+
- [`openpyxl`](https://pypi.org/project/openpyxl/)

```bash
pip install -r requirements.txt   # or: pip install openpyxl
```

## Usage

```bash
# One file → print JSON to stdout
python xlsx2json.py path/to/en-ja_sys1_03.xlsx

# One file → write to a JSON file
python xlsx2json.py path/to/01.xlsx -o 01.json

# A whole directory of .xlsx → one <name>.json per file in an output directory
python xlsx2json.py path/to/dir --outdir path/to/out

# A whole directory → a single combined JSON ({"files": [...]})
python xlsx2json.py path/to/dir -o all.json
```

Options:

| Option | Description |
|---|---|
| `-o, --output FILE` | Write a single JSON file (combined when the input is a directory). |
| `--outdir DIR` | Write one `<basename>.json` per input `.xlsx` into `DIR`. |
| `--compact` | Emit single-line JSON (no indentation). |
| `--highlight {auto,rows,grid,ref}` | Highlight layout (default `auto`; see below). |
| `-q, --quiet` | Suppress progress messages. |

Files beginning with `~$` (Excel lock files) and workbooks that fail to open are
skipped automatically.

## Highlight layout (`--highlight`)

The task's Excel files use a few different ways of marking errors. `auto` (the
default) picks the right one per file; you normally do not need to set this.

- **`rows`** — labels such as `訳抜け／大` sit in a yellow cell on a separate
  annotation row below the character grid (WAT2025 human-evaluation / checked data).
- **`ref`** — error spans are highlighted in color directly on the character grid,
  and a label `severity（category）：explanation` sits on the row just below, aligned
  to the start of the span (organizer reference data).
- **`grid`** — spans are highlighted on the character grid with no label.

## Output JSON schema

```jsonc
{
  "file": "en-ja_sys1_03.xlsx",
  "blocks": [
    {
      "index": 1,                 // 1-based block (claim) number in the file
      "score": 80,                // quality score (integer) or null
      "input_text": "...",        // source text, verbatim from the Input cell
      "output_text": "...",       // translation, verbatim from the Output cell
      "broad_error": "",          // "broad / wide-ranging error" field ("" if empty)
      "correction": "",           // "corrected translation" field ("" if empty)
      "input_text_grid": "...",   // text reconstructed from the character grid
      "output_text_grid": "...",  //   (start/end are offsets into these grid strings)
      "annotations": [
        {
          "side": "output",       // "input" (source) or "output" (translation)
          "start": 12,            // 0-based start offset into *_text_grid (inclusive)
          "end": 15,              // 0-based end offset (exclusive); text == grid[start:end]
          "text": "for",          // the highlighted substring
          "error": "訳抜け",       // category label ("" if the span has no label)
          "severity": "小",        // "大" / "中" / "小" / "" (or a range such as "中～大")
          "explanation": "",      // free-text note after "：" (ref data only; else "")
          "color": "yellow",      // highlight color (ref/grid data; else "")
          "cells": ["M7:O7"]      // source Excel cell range(s); multiple when a span wraps lines
        }
      ]
    }
  ]
}
```

### Notes on positions

- Annotation `start`/`end` are offsets into the same side's `*_text_grid` string,
  so `grid[start:end] == text` always holds. Slice the `*_text_grid` field, not
  `*_text`, to recover a span.
- `*_text_grid` equals the original cell text on the CJK side; on the Latin side it
  differs only in that line breaks and runs of spaces are normalized to a single
  space. Both the original (`*_text`) and the grid form (`*_text_grid`) are kept.
- Character counts are Unicode code points (a full-width character counts as one).
- When a single highlight wraps from the end of one display line to the start of
  the next, `cells` lists the rectangular ranges in reading order
  (e.g. `["ER14:ET14", "DG18:DG18"]`).

## Example

```bash
python xlsx2json.py ../2026/ref/ja-en --outdir ../2026/ref/ja-en/json
```

## See also

- [`mqm_score.py`](./README_mqm_score.md) — derive a GEMBA-MQM-style 0–100
  quality score for each claim from the error annotations in the JSON produced
  by this converter.
- [`evaluate.py`](./README_evaluate.md) — evaluate participant submissions
  against the gold reference: error-detection F1, severity agreement,
  per-category recall, and quality-score correlation (Pearson / Spearman).
