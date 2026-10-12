# Tools for the Patent Claims Translation/Evaluation Task

This directory contains three command-line tools that form a small pipeline for
the evaluation task:

```
Excel (.xlsx)  ──xlsx2json.py──▶  JSON  ──┬── mqm_score.py ──▶  per-claim quality score
                                          └── evaluate.py  ──▶  submission vs. gold metrics
```

1. [`xlsx2json.py`](#1-xlsx2jsonpy--excel--json) — convert the annotated Excel files into JSON.
2. [`mqm_score.py`](#2-mqm_scorepy--gemba-mqm-style-quality-score) — derive a GEMBA-MQM-style 0–100 quality score from the annotations.
3. [`evaluate.py`](#3-evaluatepy--evaluate-submissions-against-the-gold) — evaluate participant submissions against the gold reference.

## Requirements

- Python 3.8+
- [`openpyxl`](https://pypi.org/project/openpyxl/) — required by `xlsx2json.py` only.
- [`scipy`](https://scipy.org/) — optional; if present, `evaluate.py` reports correlation p-values.

```bash
pip install -r requirements.txt   # installs openpyxl
```

`mqm_score.py` and `evaluate.py` use only the standard library (SciPy optional).

---

## 1. `xlsx2json.py` — Excel → JSON

Converts the annotated Excel (`.xlsx`) files into one JSON file per workbook,
holding the source text, the translation, the quality score, and the error
annotations (spans, categories, severities). The direction (en-ja / ja-en) and
the highlight layout are auto-detected, so the same command works for every file.

### Usage

```bash
# One file → print JSON to stdout
python xlsx2json.py path/to/en-ja_sys1_03.xlsx

# One file → write a JSON file
python xlsx2json.py path/to/01.xlsx -o 01.json

# A directory of .xlsx → one <name>.json per file in an output directory
python xlsx2json.py path/to/dir --outdir path/to/out

# A directory → a single combined JSON ({"files": [...]})
python xlsx2json.py path/to/dir -o all.json
```

| Option | Description |
|---|---|
| `-o, --output FILE` | Write a single JSON file (combined when the input is a directory). |
| `--outdir DIR` | Write one `<basename>.json` per input `.xlsx` into `DIR`. |
| `--compact` | Emit single-line JSON (no indentation). |
| `--highlight {auto,rows,grid,ref}` | Highlight layout (default `auto`; see below). |
| `-q, --quiet` | Suppress progress messages. |

Files beginning with `~$` (Excel lock files) and workbooks that fail to open are
skipped automatically.

### Highlight layout (`--highlight`)

The task's Excel files mark errors in a few different ways; `auto` picks the
right one per file, so you normally do not set this.

- **`rows`** — labels such as `訳抜け／大` sit in a yellow cell on a separate
  annotation row below the character grid (WAT2025 human-evaluation / checked data).
- **`ref`** — error spans are highlighted in color directly on the character grid,
  and a label `severity（category）：explanation` sits on the row just below, aligned
  to the start of the span (organizer reference data).
- **`grid`** — spans are highlighted on the character grid with no label.

### Output JSON schema

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

**Notes on positions.** Annotation `start`/`end` are offsets into the same side's
`*_text_grid` string, so `grid[start:end] == text` always holds — slice
`*_text_grid`, not `*_text`, to recover a span. `*_text_grid` equals the original
cell text on the CJK side; on the Latin side it differs only in that line breaks
and runs of spaces are normalized to a single space. Character counts are Unicode
code points. When a highlight wraps from the end of one display line to the start
of the next, `cells` lists the rectangular ranges in reading order
(e.g. `["ER14:ET14", "DG18:DG18"]`).

---

## 2. `mqm_score.py` — GEMBA-MQM-style quality score

Derives a single 0–100 quality score for each claim from its error annotations,
reading the JSON produced by `xlsx2json.py`. Useful as a **reference quality
score** computed consistently from gold (or any) annotations.

### Usage

```bash
# Score every claim in a directory (prints a table)
python mqm_score.py ../2026/ref/ja-en

# Score one file and save per-block results as JSON
python mqm_score.py ../2026/ref/ja-en/01.json -o 01.score.json

# Change severity weights or disable clipping
python mqm_score.py ../2026/ref/ja-en --weight-medium 5 --no-clip
```

| Option | Default | Description |
|---|---|---|
| `--weight-small F`  | `1` | Weight of a 小 (minor) error. |
| `--weight-medium F` | `3` | Weight of a 中 (medium) error. |
| `--weight-large F`  | `5` | Weight of a 大 (major) error. |
| `--empty POLICY` | `minor` | Weight of spans with **no** severity label: `minor` / `medium` / `major` / `ignore`. |
| `--side {both,input,output}` | `both` | Which side's annotations to count. |
| `--no-clip` | off | Report raw `100 + 4s` without the −25 floor (scores may go negative). |
| `-o FILE` / `--compact` / `-q` | — | JSON output / single line / quiet. |

### Scoring method

```
penalty = Σ weight(error)
s       = -penalty              # optionally clipped so s ≥ -25
score   = 100 + 4 · s           # clipped score ∈ [0, 100]
```
following GEMBA-MQM (Kocmi and Federmann, 2023). The coefficient `4` is `100/25`:
with the per-segment floor `s = -25`, the penalty range `[0, 25]` maps onto
`[100, 0]`.

### Design rationale and decisions

GEMBA-MQM assumes **two** severity levels (major = 5, minor = 1), but the task's
gold uses **three** (小 / 中 / 大) plus ranges and some unlabeled spans. The tool
resolves this as follows; each choice is a CLI option so results are reproducible.

- **Three-level mapping (小=1, 中=3, 大=5).** Keep the GEMBA-MQM anchors and place
  中 at the midpoint (3) rather than collapsing it onto major/minor; this preserves
  小 < 中 < 大. `--weight-medium` is the most influential setting, so we recommend a
  sensitivity check (e.g. 中 as `1` or `5`).
- **Ranges (中～大, 小～中).** Scored as the **midpoint** of the endpoint weights
  (e.g. 中～大 → 4).
- **Unlabeled spans (`severity == ""`).** Counted as **minor** by default. In the
  gold these are highlighted error regions whose severity was conveyed by color
  grouping rather than a label, so they are genuine errors; counting them as minor
  is conservative. Use `--empty ignore` to drop them.
- **Clipping at −25.** On by default, matching GEMBA-MQM and the `100 + 4s`
  coefficient; a consequence is that heavily penalized claims saturate at 0.
  `--no-clip` disables it for inspection.
- **Both sides; per claim.** All annotations (both `input` and `output`) are
  counted by default; the unit is one block (claim = one segment).

The produced score is a **construct** of the formula and weights, not a human
judgment; report the settings used when comparing scores.

---

## 3. `evaluate.py` — Evaluate submissions against the gold

Scores each participant submission against the gold reference: how well a system
**locates** errors, agrees on **severity**, and whether its **quality score**
tracks the reference MQM score. It reads `xlsx2json.py` output and reuses
`mqm_score.py` for the reference score.

### Input layout

```
<gold_dir>/<direction>/<name>.json               # e.g. 2026/ref/ja-en/01.json
<submissions_dir>/<direction>/<team>/<name>.json  # e.g. 2026/submissions/ja-en/EHIME-U-GPT6/01.json
```
`<direction>` is `ja-en` or `en-ja`. Files match by basename, blocks by `index`.

### Usage

```bash
python evaluate.py ../2026/ref ../2026/submissions                       # summary table
python evaluate.py ../2026/ref ../2026/submissions --outdir ../2026/evaluation
```

| Option | Default | Description |
|---|---|---|
| `--outdir DIR` | — | Write `<direction>/<team>.json` and `summary.json`. |
| `--iou F` | `0.5` | IoU threshold for span-level matching. |
| `--weight-small/medium/large F` | `1/3/5` | Severity weights for the reference MQM score. |
| `--empty POLICY` | `minor` | Weight of unlabeled gold spans in the reference score. |
| `--no-clip` | off | Do not clip the reference MQM score at −25. |
| `-q, --quiet` | — | Suppress the console table. |

Score options are passed through to `mqm_score.py`.

### Metrics

- **Error detection (character level, primary).** Precision / recall / F1 with
  each character treated as "marked as an error or not," for `input`, `output`,
  and combined.
- **Error detection (span level).** Greedy matching at IoU ≥ `--iou`, reported as F1.
- **Severity agreement.** Over characters marked by both sides, the fraction whose
  severity matches.
- **Category.** Per-gold-category **recall** (location only) and a gold→submission
  **confusion** mapping.
- **Quality score.** Pearson r and Spearman ρ (with p-values if SciPy is present)
  between the reference MQM score and the submitted score, plus MAE and both means.

### Implementation policy / design rationale

- **Re-anchoring to gold coordinates.** Submissions differ from the gold in
  coordinate system (raw-text vs. character-grid offsets), claim-number handling,
  and whitespace, so raw `start`/`end` are not comparable. Each annotation's
  verbatim `text` and offsets are located in its own non-whitespace character
  sequence and mapped onto the gold's by aligning the two with `difflib`.
  Evaluating the gold against itself yields F1 = 1.0, validating the procedure.
- **Character level is primary.** It is robust to boundary disagreements; span
  F1 (IoU ≥ 0.5) is a stricter secondary view.
- **Category identity is not scored directly.** The gold uses abstract MQM
  categories while submissions use finer labels, so we report location-only recall
  per gold category plus a confusion mapping (which also reveals how the two
  inventories correspond).
- **Scoring.** The reference comes from the gold annotations via `mqm_score.py`, so
  it depends on the severity weights and `--empty`/clip settings (report them).
  Pearson/Spearman are invariant to offset/scale, so they measure rank consistency;
  submitted scores are typically far more lenient (large MAE), reported separately.
- **Input-side signal is weak.** Almost all gold errors are on the output side, so
  input-side figures rest on very few annotations; read the combined/output figures.

### Output JSON (per team, under `--outdir`)

```jsonc
{
  "direction": "ja-en", "team": "EHIME-U-GPT6", "files": 10,
  "character_level": { "input": {...}, "output": {...}, "all": {"precision":...,"recall":...,"f1":...} },
  "span_level":      { "input": {...}, "output": {...}, "all": {...} },
  "severity_agreement_on_overlap": { "overlap_chars": 4726, "match_rate": 0.9124 },
  "per_gold_category_recall": { "正確性": {"recall":0.40,"gold_chars":1062,"detected":427}, ... },
  "confusion_gold_to_sub":   { "正確性": {"誤訳":190, "誤訳:係り先":157, ...}, ... },
  "score_correlation": { "n":10, "pearson":0.6487, "spearman":0.6333, "mae":42.1, "gold_mean":34.4, "sub_mean":76.5 }
}
```
`summary.json` collects one such object per `"<direction>/<team>"` key.

---

## Caveats

- The current test set is small (10 JA–EN, 12 EN–JA claims), so detection scores
  and especially correlations are indicative rather than definitive.
- The reference MQM quality score is a construct of the formula and weighting
  choices above, not a human judgment; report the settings (weights, `--empty`,
  clipping) and prefer Spearman (with a 中-weight sensitivity check) for score
  comparisons.

## Reference

Tom Kocmi and Christian Federmann. 2023. *GEMBA-MQM: Detecting Translation
Quality Error Spans with GPT-4.* Proceedings of the Eighth Conference on Machine
Translation (WMT).
