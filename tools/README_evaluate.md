# `evaluate.py` — Evaluate submissions against the gold reference

This tool scores each participant submission against the organizers' gold
reference annotations. It reports, per team and direction, how well a system
**locates** translation errors, how well it agrees on **severity**, and how well
its **quality score** tracks a reference score derived from the gold.

It operates on the JSON produced by [`xlsx2json.py`](./README.md) and reuses
[`mqm_score.py`](./README_mqm_score.md) to compute the reference quality score.

## Requirements

- Python 3.8+ (standard library).
- Optional: [`scipy`](https://scipy.org/) — if installed, correlation **p-values**
  are reported; otherwise the coefficients are still computed (pure-Python
  fallback).

## Input layout

```
<gold_dir>/<direction>/<name>.json              # e.g. 2026/ref/ja-en/01.json
<submissions_dir>/<direction>/<team>/<name>.json # e.g. 2026/submissions/ja-en/EHIME-U-GPT6/01.json
```
where `<direction>` is `ja-en` or `en-ja`. Files are matched by basename and
blocks by their `index`.

## Usage

```bash
# Print a summary table for every team
python evaluate.py ../2026/ref ../2026/submissions

# Also write per-team JSON and summary.json
python evaluate.py ../2026/ref ../2026/submissions --outdir ../2026/evaluation
```

### Options

| Option | Default | Description |
|---|---|---|
| `--outdir DIR` | — | Write `<direction>/<team>.json` and `summary.json`. |
| `--iou F` | `0.5` | IoU threshold for span-level matching. |
| `--weight-small/medium/large F` | `1/3/5` | Severity weights for the reference MQM score. |
| `--empty POLICY` | `minor` | How unlabeled gold spans are weighted in the reference score. |
| `--no-clip` | off | Do not clip the reference MQM score at −25. |
| `-q, --quiet` | — | Suppress the console table. |

The score options are passed through to `mqm_score.py`; see its README for their
meaning.

## Metrics

- **Error detection (character level, primary).** Each character is treated as
  "marked as an error or not." We report precision, recall, and F1 for the
  `input` side, the `output` side, and combined. Precision = fraction of the
  system's marked characters that are also marked in the gold; recall = fraction
  of gold-marked characters recovered.
- **Error detection (span level).** Gold and predicted spans are greedily
  matched when their intersection-over-union is at least `--iou`; we report the
  resulting F1.
- **Severity agreement.** Over characters marked by both sides, the fraction
  whose severity label matches.
- **Category.** Per-gold-category **recall** (how well each gold category is
  *located*, regardless of the predicted label) and a gold→submission category
  **confusion** mapping.
- **Quality score.** Pearson r and Spearman ρ (with p-values if SciPy is
  available) between the reference MQM score and the submitted score, plus the
  mean absolute error (MAE) and the two means.

## Implementation policy / design rationale

- **Re-anchoring to gold coordinates.** Submissions differ from the gold in
  coordinate system (raw-text offsets vs. character-grid offsets), in whether the
  leading claim number is present, and in whitespace, so raw `start`/`end` offsets
  are not directly comparable. For every annotation we take its verbatim `text`
  and reported offsets, locate the span in its own non-whitespace character
  sequence, and map it onto the gold's non-whitespace sequence by aligning the two
  with `difflib`. This absorbs coordinate, claim-number, and whitespace
  differences. Evaluating the gold against itself yields F1 = 1.0 on every metric,
  which validates the procedure.
- **Character level is the primary detection metric.** It is robust to boundary
  disagreements, which are common; span-level F1 (exact-ish, IoU≥0.5) is reported
  as a stricter secondary view. In practice systems locate the approximate region
  of an error more reliably than its exact boundaries.
- **Category identity is *not* scored directly.** The gold uses abstract MQM
  categories (accuracy, fluency, terminology) while submissions use finer labels
  (omission, article, dependency, …). Scoring label identity across different
  inventories would be meaningless, so we report location-only recall per gold
  category plus a confusion mapping, which also reveals how the two inventories
  correspond.
- **Reference quality score.** Because the test set has no human score, the
  reference is derived from the gold error annotations with the GEMBA-MQM formula
  of `mqm_score.py` (`score = 100 + 4·max(−25, −penalty)`). It is therefore a
  construct that depends on the severity weights and the `--empty`/clip settings;
  these are exposed as options and should be reported alongside any result.
- **Why correlation, not absolute error, is the headline for scoring.** Pearson
  and Spearman are invariant to a constant offset/scale, so they measure whether a
  system ranks translations by quality consistently with the gold. Submitted
  scores are typically far more lenient than the reference MQM score (large MAE),
  so absolute agreement and rank agreement are reported separately.
- **Input-side signal is weak.** In the gold data almost all errors are on the
  output (translation) side, so input-side detection figures rest on very few gold
  annotations and should be read with care; the combined and output-side figures
  are the informative ones.

## Output JSON (per team, under `--outdir`)

```jsonc
{
  "direction": "ja-en",
  "team": "EHIME-U-GPT6",
  "files": 10,
  "character_level": { "input": {...}, "output": {...}, "all": {"precision":...,"recall":...,"f1":...,"tp":...,"fp":...,"fn":...} },
  "span_level":      { "input": {...}, "output": {...}, "all": {...} },
  "severity_agreement_on_overlap": { "overlap_chars": 4726, "match_rate": 0.9124 },
  "per_gold_category_recall": { "正確性": {"recall":0.40,"gold_chars":1062,"detected":427}, ... },
  "confusion_gold_to_sub":   { "正確性": {"誤訳":190, "誤訳:係り先":157, ...}, ... },
  "score_correlation": { "n":10, "pearson":0.6487, "pearson_p":..., "spearman":0.6333, "spearman_p":..., "mae":42.1, "gold_mean":34.4, "sub_mean":76.5 }
}
```
`summary.json` collects one such object per `"<direction>/<team>"` key.

## Caveat

The test set is small (10 JA–EN and 12 EN–JA claims), so detection scores and
especially correlations are indicative rather than definitive. Report the MQM
scoring settings used, and prefer Spearman (and a 中-weight sensitivity check) for
the score comparison.
