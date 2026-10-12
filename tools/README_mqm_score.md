# `mqm_score.py` — GEMBA-MQM-style quality score from error annotations

This tool derives a single 0–100 quality score for each annotated patent claim
from its error annotations. It reads the JSON produced by
[`xlsx2json.py`](./README.md) and applies the GEMBA-MQM scoring convention.

It is useful as a **reference quality score** computed consistently from gold (or
any) error annotations — for example, to correlate a participant's own quality
scores against a score implied by the gold error spans.

## Requirements

- Python 3.8+ (standard library only; no third-party packages).

## Usage

```bash
# Score every claim in a directory of xlsx2json outputs (prints a table)
python mqm_score.py ../2026/ref/ja-en

# Score one file and save the per-block results as JSON
python mqm_score.py ../2026/ref/ja-en/01.json -o 01.score.json

# Change the severity weights or disable clipping
python mqm_score.py ../2026/ref/ja-en --weight-medium 5 --no-clip
```

Input is the JSON emitted by `xlsx2json.py` (a file or a directory of them). If
you start from Excel, run `xlsx2json.py` first.

### Options

| Option | Default | Description |
|---|---|---|
| `--weight-small F`  | `1` | Weight of a 小 (minor) error. |
| `--weight-medium F` | `3` | Weight of a 中 (medium) error. |
| `--weight-large F`  | `5` | Weight of a 大 (major) error. |
| `--empty POLICY` | `minor` | How to weight spans with **no** severity label: `minor` / `medium` / `major` / `ignore`. |
| `--side {both,input,output}` | `both` | Which side's annotations to count. |
| `--no-clip` | off | Report the raw `100 + 4s` without the −25 floor (scores may go negative). |
| `-o FILE` | — | Write per-block results as a JSON array. |
| `--compact`, `-q` | — | Single-line JSON / quiet. |

### Output

Per block: `file`, `block`, `n_annotations`, `penalty`, `score`, and a
`severity_counts` breakdown. The console view also prints the mean score.

## Scoring method

For each claim (block), every error contributes a **weight** determined by its
severity. The **penalty** is the sum of these weights, and the score follows the
GEMBA-MQM mapping of Kocmi and Federmann (2023):

```
penalty = Σ weight(error)
s       = -penalty              # optionally clipped so s ≥ -25
score   = 100 + 4 · s           # clipped score ∈ [0, 100]
```

The coefficient `4` is `100 / 25`: with the per-segment floor of `s = -25`, the
score maps the penalty range `[0, 25]` onto `[100, 0]`.

## Design rationale and decisions

The GEMBA-MQM scheme assumes **two** severity levels (major = 5, minor = 1). The
task's gold annotations, however, use **three** levels (小 / 中 / 大) plus
ranges (e.g. 中～大), and some spans carry no severity. The tool resolves this as
follows; each choice is a CLI option so results are reproducible and auditable.

- **Three-level mapping (小=1, 中=3, 大=5).** We keep the GEMBA-MQM anchors
  (minor = 1, major = 5) and place 中 at the midpoint (3) rather than collapsing
  it onto major or minor. This preserves the ordering 小 < 中 < 大 and avoids an
  arbitrary binary cutoff. Use `--weight-medium` to explore the alternatives
  (e.g. `1` to treat 中 as minor, `5` as major); this is the single most
  influential setting, so we recommend reporting a sensitivity check.
- **Severity ranges (中～大, 小～中).** Scored as the **midpoint** of the two
  endpoint weights (e.g. 中～大 → (3+5)/2 = 4).
- **Unlabeled spans (`severity == ""`).** Counted as **minor** by default
  (`--empty minor`). In the gold data these are highlighted error regions whose
  severity was conveyed by color grouping rather than an explicit label, so they
  are genuine errors and should not be dropped; counting them as minor is the
  conservative choice. Use `--empty ignore` to exclude them, or
  `medium`/`major` to weight them more heavily.
- **Per-segment clipping at −25.** On by default, matching GEMBA-MQM and the
  `100 + 4s` coefficient. A consequence is that heavily penalized claims
  saturate at 0; `--no-clip` disables this for inspection.
- **Both sides counted.** By default, errors on both the source (`input`) and
  the translation (`output`) side contribute. Strict target-side MQM would use
  `--side output`; in the task's gold data almost all errors are on the output
  side, so the choice rarely matters, but it is exposed for completeness.
- **Granularity: per claim (block).** The score is computed per block (one
  claim = one segment), which is the unit GEMBA-MQM scores and the unit the task
  evaluates.

## Caveat

The produced score is a **construct** of the formula and the weighting choices
above; it is not a human quality judgment. When comparing it against other
scores, report the settings used (weights, `--empty`, clipping) and, ideally, a
sensitivity analysis over the 中 weight.

## Reference

Tom Kocmi and Christian Federmann. 2023. *GEMBA-MQM: Detecting Translation
Quality Error Spans with GPT-4.* Proceedings of the Eighth Conference on Machine
Translation (WMT).
