#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
mqm_score.py -- Compute a GEMBA-MQM-style normalized quality score for each
annotated patent claim.

It reads the JSON produced by ``xlsx2json.py`` (one or more files, or a
directory) and, for every block (claim), turns the error annotations into a
single 0-100 quality score following the GEMBA-MQM convention
(Kocmi and Federmann, 2023):

    penalty = sum of per-error weights
    s       = -penalty          (optionally clipped at -25)
    score   = 100 + 4 * s       (so a clipped score lies in [0, 100])

Severity weights are configurable. By default they are:

    small (小) = 1,  medium (中) = 3,  large (大) = 5

A severity range such as ``中～大`` is given the midpoint of the two endpoint
weights. Spans with no severity label (``""``) are, by default, counted as a
minor error (weight = small); this is configurable with ``--empty``.

With ``--no-clip`` the raw ``100 + 4 * s`` is reported without the -25 floor,
so heavily penalized claims can go negative (useful for inspection).

See ``README_mqm_score.md`` for the design rationale and the exact mapping of
the task's three-level severities onto the two-level MQM scheme.

Usage:
    python mqm_score.py <file.json | dir> [-o out.json] [options]
"""
import argparse
import glob
import json
import os
import re
import sys
from collections import Counter

# ---- severity weighting ----------------------------------------------------

SEV_SPLIT = re.compile(r"[～〜~\-]")  # range separators: 中～大, 小-中, ...


def severity_weight(sev, w, empty_policy):
    """Return the weight of a severity string, or None if it should be ignored."""
    sev = (sev or "").strip()
    if sev == "":
        if empty_policy == "ignore":
            return None
        if empty_policy == "medium":
            return w["中"]
        if empty_policy == "major":
            return w["大"]
        return w["小"]  # default: minor
    if sev in w:
        return w[sev]
    # range like "中～大": average the endpoints we recognize
    parts = [p.strip() for p in SEV_SPLIT.split(sev) if p.strip()]
    vals = [w[p] for p in parts if p in w]
    if vals:
        return sum(vals) / len(vals)
    return None  # unknown severity -> ignored


def block_penalty(annotations, w, empty_policy, sides):
    total = 0.0
    for a in annotations:
        if sides != "both" and a.get("side") != sides:
            continue
        wt = severity_weight(a.get("severity", ""), w, empty_policy)
        if wt is not None:
            total += wt
    return total


def to_score(penalty, clip=True):
    s = -penalty
    if clip:
        s = max(-25.0, s)
    return 100 + 4 * s


# ---- IO --------------------------------------------------------------------

def iter_json(target):
    if os.path.isdir(target):
        for f in sorted(glob.glob(os.path.join(target, "*.json"))):
            yield f
    else:
        yield target


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="Compute GEMBA-MQM-style quality scores from xlsx2json output.")
    ap.add_argument("target", help="A .json file produced by xlsx2json.py, or a directory of them.")
    ap.add_argument("-o", "--output", help="Write the per-block results as a JSON array.")
    ap.add_argument("--weight-small", type=float, default=1.0, help="Weight of a 小 (minor) error (default 1).")
    ap.add_argument("--weight-medium", type=float, default=3.0, help="Weight of a 中 (medium) error (default 3).")
    ap.add_argument("--weight-large", type=float, default=5.0, help="Weight of a 大 (major) error (default 5).")
    ap.add_argument("--empty", choices=["minor", "medium", "major", "ignore"], default="minor",
                    help="How to weight spans with no severity label (default minor).")
    ap.add_argument("--side", choices=["both", "input", "output"], default="both",
                    help="Which side's annotations to count (default both).")
    ap.add_argument("--no-clip", action="store_true", help="Do not clip the raw score at the -25 floor.")
    ap.add_argument("--compact", action="store_true", help="Single-line JSON output.")
    ap.add_argument("-q", "--quiet", action="store_true")
    args = ap.parse_args(argv)

    w = {"小": args.weight_small, "中": args.weight_medium, "大": args.weight_large}
    clip = not args.no_clip

    rows = []
    for f in iter_json(args.target):
        try:
            doc = json.load(open(f, encoding="utf-8"))
        except Exception as e:
            if not args.quiet:
                print(f"[skip] {f}: {e}", file=sys.stderr)
            continue
        for b in doc.get("blocks", []):
            anns = b.get("annotations", [])
            pen = block_penalty(anns, w, args.empty, args.side)
            counts = Counter((a.get("severity", "") or "(none)") for a in anns
                             if args.side == "both" or a.get("side") == args.side)
            rows.append({
                "file": doc.get("file", os.path.basename(f)),
                "block": b.get("index"),
                "n_annotations": sum(counts.values()),
                "penalty": round(pen, 3),
                "score": round(to_score(pen, clip), 1),
                "severity_counts": dict(counts),
            })

    if args.output:
        with open(args.output, "w", encoding="utf-8") as fp:
            json.dump(rows, fp, ensure_ascii=False, indent=None if args.compact else 2)

    if not args.quiet:
        print(f"{'file':28} {'blk':>3} {'#ann':>5} {'penalty':>8} {'score':>6}")
        for r in rows:
            print(f"{str(r['file'])[:28]:28} {str(r['block']):>3} {r['n_annotations']:>5} "
                  f"{r['penalty']:>8} {r['score']:>6}")
        if rows:
            mean = sum(r["score"] for r in rows) / len(rows)
            print(f"{'mean':28} {'':>3} {'':>5} {'':>8} {round(mean,1):>6}")
        if args.output:
            print(f"wrote {args.output}", file=sys.stderr)
    elif not args.output:
        print(json.dumps(rows, ensure_ascii=False, indent=None if args.compact else 2))


if __name__ == "__main__":
    main()
