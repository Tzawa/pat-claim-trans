#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
evaluate.py -- Evaluate participant submissions against the gold reference
annotations for the Patent Claims Evaluation Task.

Given a directory of gold JSON files and a directory of submissions (one
sub-directory per team), both organized by direction, this tool reports, per
team and direction:

  * Error detection  : character-level precision / recall / F1 (input / output /
                       combined) and span-level F1 (IoU >= 0.5).
  * Severity         : agreement rate over characters marked by both sides.
  * Category         : recall per gold category (location only) and a
                       gold->submission category confusion mapping.
  * Quality score    : Pearson r and Spearman rho between a GEMBA-MQM reference
                       score derived from the gold annotations (via mqm_score)
                       and the team's submitted score, plus MAE.

Submissions may differ from the gold in coordinate system (raw-text vs.
character-grid offsets), claim-number handling, and whitespace. Every annotation
is therefore RE-ANCHORED to a non-whitespace character position in the gold text
by aligning the two non-whitespace character sequences with difflib. (Evaluating
the gold against itself yields F1 = 1.0, confirming the re-anchoring.)

All inputs are the JSON produced by xlsx2json.py.

Expected layout:
    <gold_dir>/<direction>/<name>.json
    <submissions_dir>/<direction>/<team>/<name>.json
where <direction> is e.g. "ja-en" / "en-ja".

Usage:
    python evaluate.py <gold_dir> <submissions_dir> [--outdir DIR] [options]

Correlation p-values are reported when SciPy is installed; otherwise the
coefficients are still computed (pure-Python fallback).
"""
import argparse
import difflib
import glob
import json
import math
import os
import re
from collections import Counter, defaultdict

import mqm_score  # local module (same tools/ directory)

try:
    from scipy import stats as _scipy_stats
except Exception:
    _scipy_stats = None

DIRECTIONS = ("ja-en", "en-ja")
SIDES = ("input", "output")


# ---- re-anchoring ----------------------------------------------------------

def nws(s):
    return re.sub(r"\s+", "", s or "")


def nws_prefix(ref, idx):
    return sum(1 for ch in ref[:idx] if not ch.isspace())


def choose_ref(a, raw, grid):
    """Pick the reference string (grid preferred, then raw) whose [start:end] == text."""
    st, en = a.get("start"), a.get("end")
    if not isinstance(st, int) or not isinstance(en, int):
        return None
    for ref in ((grid,) if grid else ()) + ((raw,) if raw is not None else ()):
        if 0 <= st <= en <= len(ref) and ref[st:en] == a.get("text", ""):
            return ref
    return None


def build_map(sub_nws, gold_nws):
    """Map sub non-whitespace index -> gold non-whitespace index via difflib."""
    sm = difflib.SequenceMatcher(None, sub_nws, gold_nws, autojunk=False)
    m = [None] * (len(sub_nws) + 1)
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal":
            for k in range(i2 - i1):
                m[i1 + k] = j1 + k
        else:
            for k in range(i1, i2):
                m[k] = j1
    m[len(sub_nws)] = len(gold_nws)
    last = 0
    for k in range(len(m)):
        if m[k] is None:
            m[k] = last
        else:
            last = m[k]
    return m


# ---- category normalization ------------------------------------------------

_CO = [("訳抜け", "訳抜け"), ("未翻訳", "訳抜け"), ("湧き出し", "湧き出し"),
       ("冠詞", "誤訳:冠詞"), ("数値", "誤訳:数値記号"), ("記号", "誤訳:数値記号"),
       ("係り先", "誤訳:係り先"), ("スペリング", "誤訳:綴り"),
       ("誤訳", "誤訳"), ("正確性", "正確性"),
       ("文脈上不適切", "文脈/訳語選択"), ("訳語選択", "文脈/訳語選択"),
       ("用語", "用語一貫性"), ("一貫性", "用語一貫性"),
       ("流暢", "流暢性"), ("ぎこちな", "流暢性"), ("文法", "文法"),
       ("句読点", "句読点"), ("解釈", "誤訳:曖昧")]


def coarse(err):
    for k, name in _CO:
        if k in (err or ""):
            return name
    return (err or "(none)")


# ---- metric helpers --------------------------------------------------------

def prf(tp, fp, fn):
    p = tp / (tp + fp) if (tp + fp) else 0.0
    r = tp / (tp + fn) if (tp + fn) else 0.0
    f = 2 * p * r / (p + r) if (p + r) else 0.0
    return {"precision": round(p, 4), "recall": round(r, 4), "f1": round(f, 4),
            "tp": tp, "fp": fp, "fn": fn}


def iou(a, b):
    lo = max(a[0], b[0]); hi = min(a[1], b[1]); inter = max(0, hi - lo)
    uni = (a[1] - a[0]) + (b[1] - b[0]) - inter
    return inter / uni if uni else 0.0


def span_gold(a, raw, grid, mapper):
    """Return the gold-coordinate [start, end) of annotation a, or None."""
    ref = choose_ref(a, raw, grid)
    text = a.get("text", "")
    if ref is None:
        base = nws(raw); t = nws(text)
        if not t:
            return None
        pos = base.find(t)
        if pos < 0:
            return None
        s, e = pos, pos + len(t)
    else:
        s = nws_prefix(ref, a["start"]); e = s + len(nws(text))
    gs = mapper[min(s, len(mapper) - 1)]
    ge = mapper[min(e, len(mapper) - 1)]
    return (gs, max(ge, gs))


def side_spans(block, side, raw, grid, mapper):
    out = []
    for a in block["annotations"]:
        if a["side"] != side:
            continue
        sp = span_gold(a, raw, grid, mapper)
        if sp is not None:
            out.append((sp[0], sp[1], coarse(a.get("error", "")), a.get("severity", "") or ""))
    return out


def masks(spans, n):
    cov = [False] * n; pri = [None] * n; sev = [None] * n
    for s, e, c, v in sorted(spans, key=lambda x: -(x[1] - x[0])):
        for i in range(max(0, s), min(n, e)):
            cov[i] = True; pri[i] = c; sev[i] = v
    return cov, pri, sev


# ---- correlations (pure-Python fallback) -----------------------------------

def _pearson(x, y):
    n = len(x); mx = sum(x) / n; my = sum(y) / n
    num = sum((a - mx) * (b - my) for a, b in zip(x, y))
    den = math.sqrt(sum((a - mx) ** 2 for a in x) * sum((b - my) ** 2 for b in y))
    return num / den if den else float("nan")


def _ranks(v):
    order = sorted(range(len(v)), key=lambda i: v[i])
    r = [0.0] * len(v); i = 0
    while i < len(v):
        j = i
        while j + 1 < len(v) and v[order[j + 1]] == v[order[i]]:
            j += 1
        avg = (i + j) / 2.0 + 1
        for k in range(i, j + 1):
            r[order[k]] = avg
        i = j + 1
    return r


def correlations(gold, sub):
    out = {"n": len(gold)}
    if _scipy_stats is not None and len(gold) >= 2:
        pr = _scipy_stats.pearsonr(gold, sub); sp = _scipy_stats.spearmanr(gold, sub)
        out.update(pearson=round(float(pr.statistic), 4), pearson_p=round(float(pr.pvalue), 4),
                   spearman=round(float(sp.statistic), 4), spearman_p=round(float(sp.pvalue), 4))
    else:
        out.update(pearson=round(_pearson(gold, sub), 4),
                   spearman=round(_pearson(_ranks(gold), _ranks(sub)), 4),
                   pearson_p=None, spearman_p=None)
    n = len(gold)
    out["mae"] = round(sum(abs(a - b) for a, b in zip(gold, sub)) / n, 2) if n else None
    out["gold_mean"] = round(sum(gold) / n, 2) if n else None
    out["sub_mean"] = round(sum(sub) / n, 2) if n else None
    return out


# ---- per-team evaluation ---------------------------------------------------

def evaluate_team(gold_files, sub_dir, w, empty, clip, iou_thr):
    char = {s: Counter() for s in SIDES}
    span = {s: Counter() for s in SIDES}
    sev = Counter(); conf = Counter(); gcat = defaultdict(Counter)
    gold_scores = []; sub_scores = []
    nfiles = 0
    for bn, gf in sorted(gold_files.items()):
        sf = os.path.join(sub_dir, bn)
        if not os.path.exists(sf):
            continue
        nfiles += 1
        gdoc = json.load(open(gf, encoding="utf-8"))
        sdoc = json.load(open(sf, encoding="utf-8"))
        for gb, sb in zip(gdoc["blocks"], sdoc["blocks"]):
            # quality score correlation
            gp = mqm_score.block_penalty(gb["annotations"], w, empty, "both")
            gs = mqm_score.to_score(gp, clip)
            ss = sb.get("score")
            if isinstance(ss, (int, float)):
                gold_scores.append(gs); sub_scores.append(ss)
            # detection
            raw = {"input": (gb.get("input_text"), sb.get("input_text")),
                   "output": (gb.get("output_text"), sb.get("output_text"))}
            grid = {"input": (gb.get("input_text_grid"), sb.get("input_text_grid")),
                    "output": (gb.get("output_text_grid"), sb.get("output_text_grid"))}
            for side in SIDES:
                gnwss = nws(raw[side][0])
                mp = build_map(nws(raw[side][1]), gnwss)
                n = len(gnwss)
                gsp = side_spans(gb, side, raw[side][0], grid[side][0], list(range(n + 1)))
                ssp = side_spans(sb, side, raw[side][1], grid[side][1], mp)
                gc, gpri, gsev = masks(gsp, n)
                sc, spri, ssev = masks(ssp, n)
                cc = char[side]
                for i in range(n):
                    if gc[i] and sc[i]:
                        cc["tp"] += 1
                        if gsev[i] and ssev[i]:
                            sev["total"] += 1; sev["match"] += 1 if gsev[i] == ssev[i] else 0
                        if gpri[i] is not None and spri[i] is not None:
                            conf[(gpri[i], spri[i])] += 1
                        if gpri[i] is not None:
                            gcat[gpri[i]]["tp"] += 1
                    elif sc[i]:
                        cc["fp"] += 1
                    elif gc[i]:
                        cc["fn"] += 1
                        if gpri[i] is not None:
                            gcat[gpri[i]]["fn"] += 1
                pairs = sorted(((iou((s[0], s[1]), (t[0], t[1])), gi, ti)
                                for gi, s in enumerate(gsp) for ti, t in enumerate(ssp)
                                if iou((s[0], s[1]), (t[0], t[1])) > 0), reverse=True)
                ug = set(); us = set(); tp = 0
                for ov, gi, ti in pairs:
                    if gi in ug or ti in us or ov < iou_thr:
                        continue
                    ug.add(gi); us.add(ti); tp += 1
                span[side]["tp"] += tp; span[side]["fp"] += len(ssp) - tp; span[side]["fn"] += len(gsp) - tp

    cl = {s: prf(char[s]["tp"], char[s]["fp"], char[s]["fn"]) for s in SIDES}
    callc = Counter(); [callc.update(char[s]) for s in SIDES]
    cl["all"] = prf(callc["tp"], callc["fp"], callc["fn"])
    spl = {s: prf(span[s]["tp"], span[s]["fp"], span[s]["fn"]) for s in SIDES}
    sallc = Counter(); [sallc.update(span[s]) for s in SIDES]
    spl["all"] = prf(sallc["tp"], sallc["fp"], sallc["fn"])
    percat = {}
    for c, v in sorted(gcat.items(), key=lambda kv: -(kv[1]["tp"] + kv[1]["fn"])):
        tot = v["tp"] + v["fn"]
        percat[c] = {"recall": round(v["tp"] / tot, 4) if tot else None,
                     "gold_chars": tot, "detected": v["tp"]}
    confm = defaultdict(dict)
    for (cg, cs), nn in conf.items():
        confm[cg][cs] = nn
    confm = {cg: dict(sorted(v.items(), key=lambda kv: -kv[1])) for cg, v in confm.items()}
    return {
        "files": nfiles,
        "character_level": cl,
        "span_level": spl,
        "severity_agreement_on_overlap": {
            "overlap_chars": sev["total"],
            "match_rate": round(sev["match"] / sev["total"], 4) if sev["total"] else None},
        "per_gold_category_recall": percat,
        "confusion_gold_to_sub": confm,
        "score_correlation": correlations(gold_scores, sub_scores) if gold_scores else None,
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description="Evaluate submissions against gold reference annotations.")
    ap.add_argument("gold_dir")
    ap.add_argument("submissions_dir")
    ap.add_argument("--outdir", help="Write per-team JSON and summary.json here.")
    ap.add_argument("--iou", type=float, default=0.5, help="IoU threshold for span matching (default 0.5).")
    ap.add_argument("--weight-small", type=float, default=1.0)
    ap.add_argument("--weight-medium", type=float, default=3.0)
    ap.add_argument("--weight-large", type=float, default=5.0)
    ap.add_argument("--empty", choices=["minor", "medium", "major", "ignore"], default="minor",
                    help="Severity for unlabeled gold spans in the reference score (default minor).")
    ap.add_argument("--no-clip", action="store_true", help="Do not clip the reference MQM score at -25.")
    ap.add_argument("-q", "--quiet", action="store_true")
    args = ap.parse_args(argv)

    w = {"小": args.weight_small, "中": args.weight_medium, "大": args.weight_large}
    clip = not args.no_clip
    summary = {}

    for direction in DIRECTIONS:
        gdir = os.path.join(args.gold_dir, direction)
        sdir = os.path.join(args.submissions_dir, direction)
        if not (os.path.isdir(gdir) and os.path.isdir(sdir)):
            continue
        gold_files = {os.path.basename(f): f for f in glob.glob(os.path.join(gdir, "*.json"))}
        teams = sorted(os.path.basename(x) for x in glob.glob(os.path.join(sdir, "*")) if os.path.isdir(x))
        for team in teams:
            res = evaluate_team(gold_files, os.path.join(sdir, team), w, args.empty, clip, args.iou)
            res["direction"] = direction; res["team"] = team
            summary[f"{direction}/{team}"] = res
            if args.outdir:
                os.makedirs(os.path.join(args.outdir, direction), exist_ok=True)
                json.dump(res, open(os.path.join(args.outdir, direction, team + ".json"), "w", encoding="utf-8"),
                          ensure_ascii=False, indent=2)

    if args.outdir:
        json.dump(summary, open(os.path.join(args.outdir, "summary.json"), "w", encoding="utf-8"),
                  ensure_ascii=False, indent=2)

    if not args.quiet:
        hdr = f"{'team':28}{'files':>6}{'detF1':>7}{'detP':>6}{'detR':>6}{'spanF1':>7}{'sevAgr':>7}{'Pear':>6}{'Spear':>6}"
        print(hdr)
        for k, o in summary.items():
            d = o["character_level"]["all"]; sp = o["span_level"]["all"]
            sev = o["severity_agreement_on_overlap"]["match_rate"]
            c = o["score_correlation"] or {}
            print(f"{k:28}{o['files']:>6}{d['f1']:>7}{d['precision']:>6}{d['recall']:>6}"
                  f"{sp['f1']:>7}{str(sev):>7}{str(c.get('pearson','-')):>6}{str(c.get('spearman','-')):>6}")
        if _scipy_stats is None:
            print("(install scipy for correlation p-values)")


if __name__ == "__main__":
    main()
