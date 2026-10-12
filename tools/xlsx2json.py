#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
翻訳品質アノテーション付き Excel ファイルから、入力テキスト・翻訳テキスト・
広範囲にわたるエラー・修正訳・アノテーション情報を抽出して JSON 化する。

翻訳方向は en-ja（入力=英語 / 出力=日本語）と ja-en（入力=日本語 / 出力=英語）の
両方に対応する。列範囲は結合セルから、折り返し行の連結方法は文字種から自動判定する
ため、方向を指定するオプションは不要。

対象ファイルの構造（1 ブロック分）:

    - ヘッダ行  : A列 = "Input", 出力列の直前 = 翻訳スコア, 出力列 = "Output"
    - ヘッダ+1行: 入力側の結合セル = 入力テキスト, 出力側の結合セル = 翻訳テキスト
    - 文字グリッド行(3行おき): 入力テキストと翻訳テキストが 1 セル 1 文字で並ぶ。
        * 左側 (入力の結合セルと同じ列範囲)   = 入力テキスト
          例) en-ja では A:CV, ja-en では A:AN
        * 右側 (出力の結合セルと同じ列範囲)   = 翻訳テキスト
          例) en-ja では DG:ET, ja-en では AY:ET
        * 空セルはスペースに対応する
        * 入力が翻訳より短い等で片側が尽きると「片側のみの文字行」が続く
    - 文字グリッド行の直下(1〜2行)にアノテーション行がある。
        * 背景が黄色に塗られた連続セルの範囲が「アノテーション対象スパン」。
        * その範囲の先頭セルにアノテーション内容(文字列)が入っている
          (ラベルの無いハイライトのみのスパンもあり、その場合 content は "")。
        * 対象スパンの列位置 → 直上の文字グリッド行の文字位置 に変換して
          「開始文字位置(start)・終了文字位置(end)」を求める。
        * 1 つの連続した黄色範囲の中にラベル文字セルが複数ある場合は、各ラベルを
          起点に別々のアノテーションへ分割する。
        * 逆に、黄色範囲が行末(その側の右端列)まで達し、後続のラベル無し範囲が
          次の表示行の先頭(左端列)から連続している場合は、表示行の折り返しとして
          1 つのアノテーションに結合する(cells が複数要素になる)。
        * 各アノテーションが占める元セル範囲は cells(範囲文字列のリスト)で示す。
          通常は 1 要素、折り返し時のみ複数(例 ["EQ33:ET33","DG37:DI37"])。
    - ブロック末尾: 出力列に "↓広範囲にわたるエラーがある場合は…"(その次行の結合セルが本文)、
      "↓修正訳を入力してください"(その次行の結合セルが本文)。
    - 以上のブロックが 1 ファイルに 1 回以上繰り返される。

出力される各アノテーションの start/end は、そのブロックの
input_text_grid / output_text_grid（文字グリッドから復元した文字列）に対する
0 始まりのオフセットで、end は排他的（text == grid[start:end]）。
CJK 側(日本語)の復元結果は元セルと完全一致する。ラテン側(英語)は元セルと
改行/連続空白の扱いのみが異なる（グリッド上は単一スペース）。

Usage:
    python extract_annotations.py <xlsx ファイル または ディレクトリ> [options]

Options:
    -o, --output FILE   まとめて 1 つの JSON ファイルに書き出す
    --outdir DIR        入力ファイルごとに <name>.json を DIR に書き出す
    --pretty            インデント付きで整形出力（既定で有効）
    --compact           整形せず 1 行で出力
    -q, --quiet         進捗メッセージを表示しない

いずれの出力先も指定しない場合は標準出力に JSON を表示する。
"""

import argparse
import glob
import json
import os
import re
import sys

import openpyxl
from openpyxl.utils import get_column_letter

# アノテーション文字列を「エラー内容」と「重要度(大/小)」に分割する。
# 重要度は "／大" "／小"(全角スラッシュ) または "/大" "/小"(半角) の形で付く。
# 見つからなければ severity は ""。error は重要度マーカーを除いた残り。
_SEVERITY_RE = re.compile(r"[／/]([大小])")


def _split_content(raw):
    if not raw:
        return "", ""
    m = _SEVERITY_RE.search(raw)
    if not m:
        return raw.strip(), ""
    severity = m.group(1)
    error = (raw[:m.start()] + raw[m.end():]).strip()
    return error, severity


# 参照(正解)データの注釈ラベル: 「重要度（カテゴリ）：詳細説明」形式。
#   重要度 = 大/中/小(「中～大」等の範囲も可) / カテゴリ = （）内 / 説明 = ：以降
_REF_LABEL_RE = re.compile(
    r"^\s*([大中小][～\-〜]?[大中小]?)\s*[（(]([^）)]*)[）)]\s*[：:]\s*(.*)$", re.S)


def _parse_ref_label(text):
    if not isinstance(text, str):
        return None
    m = _REF_LABEL_RE.match(text)
    if not m:
        return None
    return m.group(1).strip(), m.group(2).strip(), m.group(3).strip()


# ---- 塗りつぶし色 ----------------------------------------------------------
_COLOR_NAMES = {"FFFFFF00": "yellow", "FFFB8181": "red", "FFFFC000": "orange",
                "FF00B050": "green", "FF0070C0": "blue", "FF00B0F0": "blue",
                "FFFFFF99": "yellow"}


def _fill_color_key(cell):
    """塗りつぶし色の識別キー(rgb 文字列 / theme 番号)。無地なら None。"""
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return None
    fg = fill.fgColor
    rgb = getattr(fg, "rgb", None)
    if isinstance(rgb, str) and rgb not in ("00000000",):
        return rgb
    th = getattr(fg, "theme", None)
    if isinstance(th, int):
        return f"theme{th}"
    return "solid"


def _color_name(key):
    if key is None:
        return ""
    return _COLOR_NAMES.get(key, key)


# ---- 黄色判定 --------------------------------------------------------------

def _is_yellow(cell):
    """セル背景が黄色(FFFF00 系)かどうか。"""
    fill = cell.fill
    if fill is None or fill.patternType != "solid":
        return False
    for color in (fill.fgColor, fill.start_color):
        rgb = getattr(color, "rgb", None)
        if isinstance(rgb, str) and rgb.upper().endswith("FFFF00"):
            return True
    return False


# ---- グリッド文字列の復元 --------------------------------------------------

def _segment(ws, row, c1, c2):
    """行 row の 列 c1..c2 を 1 セル 1 文字として連結。空セルはスペース。
    末尾の空セルは切り詰める。"""
    last = c1 - 1
    for c in range(c1, c2 + 1):
        if ws.cell(row=row, column=c).value not in (None, ""):
            last = c
    out = []
    for c in range(c1, last + 1):
        v = ws.cell(row=row, column=c).value
        out.append(str(v) if v not in (None, "") else " ")
    return "".join(out)


def _count_singlechar(ws, row, c1, c2):
    n = 0
    for c in range(c1, c2 + 1):
        v = ws.cell(row=row, column=c).value
        if v is not None and len(str(v)) == 1:
            n += 1
    return n


def _join_sep(text):
    """その言語の折り返し行を連結するときのセパレータを返す。
    日本語などの CJK は連結なし("")、英語などのラテン系は単一スペース(" ")。
    翻訳方向 (en-ja / ja-en) に依存せず、テキストの文字種で自動判定する。"""
    if not text:
        return ""
    cjk = latin = 0
    for ch in text:
        if ("　" <= ch <= "鿿") or ("＀" <= ch <= "￯"):
            cjk += 1
        elif ch.isascii() and ch.isalpha():
            latin += 1
    return "" if cjk >= latin else " "


# ---- 列範囲(入力/出力)を結合セルから取得 -----------------------------------

def _col_ranges(ws, text_row):
    """text_row(ヘッダ+1 行)にある結合セルから、入力側・出力側の列範囲を求める。
    戻り値: (in1, in2, out1, out2)。見つからない場合は None。"""
    ranges = [m for m in ws.merged_cells.ranges if m.min_row == text_row]
    in_r = out_r = None
    for m in sorted(ranges, key=lambda x: x.min_col):
        if m.min_col == 1:
            in_r = (m.min_col, m.max_col)
        elif out_r is None and in_r is not None and m.min_col > in_r[1]:
            out_r = (m.min_col, m.max_col)
    if in_r is None or out_r is None:
        return None
    return in_r[0], in_r[1], out_r[0], out_r[1]


# ---- 1 ブロックの抽出 ------------------------------------------------------

def _extract_block(ws, header_row, block_end, index, highlight="auto"):
    text_row = header_row + 1

    ranges = _col_ranges(ws, text_row)
    if ranges is None:
        return None
    IN1, IN2, OUT1, OUT2 = ranges

    input_text = ws.cell(row=text_row, column=IN1).value
    output_text = ws.cell(row=text_row, column=OUT1).value
    # スコアはヘッダ行の "Output"(DG) の 1 つ左 = DF。実際には Input と Output の
    # 間の列に入っているので、ヘッダ行で数値が入っているセルを採用する。
    score = None
    for c in range(IN2 + 1, OUT1):
        v = ws.cell(row=header_row, column=c).value
        if isinstance(v, (int, float)):
            score = v
            break

    # 文字グリッド行: 入力/出力の列範囲全体で 1 文字セルが十分に多い行。
    # 英語入力より日本語訳が長い場合、英語が尽きた後ろに「日本語のみ」の
    # 文字行が続くため、A 列だけでなく全範囲で判定する。
    # アノテーション行はラベル(複数文字)しか持たないので 1 文字セルは実質 0。
    char_rows = []
    for r in range(text_row + 1, block_end + 1):
        if _count_singlechar(ws, r, IN1, OUT2) >= 2:
            char_rows.append(r)

    # グリッドから入力/出力テキストを復元し、各行の開始オフセットを記録する。
    # 折り返し行の連結セパレータは言語(文字種)から自動判定する:
    #   英語などラテン系 -> 単一スペース / 日本語など CJK -> 連結なし。
    # これにより翻訳方向 (en-ja / ja-en) の両方に対応する。
    # 空行(片側だけ文字がある行)は各側で独立にスキップする。
    in_sep = _join_sep(input_text)
    out_sep = _join_sep(output_text)

    in_text = ""
    in_start = {}
    for r in char_rows:
        s = _segment(ws, r, IN1, IN2)
        if s == "":
            continue
        if in_text:
            in_text += in_sep
        in_start[r] = len(in_text)
        in_text += s

    out_text = ""
    out_start = {}
    for r in char_rows:
        s = _segment(ws, r, OUT1, OUT2)
        if s == "":
            continue
        if out_text:
            out_text += out_sep
        out_start[r] = len(out_text)
        out_text += s

    def nearest_row(starts, r):
        prev = [cr for cr in starts if cr <= r]
        return prev[-1] if prev else None

    # 1 つの黄色スパン(列 sc1..sc2, 行 r)を 1 アノテーションに変換する。
    # content を渡すと「エラー内容 / 重要度」に分割する。error / severity を直接
    # 指定した場合はそれを使う(参照データの label 形式用)。
    def _make(sc1, sc2, content, r, error=None, severity=None,
              explanation="", color=""):
        if sc1 <= IN2:              # 入力側
            side = "input"
            cr = nearest_row(in_start, r)
            if cr is None:
                return None
            base = in_start[cr]
            start = base + (sc1 - IN1)
            end = base + (sc2 - IN1) + 1
            text = in_text[start:end]
        else:                       # 出力側
            side = "output"
            cr = nearest_row(out_start, r)
            if cr is None:
                return None
            base = out_start[cr]
            start = base + (sc1 - OUT1)
            end = base + (sc2 - OUT1) + 1
            text = out_text[start:end]
        if error is None:
            error, severity = _split_content(content)
        return {
            "side": side,
            "start": start,
            "end": end,
            "text": text,
            # content を「エラー内容(error)」と「重要度(severity: 大/中/小/"")」に分割。
            "error": error,
            "severity": severity,
            # 詳細説明(参照データの label の「：」以降)。無ければ ""。
            "explanation": explanation,
            # ハイライト色(参照データは複数色を使う)。無ければ ""。
            "color": color,
            # 折り返し(L 字型)を表せるよう cell 範囲はリストで保持する。
            # 通常は 1 要素、行をまたぐ場合のみ複数になる。
            "cells": [f"{get_column_letter(sc1)}{r}:{get_column_letter(sc2)}{r}"],
            "_c1": sc1,
            "_c2": sc2,
        }

    # ハイライトの方式を決める。
    #   rows : ラベル付きの黄色が文字グリッドの「下の別行」に置かれる方式
    #          (WAT / checked データ。文字行の黄色は無視する)
    #   ref  : 誤り箇所を色ハイライトで文字行に塗り、直下の行に「重要度（カテゴリ）：
    #          説明」形式のラベルを、ハイライト開始列に揃えて置く方式(参照(正解)データ)
    #   grid : 色ハイライトが文字行に直接あるがラベルが無い方式(スパンのみ抽出)
    # auto : 参照ラベルがあれば ref、別行にラベル付き黄色があれば rows、文字行にのみ
    #        ハイライトがあれば grid。既存データは rows のままで挙動が変わらない。
    char_set = set(char_rows)

    def _row_has_yellow(r):
        return any(_is_yellow(ws.cell(row=r, column=c)) for c in range(IN1, OUT2 + 1))

    def _row_has_fill(r):
        return any(_fill_color_key(ws.cell(row=r, column=c)) for c in range(IN1, OUT2 + 1))

    mode = highlight
    if highlight == "auto":
        has_ref_label = any(
            _parse_ref_label(ws.cell(row=r, column=c).value)
            for r in range(text_row + 1, block_end + 1) if r not in char_set
            for c in range(IN1, OUT2 + 1))
        nonchar_yellow = any(_row_has_yellow(r) for r in range(text_row + 1, block_end + 1)
                             if r not in char_set)
        if has_ref_label:
            mode = "ref"
        elif nonchar_yellow:
            mode = "rows"
        elif any(_row_has_fill(r) for r in char_rows):
            mode = "grid"
        else:
            mode = "rows"

    annotations = []
    if mode == "ref":
        # 参照(正解)データ: 色ハイライト(文字行)＋直下のラベル「重要度（カテゴリ）：説明」。
        # ラベルは (直上の文字行, 同じ列) のハイライト開始位置に対応づく。
        labelmap = {}
        for r in range(text_row + 1, block_end + 1):
            if r in char_set:
                continue
            above = [cr for cr in char_rows if cr < r]
            if not above:
                continue
            cr = above[-1]
            for c in range(IN1, OUT2 + 1):
                p = _parse_ref_label(ws.cell(row=r, column=c).value)
                if p:
                    labelmap[(cr, c)] = p
        for r in char_rows:
            c = IN1
            while c <= OUT2:
                key = _fill_color_key(ws.cell(row=r, column=c))
                if key is not None:
                    c1 = c
                    while c <= OUT2 and _fill_color_key(ws.cell(row=r, column=c)) == key:
                        c += 1
                    c2 = c - 1
                    # このハイライト範囲に含まれるラベル(列)を集める。
                    lcols = sorted(lc for (rr, lc) in labelmap if rr == r and c1 <= lc <= c2)
                    color = _color_name(key)
                    if not lcols:
                        a = _make(c1, c2, "", r, error="", severity="", color=color)
                        if a:
                            annotations.append(a)
                    else:
                        for i, lc in enumerate(lcols):
                            sc1 = c1 if i == 0 else lc
                            sc2 = (lcols[i + 1] - 1) if i + 1 < len(lcols) else c2
                            sev, cat, expl = labelmap[(r, lc)]
                            a = _make(sc1, sc2, "", r, error=cat, severity=sev,
                                      explanation=expl, color=color)
                            if a:
                                annotations.append(a)
                else:
                    c += 1
    elif mode == "grid":
        # 文字グリッド行に直接塗られたハイライトの連続範囲を、ラベル無しスパンとして抽出。
        for r in char_rows:
            c = IN1
            while c <= OUT2:
                key = _fill_color_key(ws.cell(row=r, column=c))
                if key is not None:
                    c1 = c
                    while c <= OUT2 and _fill_color_key(ws.cell(row=r, column=c)) == key:
                        c += 1
                    a = _make(c1, c - 1, "", r, color=_color_name(key))
                    if a:
                        annotations.append(a)
                else:
                    c += 1
    else:
        # 別行にあるラベル付き黄色の抽出。連続した黄色範囲の中にラベル文字セルが
        # 複数ある場合は、各ラベルセルを起点に別々のアノテーションへ分割する。各サブ
        # スパンはそのラベルセルの列から次のラベルセルの直前まで(先頭ラベルより前の
        # 空セルは先頭サブスパンに含める)。ラベル文字が無い黄色範囲は content 空で出力。
        for r in range(text_row + 1, block_end + 1):
            if r in char_set:
                continue
            c = IN1
            while c <= OUT2:
                if _is_yellow(ws.cell(row=r, column=c)):
                    c1 = c
                    while c <= OUT2 and _is_yellow(ws.cell(row=r, column=c)):
                        c += 1
                    c2 = c - 1
                    label_cols = [cc for cc in range(c1, c2 + 1)
                                  if ws.cell(row=r, column=cc).value not in (None, "")]
                    if not label_cols:
                        a = _make(c1, c2, "", r)
                        if a:
                            annotations.append(a)
                    else:
                        for i, lc in enumerate(label_cols):
                            sc1 = c1 if i == 0 else lc
                            sc2 = (label_cols[i + 1] - 1) if i + 1 < len(label_cols) else c2
                            a = _make(sc1, sc2, str(ws.cell(row=r, column=lc).value), r)
                            if a:
                                annotations.append(a)
                else:
                    c += 1

    # 行をまたいで連続する黄色スパンの結合。
    # あるスパンがその側の右端列(IN2 / OUT2)まで達し、後続の(ラベルの無い)スパンが
    # その側の左端列(IN1 / OUT1)から始まり、復元テキスト上で連続(前.end == 後.start)
    # している場合、それは表示行の折り返しであり 1 つのアノテーションとして結合する。
    side_range = {"input": (IN1, IN2), "output": (OUT1, OUT2)}
    merged = []
    for a in annotations:
        mn, mx = side_range[a["side"]]
        if a["error"] == "" and a["severity"] == "" and a["_c1"] == mn:
            # 折り返しの継続スパンは「次の表示行の左端(mn)から始まり(a._c1==mn)」、
            # 復元テキスト上で直前スパンの末尾に連結セパレータ分だけ空けて続く。
            # CJK 側はセパレータ無し(0)、ラテン側は単一スペース(1)。英語は語境界で
            # 折り返すため直前スパンが行末列まで届かない点に注意(右端条件は課さない)。
            sep_len = len(in_sep) if a["side"] == "input" else len(out_sep)
            cont = None
            for p in reversed(merged):
                if p["side"] == a["side"] and a["start"] - p["end"] == sep_len:
                    cont = p
                    break
            if cont is not None:
                grid = in_text if a["side"] == "input" else out_text
                cont["end"] = a["end"]
                cont["text"] = grid[cont["start"]:cont["end"]]
                cont["cells"] += a["cells"]
                cont["_c2"] = a["_c2"]
                continue
        merged.append(a)
    annotations = merged
    for a in annotations:
        a.pop("_c1", None)
        a.pop("_c2", None)

    annotations.sort(key=lambda a: (a["side"], a["start"]))

    # 広範囲にわたるエラー / 修正訳
    broad_error = None
    correction = None
    for r in range(text_row + 1, block_end + 1):
        v = ws.cell(row=r, column=OUT1).value
        if isinstance(v, str) and "広範囲にわたるエラー" in v:
            broad_error = ws.cell(row=r + 1, column=OUT1).value
        elif isinstance(v, str) and "修正訳" in v and "入力" in v:
            correction = ws.cell(row=r + 1, column=OUT1).value

    return {
        "index": index,
        "score": score,
        "input_text": input_text,
        "output_text": output_text,
        "broad_error": broad_error if broad_error not in (None, "") else "",
        "correction": correction if correction not in (None, "") else "",
        "annotations": annotations,
        # start/end が指す復元文字列（参照用）。output は元の Output と一致する。
        "input_text_grid": in_text,
        "output_text_grid": out_text,
    }


# ---- 1 ファイルの抽出 ------------------------------------------------------

def _detect_mode(ws):
    """ファイル全体でハイライト方式を判定する(ブロックごとに揺れないよう file 単位)。
    優先: ref(重要度（カテゴリ）：説明ラベル有) > rows(非文字行に黄色) > grid(文字行に色) > rows。"""
    headers = [r for r in range(1, ws.max_row + 1)
               if str(ws.cell(r, 1).value).strip() == "Input"]
    if not headers:
        return "rows"
    rng = _col_ranges(ws, headers[0] + 1)
    if not rng:
        return "rows"
    IN1, IN2, OUT1, OUT2 = rng
    charset = set(r for r in range(1, ws.max_row + 1)
                  if _count_singlechar(ws, r, IN1, OUT2) >= 2)
    for r in range(1, ws.max_row + 1):
        if r in charset:
            continue
        for c in range(IN1, OUT2 + 1):
            if _parse_ref_label(ws.cell(row=r, column=c).value):
                return "ref"
    for r in range(1, ws.max_row + 1):
        if r in charset:
            continue
        if any(_is_yellow(ws.cell(row=r, column=c)) for c in range(IN1, OUT2 + 1)):
            return "rows"
    for r in charset:
        if any(_fill_color_key(ws.cell(row=r, column=c)) for c in range(IN1, OUT2 + 1)):
            return "grid"
    return "rows"


def extract_file(path, highlight="auto"):
    ws = openpyxl.load_workbook(path, data_only=True).active

    mode = _detect_mode(ws) if highlight == "auto" else highlight

    # ヘッダ行 = A 列が "Input"
    headers = [r for r in range(1, ws.max_row + 1)
               if str(ws.cell(r, 1).value).strip() == "Input"]

    blocks = []
    for i, h in enumerate(headers):
        block_end = (headers[i + 1] - 1) if i + 1 < len(headers) else ws.max_row
        blk = _extract_block(ws, h, block_end, i + 1, highlight=mode)
        if blk is not None:
            blocks.append(blk)

    return {
        "file": os.path.basename(path),
        "blocks": blocks,
    }


# ---- CLI -------------------------------------------------------------------

def _iter_inputs(target):
    if os.path.isdir(target):
        for f in sorted(glob.glob(os.path.join(target, "*.xlsx"))):
            if not os.path.basename(f).startswith("~$"):
                yield f
    else:
        yield target


def main(argv=None):
    ap = argparse.ArgumentParser(
        description="翻訳アノテーション付き Excel を JSON 抽出する")
    ap.add_argument("target", help="xlsx ファイル または ディレクトリ")
    ap.add_argument("-o", "--output", help="まとめて書き出す JSON ファイル")
    ap.add_argument("--outdir", help="ファイルごとに <name>.json を書き出すディレクトリ")
    ap.add_argument("--compact", action="store_true", help="整形せず 1 行で出力")
    ap.add_argument("-q", "--quiet", action="store_true", help="進捗を表示しない")
    ap.add_argument("--highlight", choices=["auto", "rows", "grid"], default="auto",
                    help="黄色ハイライトの方式。auto(既定)/rows(別行ラベル)/grid(文字行直接)")
    args = ap.parse_args(argv)

    indent = None if args.compact else 2
    results = []
    for f in _iter_inputs(args.target):
        try:
            res = extract_file(f, highlight=args.highlight)
        except Exception as e:  # 壊れたファイル等はスキップ
            if not args.quiet:
                print(f"[skip] {f}: {e}", file=sys.stderr)
            continue
        results.append(res)
        if not args.quiet:
            n = sum(len(b["annotations"]) for b in res["blocks"])
            print(f"[ok] {res['file']}: {len(res['blocks'])} blocks, "
                  f"{n} annotations", file=sys.stderr)
        if args.outdir:
            os.makedirs(args.outdir, exist_ok=True)
            stem = os.path.splitext(res["file"])[0]
            with open(os.path.join(args.outdir, stem + ".json"), "w",
                      encoding="utf-8") as fp:
                json.dump(res, fp, ensure_ascii=False, indent=indent)

    if args.outdir:
        return

    payload = results[0] if (len(results) == 1 and not os.path.isdir(args.target)) \
        else {"files": results}
    text = json.dumps(payload, ensure_ascii=False, indent=indent)
    if args.output:
        with open(args.output, "w", encoding="utf-8") as fp:
            fp.write(text)
        if not args.quiet:
            print(f"wrote {args.output}", file=sys.stderr)
    else:
        print(text)


if __name__ == "__main__":
    main()
