#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_cp932.py -- src/ の VBA ソースに「ブックへ入れると壊れる文字」が無いか検査する

    python tools/check_cp932.py                 # src/ と src/document/ を検査
    python tools/check_cp932.py path\to\file    # ファイル/フォルダを指定して検査

なぜ要るか:
    VBA のソースはプロジェクトのコードページ (日本語環境では CP932) で保持される。
    CP932 で扱えない文字は import した時点で失われ、二度と戻らない。
    失敗モードは 2 つある。

      (1) encode 自体が不可          -> ? になる            例: 歯車 U+2699, チェック U+2713
      (2) encode は通るが別の字に化ける -> 黙って置き換わる    例: U+301C -> U+FF5E

    (2) は目視では気づけない。だから「encode できるか」ではなく
    「往復して同じ字に戻るか」で判定する。

    なお Python の cp932 コーデックは Windows より寛容で、U+301C は Python では
    通るが VBA では ? になる。往復検査ならどちらのモードも捕まる。

終了コード:
    0 = 問題なし / 1 = 問題あり
"""

import os
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# 化けた文字を書き直すときの候補 (CP932 で往復できると確認済みのもの)
SUGGEST = {
    "⚙": "[設定]",      # 歯車
    "⟳": "[更新]",      # 回転矢印
    "⇄": "⇔",      # 左右矢印 -> ⇔
    "↔": "⇔",      # ↔ -> ⇔
    "✓": "○",      # ✓ -> ○
    "✗": "×",      # ✗ -> ×
    "➡": "→",      # ➡ -> →
    "—": "―",      # — (EM DASH) -> ― (HORIZONTAL BAR)
    "〜": "～",      # 〜 (WAVE DASH) -> ～ (FULLWIDTH TILDE)
    "−": "－",      # − (MINUS SIGN) -> －
    "¢": "￠",      # ¢ -> ￠
    "£": "￡",
    "¬": "￢",
}

EXTS = (".bas", ".cls", ".frm")


def classify(ch):
    """ch を CP932 へ往復させて判定する。問題なければ None を返す。"""
    try:
        enc = ch.encode("cp932")
    except UnicodeEncodeError:
        return "encode 不可 (? になる)"
    back = enc.decode("cp932")
    if back != ch:
        return "別の字に化ける -> %s (U+%04X)" % (back, ord(back))
    return None


def collect(paths):
    files = []
    for p in paths:
        if os.path.isdir(p):
            for ext in EXTS:
                files.extend(glob.glob(os.path.join(p, "**", "*" + ext), recursive=True))
        elif os.path.isfile(p):
            files.append(p)
    return sorted(set(files))


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    targets = argv[1:] if len(argv) > 1 else [os.path.join(repo, "src")]

    files = collect(targets)
    if not files:
        print("検査対象が見つからない: %s" % ", ".join(targets))
        return 1

    problems = []
    for f in files:
        # newline="" を付けないと CRLF が LF に正規化されて桁がずれる
        with open(f, encoding="utf-8", newline="") as fh:
            text = fh.read()
        for lineno, line in enumerate(text.split("\r\n"), 1):
            for col, ch in enumerate(line, 1):
                if ord(ch) < 0x80:
                    continue
                why = classify(ch)
                if why:
                    problems.append((f, lineno, col, ch, why, line.strip()))

    rel = lambda p: os.path.relpath(p, repo).replace("\\", "/")

    print("検査対象: %d ファイル" % len(files))
    print()

    if not problems:
        print("問題なし。CP932 で往復できない文字は 1 個も無い。")
        return 0

    print("★問題 %d 件★ ブックへ import すると失われる文字がある" % len(problems))
    print()
    for f, lineno, col, ch, why, src in problems:
        sug = SUGGEST.get(ch)
        tail = ("  代替候補: %s" % sug) if sug else ""
        print("%s:%d:%d" % (rel(f), lineno, col))
        print("    %s (U+%04X)  %s%s" % (ch, ord(ch), why, tail))
        print("    | %s" % src[:110])
    print()

    # 文字ごとの集計
    from collections import Counter
    tally = Counter(p[3] for p in problems)
    print("文字ごとの件数:")
    for ch, n in tally.most_common():
        sug = SUGGEST.get(ch)
        tail = ("  -> %s" % sug) if sug else ""
        print("  %s (U+%04X) x%d%s" % (ch, ord(ch), n, tail))
    return 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
