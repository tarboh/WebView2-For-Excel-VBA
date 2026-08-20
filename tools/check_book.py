#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_book.py -- ブック内部に格納されている VBA の実体を直接検査する

    python tools/check_book.py                    # book/ の .xlsm
    python tools/check_book.py path\to\book.xlsm

★なぜ COM ではなくブックを直接読むのか★
    VBIDE の CodeModule.Lines は ★Attribute 行を一切返さない★。
    そのため COM 経由では、属性行がコードとして露出していても検出できない。
    K-2 で実際にこれを踏んだ:

      Private WithEvents m_newTabTimer As SafeTimer
      Attribute m_newTabTimer.VB_VarHelpID = -1    <- VBE が持つ本物 (非表示)
      Attribute m_newTabTimer.VB_VarHelpID = -1    <- AddFromString が入れた重複 (露出)

    重複した方は VBE のコードペインに現れ、コンパイルエラーになる。
    しかし CodeModule 経由では両方とも見えないので「異常なし」と嘘をつく。

    olevba でブック内部のモジュールストリームを読めば実体が分かる。

見るもの:
    ・メンバー属性行 (Attribute <名前>.<属性>) の重複
    ・宣言行の直後にないメンバー属性行
    ・VB_VarHelpID 以外のメンバー属性 (VBE が自動生成しないので import で失われる)

前提: pip install oletools

終了コード: 0 = 問題なし / 1 = 問題あり
"""

import os
import re
import sys
import glob
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

try:
    from oletools.olevba import VBA_Parser
except ImportError:
    print("oletools が要る:  pip install oletools")
    sys.exit(1)

MEMBER_ATTR = re.compile(r"^Attribute\s+(\w+)\.(\w+)\s*=")


def check(path):
    problems = []
    vp = VBA_Parser(path)
    try:
        rows = []
        for (_fname, _stream, vba_name, code) in vp.extract_macros():
            lines = code.replace("\r", "").split("\n")
            hits = []
            for i, l in enumerate(lines, 1):
                m = MEMBER_ATTR.match(l)
                if m:
                    hits.append((i, l.strip(), m.group(1), m.group(2)))
            if not hits:
                continue

            dup = sum(v - 1 for v in Counter(h[1] for h in hits).values())
            rows.append((vba_name, len(hits), dup))

            if dup:
                problems.append("%s : メンバー属性行が %d 本 重複している "
                                "(片方が VBE でコードとして露出する)" % (vba_name, dup))
                seen = set()
                for i, text, name, attr in hits:
                    if text in seen:
                        problems.append("    %d 行目: %s  <- 重複" % (i, text))
                    seen.add(text)

            for i, text, name, attr in hits:
                prev = lines[i - 2] if i >= 2 else ""
                if not re.search(r"\b" + re.escape(name) + r"\b", prev):
                    if text not in [h[1] for h in hits[:hits.index((i, text, name, attr))]]:
                        problems.append("%s : %d 行目の %s が宣言行の直後にない"
                                        % (vba_name, i, text))
                if attr != "VB_VarHelpID":
                    problems.append("%s : %d 行目の %s は VBE が自動生成しない属性。"
                                    "import で失われる可能性がある" % (vba_name, i, text))

        print("%-24s %8s %8s" % ("モジュール", "属性行", "重複"))
        print("-" * 44)
        for nm, cnt, dup in rows:
            mark = "  ★" if dup else ""
            print("%-24s %8d %8d%s" % (nm, cnt, dup, mark))
        if not rows:
            print("(メンバー属性を持つモジュールなし)")
        print()
    finally:
        vp.close()
    return problems


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    if len(argv) > 1:
        books = argv[1:]
    else:
        books = sorted(glob.glob(os.path.join(repo, "book", "*.xlsm")))
    if not books:
        print("検査対象の .xlsm が見つからない")
        return 1

    total = 0
    for b in books:
        print("=== %s ===" % os.path.basename(b))
        problems = check(b)
        if problems:
            print("★問題★")
            for p in problems:
                print("  " + p)
            total += len(problems)
        else:
            print("問題なし")
        print()

    if total:
        print("★合計 %d 件★" % total)
        print()
        print("重複の直し方: CodeModule からは見えないので DeleteLines では消せない。")
        print("  tools/repair_book.ps1 でコンポーネントを削除して Import し直す。")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
