#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_idents.py -- 新規識別子が VBA 予約語や既存識別子と衝突しないか調べる

    python tools/check_idents.py LogE LogW LogI LogD LogPath
    python tools/check_idents.py --dump            # 既存識別子を一覧にする

なぜ要るか:
    (1) VBA は識別子の大小を区別しない。eNum は予約語 Enum と衝突する (設計原則90)。
    (2) ★仕様事実43★ 新規モジュールが既存と同じ綴りを別の大小で宣言すると、
        VBA がプロジェクト全体の綴りを統一し、触っていないモジュールまで書き換わる。

    したがって新規識別子は「予約語と小文字化して照合」だけでは足りず、
    「既存の全識別子と小文字化して照合」する必要がある。

判定:
    NG   予約語と衝突
    NG   既存に同綴りがあり、大小が違う  -> プロジェクト全体が書き換わる
    OK*  既存に同綴りがあり、大小も同じ  -> 書き換えは起きないが名前が被る
    OK   どこにも無い

終了コード: 0 = 問題なし / 1 = NG あり
"""

import os
import re
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# VBA の予約語・組み込み (小文字で保持)
RESERVED = set("""
abs addressof and any array as attribute base binary boolean byref byte byval call
case cbool cbyte ccur cdate cdbl cdec chdir chdrive cint clng clngptr clnglng close
collection compare const cos csng cstr currency cvar cvdate cverr date debug decimal
declare defbool defbyte defcur defdate defdbl defdec defint deflng defobj defsng
defstr defvar dim dir do doevents double each else elseif empty end endif enum eqv
erase err error event exit exp explicit false fix for friend function get global
gosub goto if imp implements in input instr int integer is kill lbound lcase left
len let lib like line lock long longlong longptr loop lset ltrim me mid mkdir mod
module name new next not nothing null object on open option optional or paramarray
preserve print private property ptrsafe public put raiseevent randomize redim rem
reset resume return rgb right rmdir rnd rset rtrim seek select set sgn shell single
sin space spc sqr static step stop str strcomp strconv string sub switch tab tan
text then time timer to trim true type typeof ubound ucase unlock until val variant
vartype wend while width with withevents write xor
""".split())

DECL_PATTERNS = [
    # Sub / Function / Property の名前
    re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
               r'(?:Sub|Function)\s+([A-Za-z_]\w*)', re.I),
    re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*'
               r'Property\s+(?:Get|Let|Set)\s+([A-Za-z_]\w*)', re.I),
    # Dim / Const / Public / Private の変数
    re.compile(r'^\s*(?:Dim|Const|Public|Private|Global|Static)\s+'
               r'(?:WithEvents\s+)?([A-Za-z_]\w*)', re.I),
    # Enum / Type の名前
    re.compile(r'^\s*(?:Public\s+|Private\s+)?(?:Enum|Type)\s+([A-Za-z_]\w*)', re.I),
    # Declare
    re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\s+(?:PtrSafe\s+)?'
               r'(?:Sub|Function)\s+([A-Za-z_]\w*)', re.I),
]

# 引数リストの ByVal / ByRef / Optional つき仮引数
PARAM_RE = re.compile(r'(?:ByVal|ByRef|Optional|ParamArray)\s+([A-Za-z_]\w*)', re.I)


def collect(src_dir):
    """既存識別子を {小文字: set(実際の綴り)} で返す"""
    table = {}

    def add(name, where):
        table.setdefault(name.lower(), set()).add(name)

    files = sorted(glob.glob(os.path.join(src_dir, "*.bas")) +
                   glob.glob(os.path.join(src_dir, "*.cls")))
    for f in files:
        mod = os.path.splitext(os.path.basename(f))[0]
        add(mod, f)
        with open(f, encoding="utf-8", newline="") as fh:
            text = fh.read()
        for line in text.split("\r\n"):
            s = line.strip()
            if s.startswith("'"):
                continue
            for pat in DECL_PATTERNS:
                m = pat.match(line)
                if m:
                    add(m.group(1), f)
            for m in PARAM_RE.finditer(line):
                add(m.group(1), f)
    return table, files


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    src = os.path.join(repo, "src")

    args = [a for a in argv[1:] if not a.startswith("--")]
    dump = "--dump" in argv

    table, files = collect(src)
    print("既存識別子を %d 本のファイルから収集: %d 種" % (len(files), len(table)))
    print()

    if dump:
        for k in sorted(table):
            sp = sorted(table[k])
            mark = "  <- 綴りが複数" if len(sp) > 1 else ""
            print("  %s%s" % (" / ".join(sp), mark))
        return 0

    if not args:
        print("調べたい識別子を引数で渡すこと。例: python tools/check_idents.py LogE LogD")
        return 1

    ng = 0
    print("%-16s %-10s %s" % ("候補", "判定", "理由"))
    print("-" * 72)
    for name in args:
        low = name.lower()
        if low in RESERVED:
            print("%-16s %-10s VBA 予約語" % (name, "NG"))
            ng += 1
            continue
        if low in table:
            spellings = table[low]
            if name in spellings and len(spellings) == 1:
                print("%-16s %-10s 既存に同じ綴りがある (%s)" % (name, "OK*", " / ".join(sorted(spellings))))
            else:
                print("%-16s %-10s ★既存に別の大小がある: %s -> プロジェクト全体が書き換わる★"
                      % (name, "NG", " / ".join(sorted(spellings))))
                ng += 1
            continue
        if len(name) <= 2:
            print("%-16s %-10s 衝突なし。ただし短すぎる (仕様事実43 の的になりやすい)" % (name, "注意"))
            continue
        print("%-16s %-10s 衝突なし" % (name, "OK"))
    print()
    if ng:
        print("★NG %d 件★ 名前を変えること。" % ng)
    else:
        print("NG なし。")
    return 1 if ng else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
