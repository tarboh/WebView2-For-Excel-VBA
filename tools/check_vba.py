#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
check_vba.py -- VBA ソースの整形・構文検査 (定型検査の 1 番)

    python tools/check_vba.py                 # src/ を全部
    python tools/check_vba.py src/Wv2Log.bas  # ファイル指定

見るもの:
    ・BOM / CRLF / 裸 LF / 裸 CR
    ・行末の空白 (VBE が落とすので差分の元になる)
    ・引用符の偶奇
    ・行継続記号の右側に同じ行内でコメント (構文エラーになる)
    ・Sub / Function / Property / Enum / Type の開閉一致
    ・If / Do / Select / For / With の均衡
    ・Exit の対象が、そのとき開いているブロックと合っているか
    ・&H リテラルが 0x8000〜0xFFFF なのに & サフィックスが無い (設計原則29)
    ・コロン区切りの 1 行完結 Sub / Enum も正しく閉じとみなす

終了コード: 0 = 問題なし / 1 = 問題あり
"""

import os
import re
import sys
import glob

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

OPENERS = {
    "sub": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Sub\s+\w+', re.I),
    "function": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Function\s+\w+', re.I),
    "property": re.compile(r'^\s*(?:Public\s+|Private\s+|Friend\s+|Static\s+)*Property\s+(?:Get|Let|Set)\s+\w+', re.I),
    "enum": re.compile(r'^\s*(?:Public\s+|Private\s+)?Enum\s+\w+', re.I),
    "type": re.compile(r'^\s*(?:Public\s+|Private\s+)?Type\s+\w+', re.I),
    "with": re.compile(r'^\s*With\s+\S', re.I),
    "select": re.compile(r'^\s*Select\s+Case\b', re.I),
}
CLOSERS = {
    "sub": re.compile(r'^\s*End\s+Sub\b', re.I),
    "function": re.compile(r'^\s*End\s+Function\b', re.I),
    "property": re.compile(r'^\s*End\s+Property\b', re.I),
    "enum": re.compile(r'^\s*End\s+Enum\b', re.I),
    "type": re.compile(r'^\s*End\s+Type\b', re.I),
    "with": re.compile(r'^\s*End\s+With\b', re.I),
    "select": re.compile(r'^\s*End\s+Select\b', re.I),
}
DECLARE_RE = re.compile(r'^\s*(?:Public\s+|Private\s+)?Declare\b', re.I)
IF_OPEN_RE = re.compile(r'^\s*(?:\}|)\s*(?:Else)?If\b.*\bThen\s*$', re.I)
IF_END_RE = re.compile(r'^\s*End\s+If\b', re.I)
FOR_OPEN_RE = re.compile(r'^\s*For\b', re.I)
FOR_END_RE = re.compile(r'^\s*Next\b', re.I)
DO_OPEN_RE = re.compile(r'^\s*Do\b', re.I)
DO_END_RE = re.compile(r'^\s*Loop\b', re.I)
EXIT_RE = re.compile(r'^\s*Exit\s+(Sub|Function|Property|Do|For)\b', re.I)


def strip_comment(line):
    """文字列リテラルの外にあるアポストロフィ以降を落とす。引用符の偶奇も返す。"""
    out = []
    in_str = False
    quotes = 0
    i = 0
    while i < len(line):
        c = line[i]
        if c == '"':
            quotes += 1
            in_str = not in_str
            out.append(c)
        elif c == "'" and not in_str:
            break
        else:
            out.append(c)
        i += 1
    return "".join(out), quotes


ONE_LINER = {}
for _k in ("sub", "function", "property", "enum", "type", "with", "select"):
    ONE_LINER[_k] = re.compile(r":\s*End\s+" + _k + r"\b", re.I)


def is_one_liner(code, kind, pat):
    """Private Sub EntryPoint(): End Sub のようにコロンで 1 行完結しているか"""
    m = pat.match(code)
    if not m:
        return False
    return bool(ONE_LINER[kind].search(code))


def check_file(path):
    problems = []

    def bad(lineno, msg, text=""):
        problems.append((lineno, msg, text))

    raw = open(path, "rb").read()
    if raw[:3] == b"\xef\xbb\xbf":
        bad(0, "BOM が付いている")
        raw = raw[3:]
    crlf = raw.count(b"\r\n")
    bare_lf = raw.count(b"\n") - crlf
    bare_cr = raw.count(b"\r") - crlf
    if bare_lf:
        bad(0, "裸 LF が %d 個" % bare_lf)
    if bare_cr:
        bad(0, "裸 CR が %d 個" % bare_cr)

    text = raw.decode("utf-8")
    lines = text.split("\r\n")
    if lines and lines[-1] == "":
        lines = lines[:-1]

    # --- 行継続 (_) を連結して「論理行」にする ---
    # ★これをやらないと、複数行に跨る If ... Then を開きとして拾えず、
    #   End If が「対応する If が無い」と誤検知になる★
    logical = []        # (先頭の物理行番号, 連結したコード, 元の行)
    buf = ""
    buf_line = 0
    for i, line in enumerate(lines, 1):
        if line.rstrip() != line:
            bad(i, "行末に空白がある", line)
        code, quotes = strip_comment(line)
        if quotes % 2 != 0:
            bad(i, "引用符が奇数個", line)
        s = code.rstrip()
        if buf == "":
            buf_line = i
        if s.endswith("_"):
            # ★行継続の右側に同じ行内でコメントを書くと構文エラーになる★
            #   このチェックは「物理行」でしかできない。論理行に畳んだ後だと
            #   継続行そのものが手元に来ないため素通しになる。
            pos = line.rfind("_")
            if pos >= 0 and "'" in line[pos:]:
                bad(i, "行継続の右側に同じ行内でコメント (構文エラー)", line)
            buf = buf + s[:-1] + " "
            continue
        logical.append((buf_line, buf + s, line))
        buf = ""
    if buf:
        logical.append((buf_line, buf, ""))

    # --- メンバー属性行が宣言から引き離されていないか ---
    #   ★Attribute m_foo.VB_VarHelpID = -1 は「宣言行の直後」でなければならない★
    #   間に行を差し込むと VBE が属性として吸収できず、コードとして露出する。
    #   K-2 でパッチが宣言と属性の間に挿入してしまい、これを踏みかけた。
    #   Export したファイルを編集するときの定番の事故なので常設で見る。
    prev_code = ""
    for i, code, line in logical:
        s = code.strip()
        m = re.match(r'Attribute\s+(\w+)\.\w+\s*=', s)
        if m:
            name = m.group(1)
            if not re.search(r'\b' + re.escape(name) + r'\b', prev_code):
                bad(i, "メンバー属性 %s が宣言行の直後にない (VBE がコードとして扱う)"
                    % name, line)
        if s:
            prev_code = s

    stack = []          # (kind, lineno)
    for i, code, line in logical:

        # &H リテラル (設計原則29)
        # ★危険なのは 0x8000..0xFFFF だけ★ この範囲は Integer と解釈されて
        #   負の値になる。0x10000 以上は最初から Long なので問題ない。
        #   & (Long) / ^ (LongLong) のサフィックスが付いていれば明示済み。
        for mh in re.finditer(r'&H([0-9A-Fa-f]+)([&^]?)', code):
            v = int(mh.group(1), 16)
            if 0x8000 <= v <= 0xFFFF and mh.group(2) == "":
                bad(i, "&H%s は Integer に化ける。& サフィックスが要る (設計原則29)"
                    % mh.group(1), line)

        s = code.strip()
        if not s:
            continue
        if DECLARE_RE.match(code):
            continue

        # 開き
        matched = False
        for kind, pat in OPENERS.items():
            if pat.match(code):
                # ★コロン区切りの 1 行完結★ 例: Private Sub EntryPoint(): End Sub
                if is_one_liner(code, kind, pat):
                    matched = True
                    break
                stack.append((kind, i))
                matched = True
                break
        if matched:
            continue

        # 閉じ
        for kind, pat in CLOSERS.items():
            if pat.match(code):
                if not stack:
                    bad(i, "End %s に対応する開きが無い" % kind, line)
                elif stack[-1][0] != kind:
                    bad(i, "End %s だが直近の開きは %s (%d 行目)" % (kind, stack[-1][0], stack[-1][1]), line)
                    stack.pop()
                else:
                    stack.pop()
                matched = True
                break
        if matched:
            continue

        if IF_OPEN_RE.match(code) and not re.match(r'^\s*ElseIf\b', code, re.I):
            stack.append(("if", i))
        elif IF_END_RE.match(code):
            if stack and stack[-1][0] == "if":
                stack.pop()
            else:
                bad(i, "End If に対応する If が無い", line)
        elif FOR_OPEN_RE.match(code):
            stack.append(("for", i))
        elif FOR_END_RE.match(code):
            if stack and stack[-1][0] == "for":
                stack.pop()
            else:
                bad(i, "Next に対応する For が無い", line)
        elif DO_OPEN_RE.match(code):
            stack.append(("do", i))
        elif DO_END_RE.match(code):
            if stack and stack[-1][0] == "do":
                stack.pop()
            else:
                bad(i, "Loop に対応する Do が無い", line)
        else:
            me = EXIT_RE.match(code)
            if me:
                want = me.group(1).lower()
                kinds = [k for k, _ in stack]
                if want not in kinds:
                    bad(i, "Exit %s だが %s ブロックの中に居ない (開いているのは %s)"
                        % (me.group(1), want, "/".join(kinds) or "なし"), line)

    for kind, lineno in stack:
        bad(lineno, "%s が閉じていない" % kind)

    return problems


def main(argv):
    repo = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    args = argv[1:]
    if args:
        files = []
        for a in args:
            if os.path.isdir(a):
                files += glob.glob(os.path.join(a, "*.bas")) + glob.glob(os.path.join(a, "*.cls"))
            else:
                files.append(a)
    else:
        src = os.path.join(repo, "src")
        files = glob.glob(os.path.join(src, "*.bas")) + glob.glob(os.path.join(src, "*.cls"))
    files = sorted(set(files))

    total = 0
    for f in files:
        problems = check_file(f)
        rel = os.path.relpath(f, repo).replace("\\", "/")
        if problems:
            print("=== %s : %d 件 ===" % (rel, len(problems)))
            for lineno, msg, txt in problems:
                print("  %s:%d  %s" % (rel, lineno, msg))
                if txt:
                    print("      | %s" % txt.strip()[:110])
            total += len(problems)
        else:
            print("OK  %s" % rel)
    print()
    if total:
        print("★問題 %d 件★" % total)
    else:
        print("問題なし (%d ファイル)" % len(files))
    return 1 if total else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
