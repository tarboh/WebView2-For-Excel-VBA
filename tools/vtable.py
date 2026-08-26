# -*- coding: utf-8 -*-
"""WebView2 SDK のヘッダから vtable の並びを取り出す。

★COM の vtable index は推測で書かない★ (N-1 の教訓)
  N-1 の準備で add_WebResourceRequested を 54 と推測したが、正解は 55 だった。
  OpenDevToolsWindow を数え落としていた。★実機で試していたら別のメソッドを呼んで
  クラッシュしていた。★ index は必ずこのスクリプトで確定させること。

使い方:
    python tools/vtable.py ICoreWebView2
    python tools/vtable.py ICoreWebView2WebResourceRequest
    python tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT
    python tools/vtable.py --list            (IF の名前を一覧)
    python tools/vtable.py --find WebResource (名前で探す)

ヘッダの場所:
    既定は下記の候補を順に探す。環境変数 WEBVIEW2_H で明示もできる。
    ★Downloads の中は消えうる★ ので、動かすときは環境変数を使うか
    このファイルの HEADER_CANDIDATES に足すこと。
"""
import io
import os
import re
import sys

HEADER_CANDIDATES = [
    os.environ.get("WEBVIEW2_H", ""),
    r"C:\Users\gugug\Downloads\新しいフォルダー (4)\WebView2.h",
    r"C:\Users\gugug\GitHub\WebView2-For-Excel-VBA\tools\WebView2.h",
]


def find_header():
    for p in HEADER_CANDIDATES:
        if p and os.path.isfile(p):
            return p
    return None


def load(path):
    with io.open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read()


def vtable(src, iface):
    """IF の vtable を宣言順に返す。IUnknown の 3 本も含めた通し番号。"""
    i = src.find("typedef struct " + iface + "Vtbl")
    if i < 0:
        return None
    j = src.find("} " + iface + "Vtbl", i)
    if j < 0:
        j = i + 200000
    # 関数ポインタ宣言 "... ( STDMETHODCALLTYPE *Name )( ..." から名前を拾う
    return re.findall(r"\*\s*(\w+)\s*\)\s*\(", src[i:j])


def enum_values(src, name):
    """enum の値を (識別子, 値) で返す。SDK は前の値 + 1 の形で書いてある。"""
    i = src.find(name)
    while i >= 0:
        j = src.find("}", i)
        blk = src[i:j]
        if "=" in blk:
            break
        i = src.find(name, i + 1)
    if i < 0:
        return []

    out = []
    n = 0
    for ln in blk.split("\n"):
        m = re.match(r"\s*(\w+)\s*=\s*(.+?),?\s*$", ln)
        if not m:
            continue
        ident, expr = m.group(1), m.group(2)
        m2 = re.match(r"^\s*(-?\d+)\s*$", expr)
        if m2:
            n = int(m2.group(1))
        else:
            n = n + 1 if out else 0
        out.append((ident, n))
    return out


def main():
    path = find_header()
    if not path:
        print("WebView2.h が見つからない。環境変数 WEBVIEW2_H で場所を指定すること。")
        print("候補として探した場所:")
        for p in HEADER_CANDIDATES:
            if p:
                print("   ", p)
        return 1

    src = load(path)
    args = sys.argv[1:]
    if not args:
        print(__doc__)
        print("ヘッダ: %s" % path)
        return 0

    if args[0] == "--list":
        names = sorted(set(re.findall(r"typedef struct (\w+)Vtbl", src)))
        print("ヘッダ: %s  (%d 個)" % (path, len(names)))
        for n in names:
            print("   ", n)
        return 0

    if args[0] == "--find":
        key = args[1].lower()
        names = sorted(set(re.findall(r"typedef struct (\w+)Vtbl", src)))
        hit = [n for n in names if key in n.lower()]
        print("ヘッダ: %s" % path)
        print("「%s」を含む IF: %d 個" % (args[1], len(hit)))
        for n in hit:
            print("   ", n)
        return 0

    if args[0] == "--enum":
        vals = enum_values(src, args[1])
        if not vals:
            print("enum %s が見つからない" % args[1])
            return 1
        print("ヘッダ: %s" % path)
        print("=== %s" % args[1])
        for ident, n in vals:
            print("   %3d  %s" % (n, ident))
        return 0

    iface = args[0]
    v = vtable(src, iface)
    if v is None:
        print("IF %s が見つからない (--find で探せる)" % iface)
        return 1

    print("ヘッダ: %s" % path)
    print("=== %s  (vtable %d 本。0-2 は IUnknown)" % (iface, len(v)))
    for k, n in enumerate(v):
        mark = "   " if k >= 3 else " * "
        print("%s%3d  %s" % (mark, k, n))
    return 0


if __name__ == "__main__":
    sys.exit(main())
