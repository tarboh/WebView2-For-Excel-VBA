Attribute VB_Name = "Wv2Json"
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  D-2 段階 (JSON 文字列の復号を共通化) ---
'
'   D-2 の追加事項 (論点6 案b):
'     ★JsonUnescapeAt / JsonUnescape / JsonPickStr を新設した★
'       D-1 で Wv2Pane に Private として置いた JsonPickStr を、このモジュールへ
'       移設して Public にした。D-2 では Wv2Pane (EvalSync の結果ほどき) と
'       Wv2Element (プロパティ値の取り出し) の 2 箇所が同じ復号を必要とするため、
'       状態を持たない純粋な変換としてここに集約する (設計原則70)。
'
'     JsonUnescapeAt(json, startPos, outVal) As Long
'       JSON 文字列リテラルの★中身★を startPos から走査し、エスケープを解いて
'       outVal に入れる。戻り値は閉じ引用符の位置 (見つからなければ 0)。
'       startPos は「開き引用符の★次★の文字」を指すこと。
'
'     JsonUnescape(raw) As String
'       まるごと 1 個の JSON 値を受け取り、それが文字列リテラルなら引用符を外して
'       エスケープを解く。文字列でなければ (数値 / true / null / オブジェクト等)
'       ★そのまま素通しする★。EvalSync の戻り値をそのまま渡せる形にしてある。
'
'     JsonPickStr(json, key, outVal) As Boolean
'       "key":"value" の value を、エスケープを解いて取り出す。
'
'   ★JsonGetStr との使い分け★
'     JsonGetStr は★エスケープを解かない★軽い実装で、自前 JS が送る cmd / index の
'     ように「エスケープが入りえない値」を取り出すためのもの (設計原則64)。
'     ページ由来の文字列 (innerText / innerHTML / 属性値) は必ずエスケープを含むので、
'     そちらには JsonPickStr / JsonUnescape を使うこと。JsonGetStr は既存の
'     呼び出し元があるため★一切変更していない★。
'
'   ★復号が必須である理由 (仕様事実30)★
'     結果 JSON では < が \u003C にエスケープされて届く。D-2 で innerHTML を取ると
'     全タグがこの形で来るため、\uXXXX の復号は保険ではなく必須。
'
'   ★サロゲートペア★
'     \uD83D\uDE00 のような 2 個組は、ChrW$ を 2 回呼んで連結すれば正しく戻る
'     (VBA の String は UTF-16 コードユニット列なので、片割れ同士を足せばペアになる)。
'     ★Python でシミュレートするときは注意★ Python の str 連結では再結合しないため、
'     コードユニット列として持って最後に一括デコードしないと偽の NG が出る。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Json.bas  第9.13 段階 (整理: JSON ヘルパーの共通化) ---
'
'   ★VBA?JS 通信で使う最小限の JSON ヘルパー★
'
'   第9.11c?9.12d で Wv2TabBar と Wv2NavBar が各々 Private に同じ関数を持って
'   しまっていた (JsonEscape / GetJsonStr / GetJsonNum / BoolToJson)。3 つ目の UI が
'   出る前に共通モジュールへ抽出する (合意済みのクリーンアップ)。
'
'   ★なぜ Wv2Thunks.bas に入れないか★
'     Wv2Thunks.bas は「マシンコードサンク・手製 vtable・SAFEARRAY メモリプリミティブ」
'     という低レベル層のモジュールであり、既に 2600 行超。JSON 文字列処理は明らかに
'     別の関心事なので、責務を混ぜず独立したモジュールにする。
'
'   ★設計原則 64★ 自前 HTML の JS とだけ通信するので、汎用 JSON パーサは作らない。
'     送信側は自分で書いた JS に限られるため入力は完全に信頼でき、ネスト・配列・
'     エスケープを含む一般形への対応は不要。必要なキーを取り出す最小関数で足りる。
'     (VBA に標準 JSON パーサがなく、外部 COM も本プロジェクトの方針で使えない)
'
'   ★命名規則★ 標準モジュールの Public 関数なので、他モジュールとの名前衝突を
'     避けるため Json プレフィックスで統一する:
'       JsonEscape  … 文字列を JSON 文字列リテラルに安全に埋め込む
'       JsonGetStr  … "key":"value" の value を取り出す
'       JsonGetNum  … "key":123     の 123 を取り出す
'       JsonBool    … Boolean を true/false リテラルにする
'
'   旧名との対応 (9.13 で改名):
'       JsonEscape   ← JsonEscape  (変更なし)
'       JsonGetStr   ← GetJsonStr
'       JsonGetNum   ← GetJsonNum
'       JsonBool     ← BoolToJson
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' JsonEscape ? 文字列を JSON 文字列リテラルに安全に埋め込む
'
'   " と \ と制御文字 (0x00-0x1F) をエスケープする。
'   タブタイトルや URL を JSON に埋め込むための最小限。
'
'   ★第9.13 で修正: 高位文字を素通しするようにした (仕様事実 16)★
'     VBA の AscW は 0x8000 以上の文字を符号付き Integer として返す。
'     例: 「進」U+9032 = 36914 → AscW は -28622 を返す。
'     旧実装は制御文字判定を Case Is < 32 と書いていたため、日本語などの高位文字が
'     負値としてこの分岐に落ち、\uXXXX 形式にエスケープされていた。
'     ★結果は JSON として正しく、JS 側で正しい文字に戻るため実害はなかった★が、
'     意図しない経路であり、ログが読みにくく送信バイト数も無駄に増えていた。
'     → 共通化のタイミングで ch >= 0 And ch < 32 に直し、高位文字は素通しする。
'       挙動 (JS 側での見え方) は変わらない。
' ============================================================
Public Function JsonEscape(ByVal s As String) As String
    Dim out As String
    Dim i As Long
    Dim ch As Long
    Dim c As String

    out = ""
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        ch = AscW(c)
        Select Case ch
            Case 34            ' " (二重引用符)
                out = out & "\"""
            Case 92            ' \ (バックスラッシュ)
                out = out & "\\"
            Case 8             ' backspace
                out = out & "\b"
            Case 9             ' tab
                out = out & "\t"
            Case 10            ' LF
                out = out & "\n"
            Case 12            ' FF
                out = out & "\f"
            Case 13            ' CR
                out = out & "\r"
            Case Else
                ' ★第9.13 修正★ 制御文字だけを \uXXXX にする。
                '   ch >= 0 の条件が要 (AscW が高位文字を負値で返すため。仕様事実 16)。
                '   これがないと日本語などが \uXXXX に化ける。
                If ch >= 0 And ch < 32 Then
                    out = out & "\u" & Right$("0000" & Hex$(ch), 4)
                Else
                    out = out & c
                End If
        End Select
    Next i

    JsonEscape = out
End Function


' ============================================================
' JsonGetStr ? "key":"value" の value を取り出す
'
'   例: JsonGetStr("{""cmd"":""activate"",""index"":3}", "cmd") → "activate"
'   見つからなければ空文字を返す。
'
'   ★エスケープは解かない★ 送信元は自前の JS で、cmd や url にエスケープが必要な
'     文字は実質入らないため。厳密な復号が必要になったら拡張する (設計原則 64)。
' ============================================================
Public Function JsonGetStr(ByVal json As String, ByVal key As String) As String
    Dim pat As String
    pat = """" & key & """:"""
    ' pat の中身は  "key":"  という 7 文字 + キー長

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)                ' 値の先頭

    Dim en As Long
    en = InStr(st, json, """")       ' 閉じ引用符
    If en = 0 Then
        JsonGetStr = ""
        Exit Function
    End If

    JsonGetStr = Mid$(json, st, en - st)
End Function


' ============================================================
' JsonGetNum ? "key":123 の 123 を取り出す
'
'   例: JsonGetNum("{""cmd"":""close"",""index"":2}", "index") → 2
'   見つからなければ 0 を返す。
'
'   ★負数・小数は扱わない★ 現状 index (1-based の正整数) にしか使わないため。
'     必要になったら拡張する (設計原則 64)。
' ============================================================
Public Function JsonGetNum(ByVal json As String, ByVal key As String) As Long
    Dim pat As String
    pat = """" & key & """:"
    ' pat の中身は  "key":  という 4 文字 + キー長

    Dim p As Long
    p = InStr(json, pat)
    If p = 0 Then
        JsonGetNum = 0
        Exit Function
    End If

    Dim st As Long
    st = p + Len(pat)

    ' 空白をスキップ (JSON.stringify は出さないが念のため)
    Do While st <= Len(json)
        If Mid$(json, st, 1) <> " " Then Exit Do
        st = st + 1
    Loop

    ' 数字が続く限り読む
    Dim buf As String
    Dim c As String
    Do While st <= Len(json)
        c = Mid$(json, st, 1)
        If c < "0" Or c > "9" Then Exit Do
        buf = buf & c
        st = st + 1
    Loop

    If Len(buf) = 0 Then
        JsonGetNum = 0
    Else
        JsonGetNum = CLng(buf)
    End If
End Function


' ============================================================
' JsonBool ? Boolean を JSON の true/false リテラルにする
'
'   ★CStr(True) を使わない理由★ CStr(True) は "True" (先頭大文字) を返すうえ、
'     ロケール依存の懸念もある。JSON は小文字の true/false でなければならないので
'     明示的に変換する。
' ============================================================
Public Function JsonBool(ByVal b As Boolean) As String
    If b Then
        JsonBool = "true"
    Else
        JsonBool = "false"
    End If
End Function


' ============================================================
' JsonUnescapeAt (D-2) - JSON 文字列リテラルの中身を復号しながら読む
'
'   引数:
'     json     : 対象の JSON 文字列 (全体)
'     startPos : 走査開始位置。★開き引用符の次の文字★を指すこと (1 始まり)
'     outVal   : 取り出した値 (エスケープを解いたもの)
'   戻り値:
'     閉じ引用符の位置。閉じられていなければ 0 (このとき outVal は空文字)
'
'   対応するエスケープ: 引用符 / バックスラッシュ / スラッシュ / b / f / n / r / t / uXXXX
'   (いずれもバックスラッシュに続く 1 文字。uXXXX のみ 4 桁の 16 進が続く)
'   それ以外の \x は x をそのまま採用する (JSON としては不正だが、落とさない)。
' ============================================================
Public Function JsonUnescapeAt(ByVal json As String, _
                               ByVal startPos As Long, _
                               ByRef outVal As String) As Long
    Dim i As Long
    Dim buf As String
    Dim c As String
    Dim d As String

    outVal = ""
    JsonUnescapeAt = 0

    If startPos < 1 Then Exit Function

    i = startPos
    Do While i <= Len(json)
        c = Mid$(json, i, 1)

        If c = "\" Then
            d = Mid$(json, i + 1, 1)
            Select Case d
                Case """"
                    buf = buf & """"
                Case "\"
                    buf = buf & "\"
                Case "/"
                    buf = buf & "/"
                Case "b"
                    buf = buf & Chr$(8)
                Case "f"
                    buf = buf & Chr$(12)
                Case "n"
                    buf = buf & vbLf
                Case "r"
                    buf = buf & vbCr
                Case "t"
                    buf = buf & vbTab
                Case "u"
                    ' ★4 桁を超える値も正しく戻る★ CLng("&HD83D") は 16 ビットの
                    '   符号付きとして -10179 になるが、ChrW$ は負値を 65536 + n と
                    '   解釈するので 55357 (= D83D) の文字が得られる。
                    '   サロゲートの片割れ同士を連結すればペアとして成立する。
                    buf = buf & ChrW$(CLng("&H" & Mid$(json, i + 2, 4)))
                    i = i + 4
                Case Else
                    buf = buf & d
            End Select
            i = i + 2

        ElseIf c = """" Then
            outVal = buf
            JsonUnescapeAt = i
            Exit Function

        Else
            buf = buf & c
            i = i + 1
        End If
    Loop
End Function


' ============================================================
' JsonUnescape (D-2) - JSON 値 1 個を素の文字列に戻す
'
'   文字列リテラル (先頭が引用符) なら引用符を外してエスケープを解く。
'   それ以外 (数値 / true / false / null / undefined / オブジェクト / 配列) は
'   ★そのまま返す★。EvalSync の戻り値をそのまま渡せる。
'
'   閉じ引用符が見つからない壊れた入力は、原文をそのまま返す (握り潰さず、
'   呼び出し側が異常に気づけるようにする)。
' ============================================================
Public Function JsonUnescape(ByVal raw As String) As String
    Dim v As String
    Dim en As Long

    If Len(raw) = 0 Then
        JsonUnescape = ""
        Exit Function
    End If

    If Left$(raw, 1) <> """" Then
        JsonUnescape = raw
        Exit Function
    End If

    en = JsonUnescapeAt(raw, 2, v)
    If en > 0 Then
        JsonUnescape = v
    Else
        JsonUnescape = raw
    End If
End Function


' ============================================================
' JsonPickStr (D-2、D-1 で Wv2Pane に置いたものを移設)
'   "key":"value" の value を、エスケープを解いて取り出す。
'
'   引数:
'     json   : 対象の JSON 文字列
'     key    : 取り出すキー (値が文字列であること)
'     outVal : 取り出した値 (エスケープを解いたもの)
'   戻り値:
'     True  = キーが見つかり、閉じ引用符まで読めた
'     False = キーが無い、または閉じられていない
' ============================================================
Public Function JsonPickStr(ByVal json As String, _
                            ByVal key As String, _
                            ByRef outVal As String) As Boolean
    Dim pat As String
    Dim p As Long
    Dim en As Long

    outVal = ""
    JsonPickStr = False

    pat = """" & key & """:"""
    p = InStr(1, json, pat)
    If p = 0 Then Exit Function

    en = JsonUnescapeAt(json, p + Len(pat), outVal)
    JsonPickStr = (en > 0)
    If Not JsonPickStr Then outVal = ""
End Function


