Attribute VB_Name = "Wv2Url"
''''''''''''''''''''''''''''''''''
' --- Wv2Url.bas  第9.16 段階 (整理: URL 変換ユーティリティの切り出し) ---
'
'   URL 処理のうち「入力だけで出力が決まる純粋なデータ変換」を Wv2Browser から
'   切り出した標準モジュール。Wv2Json.bas (JSON 変換) と対をなす「状態を持たない
'   変換ユーティリティ群」。整理段階で作成 (9.13 の JSON 抽出と同じリズム)。
'
'   ★何を持つか (メカニズム = 純関数のみ)★
'     ・IsAlphaOnly(s)          : 文字列が英字のみか (スキーム名判定に使う)
'     ・UrlEncodeComponent(s)   : encodeURIComponent 相当の percent-encoding
'     ・Utf8Bytes(s)            : String(UTF-16LE) → UTF-8 バイト列 (ADODB.Stream)
'
'   ★何を持たないか (ポリシー = 状態依存は Wv2Browser に残す)★
'     NormalizeUrl / BuildSearchUrl は検索エンジン設定 (m_searchTemplate) という
'     インスタンス状態に依存する「ブラウザの振る舞い」なので Wv2Browser に残し、
'     この Wv2Url の純関数を呼び出す。ポリシーとメカニズムの分離。
'
'   ★ADODB.Stream の扱い★ Utf8Bytes は ADODB.Stream を使うが、これは WebView2
'     制御 (サンク・vtable・メモリプリミティブ) とは無関係な単なるデータ変換であり、
'     既に HTML 書き出しで採用済み。本プロジェクトの「WebView2 を外部 COM なしで
'     叩く」核心思想は侵さない (9.14 の判断を踏襲)。
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' IsAlphaOnly (第9.16 に Wv2Browser から移設、ロジック無変更)
'
'   文字列が英字 (A-Z / a-z) のみで構成されるかを返す。
'   NormalizeUrl が "about:" 等のスキーム名を判定するのに使う
'   ("example.com:8080" を誤ってスキーム扱いしないよう、数字を含むものは False)。
' ============================================================
Public Function IsAlphaOnly(ByVal s As String) As Boolean
    If Len(s) = 0 Then Exit Function
    Dim i As Long
    Dim c As String
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If c < "a" Or c > "z" Then Exit Function
    Next i
    IsAlphaOnly = True
End Function


' ============================================================
' UrlEncodeComponent (第9.16 に Wv2Browser から移設、ロジック無変更)
'
'   文字列を RFC3986 に沿って percent-encoding する
'   (JS の encodeURIComponent 相当。非予約文字だけ素通し、他は %XX)。
'
'   ★非予約文字 (unreserved)★ A-Z a-z 0-9 と "-" "_" "." "~" はそのまま。
'     それ以外はすべて UTF-8 バイト列にして各バイトを %XX にする。
'     (空白は %20。encodeURIComponent と同じく "+" にはしない)
'
'   ★UTF-8 化は Utf8Bytes (ADODB.Stream 経由) が担当★ 日本語等の高位文字を
'     正しく複数バイトに展開するため。1 文字ずつ AscW で処理すると
'     サロゲートペア (絵文字等) を壊すので、文字列全体を一括で UTF-8 化してから
'     バイト単位でエスケープする。
' ============================================================
Public Function UrlEncodeComponent(ByVal s As String) As String
    If Len(s) = 0 Then Exit Function

    Dim bytes() As Byte
    bytes = Utf8Bytes(s)

    ' Utf8Bytes が空 (Len(s)=0 は上で弾いているので通常起きない) なら空を返す
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(bytes)
    ub = UBound(bytes)
    If Err.Number <> 0 Then
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Dim sb As String
    Dim i As Long
    Dim b As Integer
    Dim c As String
    For i = lb To ub
        b = bytes(i)            ' 0..255
        c = Chr$(b)             ' ASCII 範囲の判定用 (b<128 のときのみ意味を持つ)
        If (b >= 65 And b <= 90) _
        Or (b >= 97 And b <= 122) _
        Or (b >= 48 And b <= 57) _
        Or b = 45 Or b = 95 Or b = 46 Or b = 126 Then
            ' unreserved: A-Z a-z 0-9 - _ . ~
            sb = sb & c
        Else
            sb = sb & "%" & Right$("0" & Hex$(b), 2)
        End If
    Next i

    UrlEncodeComponent = sb
End Function


' ============================================================
' Utf8Bytes (第9.16 に Wv2Browser から移設、ロジック無変更)
'
'   VBA の String (UTF-16LE) を UTF-8 のバイト列に変換する。
'   ADODB.Stream に Charset="utf-8" で書き込み、Type=adTypeBinary で読み戻す。
'
'   ★先頭 BOM は付かない★ Charset="utf-8" のテキスト書き込みでは BOM は付与
'     されない (adTypeBinary で読み戻したバイト列に EF BB BF は含まれない)。
'     万一の将来変更に備え、先頭 3 バイトが BOM だった場合は除去する保険を入れる。
' ============================================================
Public Function Utf8Bytes(ByVal s As String) As Byte()
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "utf-8"
    st.Open
    st.WriteText s
    st.Position = 0
    st.Type = 1                 ' adTypeBinary
    ' テキストからバイナリへ切り替えると Position が先頭に戻る。
    Dim raw() As Byte
    raw = st.Read                ' 全バイト取得
    st.Close
    Set st = Nothing

    ' BOM 保険 (通常は付かないが、付いていたら EF BB BF を落とす)
    On Error Resume Next
    Dim lb As Long, ub As Long
    lb = LBound(raw)
    ub = UBound(raw)
    If Err.Number <> 0 Then
        ' 空文字入力等でバイトが無い場合。空配列を返す。
        On Error GoTo 0
        Utf8Bytes = raw
        Exit Function
    End If
    On Error GoTo 0

    If (ub - lb) >= 2 Then
        If raw(lb) = &HEF And raw(lb + 1) = &HBB And raw(lb + 2) = &HBF Then
            Dim out() As Byte
            ReDim out(0 To (ub - lb) - 3)
            Dim j As Long
            For j = lb + 3 To ub
                out(j - (lb + 3)) = raw(j)
            Next j
            Utf8Bytes = out
            Exit Function
        End If
    End If

    Utf8Bytes = raw
End Function

