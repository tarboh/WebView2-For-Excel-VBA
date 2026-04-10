Attribute VB_Name = "mWebView2Helper"
'Module mWebView2Helper


Option Explicit

' ???????????????????????????????????????????????????????????????????????
' mWebView2Helper
' WebView2 ラッパークラス用ヘルパー関数群
' ???????????????????????????????????????????????????????????????????????

#If Win64 Then
#Else
    #Error "このモジュールは 64 ビット VBA (x64) が必要です"
#End If

' --- API 宣言 (モジュール内ローカル) ---

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)


' ???????????????????????????????????????????????
' Section 10: バイト列 ⇔ Base64 変換
' ???????????????????????????????????????????????

Private Declare PtrSafe Function CryptStringToBinaryW Lib "crypt32" ( _
    ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long
Private Declare PtrSafe Function CryptBinaryToStringW Lib "crypt32" ( _
    ByVal pbBinary As LongPtr, ByVal cbBinary As Long, ByVal dwFlags As Long, _
    ByVal pszString As LongPtr, ByRef pcchString As Long) As Long

' ???????????????????????????????????????????????
' Section 1: DispCallFunc 汎用ラッパー
' ???????????????????????????????????????????????

''' <summary>
''' DispCallFunc を使って COM vtable の任意メソッドを呼び出す汎用ラッパー。
''' 可変引数 (ParamArray) で最大16個のパラメータを受け取る。
'''
''' 戻り値: COM メソッドの HRESULT。
'''         DispCallFunc 自体が失敗した場合はその HRESULT。
'''
''' ? 全パラメータは Variant として渡される。
'''   内部で VarType を判定し、vbLong / vbLongPtr / vbDouble に振り分ける。
'''   それ以外の型は vbLongPtr として扱う (ポインタ前提)。
''' </summary>
Public Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = E_POINTER
        Exit Function
    End If

    Dim pVTable As LongPtr
    CopyMemory pVTable, ByVal pInterface, LenB(pVTable)

    Dim oVft As LongPtr
    oVft = CLngPtr(vtblIndex) * LenB(pVTable)

    ' ── 引数展開 ──
    Dim argc As Long
    If UBound(args) >= LBound(args) Then
        argc = UBound(args) - LBound(args) + 1
    Else
        argc = 0
    End If

    Dim vt() As Integer
    Dim vp() As LongPtr
    Dim vals() As Variant

    If argc > 0 Then
        ReDim vt(0 To argc - 1)
        ReDim vp(0 To argc - 1)
        ReDim vals(0 To argc - 1)

        Dim i As Long
        For i = 0 To argc - 1
            vals(i) = args(LBound(args) + i)

            Select Case VarType(vals(i))
                Case vbLong
                    vt(i) = vbLong
                Case vbLongLong   ' x64 の LongPtr は vbLongLong (=20)
                    vt(i) = vbLongLong
                Case vbDouble
                    vt(i) = vbDouble
                Case vbInteger
                    vals(i) = CLng(CInt(vals(i)))
                    vt(i) = vbLong
                Case vbBoolean
                    vals(i) = CLng(IIf(CBool(vals(i)), 1, 0))
                    vt(i) = vbLong
                Case vbString
                    vals(i) = StrPtr(CStr(vals(i)))
                    vt(i) = vbLongLong
                Case Else
                    vt(i) = vbLongLong
            End Select
            vp(i) = VarPtr(vals(i))
        Next i
    End If

    Dim res As Variant
    Dim hr As Long

    If argc > 0 Then
        hr = DispCallFunc(pInterface, oVft, CC_STDCALL, vbLong, _
            argc, vt(0), vp(0), res)
    Else
        hr = DispCallFunc(pInterface, oVft, CC_STDCALL, vbLong, _
            0, 0, 0, res)
    End If

    If hr <> 0 Then
        If Len(funcName) > 0 Then
            Debug.Print "dcf CALL FAILED: " & funcName & " hr=&H" & Hex(hr)
        End If
        dcf = hr
    Else
        Dim lres As Long: lres = CLng(res)
        If lres <> 0 Then
            If Len(funcName) > 0 Then
                Debug.Print "dcf METHOD FAILED: " & funcName & " result=&H" & Hex(lres)
            End If
        End If
        dcf = lres
    End If
End Function


' ???????????????????????????????????????????????
' Section 2: Bool プロパティ ヘルパー
' ???????????????????????????????????????????????

''' <summary>
''' COM インターフェースの get_XxxBool プロパティを呼び出して Boolean を返す。
''' vtable レイアウト: HRESULT get_Xxx([out, retval] BOOL *value)
''' </summary>
''' <param name="pInterface">COM インターフェースポインタ</param>
''' <param name="vtblIndex">get_ メソッドの vtable インデックス</param>
''' <returns>プロパティ値</returns>
Public Function GetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long) As Boolean

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, "", VarPtr(value))
    If hr = S_OK Then
        GetBoolProperty = (value <> 0)
    End If
End Function

''' <summary>
''' COM インターフェースの put_XxxBool プロパティを呼び出す。
''' vtable レイアウト: HRESULT put_Xxx([in] BOOL value)
''' </summary>
''' <param name="pInterface">COM インターフェースポインタ</param>
''' <param name="vtblIndex">put_ メソッドの vtable インデックス</param>
''' <param name="value">設定する値</param>
Public Sub LetBoolProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal value As Boolean)

    If pInterface = 0 Then Exit Sub

    Dim boolVal As Long
    boolVal = IIf(value, 1, 0)
    dcf pInterface, vtblIndex, "", boolVal
End Sub


' ???????????????????????????????????????????????
' Section 3: Long プロパティ ヘルパー
' ???????????????????????????????????????????????

''' <summary>
''' COM インターフェースの get_XxxLong プロパティを呼び出して Long を返す。
''' </summary>
Public Function GetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long) As Long

    If pInterface = 0 Then Exit Function

    Dim value As Long
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, "", VarPtr(value))
    If hr = S_OK Then GetLongProperty = value
End Function

''' <summary>
''' COM インターフェースの put_XxxLong プロパティを呼び出す。
''' </summary>
Public Sub LetLongProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal value As Long)

    If pInterface = 0 Then Exit Sub
    dcf pInterface, vtblIndex, "", value
End Sub


' ???????????????????????????????????????????????
' Section 4: LongPtr (ポインタ) プロパティ ヘルパー
' ???????????????????????????????????????????????

''' <summary>
''' COM インターフェースの get_XxxPtr プロパティを呼び出して LongPtr を返す。
''' 戻り値のポインタは AddRef 済み。呼び出し元が Release する責任を持つ。
''' </summary>
Public Function GetPtrProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long) As LongPtr

    If pInterface = 0 Then Exit Function

    Dim value As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, "", VarPtr(value))
    If hr = S_OK Then GetPtrProperty = value
End Function


' ???????????????????????????????????????????????
' Section 5: 文字列プロパティ ヘルパー
' ???????????????????????????????????????????????

''' <summary>
''' COM インターフェースの get_XxxString プロパティを呼び出して String を返す。
''' vtable レイアウト: HRESULT get_Xxx([out, retval] LPWSTR *value)
''' 取得した LPWSTR は CoTaskMemFree で解放する。
''' </summary>
Public Function GetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    Optional ByVal funcName As String = "") As String

    If pInterface = 0 Then Exit Function

    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, vtblIndex, funcName, VarPtr(pStr))
    If hr = S_OK And pStr <> 0 Then
        GetStringProperty = PtrToString(pStr)
        CoTaskMemFree pStr
    End If
End Function

''' <summary>
''' COM インターフェースの put_XxxString プロパティを呼び出す。
''' vtable レイアウト: HRESULT put_Xxx([in] LPCWSTR value)
''' </summary>
Public Sub LetStringProperty( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal value As String, _
    Optional ByVal funcName As String = "")

    If pInterface = 0 Then Exit Sub
    dcf pInterface, vtblIndex, funcName, StrPtr(value)
End Sub


' ???????????????????????????????????????????????
' Section 6: ポインタ ⇔ 文字列 変換
' ???????????????????????????????????????????????

''' <summary>
''' LPWSTR (null終端Unicode文字列) を VBA の String に変換する。
''' </summary>
Public Function PtrToString(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long: cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToString = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToString), p, CLngPtr(cch * 2)
End Function


' ???????????????????????????????????????????????
' Section 7: COM Release ヘルパー
' ???????????????????????????????????????????????

''' <summary>
''' COM インターフェースポインタを安全に Release する。
''' 呼び出し後、ByRef で渡されたポインタは 0 にリセットされる。
''' </summary>
Public Sub SafeRelease(ByRef pInterface As LongPtr)
    If pInterface = 0 Then Exit Sub
    dcf pInterface, 2, ""
    pInterface = 0
End Sub

''' <summary>
''' COM インターフェースポインタを安全に AddRef する。
''' </summary>
Public Sub SafeAddRef(ByVal pInterface As LongPtr)
    If pInterface = 0 Then Exit Sub
    dcf pInterface, 1, ""
End Sub


' ???????????????????????????????????????????????
' Section 8: HRESULT ユーティリティ
' ???????????????????????????????????????????????

''' <summary>HRESULT が成功かどうかを判定する</summary>
Public Function SUCCEEDED(ByVal hr As Long) As Boolean
    SUCCEEDED = (hr >= 0)
End Function

''' <summary>HRESULT が失敗かどうかを判定する</summary>
Public Function FAILED(ByVal hr As Long) As Boolean
    FAILED = (hr < 0)
End Function

''' <summary>HRESULT を人間が読める文字列に変換する</summary>
Public Function HResultToString(ByVal hr As Long) As String
    Select Case hr
        Case S_OK:              HResultToString = "S_OK"
        Case S_FALSE:           HResultToString = "S_FALSE"
        Case E_FAIL:            HResultToString = "E_FAIL"
        Case E_POINTER:         HResultToString = "E_POINTER"
        Case E_NOINTERFACE:     HResultToString = "E_NOINTERFACE"
        Case E_INVALIDARG:      HResultToString = "E_INVALIDARG"
        Case E_PENDING:         HResultToString = "E_PENDING"
        Case E_ABORT:           HResultToString = "E_ABORT"
        Case E_ACCESSDENIED:    HResultToString = "E_ACCESSDENIED"
        Case E_OUTOFMEMORY:     HResultToString = "E_OUTOFMEMORY"
        Case E_UNEXPECTED:      HResultToString = "E_UNEXPECTED"
        Case E_NOTIMPL:         HResultToString = "E_NOTIMPL"
        Case Else:              HResultToString = "&H" & Hex(hr)
    End Select
End Function


' ???????????????????????????????????????????????
' Section 9: Variant 配列ユーティリティ
' ???????????????????????????????????????????????

''' <summary>
''' Variant 配列が初期化 (Allocate) 済みかどうかを判定する。
''' ParamArray が空の場合でも安全に使える。
''' </summary>
Public Function IsArrayAllocated(ByRef arr As Variant) As Boolean
    On Error GoTo NotAllocated
    Dim lb As Long: lb = LBound(arr)
    Dim ub As Long: ub = UBound(arr)
    IsArrayAllocated = (ub >= lb)
    Exit Function
NotAllocated:
    IsArrayAllocated = False
End Function



''' <summary>Base64 文字列をバイト配列にデコードする</summary>
Public Function Base64Decode(ByVal base64 As String) As Byte()
    Dim cbBinary As Long
    Dim dwSkip As Long, dwFlags As Long

    ' 1st call: サイズ取得
    If CryptStringToBinaryW(StrPtr(base64), Len(base64), _
        CRYPT_STRING_BASE64, 0, cbBinary, dwSkip, dwFlags) = 0 Then
        Exit Function
    End If

    Dim buf() As Byte
    ReDim buf(0 To cbBinary - 1)

    ' 2nd call: デコード
    CryptStringToBinaryW StrPtr(base64), Len(base64), _
        CRYPT_STRING_BASE64, VarPtr(buf(0)), cbBinary, dwSkip, dwFlags

    Base64Decode = buf
End Function

''' <summary>バイト配列を Base64 文字列にエンコードする</summary>
Public Function Base64Encode(ByRef data() As Byte) As String
    Dim cbData As Long: cbData = UBound(data) - LBound(data) + 1
    Dim cchString As Long

    ' 1st call: サイズ取得
    If CryptBinaryToStringW(VarPtr(data(LBound(data))), cbData, _
        CRYPT_STRING_BASE64, 0, cchString) = 0 Then
        Exit Function
    End If

    Base64Encode = String$(cchString, vbNullChar)

    ' 2nd call: エンコード
    CryptBinaryToStringW VarPtr(data(LBound(data))), cbData, _
        CRYPT_STRING_BASE64, StrPtr(Base64Encode), cchString

    ' 末尾の CRLF を除去
    Base64Encode = Replace(Base64Encode, vbCrLf, "")
    Base64Encode = Replace(Base64Encode, vbNullChar, "")
End Function


' ???????????????????????????????????????????????
' Section 11: JSON 簡易パーサー
' ???????????????????????????????????????????????

''' <summary>
''' JSON 文字列から指定キーの文字列値を取り出す (簡易版)。
''' ネストした JSON やエスケープシーケンスは非対応。
'''
''' 例: JsonGetString("{""name"":""hello""}", "name") → "hello"
''' </summary>
Public Function JsonGetString(ByVal json As String, ByVal Key As String) As String
    Dim searchKey As String
    searchKey = """" & Key & """:"

    Dim pos As Long: pos = InStr(1, json, searchKey, vbTextCompare)
    If pos = 0 Then Exit Function
    pos = pos + Len(searchKey)

    ' 空白スキップ
    Do While pos <= Len(json) And Mid$(json, pos, 1) = " "
        pos = pos + 1
    Loop
    If pos > Len(json) Then Exit Function

    Dim ch As String: ch = Mid$(json, pos, 1)

    If ch = """" Then
        ' 文字列値
        pos = pos + 1
        Dim endPos As Long: endPos = InStr(pos, json, """")
        If endPos = 0 Then Exit Function
        JsonGetString = Mid$(json, pos, endPos - pos)
    Else
        ' 数値・true・false・null
        Dim endPos2 As Long
        endPos2 = pos
        Do While endPos2 <= Len(json)
            ch = Mid$(json, endPos2, 1)
            If ch = "," Or ch = "}" Or ch = "]" Or ch = " " Then Exit Do
            endPos2 = endPos2 + 1
        Loop
        JsonGetString = Mid$(json, pos, endPos2 - pos)
    End If
End Function

''' <summary>
''' JSON 文字列から指定キーの Long 値を取り出す (簡易版)。
''' </summary>
Public Function JsonGetLong(ByVal json As String, ByVal Key As String) As Long
    Dim s As String: s = JsonGetString(json, Key)
    If Len(s) > 0 And IsNumeric(s) Then JsonGetLong = CLng(s)
End Function

''' <summary>
''' JSON 文字列から指定キーの Boolean 値を取り出す (簡易版)。
''' </summary>
Public Function JsonGetBool(ByVal json As String, ByVal Key As String) As Boolean
    Dim s As String: s = LCase$(JsonGetString(json, Key))
    JsonGetBool = (s = "true" Or s = "1")
End Function


' ???????????????????????????????????????????????
' Section 12: デバッグ支援
' ???????????????????????????????????????????????

''' <summary>
''' LongPtr のバイト列を 16 進数文字列としてダンプする。
''' デバッグ用。
''' </summary>
Public Function HexDump(ByVal p As LongPtr, ByVal byteCount As Long) As String
    If p = 0 Then HexDump = "(null)": Exit Function

    Dim buf() As Byte
    ReDim buf(0 To byteCount - 1)
    RtlMoveMemory VarPtr(buf(0)), p, CLngPtr(byteCount)

    Dim sb As String
    Dim i As Long
    For i = 0 To byteCount - 1
        If i > 0 And i Mod 16 = 0 Then sb = sb & vbCrLf
        sb = sb & Right$("0" & Hex(buf(i)), 2) & " "
    Next i
    HexDump = sb
End Function

''' <summary>
''' COM インターフェースの vtable の最初の numEntries 個のアドレスを
''' デバッグ出力する。
''' </summary>
Public Sub DumpVTable(ByVal pInterface As LongPtr, _
    Optional ByVal numEntries As Long = 10, _
    Optional ByVal label As String = "")

    If pInterface = 0 Then
        Debug.Print "DumpVTable: null"
        Exit Sub
    End If

    Dim pVTable As LongPtr
    CopyMemory pVTable, ByVal pInterface, LenB(pVTable)

    If Len(label) > 0 Then Debug.Print "=== VTable: " & label & " ==="
    Debug.Print "  Interface: &H" & Hex(pInterface)
    Debug.Print "  VTable:    &H" & Hex(pVTable)

    Dim i As Long
    For i = 0 To numEntries - 1
        Dim pEntry As LongPtr
        CopyMemory pEntry, ByVal (pVTable + i * LenB(pVTable)), LenB(pEntry)
        Debug.Print "  [" & format(i, "00") & "] &H" & Hex(pEntry)
    Next i
End Sub


