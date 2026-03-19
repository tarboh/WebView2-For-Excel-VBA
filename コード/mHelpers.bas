Attribute VB_Name = "mHelpers"
Option Explicit

' IID_IDispatch の定義 (16バイトのGUID構造体)
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' IID_IDispatch: {00020400-0000-0000-C000-000000000046}
Public Function IID_IDispatch() As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Function



Public Function IsIDispatchSupported(ByVal pUnk As IUnknown) As Boolean
    
    Dim ppUnk As LongPtr
    Dim iid As GUID
    Dim iidDisp As LongPtr
    Dim hr As Long
    Dim res As Variant
    
    ppUnk = ObjPtr(pUnk)
    iid = IID_IDispatch()
    
    ' Object(IUnknown) から IDispatch を QueryInterface する
    ' IUnknown::QueryInterface は VTable Index 0 です
    ' 引数は2つ
    ' [in]  REFIID riid      -> VarPtr(iidDisp)
    ' [out] void **ppvObject -> VarPtr(pDisp)
    
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
    
    args(0) = VarPtr(iid) ' GUID構造体の場所
    argTypes(0) = vbLongPtr
    
    args(1) = VarPtr(iidDisp)   ' ポインタ変数の場所
    argTypes(1) = vbLongPtr
    
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))
    
    ' Index 0 (QueryInterface) をコール
    hr = DispCallFunc(ppUnk, 0, CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then ' S_OK
        If res = 0 Then
            IsIDispatchSupported = True
            ' QueryInterface で参照カウントが増えるので、Release が必要（これも DispCallFunc 等で）
            ' Call Release(pDisp)
            hr = DispCallFunc(ppUnk, 2 * LenB(ppUnk), CC_STDCALL, vbLong, 0, argTypes(0), argPtrs(0), res)
            If hr = 0 Then
            ' res には Release 後の参照カウントが返ってきますが、通常は無視してOK
                Debug.Print "Release 成功。残りの参照カウント: " & res
            Else
                Debug.Print "Release 失敗: " & hr
            End If
        Else
            IsIDispatchSupported = False
        End If
    Else
        IsIDispatchSupported = False
    End If
End Function

'DispCallFunc引数無しで文字列を取得
Public Function DCF_引数無しで文字列を取得(pObj As LongPtr, vIndex As Long, strFuncName As String) As String
    Dim hr As Long, res As Variant, pStr As LongPtr
    Dim args(0) As Variant: args(0) = VarPtr(pStr)
    Dim argTypes(0) As Integer: argTypes(0) = vbLongPtr
    Dim argPtrs(0) As LongPtr: argPtrs(0) = VarPtr(args(0))
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            If pStr <> 0 Then
                DCF_引数無しで文字列を取得 = PtrToStrW(pStr)
                Call CoTaskMemFree(pStr)
            Else
                Debug.Print strFuncName & "成功、しかしpStr取得失敗、pStr:" & pStr
            End If
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
        End If
    Else
        Debug.Print strFuncName & "_dispcallfunc失敗、hr:" & hr
    End If
End Function
Public Function DCF_引数1つで文字列を渡す(pObj As LongPtr, vIndex As Long, strFuncName As String, ByVal str As String) As Long
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(str)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    Dim hr As Long, res As Variant
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            DCF_引数1つで文字列を渡す = res
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
            DCF_引数1つで文字列を渡す = res
        End If
    Else
        Debug.Print strFuncName & "_dispcallfunc失敗、hr:" & hr
        DCF_引数1つで文字列を渡す = hr
    End If
End Function
Public Function DCF_引数2つ_StringとObject(pObj As LongPtr, vIndex As Long, strFuncName As String, str As String, Obj As Object) As Long
    Dim hr As Long
    Dim res As Variant
    Dim vObj As Variant

    ' 1. インスタンスを VARIANT 型に格納する
    ' これにより、VBA内部で IDispatch インターフェースとしての正装が整います
    Set vObj = Obj

    ' 2. DispCallFunc のための引数準備
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
 
    ' 第1引数: 文字列のポインタ (LPCWSTR)
    args(0) = StrPtr(str)
    argTypes(0) = vbLongPtr ' 64bit: 20(vbLongLong), 32bit: 3(vbLong)
   
    ' 第2引数: VARIANT構造体へのポインタ (VARIANT*)
    ' VarPtr(vObj) で、vObj変数そのもののメモリアドレスを渡します
    args(1) = VarPtr(vObj)
    argTypes(1) = vbLongPtr

    ' 引数ポインタ配列の構築
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(pObj, _
                      vIndex * LenB(pObj), _
                      CC_STDCALL, _
                      vbLong, _
                      2, _
                      argTypes(0), _
                      argPtrs(0), _
                      res)

    ' 4. 結果判定
    If hr = S_OK Then
        If res = S_OK Then
            DCF_引数2つ_StringとObject = res
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
            DCF_引数2つ_StringとObject = res
        End If
    Else
        Debug.Print strFuncName & "_dispcallfunc失敗、hr:" & hr
        DCF_引数2つ_StringとObject = hr
    End If
End Function

Public Function DCF_ハンドラ登録(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, FuncPtr As LongPtr) As Long
    
    Dim Handler As c4_Handler: Set Handler = New c4_Handler
    Handler.CreateVTble FuncPtr, WB2.ppWebView2
    WB2.Col_Handler.Add Handler
    
    Dim pObj As LongPtr, Token As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    args(0) = Handler.Pointer
    args(1) = VarPtr(Token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))
    
    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            Handler.Token = Token
            RegisterInstance Handler.Pointer, WB2
            DCF_ハンドラ登録 = res
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
        End If
    Else
        Debug.Print strFuncName & "_失敗、hr:" & hr
    End If
End Function

Public Function DFC_ハンドラ登録_文字列渡しあり(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, str As String, FuncPtr As LongPtr)
    
    Dim Handler As c4_Handler: Set Handler = New c4_Handler
    Handler.CreateVTble FuncPtr, WB2.ppWebView2
    WB2.Col_Handler.Add Handler
    RegisterInstance Handler.Pointer, WB2

    Dim pObj As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    
    args(0) = StrPtr(str)
    args(1) = Handler.Pointer
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            DFC_ハンドラ登録_文字列渡しあり = res
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
        End If
    Else
        Debug.Print strFuncName & "_失敗、hr:" & hr
    End If
End Function
