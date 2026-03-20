Attribute VB_Name = "mHelpers"
Option Explicit


' FUNCDESC 構造体 (ITypeInfo::GetFuncDesc で取得)
Private Type FUNCDESC
    memid As Long
    lprgscode As LongPtr
    lprgelemdescParam As LongPtr
    funckind As Long      ' 0=Virtual, 1=PureVirtual...
    invkind As Long       ' 1=Method, 2=PropGet, 4=PropPut
    callconv As Long      ' 4=STDCALL
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer       ' <--- これが VTable のオフセット (バイト単位)
    wReserved1 As Integer
    varkind As Long
    resW32 As Long
End Type

' ITypeInfo 用の構造体定義
Private Type TYPEATTR
    guid(15) As Byte
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Long
    typekind As Long
    cFuncs As Integer
        
    'GetClassMethodPtr で関数の数（cFuncs）を知りたいだけであれば、
    'これ以降のデータは読み飛ばしても良い。
    '（CopyMemory で cFuncs までのサイズ分だけコピーすれば良いため）
'    cVars As Integer
'    cImplTypes As Integer
'    cbSizeVft As Integer
'    cbAlignment As Integer
'    wTypeFlags As Integer
'    wMajorVerNum As Integer
'    wMinorVerNum As Integer
'    tdescAlias As TYPEDESC
'    idldescType As IDLDESC
End Type

' IID_IDispatch の定義 (16バイトのGUID構造体)
Public Type guid
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' IID_IDispatch: {00020400-0000-0000-C000-000000000046}
Public Function IID_IDispatch() As guid
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Function



Public Function IsIDispatchSupported(ByVal pUnk As IUnknown) As Boolean
    
    Dim ppUnk As LongPtr
    Dim iid As guid
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
Public Function DCF_引数2つ_StringとObject(pObj As LongPtr, vIndex As Long, strFuncName As String, str As String, obj As Object) As Long
    Dim hr As Long
    Dim res As Variant
    Dim vObj As Variant

    ' 1. インスタンスを VARIANT 型に格納する
    ' これにより、VBA内部で IDispatch インターフェースとしての正装が整います
    Set vObj = obj

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

Public Function DCF_ハンドラ登録(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, funcPtr As LongPtr) As Long
    
    Dim handler As c4_Handler: Set handler = New c4_Handler
    handler.CreateVTble funcPtr, WB2.ppWebView2
    WB2.Col_Handler.Add handler
    
    Dim pObj As LongPtr, Token As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    args(0) = handler.Pointer
    args(1) = VarPtr(Token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))
    
    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            handler.Token = Token
            RegisterInstance handler.Pointer, WB2
            DCF_ハンドラ登録 = res
        Else
            Debug.Print strFuncName & "_失敗、res:" & res
        End If
    Else
        Debug.Print strFuncName & "_失敗、hr:" & hr
    End If
End Function

Public Function DFC_ハンドラ登録_文字列渡しあり(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, str As String, funcPtr As LongPtr)
    
    Dim handler As c4_Handler: Set handler = New c4_Handler
    handler.CreateVTble funcPtr, WB2.ppWebView2
    WB2.Col_Handler.Add handler
    RegisterInstance handler.Pointer, WB2

    Dim pObj As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    
    args(0) = StrPtr(str)
    args(1) = handler.Pointer
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

Public Function GetClassMethodPtr(ByVal TargetObj As Object, ByVal MethodName As String, ByRef vtbloffset As Long) As LongPtr
    If TargetObj Is Nothing Then Exit Function

    Dim pDisp As LongPtr: pDisp = ObjPtr(TargetObj)
    Dim pTInfo As LongPtr
    Dim hr As Long, res As Variant
    
    ' 1. IDispatch::GetTypeInfo (Index 4) を叩いて ITypeInfo を取得
    ' HRESULT GetTypeInfo(UINT iTInfo, LCID lcid, ITypeInfo** ppTInfo)
    'If DcfCall(pDisp, 4, vbLong, pTInfo, 0&, 0&) <> 0 Then Exit Function
    Dim args(2) As Variant
    args(0) = 0& '[in]  UINT      iTInfo,
    args(1) = 0& '[in]  LCID      lcid
    args(2) = VarPtr(pTInfo) '[out] ITypeInfo **ppTInfo
    
    Dim argTypes(2) As Integer
    argTypes(0) = vbLong
    argTypes(1) = vbLong
    argTypes(2) = vbLongPtr
    
    Dim argPtrs(2) As LongPtr
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))
    argPtrs(2) = VarPtr(args(2))
    
    hr = DispCallFunc(pDisp, 4 * LenB(pDisp), CC_STDCALL, vbLong, 3, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            'Debug.Print pTInfo
        Else
            'Debug.Print "GetTypeInfo Error! res:" & res
        End If
    Else
        'Debug.Print "GetTypeInfo Error! hr:" & hr
    End If
    
    If res <> 0 Then Exit Function
    
    ' 2. ITypeInfo::GetTypeAttr (Index 3)
    ' HRESULT GetTypeAttr([out] TYPEATTR **ppTypeAttr)
    ' 引数は「ポインタのポインタ」1つだけです。
    
    Dim pTypeAttr As LongPtr ' 構造体のアドレスを受け取る変数
    Dim args_t(0) As Variant, argsType_t(0) As Integer, argsPtr_t(0) As LongPtr
    args_t(0) = VarPtr(pTypeAttr)
    argsType_t(0) = vbLongPtr
    argsPtr_t(0) = VarPtr(args_t(0))
    
    ' DispCallFunc の第2引数は Index 3 * 8(または4) です
    hr = DispCallFunc(pTInfo, 3 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_t(0), argsPtr_t(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            'Debug.Print "GetTypeAttr Success! pTypeAttr:" & pTypeAttr
        Else
            'Debug.Print "GetTypeAttr Error! res:" & res
        End If
    Else
        'Debug.Print "GetTypeAttr Error! hr:" & hr
    End If
    
    If res <> 0 Then GoTo 後片付け②
    
    ' --- ここからループ開始 ---
    Dim uTypeAttr As TYPEATTR
    CopyMemory uTypeAttr, ByVal pTypeAttr, LenB(uTypeAttr)
    
    Dim i As Long
    Dim pFuncDesc As LongPtr
    Dim uFuncDesc As FUNCDESC
    Dim bstrName As String
    Dim pBstr As LongPtr
    
    'Debug.Print "関数スキャン開始: " & uTypeAttr.cFuncs & " 個の定義が見つかりました"
    
    For i = 0 To uTypeAttr.cFuncs - 1
    
        ' --- ① GetFuncDesc (Index 5) ---
        ' HRESULT GetFuncDesc([in] UINT index, [out] FUNCDESC** ppFuncDesc)
        Dim args_Gf(1) As Variant, argTypes_Gf(1) As Integer, argPtrs_Gf(1) As LongPtr
        args_Gf(0) = i: args_Gf(1) = VarPtr(pFuncDesc)
        argTypes_Gf(0) = vbLong: argTypes_Gf(1) = vbLongPtr
        argPtrs_Gf(0) = VarPtr(args_Gf(0)): argPtrs_Gf(1) = VarPtr(args_Gf(1))
        
        hr = DispCallFunc(pTInfo, 5 * LenB(pTInfo), CC_STDCALL, vbLong, 2, argTypes_Gf(0), argPtrs_Gf(0), res)
        
        If hr = 0 And res = 0 Then
            CopyMemory uFuncDesc, ByVal pFuncDesc, LenB(uFuncDesc)
            
        ' --- ② GetDocumentation (Index 12) の修正版 ---
        Dim args_Gd(4) As Variant, argTypes_Gd(4) As Integer, argPtrs_Gd(4) As LongPtr
        pBstr = 0
        
        args_Gd(0) = uFuncDesc.memid: argTypes_Gd(0) = vbLong
        args_Gd(1) = VarPtr(pBstr):   argTypes_Gd(1) = vbLongPtr
        
        ' ここからが重要：第3?5引数は「変数のアドレス」ではなく、
        ' 「0 (NULLポインタ)」そのものをスタックに積ませる
        args_Gd(2) = 0:            argTypes_Gd(2) = vbLongPtr
        args_Gd(3) = 0:            argTypes_Gd(3) = vbLongPtr
        args_Gd(4) = 0:            argTypes_Gd(4) = vbLongPtr
        
        ' argPtrs への格納
        argPtrs_Gd(0) = VarPtr(args_Gd(0)) ' [in] memid
        argPtrs_Gd(1) = VarPtr(args_Gd(1)) ' [out] BSTR* (ポインタを書き込んでもらう場所のアドレス)
        argPtrs_Gd(2) = VarPtr(args_Gd(2)) ' [out] NULL
        argPtrs_Gd(3) = VarPtr(args_Gd(3)) ' [out] NULL
        argPtrs_Gd(4) = VarPtr(args_Gd(4)) ' [out] NULL
        
        hr = DispCallFunc(pTInfo, 12 * LenB(pTInfo), CC_STDCALL, vbLong, 5, argTypes_Gd(0), argPtrs_Gd(0), res)
            
            If hr = 0 And res = 0 Then
                'Debug.Print "GetFuncDescメソッド成功"
                
                ' pBstr は BSTRポインタそのものなので、直接 String 変数の「中身」として代入する
                ' VBAの String 型変数 (bstrName) の実体はポインタなので、そこに pBstr を書き込む
                CopyMemory ByVal VarPtr(bstrName), pBstr, LenB(pBstr)
                
                ' これで bstrName に名前が入ります。
                ' ★重要：bstrName は VBAが管理するようになるので、SysFreeString pBstr は「不要」になります。
                ' (VBAが関数の終わりに bstrName を解放するときに一緒に消えるため)
                
                'Debug.Print "Found Method: " & bstrName
                
                ' --- ③ 名前が一致したらポインタ計算 ---
                If LCase$(bstrName) = LCase$(MethodName) Then
                    'Debug.Print "LCase$(bstrName) = LCase$(MethodName)がTrueでした"
                    Dim pVTable As LongPtr, pRealAddr As LongPtr
                    CopyMemory pVTable, ByVal pDisp, LenB(pVTable)
                    CopyMemory pRealAddr, ByVal (pVTable + uFuncDesc.oVft), LenB(pRealAddr)
                    
                    GetClassMethodPtr = pRealAddr
                    vtbloffset = uFuncDesc.oVft
                    'Debug.Print "★発見: " & bstrName & " -> Addr: " & pRealAddr
                Else
                    'Debug.Print "LCase$(bstrName) = LCase$(MethodName)がFalseでした"
                End If
            End If
            
            ' --- ④ GetFuncDesc で確保されたメモリを解放 (Index 20) ---
            ' HRESULT ReleaseFuncDesc([in] FUNCDESC* pFuncDesc)
            Dim args_Rf(0) As Variant, argTypes_Rf(0) As Integer, argPtrs_Rf(0) As LongPtr
            args_Rf(0) = pFuncDesc: argTypes_Rf(0) = vbLongPtr: argPtrs_Rf(0) = VarPtr(args_Rf(0))
            Call DispCallFunc(pTInfo, 20 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argTypes_Rf(0), argPtrs_Rf(0), res)
        End If
        
        ' 目的のポインタが見つかったらループを抜ける（後片付けはループ外で行う）
        If GetClassMethodPtr <> 0 Then
            'Debug.Print "GetClassMethodPtr <> 0だったのでループを抜けます"
            Exit For
        End If
    Next i
    
    
    ' 後片付け①TYPEATTR の解放 (ITypeInfo Index 19)
    ' HRESULT ReleaseTypeAttr([in] TYPEATTR* pTypeAttr)
    ' Index 19
    Dim args_r(0) As Variant, argsType_r(0) As Integer, argsPtr_r(0) As LongPtr
    args_r(0) = pTypeAttr ' ポインタそのものを渡す
    argsType_r(0) = vbLongPtr
    argsPtr_r(0) = VarPtr(args_r(0))

    hr = DispCallFunc(pTInfo, 19 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_r(0), argsPtr_r(0), res)
    If hr = 0 Then
        'Debug.Print "TYPEATTR 解放成功"
    End If
    
後片付け②:
    ' 後片付け②ITypeInfo の解放 (IUnknown Index 2)
    ' IUnknown::Release (Index 2)
    ' 引数なし
    hr = DispCallFunc(pTInfo, 2 * LenB(pTInfo), CC_STDCALL, vbLong, 0, 0, 0, res)
    If hr = 0 Then
        'Debug.Print "ITypeInfo リリース成功"
    End If

End Function
