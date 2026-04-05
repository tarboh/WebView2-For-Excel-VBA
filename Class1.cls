VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Public Class Class1 （このクラス単独でローダー実行からICoreWebView2のイベントハンドラ登録、Invoke受信まで全部やりたいけど・・・）
Option Explicit

Public WithEvents wv2 As c3_WebView2
Attribute wv2.VB_VarHelpID = -1

Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function KillTimer Lib "user32" (ByVal hWnd As LongPtr, ByVal uIDEvent As LongPtr) As Long

Private m_pStubQI As LongPtr
Private m_pStubAR As LongPtr
Private m_fakeObj_ctrl As LongPtr ' Controller用
Private m_vtbl_ctrl As LongPtr


Private m_lastErrorCode As Long
Private m_pEnv As LongPtr

Private m_pController As LongPtr

Private m_pWebView As LongPtr ' ICoreWebView2 ポインタ

' RECT構造体 (put_Bounds用)
Private Type RECT
    Left   As Long
    Top    As Long
    Right  As Long
    Bottom As Long
End Type


Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)

Private Const MEM_COMMIT  As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private targetHWnd As LongPtr 'フレームのhwnd

Private m_vtbl    As LongPtr
Private m_fakeObj As LongPtr

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cArgs As Long, _
    ByRef rgvt As Integer, _
    ByRef rgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

#If Win64 Then
    Private Const vbLongPtr = vbLongLong
#Else
    Private Const vbLongPtr = vbLong
#End If

Private Const CC_STDCALL As Long = 4 ' Calling convention for the 3rd argument of DispCallFunc (__stdcall)
Private Const S_OK As Long = 0

' FUNCDESC structure (Retrieved via ITypeInfo::GetFuncDesc)
Private Type FUNCDESC
    memid As Long
    lprgscode As LongPtr
    lprgelemdescParam As LongPtr
    funckind As Long      ' 0 = Virtual, 1 = PureVirtual...
    invkind As Long       ' 1 = Method, 2 = PropGet, 4 = PropPut
    callconv As Long      ' 4 = CC_STDCALL
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer       ' <-- VTable offset in bytes
    wReserved1 As Integer
    varkind As Long
    resW32 As Long
End Type

' TYPEATTR structure for ITypeInfo
Private Type TYPEATTR
    GUID(15) As Byte
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Long
    typekind As Long
    cFuncs As Integer     ' Number of functions in the interface
        
    ' NOTE: If you only need the function count (cFuncs) for GetClassMethodPtr,
    ' you can safely ignore the rest of the members and copy only up to cFuncs using CopyMemory.
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

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hWnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Private m_funcPtrs As Object  ' Scripting.Dictionary: 関数名(小文字) → LongPtr
Private m_vtbls(7 To 60)    As LongPtr  ' vTblIndex → vtble+サンク ブロック
Private m_fakeObjs(7 To 60) As LongPtr  ' vTblIndex → fakeObj ポインタ
Private m_tokens(7 To 60)   As Long  ' vTblIndex → 登録解除用トークン

Private m_pStubAddRefRelease As LongPtr  ' AddRef と Release で共用
'Private m_pStubQI            As LongPtr

'Events
Public Event NavigationStarting()       ' 7

Private Function dcf(ptr As LongPtr, vtblIndex As Long, funcName As String, ParamArray args() As Variant) As Long
    
    'Debug.Print "dcf called for " & funcName
    Dim l As Long: l = LBound(args)
    Dim u As Long: u = UBound(args)
    Dim cnt As Long: cnt = u - l + 1
    Dim hr As Long, res As Variant
    Dim args_Type() As Integer
    Dim args_Ptr() As LongPtr
    Dim localVar() As Variant
    ' IMPORTANT: Do NOT use VarPtr(args(i)) directly.
    ' ParamArray elements are temporary Variants managed by the VBA runtime stack.
    ' Their addresses become invalid by the time DispCallFunc internally reads rgpvarg,
    ' causing the COM method to receive garbage values.
    ' Copying into a heap-allocated dynamic array (localArgs) ensures the Variant
    ' addresses remain stable throughout the DispCallFunc call.
    If cnt > 0 Then
        ReDim args_Type(l To u): ReDim args_Ptr(l To u): ReDim localVar(l To u)
        Dim i As Long
        For i = l To u
            localVar(i) = args(i)
            args_Type(i) = VarType(localVar(i))
            args_Ptr(i) = VarPtr(localVar(i))
            'Debug.Print "args(" & i & ")", "Type:" & args_Type(i), "Ptr:" & Hex(args_Ptr(i)),"Value:" & localVar(i)
        Next
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, args_Type(l), args_Ptr(l), res)
    Else
        hr = DispCallFunc(ptr, vtblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, 0, 0, res)
    End If
    If hr = 0 Then
        If res <> 0 Then
            Debug.Print funcName & " failed. res:" & res
        End If
        dcf = res
    Else
        Debug.Print funcName & " failed. hr:" & hr
        dcf = hr
    End If
End Function

Private Function GetClassMethodPtr(ByVal TargetObj As Object, ByVal methodName As String, ByRef vTableOffset As Long) As LongPtr
    If TargetObj Is Nothing Then Exit Function

    Dim pDisp As LongPtr: pDisp = ObjPtr(TargetObj)
    Dim pTInfo As LongPtr
    Dim hr As Long, res As Variant
    
    ' 1. Retrieve ITypeInfo by calling IDispatch::GetTypeInfo (VTable Index 4)
    ' HRESULT GetTypeInfo([in] UINT iTInfo, [in] LCID lcid, [out] ITypeInfo** ppTInfo)
    Dim args(2) As Variant
    args(0) = 0&
    args(1) = 0&
    args(2) = VarPtr(pTInfo)
    
    Dim argTypes(2) As Integer
    argTypes(0) = vbLong
    argTypes(1) = vbLong
    argTypes(2) = vbLongPtr
    
    Dim argPtrs(2) As LongPtr
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))
    argPtrs(2) = VarPtr(args(2))
    
    hr = DispCallFunc(pDisp, 4 * LenB(pDisp), CC_STDCALL, vbLong, 3, argTypes(0), argPtrs(0), res)
    
    If hr <> 0 Or res <> 0 Then Exit Function
    
    ' 2. Retrieve TYPEATTR by calling ITypeInfo::GetTypeAttr (VTable Index 3)
    ' HRESULT GetTypeAttr([out] TYPEATTR** ppTypeAttr)
    Dim pTypeAttr As LongPtr
    Dim args_t(0) As Variant, argsType_t(0) As Integer, argsPtr_t(0) As LongPtr
    args_t(0) = VarPtr(pTypeAttr)
    argsType_t(0) = vbLongPtr
    argsPtr_t(0) = VarPtr(args_t(0))
    
    hr = DispCallFunc(pTInfo, 3 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_t(0), argsPtr_t(0), res)
    
    If hr <> 0 Or res <> 0 Then GoTo CleanUp_ITypeInfo
    
    ' --- Start Scanning Functions ---
    Dim uTypeAttr As TYPEATTR
    CopyMemory uTypeAttr, ByVal pTypeAttr, LenB(uTypeAttr)
    
    Dim i As Long
    Dim pFuncDesc As LongPtr
    Dim uFuncDesc As FUNCDESC
    Dim bstrName As String
    Dim pBstr As LongPtr
    
    For i = 0 To uTypeAttr.cFuncs - 1
    
        ' STEP 1: GetFuncDesc (VTable Index 5)
        ' HRESULT GetFuncDesc([in] UINT index, [out] FUNCDESC** ppFuncDesc)
        Dim args_Gf(1) As Variant, argTypes_Gf(1) As Integer, argPtrs_Gf(1) As LongPtr
        args_Gf(0) = i: args_Gf(1) = VarPtr(pFuncDesc)
        argTypes_Gf(0) = vbLong: argTypes_Gf(1) = vbLongPtr
        argPtrs_Gf(0) = VarPtr(args_Gf(0)): argPtrs_Gf(1) = VarPtr(args_Gf(1))
        
        hr = DispCallFunc(pTInfo, 5 * LenB(pTInfo), CC_STDCALL, vbLong, 2, argTypes_Gf(0), argPtrs_Gf(0), res)
        
        If hr = 0 And res = 0 Then
            CopyMemory uFuncDesc, ByVal pFuncDesc, LenB(uFuncDesc)
            
            ' STEP 2: GetDocumentation (VTable Index 12)
            ' HRESULT GetDocumentation([in] MEMBERID memid, [out] BSTR* pbstrName, ...)
            Dim args_Gd(4) As Variant, argTypes_Gd(4) As Integer, argPtrs_Gd(4) As LongPtr
            pBstr = 0
            
            args_Gd(0) = uFuncDesc.memid: argTypes_Gd(0) = vbLong
            args_Gd(1) = VarPtr(pBstr):    argTypes_Gd(1) = vbLongPtr
            
            ' Passing 0 (NULL pointer) directly to the stack for unused outputs
            args_Gd(2) = 0: argTypes_Gd(2) = vbLongPtr
            args_Gd(3) = 0: argTypes_Gd(3) = vbLongPtr
            args_Gd(4) = 0: argTypes_Gd(4) = vbLongPtr
            
            argPtrs_Gd(0) = VarPtr(args_Gd(0))
            argPtrs_Gd(1) = VarPtr(args_Gd(1))
            argPtrs_Gd(2) = VarPtr(args_Gd(2))
            argPtrs_Gd(3) = VarPtr(args_Gd(3))
            argPtrs_Gd(4) = VarPtr(args_Gd(4))
            
            hr = DispCallFunc(pTInfo, 12 * LenB(pTInfo), CC_STDCALL, vbLong, 5, argTypes_Gd(0), argPtrs_Gd(0), res)
                
            If hr = 0 And res = 0 Then
                ' Inject the BSTR pointer directly into the VBA String variable (it automatically takes ownership)
                CopyMemory ByVal VarPtr(bstrName), pBstr, LenB(pBstr)
                Debug.Print bstrName
                ' STEP 3: Compare names and calculate the VTable pointer
                If LCase$(bstrName) = LCase$(methodName) Then
                    Dim pVTable As LongPtr, pRealAddr As LongPtr
                    CopyMemory pVTable, ByVal pDisp, LenB(pVTable)
                    CopyMemory pRealAddr, ByVal (pVTable + uFuncDesc.oVft), LenB(pRealAddr)
                    
                    GetClassMethodPtr = pRealAddr
                    vTableOffset = uFuncDesc.oVft
                End If
            End If
            
            ' STEP 4: Release FUNCDESC memory (VTable Index 20)
            ' HRESULT ReleaseFuncDesc([in] FUNCDESC* pFuncDesc)
            Dim args_Rf(0) As Variant, argTypes_Rf(0) As Integer, argPtrs_Rf(0) As LongPtr
            args_Rf(0) = pFuncDesc: argTypes_Rf(0) = vbLongPtr: argPtrs_Rf(0) = VarPtr(args_Rf(0))
            Call DispCallFunc(pTInfo, 20 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argTypes_Rf(0), argPtrs_Rf(0), res)
        End If
        
        If GetClassMethodPtr <> 0 Then Exit For
    Next i
    
    ' CleanUp Part 1: Release TYPEATTR (ITypeInfo VTable Index 19)
    Dim args_r(0) As Variant, argsType_r(0) As Integer, argsPtr_r(0) As LongPtr
    args_r(0) = pTypeAttr
    argsType_r(0) = vbLongPtr
    argsPtr_r(0) = VarPtr(args_r(0))

    hr = DispCallFunc(pTInfo, 19 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_r(0), argsPtr_r(0), res)
    
CleanUp_ITypeInfo:
    ' CleanUp Part 2: Release ITypeInfo (IUnknown VTable Index 2)
    hr = DispCallFunc(pTInfo, 2 * LenB(pTInfo), CC_STDCALL, vbLong, 0, 0, 0, res)

End Function


Private Function CreateStubThunk_QI() As LongPtr
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function
    Dim b(7) As Byte
    b(0) = &H49: b(1) = &H89: b(2) = &H8     ' mov [r8], rcx
    b(3) = &H31: b(4) = &HC0                   ' xor eax, eax
    b(5) = &HC3                                 ' ret
    RtlMoveMemory addr, VarPtr(b(0)), 8
    CreateStubThunk_QI = addr
End Function

Private Function CreateStubThunk_AddRefRelease() As LongPtr
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function
    Dim b(7) As Byte
    b(0) = &HB8: b(1) = &H2: b(2) = &H0: b(3) = &H0: b(4) = &H0          ' mov eax, 2
    b(5) = &HC3                                                               ' ret
    RtlMoveMemory addr, VarPtr(b(0)), 8
    CreateStubThunk_AddRefRelease = addr
End Function

''' <summary>
''' WebView2の完了通知やイベントを受け取るためのFake Object（サンク）を生成します。
''' </summary>
''' <param name="methodName">VBAクラス内のコールバックメソッド名</param>
Private Function CreateWebView2Handler(ByVal methodName As String) As LongPtr
    Dim pTarget As LongPtr, offset As Long
    pTarget = GetClassMethodPtr(Me, methodName, offset)
    If pTarget = 0 Then
        Debug.Print "GetClassMethodPtr failed: " & methodName
        Exit Function
    End If

    ' メモリ確保: vtable(32) + thunk(64) = 96bytes
    Dim blockAddr As LongPtr
    blockAddr = VirtualAlloc(0, 96, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If blockAddr = 0 Then Exit Function

    Dim thunkAddr As LongPtr: thunkAddr = blockAddr + 32
    Dim b(127) As Byte
    Dim i As Long

    ' --- 鉄壁のサンクロジック ---
    ' sub rsp, 38h
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HEC: b(i + 3) = &H38: i = i + 4
    ' R9 <- R8 (pEnv/pController), R8 <- RDX (errorCode), RDX <- RCX (Handler_this)
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' RCX <- ObjPtr(Me)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    Dim objPtr_ As LongPtr: objPtr_ = ObjPtr(Me)
    RtlMoveMemory VarPtr(b(i)), VarPtr(objPtr_), 8: i = i + 8
    ' call pTarget
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    RtlMoveMemory VarPtr(b(i)), VarPtr(pTarget), 8: i = i + 8
    b(i) = &HFF: b(i + 1) = &HD0: i = i + 2
    ' xor eax, eax / add rsp, 38h / ret
    b(i) = &H31: b(i + 1) = &HC0: i = i + 2
    b(i) = &H48: b(i + 1) = &H83: b(i + 2) = &HC4: b(i + 3) = &H38: i = i + 4
    b(i) = &HC3: i = i + 1

    RtlMoveMemory thunkAddr, VarPtr(b(0)), i

    ' VTable構築 (QI, AddRef, Release, Invoke)
    RtlMoveMemory blockAddr + 0, VarPtr(m_pStubQI), 8
    RtlMoveMemory blockAddr + 8, VarPtr(m_pStubAR), 8
    RtlMoveMemory blockAddr + 16, VarPtr(m_pStubAR), 8
    RtlMoveMemory blockAddr + 24, VarPtr(thunkAddr), 8

    ' Fake Object (VTableへのポインタを保持する領域)
    Dim pFakeObj As LongPtr
    pFakeObj = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    RtlMoveMemory pFakeObj, VarPtr(blockAddr), 8

    CreateWebView2Handler = pFakeObj
End Function

Private Sub Class_Initialize()
    Call BuildFuncPtrCache
    If m_pStubQI = 0 Then m_pStubQI = CreateStubThunk_QI()
    If m_pStubAddRefRelease = 0 Then m_pStubAddRefRelease = CreateStubThunk_AddRefRelease()
End Sub

Private Sub Class_Terminate()
    If m_vtbl <> 0 Then VirtualFree m_vtbl, 0, MEM_RELEASE
    If m_fakeObj <> 0 Then VirtualFree m_fakeObj, 0, MEM_RELEASE
    If m_pStubQI <> 0 Then VirtualFree m_pStubQI, 0, MEM_RELEASE
    If m_pStubAR <> 0 Then VirtualFree m_pStubAR, 0, MEM_RELEASE
End Sub


Public Sub CreateWebView2Environment(ByVal FrameHWnd As LongPtr)
    
    targetHWnd = FrameHWnd
    
    ' 1. スタブ（QI/AddRef/Release）の初期化（一度だけ）
    If m_pStubQI = 0 Then m_pStubQI = CreateStubThunk_QI()
    If m_pStubAR = 0 Then m_pStubAR = CreateStubThunk_AddRefRelease()

    ' 2. ハンドラ（サンク）の作成
    ' ※ m_fakeObj はメンバ変数として保持
    m_fakeObj = CreateWebView2Handler("Invoke_EnvironmentCompleted")
    If m_fakeObj = 0 Then
        Debug.Print "Failed to create Environment Handler"
        Exit Sub
    End If

    ' 3. 環境作成の実行
    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, m_fakeObj)
    
    If hr <> S_OK Then
        Debug.Print "CreateCoreWebView2EnvironmentWithOptions failed. hr: " & hr
    Else
        Debug.Print "Environment creation started. Waiting for callback..."
    End If
End Sub

Public Sub Invoke_EnvironmentCompleted( _
    ByVal pHandlerThis As LongPtr, _
    ByVal errorCode As LongPtr, _
    ByVal pEnvironment As LongPtr _
)
    If errorCode <> 0 Then
        Debug.Print "Environment Creation Failed. HRESULT: &H" & Hex(errorCode)
        Exit Sub
    End If

    m_pEnv = pEnvironment
    Debug.Print "--- Environment Ready ---"
    Debug.Print "pEnv: &H" & Hex(m_pEnv)

    ' 1ミリ秒後に StartControllerCreation を実行するようにタイマーをセット
    ' (実際には次のメッセージループで実行される)
    ' 第4引数に AddressOf ではなく、自分自身のメソッドを叩くための仕組みが必要ですが、
    ' 面倒ならここで直接 Controller 作成に進んでも、今の「鉄壁サンク」なら耐えられます。
    
    ' 一旦、安全策をとらずに直接次のステップを呼んでみましょう。
    ' もし不安定なら Timer に切り替えます。
    Call StartControllerCreation
End Sub
Private Sub StartControllerCreation()
    ' 1. Controller完了用の第2サンクを作成
    m_fakeObj_ctrl = CreateWebView2Handler("Invoke_ControllerCompleted")
    
    ' 2. UserFormのhWndを取得 (mHelpersにFindWindow等があると想定)
    ' ※ 暫定的に GetActiveWindow や FindWindow で取得
'    Dim targetHWnd As LongPtr
'    targetHWnd = GetUserFormHwnd()
    
    '甘いぞ、targetHWndはフォーム側であらかじめパブリック変数にセット済みだ！
    
    Debug.Print "Target HWND: &H" & Hex(targetHWnd)

    ' 3. CreateCoreWebView2Controller を呼び出す
    ' ICoreWebView2Environment VTable Index: 3
    ' HRESULT CreateCoreWebView2Controller([in] HWND parentWindow, [in] ICoreWebView2CreateCoreWebView2ControllerCompletedHandler* handler)
    Dim hr As Long
    hr = dcf(m_pEnv, 3, "CreateCoreWebView2Controller", targetHWnd, m_fakeObj_ctrl)

    If hr = 0 Then
        Debug.Print "Controller creation requested..."
    Else
        Debug.Print "Failed to call CreateCoreWebView2Controller. hr: &H" & Hex(hr)
    End If
End Sub

Public Sub Invoke_ControllerCompleted( _
    ByVal pHandlerThis As LongPtr, _
    ByVal errorCode As LongPtr, _
    ByVal pController As LongPtr _
)
    If errorCode <> 0 Or pController = 0 Then
        Debug.Print "Controller Error: " & Hex(errorCode)
        Exit Sub
    End If

    ' --- 1. 参照カウントを増やす (既存コードの CallAddRef 相当) ---
    ' WebView2側で破棄されないよう、VBA側でしっかり掴んでおく
    m_pController = pController
    ' ※ AddRef (Index 1) を叩くのが正解
    Call dcf(m_pController, 1, "AddRef")

    ' --- 2. 可視化フラグを立てる (put_IsVisible: Index 4) ---
    ' これを True にしないと、OSレベルで「描画プロセス」が走りません。
    ' 引数は BOOL (Long) の 1
    Call dcf(m_pController, 4, "put_IsVisible", 1&)

    ' --- 3. サイズ設定 (SetBoundsAndZoomFactor: Index 11) ---
    ' put_Bounds (Index 6) は構造体の値渡しでコケやすいので、
    ' 引数が明確な Index 11 を使います。
    ' 引数: RECT(Left, Top, Right, Bottom), double(Zoom)
    ' ※x64のdcfで構造体渡しが厳しい場合、一旦既存コードのように Win32 API で強制するのも手です。
    
    Dim rc As RECT
    rc.Right = 800: rc.Bottom = 600
    
    ' 暫定的に既存コードの「Win32 APIで強制」をここでも発動させる
    ' (targetHWnd は Invoke_EnvironmentCompleted で取ったものを使用)
    Dim childHwnd As LongPtr
    childHwnd = FindWindowEx(targetHWnd, 0, "Chrome_WidgetWin_0", vbNullString)
    If childHwnd <> 0 Then
        MoveWindow childHwnd, 0, 0, 800, 600, 1
        Debug.Print "Forced visibility via Win32. HWND: " & Hex(childHwnd)
    End If

    ' --- 4. WebView2本体を取得 ---
    Call dcf(m_pController, 25, "get_CoreWebView2", VarPtr(m_pWebView))
    'これでモジュールレベル変数のm_pWebViewにICoreWebView2のポインタが入る
    
    If m_pWebView = 0 Then
        Debug.Print "IcoreWebView2インターフェースポインタの取得に失敗しました。"
        Exit Sub
    End If
    
    Set wv2 = New c3_WebView2
    wv2.ppWebView2 = m_pWebView
    
    ' --- 5. イベントハンドラの登録 ---
    Call Registar_IcoreWebView2_Handlers
    
    wv2.NavigateAsync "https://www.google.com"
    
'    ' --- 6. 初期ページ表示テスト ---
'    If m_pWebView <> 0 Then
'        Debug.Print "Navigating to Google..."
'        Call dcf(m_pWebView, 5, "Navigate", StrPtr("https://www.google.com"))
'    End If
    
    DoEvents
End Sub

Private Function RegisterHandler( _
    ByVal vtblIndex As Long, _
    ByVal methodName As String) As Long

    ' vtble 生成
    m_vtbls(vtblIndex) = CreateHandlerVTable(methodName)
    If m_vtbls(vtblIndex) = 0 Then Exit Function

    ' fakeObj 生成
    m_fakeObjs(vtblIndex) = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If m_fakeObjs(vtblIndex) = 0 Then Exit Function
    RtlMoveMemory m_fakeObjs(vtblIndex), VarPtr(m_vtbls(vtblIndex)), 8

    ' WebView2 に登録
    RegisterHandler = dcf(m_pWebView, vtblIndex, methodName, _
        m_fakeObjs(vtblIndex), VarPtr(m_tokens(vtblIndex)))

End Function


' 指定された関数名を元に該当関数のポインタをこのクラス内から検索し、
' WebView2側からの呼び出しに対し引数"Me"を追加して
' 指定された関数ポインタにジャンプさせるサンクを作成し、
' そのサンクをVtbleIndex:3(Invoke)として用意した「ハンドラのVtble」を作成し、
' そのハンドラのポインタを返す関数。
' この戻り値がそのまま「ハンドラオブジェクト」として使えるので、
' それをICoreWebView2::add_xxxx (xxxxはイベント名) の引数として渡すことで
' イベントハンドラを登録することができる。

' vTable of Handler
' 0|Invoke  | m_pStubQI
' 1|AddRef  | m_pStubAddRefRelease
' 2|Release | m_pStubAddRefRelease
' 3|Invoke  | thunkAddr (Created in this function)

Private Function CreateHandlerVTable(ByVal methodName As String) As LongPtr

    Dim pFunc As LongPtr, offset As Long
    pFunc = m_funcPtrs(methodName)
    If pFunc = 0 Then Exit Function

    ' vtble(32bytes) + thunk(48bytes) = 80bytes
    Dim blockAddr As LongPtr
    blockAddr = VirtualAlloc(0, 80, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If blockAddr = 0 Then Exit Function

    Dim vtblAddr  As LongPtr: vtblAddr = blockAddr
    Dim thunkAddr As LongPtr: thunkAddr = blockAddr + 32

    Dim b(47) As Byte
    Dim i As Long

    ' mov r9, r8
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, ObjPtr(Me)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    Dim objPtr_3 As LongPtr: objPtr_3 = ObjPtr(Me)
    RtlMoveMemory VarPtr(b(i)), VarPtr(objPtr_3), 8: i = i + 8
    ' mov rax, pFunc
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    RtlMoveMemory VarPtr(b(i)), VarPtr(pFunc), 8: i = i + 8
    ' jmp rax
    b(i) = &HFF: b(i + 1) = &HE0

    RtlMoveMemory thunkAddr, VarPtr(b(0)), 48

    RtlMoveMemory vtblAddr + (0 * 8), VarPtr(m_pStubQI), 8
    RtlMoveMemory vtblAddr + (1 * 8), VarPtr(m_pStubAddRefRelease), 8
    RtlMoveMemory vtblAddr + (2 * 8), VarPtr(m_pStubAddRefRelease), 8
    RtlMoveMemory vtblAddr + (3 * 8), VarPtr(thunkAddr), 8

    CreateHandlerVTable = vtblAddr

End Function

Public Sub Registar_IcoreWebView2_Handlers()
    'Call BuildFuncPtrCache
    'Call add_NavigationStarting

End Sub


'7
'virtual HRESULT STDMETHODCALLTYPE add_NavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_NavigationStarting() As Long
    add_NavigationStarting = RegisterHandler(7, "Invoke_NavigationStarting")
End Function
Public Function Invoke_NavigationStarting(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent NavigationStarting
    Invoke_NavigationStarting = 0
End Function

Public Sub BuildFuncPtrCache()
    ' 辞書を初期化する前に既存のものをクリア
    Set m_funcPtrs = Nothing
    Set m_funcPtrs = CreateObject("Scripting.Dictionary")
    m_funcPtrs.CompareMode = 1  ' vbTextCompare: 大文字小文字を無視

    Dim pDisp As LongPtr: pDisp = ObjPtr(Me)
    Dim pTInfo As LongPtr
    Dim hr As Long, res As Variant

    ' --- IDispatch::GetTypeInfo ---
    Dim args(2) As Variant, argTypes(2) As Integer, argPtrs(2) As LongPtr
    args(0) = 0&: args(1) = 0&: args(2) = VarPtr(pTInfo)
    argTypes(0) = vbLong: argTypes(1) = vbLong: argTypes(2) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1)): argPtrs(2) = VarPtr(args(2))
    hr = DispCallFunc(pDisp, 4 * LenB(pDisp), CC_STDCALL, vbLong, 3, argTypes(0), argPtrs(0), res)
    If hr <> 0 Or res <> 0 Then Exit Sub

    ' --- ITypeInfo::GetTypeAttr ---
    Dim pTypeAttr As LongPtr
    Dim args_t(0) As Variant, argsType_t(0) As Integer, argsPtr_t(0) As LongPtr
    args_t(0) = VarPtr(pTypeAttr): argsType_t(0) = vbLongPtr: argsPtr_t(0) = VarPtr(args_t(0))
    hr = DispCallFunc(pTInfo, 3 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_t(0), argsPtr_t(0), res)
    If hr <> 0 Or res <> 0 Then GoTo CleanUp_ITypeInfo

    Dim uTypeAttr As TYPEATTR
    CopyMemory uTypeAttr, ByVal pTypeAttr, LenB(uTypeAttr)

    Dim pVTable As LongPtr
    CopyMemory pVTable, ByVal pDisp, LenB(pVTable)

    ' --- 全関数をループして辞書に登録 ---
    Dim i As Long
    For i = 0 To uTypeAttr.cFuncs - 1

        Dim pFuncDesc As LongPtr
        Dim args_Gf(1) As Variant, argTypes_Gf(1) As Integer, argPtrs_Gf(1) As LongPtr
        args_Gf(0) = i: args_Gf(1) = VarPtr(pFuncDesc)
        argTypes_Gf(0) = vbLong: argTypes_Gf(1) = vbLongPtr
        argPtrs_Gf(0) = VarPtr(args_Gf(0)): argPtrs_Gf(1) = VarPtr(args_Gf(1))
        hr = DispCallFunc(pTInfo, 5 * LenB(pTInfo), CC_STDCALL, vbLong, 2, argTypes_Gf(0), argPtrs_Gf(0), res)
        If hr <> 0 Or res <> 0 Then GoTo NextFunc

        Dim uFuncDesc As FUNCDESC
        CopyMemory uFuncDesc, ByVal pFuncDesc, LenB(uFuncDesc)

        ' GetDocumentation で関数名取得
        Dim pBstr As LongPtr: pBstr = 0
        Dim args_Gd(4) As Variant, argTypes_Gd(4) As Integer, argPtrs_Gd(4) As LongPtr
        args_Gd(0) = uFuncDesc.memid: argTypes_Gd(0) = vbLong
        args_Gd(1) = VarPtr(pBstr): argTypes_Gd(1) = vbLongPtr
        args_Gd(2) = 0: argTypes_Gd(2) = vbLongPtr
        args_Gd(3) = 0: argTypes_Gd(3) = vbLongPtr
        args_Gd(4) = 0: argTypes_Gd(4) = vbLongPtr
        argPtrs_Gd(0) = VarPtr(args_Gd(0)): argPtrs_Gd(1) = VarPtr(args_Gd(1))
        argPtrs_Gd(2) = VarPtr(args_Gd(2)): argPtrs_Gd(3) = VarPtr(args_Gd(3))
        argPtrs_Gd(4) = VarPtr(args_Gd(4))
        hr = DispCallFunc(pTInfo, 12 * LenB(pTInfo), CC_STDCALL, vbLong, 5, argTypes_Gd(0), argPtrs_Gd(0), res)

        If hr = 0 And res = 0 Then
            Dim bstrName As String
            CopyMemory ByVal VarPtr(bstrName), pBstr, LenB(pBstr)

            ' vtble から実アドレスを取得して辞書に登録
            Dim pRealAddr As LongPtr
            CopyMemory pRealAddr, ByVal (pVTable + uFuncDesc.oVft), LenB(pRealAddr)
            m_funcPtrs(bstrName) = pRealAddr
            Debug.Print "登録", i, bstrName, pRealAddr
        End If

        ' ReleaseFuncDesc
        Dim args_Rf(0) As Variant, argTypes_Rf(0) As Integer, argPtrs_Rf(0) As LongPtr
        args_Rf(0) = pFuncDesc: argTypes_Rf(0) = vbLongPtr: argPtrs_Rf(0) = VarPtr(args_Rf(0))
        Call DispCallFunc(pTInfo, 20 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argTypes_Rf(0), argPtrs_Rf(0), res)

NextFunc:
    Next i

    ' ReleaseTypeAttr
    Dim args_r(0) As Variant, argsType_r(0) As Integer, argsPtr_r(0) As LongPtr
    args_r(0) = pTypeAttr: argsType_r(0) = vbLongPtr: argsPtr_r(0) = VarPtr(args_r(0))
    Call DispCallFunc(pTInfo, 19 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_r(0), argsPtr_r(0), res)

CleanUp_ITypeInfo:
    Call DispCallFunc(pTInfo, 2 * LenB(pTInfo), CC_STDCALL, vbLong, 0, 0, 0, res)
End Sub

' Helper function to receive AddressOf as LongPtr
Public Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' Dummy Sub to prevent VBA compiler optimization / address loss
Public Sub RegisterNavigationCompleted_()
    Static vTable As LongPtr
    'vTable = GetAddr(AddressOf Handler_QueryInterface)
End Sub

Private Sub wv2_NavigationStarting()
    Debug.Print "NavigationStarting"
End Sub

Public Sub CloseWebView2()
    Dim hr As Long
    Dim result As Variant

    If m_pController = 0 Then Exit Sub

    hr = DispCallFunc(m_pController, 24 * LenB(m_pController), CC_STDCALL, vbLong, 0, 0, 0, result)

    If hr = 0 Then
        Debug.Print "WebView2 Controller Closed."
    Else
        Debug.Print "Close Failed. hr: " & hr
    End If

    Set wv2 = Nothing
    m_pController = 0
End Sub
