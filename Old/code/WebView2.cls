VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ???????????????????????????????????????????????????????????????????????
' WebView2 Wrapper Class (x64 VBA) - Complete Edition
' ???????????????????????????????????????????????????????????????????????
'
' 必要な外部モジュール:
'   - mWebView2Api    : Win32 API宣言, 定数, 列挙型, TYPEATTR, FUNCDESC
'   - mWebView2Helper : dcf(), GetBoolProperty(), LetBoolProperty()
'
' ? 重要な制約:
'   (1) 64ビット VBA 専用
'   (2) ObjPtr(Me) がサンクにハードコードされるため、
'       VBA IDE の停止ボタンや End ステートメントは使用禁止。
'       必ず Finalize → Set obj = Nothing の順で破棄すること。
'   (3) jmp テールコールを使用しているため、
'       COM コールバックの明示的パラメータは最大3個まで。
'       (this + Me挿入後で4レジスタ: RCX, RDX, R8, R9)
'
' ???????????????????????????????????????????????????????????????????????

#If Win64 Then
#Else
    #Error "このクラスは 64 ビット VBA (x64) が必要です"
#End If

Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
    ByVal dest As LongPtr, ByVal src As LongPtr, ByVal Length As LongPtr)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" ( _
    ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, _
    ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" ( _
    ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExW" ( _
    ByVal hWndParent As LongPtr, ByVal hwndChildAfter As LongPtr, _
    ByVal lpszClass As LongPtr, ByVal lpszWindow As LongPtr) As LongPtr
Private Declare PtrSafe Function MoveWindow Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)
Private Declare PtrSafe Sub SysFreeString Lib "oleaut32" (ByVal bstr As LongPtr)
Private Declare PtrSafe Function SysStringLen Lib "oleaut32" (ByVal bstr As LongPtr) As Long
Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" ( _
    ByVal browserExecutableFolder As LongPtr, ByVal userDataFolder As LongPtr, _
    ByVal environmentOptions As LongPtr, ByVal environmentCreatedHandler As LongPtr) As Long
Private Declare PtrSafe Function SHCreateStreamOnFileW Lib "shlwapi" ( _
    ByVal pszFile As LongPtr, ByVal grfMode As Long, ByRef ppstm As LongPtr) As Long
Private Declare PtrSafe Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" ( _
    ByVal pszString As LongPtr, ByVal cchString As Long, ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, ByRef pdwFlags As Long) As Long

Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const MEM_RELEASE As Long = &H8000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const CC_STDCALL As Long = 4
Private Const S_OK As Long = 0
Private Const E_NOINTERFACE As Long = &H80004002
Private Const E_PENDING As Long = &H8000000A
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const STGM_CREATE As Long = &H1000
Private Const STGM_READWRITE As Long = &H2
Private Const STGM_SHARE_DENY_NONE As Long = &H40
Private Const CRYPT_STRING_BASE64 As Long = 1

' ???????????????????????????????????????????????
' Section 1: Event 宣言
' ???????????????????????????????????????????????

' --- ICoreWebView2Controller Events ---
Public Event ControllerZoomFactorChanged()
Public Event ControllerMoveFocusRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event ControllerGotFocus()
Public Event ControllerLostFocus()
Public Event ControllerAcceleratorKeyPressed(ByVal sender As LongPtr, ByVal args As LongPtr)

' --- ICoreWebView2 Events ---
Public Event NavigationStarting(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event ContentLoading(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event SourceChanged(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event HistoryChanged()
Public Event NavigationCompleted(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event FrameNavigationStarting(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event FrameNavigationCompleted(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event ScriptDialogOpening(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event PermissionRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event ProcessFailed(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event WebMessageReceived( _
    ByVal Source As String, ByVal jsonMessage As String, ByVal webMessage As String)
Public Event NewWindowRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event DocumentTitleChanged()
Public Event ContainsFullScreenElementChanged()
Public Event WebResourceRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event WindowCloseRequested(ByVal sender As LongPtr, ByVal args As LongPtr)

' --- Completion / Protocol Events ---
Public Event AddScriptCompleted(ByVal scriptId As String, ByVal javascript As String)
Public Event AddScriptFailed(ByVal javascript As String, ByVal errorCode As Long)
Public Event ReceiveScriptResult(ByVal jsonResult As String)
Public Event CapturePreviewCompleted(ByVal errorCode As Long)
Public Event CDPMethodCompleted( _
    ByVal requestId As Long, ByVal methodName As String, _
    ByVal errorCode As Long, ByVal result As String)
Public Event DevToolsProtocolEventReceived( _
    ByVal eventName As String, ByVal paramJson As String)


' ???????????????????????????????????????????????
' Section 2: 内部型定義
' ???????????????????????????????????????????????

''' <summary>
''' 統一ハンドラエントリ。
''' イベント購読 (add_/remove_) と完了コールバックの両方を管理する。
''' </summary>
Private Type HandlerEntry
    Key As String              ' ユーザー指定または自動生成のキー
    vtblBlock As LongPtr       ' VirtualAlloc: vtable(32) + thunk(48) = 80 bytes
    fakeObj As LongPtr         ' VirtualAlloc: 8 bytes (COM Object)

    ' --- イベント購読用 (isEvent=True のみ有効) ---
    isEvent As Boolean         ' True: add_/remove_ ペア
    pInterface As LongPtr      ' remove_ 呼び出し先 COM インターフェース
    removeIndex As Long        ' remove_ の vtable インデックス
    token As LongPtr           ' EventRegistrationToken (8 bytes)
    pExtraRelease As LongPtr   ' 追加で Release が必要な COM ポインタ

    active As Boolean          ' 有効フラグ
End Type

''' <summary>CDP メソッド呼び出し追跡用</summary>
Private Type CDPRequestEntry
    requestId As Long
    methodName As String
    vtblBlock As LongPtr
    fakeObj As LongPtr
End Type

''' <summary>登録済みスクリプト情報</summary>
Private Type ScriptEntry
    scriptId As String
    javascript As String
End Type


' ???????????????????????????????????????????????
' Section 3: メンバ変数
' ???????????????????????????????????????????????

' --- COM インターフェースポインタ ---
Private m_pEnv As LongPtr
Private m_pController As LongPtr
Private m_pWebView As LongPtr
Private m_pSettings As LongPtr

' --- ターゲットウィンドウ ---
Private m_targetFrame As MSForms.Frame
Private m_targetHWnd As LongPtr

' --- 共有スタブ (全ハンドラで共用、各1つだけ) ---
Private m_pStubQI As LongPtr
Private m_pStubAddRefRelease As LongPtr

' --- 関数ポインタキャッシュ ---
Private m_funcPtrs As Object  ' Dictionary: methodName -> pFunc (LongPtr)

' --- 統一ハンドラ管理 ---
Private m_handlers As Object        ' Dictionary: key(String) -> index(Long)
Private m_fakeObjToKey As Object    ' Dictionary: CStr(fakeObj) -> key(String)
Private m_handlerData() As HandlerEntry
Private m_handlerCount As Long

' --- CDP メソッド呼び出し ---
Private m_CDPRequests() As CDPRequestEntry
Private m_CDPRequestCount As Long
Private m_CDPNextId As Long

' --- スクリプト管理 ---
Private m_Scripts() As ScriptEntry
Private m_ScriptCount As Long
Private m_PendingScript As String
Private m_IsScriptPending As Boolean

' --- ExecuteScript 状態 ---
Private m_IsExecuteScriptPending As Boolean
Private m_IsScriptRunning As Boolean
Public LastScriptResult As String

' --- ナビゲーション状態 ---
Private m_IsNavigating As Boolean

' --- キャプチャ状態 ---
Private m_pCaptureStream As LongPtr

' --- 初期化済みフラグ ---
Private m_Initialized As Boolean


' ???????????????????????????????????????????????
' Section 4: ライフサイクル
' ???????????????????????????????????????????????

Private Sub Class_Initialize()
    Set m_handlers = CreateObject("Scripting.Dictionary")
    m_handlers.CompareMode = 1  ' vbTextCompare

    Set m_fakeObjToKey = CreateObject("Scripting.Dictionary")

    Call BuildFuncPtrCache

    m_pStubQI = CreateStubThunk_QI()
    m_pStubAddRefRelease = CreateStubThunk_AddRefRelease()

    m_Initialized = True
End Sub

Private Sub Class_Terminate()
    Call Finalize
End Sub


' ???????????????????????????????????????????????
' Section 5: 共有スタブ生成
' ???????????????????????????????????????????????

''' <summary>
''' QueryInterface スタブ: 全 IID に対して E_NOINTERFACE を返す。
''' WebView2 は IUnknown::QI を呼ぶが、Fake Object なので常に拒否。
''' </summary>
Private Function CreateStubThunk_QI() As LongPtr
    ' mov eax, E_NOINTERFACE / ret
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 16, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function

    Dim b(7) As Byte
    b(0) = &HB8  ' mov eax, imm32
    Dim hr As Long: hr = E_NOINTERFACE
    CopyMemory b(1), hr, 4
    b(5) = &HC3  ' ret
    RtlMoveMemory addr, VarPtr(b(0)), 6

    CreateStubThunk_QI = addr
End Function

''' <summary>
''' AddRef / Release スタブ: 常に 1 を返す (参照カウント固定)。
''' </summary>
Private Function CreateStubThunk_AddRefRelease() As LongPtr
    ' mov eax, 1 / ret
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 16, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function

    Dim b(7) As Byte
    b(0) = &HB8: b(1) = &H1: b(2) = &H0: b(3) = &H0: b(4) = &H0  ' mov eax, 1
    b(5) = &HC3  ' ret
    RtlMoveMemory addr, VarPtr(b(0)), 6

    CreateStubThunk_AddRefRelease = addr
End Function


' ???????????????????????????????????????????????
' Section 6: 関数ポインタキャッシュ
' ???????????????????????????????????????????????

''' <summary>
''' このクラスの全 Public メソッドのアドレスを ITypeInfo 経由で取得し、
''' Dictionary にキャッシュする。サンク生成時にメソッド名から
''' 直接関数ポインタを引くために使用。
''' </summary>
Public Sub BuildFuncPtrCache()
    Set m_funcPtrs = Nothing
    Set m_funcPtrs = CreateObject("Scripting.Dictionary")
    m_funcPtrs.CompareMode = 1

    Dim pDisp As LongPtr: pDisp = ObjPtr(Me)
    Dim pTInfo As LongPtr
    Dim hr As Long, res As Variant

    ' --- IDispatch::GetTypeInfo (vtable index 4) ---
    hr = CallCOM3(pDisp, 4, 0&, 0&, VarPtr(pTInfo))
    If hr <> 0 Then Exit Sub

    ' --- ITypeInfo::GetTypeAttr (vtable index 3) ---
    Dim pTypeAttr As LongPtr
    hr = CallCOM1(pTInfo, 3, VarPtr(pTypeAttr))
    If hr <> 0 Then GoTo Cleanup

    Dim uTypeAttr As TYPEATTR
    CopyMemory uTypeAttr, ByVal pTypeAttr, LenB(uTypeAttr)

    Dim pVTable As LongPtr
    CopyMemory pVTable, ByVal pDisp, LenB(pVTable)

    Dim i As Long
    For i = 0 To uTypeAttr.cFuncs - 1
        ' --- ITypeInfo::GetFuncDesc (vtable index 5) ---
        Dim pFuncDesc As LongPtr
        hr = CallCOM2(pTInfo, 5, CLng(i), VarPtr(pFuncDesc))
        If hr <> 0 Then GoTo NextFunc

        Dim uFuncDesc As FUNCDESC
        CopyMemory uFuncDesc, ByVal pFuncDesc, LenB(uFuncDesc)

        ' --- ITypeInfo::GetDocumentation (vtable index 12) ---
        Dim pBstr As LongPtr: pBstr = 0
        hr = CallCOM5(pTInfo, 12, uFuncDesc.memid, VarPtr(pBstr), 0, 0, 0)

        If hr = 0 And pBstr <> 0 Then
            Dim funcName As String
            funcName = PtrToStrW(pBstr)
            SysFreeString pBstr
            pBstr = 0

            Dim pRealAddr As LongPtr
            CopyMemory pRealAddr, ByVal (pVTable + uFuncDesc.oVft), LenB(pRealAddr)
            m_funcPtrs(funcName) = pRealAddr
        End If

        ' --- ITypeInfo::ReleaseFuncDesc (vtable index 20) ---
        CallCOM1V pTInfo, 20, pFuncDesc
NextFunc:
    Next i

    ' --- ITypeInfo::ReleaseTypeAttr (vtable index 19) ---
    CallCOM1V pTInfo, 19, pTypeAttr

Cleanup:
    ' --- ITypeInfo::Release (vtable index 2) ---
    CallCOM0V pTInfo, 2
End Sub

' --- DispCallFunc ラッパー群 (BuildFuncPtrCache 専用) ---

Private Function CallCOM0V(ByVal p As LongPtr, ByVal idx As Long) As Long
    Dim res As Variant
    CallCOM0V = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 0, 0, 0, res)
End Function

Private Function CallCOM1(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr) As Long
    Dim res As Variant
    Dim v(0) As Variant: v(0) = a1
    Dim t(0) As Integer: t(0) = vbLongPtr
    Dim vp(0) As LongPtr: vp(0) = VarPtr(v(0))
    CallCOM1 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), res)
    If CallCOM1 = 0 Then CallCOM1 = CLng(res)
End Function

Private Sub CallCOM1V(ByVal p As LongPtr, ByVal idx As Long, ByVal a1 As LongPtr)
    Dim res As Variant
    Dim v(0) As Variant: v(0) = a1
    Dim t(0) As Integer: t(0) = vbLongPtr
    Dim vp(0) As LongPtr: vp(0) = VarPtr(v(0))
    DispCallFunc p, idx * LenB(p), CC_STDCALL, vbLong, 1, t(0), vp(0), res
End Sub

Private Function CallCOM2(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr) As Long
    Dim res As Variant
    Dim v(1) As Variant: v(0) = a1: v(1) = a2
    Dim t(1) As Integer: t(0) = vbLong: t(1) = vbLongPtr
    Dim vp(1) As LongPtr: vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1))
    CallCOM2 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 2, t(0), vp(0), res)
    If CallCOM2 = 0 Then CallCOM2 = CLng(res)
End Function

Private Function CallCOM3(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As Long, ByVal a3 As LongPtr) As Long
    Dim res As Variant
    Dim v(2) As Variant: v(0) = a1: v(1) = a2: v(2) = a3
    Dim t(2) As Integer: t(0) = vbLong: t(1) = vbLong: t(2) = vbLongPtr
    Dim vp(2) As LongPtr: vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    CallCOM3 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 3, t(0), vp(0), res)
    If CallCOM3 = 0 Then CallCOM3 = CLng(res)
End Function

Private Function CallCOM5(ByVal p As LongPtr, ByVal idx As Long, _
    ByVal a1 As Long, ByVal a2 As LongPtr, _
    ByVal a3 As LongPtr, ByVal a4 As LongPtr, ByVal a5 As LongPtr) As Long
    Dim res As Variant
    Dim v(4) As Variant: v(0) = a1: v(1) = a2: v(2) = a3: v(3) = a4: v(4) = a5
    Dim t(4) As Integer
    t(0) = vbLong: t(1) = vbLongPtr: t(2) = vbLongPtr: t(3) = vbLongPtr: t(4) = vbLongPtr
    Dim vp(4) As LongPtr
    vp(0) = VarPtr(v(0)): vp(1) = VarPtr(v(1)): vp(2) = VarPtr(v(2))
    vp(3) = VarPtr(v(3)): vp(4) = VarPtr(v(4))
    CallCOM5 = DispCallFunc(p, idx * LenB(p), CC_STDCALL, vbLong, 5, t(0), vp(0), res)
    If CallCOM5 = 0 Then CallCOM5 = CLng(res)
End Function


' ???????????????????????????????????????????????
' Section 7: ハンドラ VTable + サンク生成
' ???????????????????????????????????????????????

''' <summary>
''' 指定メソッド名の関数ポインタを取得し、
''' COM コールバック用の VTable + サンクを生成する。
'''
''' VTable レイアウト (各 8 bytes):
'''   [0] QueryInterface   → m_pStubQI
'''   [1] AddRef           → m_pStubAddRefRelease
'''   [2] Release          → m_pStubAddRefRelease
'''   [3] Invoke           → thunkAddr
'''
''' サンクの動作:
'''   COM から (this, arg1, arg2, arg3) で呼ばれた時、
'''   引数をシフトして (ObjPtr(Me), this, arg1, arg2) に変換し、
'''   VBA メソッドへ jmp する。
'''
''' 制限: 明示的パラメータは最大3個まで (4レジスタ制約)。
''' </summary>
Private Function CreateHandlerVTable(ByVal methodName As String) As LongPtr

    If Not m_funcPtrs.Exists(methodName) Then
        Debug.Print "CreateHandlerVTable: メソッド未発見 - " & methodName
        Exit Function
    End If

    Dim pFunc As LongPtr
    pFunc = m_funcPtrs(methodName)
    If pFunc = 0 Then Exit Function

    ' メモリ確保: vtable(32) + thunk(48) = 80 bytes
    Dim blockAddr As LongPtr
    blockAddr = VirtualAlloc(0, 80, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If blockAddr = 0 Then Exit Function

    Dim vtblAddr As LongPtr:  vtblAddr = blockAddr
    Dim thunkAddr As LongPtr: thunkAddr = blockAddr + 32

    ' --- サンク生成 (x64 machine code) ---
    Dim b(47) As Byte, i As Long
    ' mov r9, r8       (arg2 → slot3)
    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx      (arg1 → slot2)
    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx     (this → slot1)
    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
    ' mov rcx, imm64   (ObjPtr(Me) → slot0)
    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
    Dim mePtr As LongPtr: mePtr = ObjPtr(Me)
    CopyMemory b(i), mePtr, 8: i = i + 8
    ' mov rax, imm64   (pFunc)
    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
    CopyMemory b(i), pFunc, 8: i = i + 8
    ' jmp rax
    b(i) = &HFF: b(i + 1) = &HE0: i = i + 2

    RtlMoveMemory thunkAddr, VarPtr(b(0)), CLngPtr(i)

    ' --- VTable 構築 ---
    RtlMoveMemory vtblAddr + 0, VarPtr(m_pStubQI), 8
    RtlMoveMemory vtblAddr + 8, VarPtr(m_pStubAddRefRelease), 8
    RtlMoveMemory vtblAddr + 16, VarPtr(m_pStubAddRefRelease), 8
    RtlMoveMemory vtblAddr + 24, VarPtr(thunkAddr), 8

    CreateHandlerVTable = vtblAddr
End Function

''' <summary>
''' VTable ポインタから Fake COM Object (8 bytes) を生成する。
''' 戻り値がそのまま COM ハンドラオブジェクトとして使える。
''' </summary>
Private Function AllocFakeObj(ByVal vtblAddr As LongPtr) As LongPtr
    Dim p As LongPtr
    p = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If p = 0 Then Exit Function
    RtlMoveMemory p, VarPtr(vtblAddr), 8
    AllocFakeObj = p
End Function

''' <summary>ハンドラメモリブロックを解放する</summary>
Private Sub FreeHandlerMemory(ByRef vtblBlock As LongPtr, ByRef fakeObj As LongPtr)
    If fakeObj <> 0 Then VirtualFree fakeObj, 0, MEM_RELEASE: fakeObj = 0
    If vtblBlock <> 0 Then VirtualFree vtblBlock, 0, MEM_RELEASE: vtblBlock = 0
End Sub


' ???????????????????????????????????????????????
' Section 8: 統一イベント登録システム
' ???????????????????????????????????????????????
'
' 使い方:
'   登録: key = AddEventHandler(pInterface, addIdx, removeIdx, "Invoke_XXX", "MyKey")
'   解除: RemoveEventHandler "MyKey"
'   確認: If HasEventHandler("MyKey") Then ...
'
' 完了コールバック (add_/remove_ なし):
'   fakeObj = EnsureCallbackHandler("__exec_script", "Invoke_ExecuteScript")
'

''' <summary>
''' イベントハンドラを登録する。
''' </summary>
''' <param name="pInterface">COMインターフェースポインタ</param>
''' <param name="addIndex">add_xxx の vtable インデックス</param>
''' <param name="removeIndex">remove_xxx の vtable インデックス</param>
''' <param name="callbackMethod">VBAコールバックメソッド名</param>
''' <param name="key">管理キー (空の場合は callbackMethod を使用)</param>
''' <returns>登録キー (失敗時は空文字列)</returns>
Public Function AddEventHandler( _
    ByVal pInterface As LongPtr, _
    ByVal addIndex As Long, _
    ByVal removeIndex As Long, _
    ByVal callbackMethod As String, _
    Optional ByVal Key As String = "") As String

    If pInterface = 0 Then
        Debug.Print "AddEventHandler: インターフェースが null です"
        Exit Function
    End If

    If Key = "" Then Key = callbackMethod

    ' 重複チェック
    If m_handlers.Exists(Key) Then
        Debug.Print "AddEventHandler: キー重複 - " & Key
        Exit Function
    End If

    ' ハンドラ生成
    Dim vtbl As LongPtr: vtbl = CreateHandlerVTable(callbackMethod)
    If vtbl = 0 Then Exit Function
    Dim fakeObj As LongPtr: fakeObj = AllocFakeObj(vtbl)
    If fakeObj = 0 Then
        VirtualFree vtbl, 0, MEM_RELEASE
        Exit Function
    End If

    ' COM に登録
    Dim token As LongPtr
    Dim hr As Long
    hr = dcf(pInterface, addIndex, callbackMethod, fakeObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "AddEventHandler: 登録失敗 - " & Key & " hr=&H" & Hex(hr)
        FreeHandlerMemory vtbl, fakeObj
        Exit Function
    End If

    ' エントリ保存
    m_handlerCount = m_handlerCount + 1
    ReDim Preserve m_handlerData(0 To m_handlerCount - 1)
    Dim idx As Long: idx = m_handlerCount - 1

    With m_handlerData(idx)
        .Key = Key
        .vtblBlock = vtbl
        .fakeObj = fakeObj
        .isEvent = True
        .pInterface = pInterface
        .removeIndex = removeIndex
        .token = token
        .pExtraRelease = 0
        .active = True
    End With

    m_handlers(Key) = idx
    m_fakeObjToKey(CStr(fakeObj)) = Key

    Debug.Print "Event Registered: " & Key & " on &H" & Hex(pInterface)
    AddEventHandler = Key
End Function

''' <summary>
''' キーを指定してイベントハンドラを解除・メモリ解放する。
''' </summary>
Public Function RemoveEventHandler(ByVal Key As String) As Long
    If Not m_handlers.Exists(Key) Then
        Debug.Print "RemoveEventHandler: 未登録 - " & Key
        RemoveEventHandler = -1
        Exit Function
    End If

    Dim idx As Long: idx = m_handlers(Key)

    With m_handlerData(idx)
        If Not .active Then Exit Function

        ' COM の remove_ 呼び出し
        If .isEvent And .pInterface <> 0 Then
            Dim hr As Long
            hr = dcf(.pInterface, .removeIndex, "remove_" & Key, .token)
            RemoveEventHandler = hr
            If hr <> 0 Then
                Debug.Print "RemoveEventHandler: remove 失敗 - " & Key & " hr=&H" & Hex(hr)
            End If
        End If

        ' 追加 COM ポインタの Release
        If .pExtraRelease <> 0 Then
            dcf .pExtraRelease, 2, "Release_extra"
            .pExtraRelease = 0
        End If

        ' 逆引き辞書から削除
        Dim fakeObjKey As String: fakeObjKey = CStr(.fakeObj)
        If m_fakeObjToKey.Exists(fakeObjKey) Then m_fakeObjToKey.Remove fakeObjKey

        ' メモリ解放
        FreeHandlerMemory .vtblBlock, .fakeObj
        .active = False
        .pInterface = 0
        .token = 0
    End With

    m_handlers.Remove Key
    Debug.Print "Event Removed: " & Key
End Function

''' <summary>登録済みかどうかを確認する</summary>
Public Function HasEventHandler(ByVal Key As String) As Boolean
    HasEventHandler = m_handlers.Exists(Key)
End Function

''' <summary>
''' 再利用可能な完了コールバックハンドラを取得する (add_/remove_ 不要)。
''' 初回呼び出し時に生成し、以降はキャッシュを返す。
''' </summary>
Private Function EnsureCallbackHandler( _
    ByVal Key As String, ByVal callbackMethod As String) As LongPtr

    ' 既存があればそれを返す
    If m_handlers.Exists(Key) Then
        Dim existIdx As Long: existIdx = m_handlers(Key)
        If m_handlerData(existIdx).active Then
            EnsureCallbackHandler = m_handlerData(existIdx).fakeObj
            Exit Function
        End If
    End If

    ' 新規生成
    Dim vtbl As LongPtr: vtbl = CreateHandlerVTable(callbackMethod)
    If vtbl = 0 Then Exit Function
    Dim fakeObj As LongPtr: fakeObj = AllocFakeObj(vtbl)
    If fakeObj = 0 Then
        VirtualFree vtbl, 0, MEM_RELEASE
        Exit Function
    End If

    m_handlerCount = m_handlerCount + 1
    ReDim Preserve m_handlerData(0 To m_handlerCount - 1)
    Dim idx As Long: idx = m_handlerCount - 1

    With m_handlerData(idx)
        .Key = Key
        .vtblBlock = vtbl
        .fakeObj = fakeObj
        .isEvent = False
        .active = True
    End With

    m_handlers(Key) = idx
    m_fakeObjToKey(CStr(fakeObj)) = Key

    EnsureCallbackHandler = fakeObj
End Function

''' <summary>fakeObj アドレスからキーを逆引きする</summary>
Private Function LookupKeyByFakeObj(ByVal fakeObj As LongPtr) As String
    Dim k As String: k = CStr(fakeObj)
    If m_fakeObjToKey.Exists(k) Then
        LookupKeyByFakeObj = m_fakeObjToKey(k)
    End If
End Function


' ???????????????????????????????????????????????
' Section 9: WebView2 ライフサイクル
' ???????????????????????????????????????????????

Public Sub CreateWebView2Environment(ByVal TargetFrame As MSForms.Frame)

    Set m_targetFrame = TargetFrame
    Dim TargetCtrl As MSForms.Control
    Set TargetCtrl = m_targetFrame
    m_targetHWnd = TargetCtrl.[_GethWnd]

    ' 環境作成完了用コールバック
    Dim fakeObj As LongPtr
    fakeObj = EnsureCallbackHandler("__env_completed", "Invoke_EnvironmentCompleted")
    If fakeObj = 0 Then
        Debug.Print "Failed to create Environment Handler"
        Exit Sub
    End If

    Dim userDataPath As String: userDataPath = "C:\Temp\VBA_WebView2"
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, fakeObj)

    If hr <> S_OK Then
        Debug.Print "CreateCoreWebView2EnvironmentWithOptions failed. hr=&H" & Hex(hr)
    Else
        Debug.Print "Environment creation started..."
    End If
End Sub



'MIDL_INTERFACE ("b96d755e-0319-4e92-a296-23436f46a1fc")
'ICoreWebView2Environment:  Public IUnknown
'{
'public:
'3
'    virtual HRESULT STDMETHODCALLTYPE CreateCoreWebView2Controller(
'        HWND parentWindow,
'        ICoreWebView2CreateCoreWebView2ControllerCompletedHandler *handler) = 0;
Private Function Environment_CreateCoreWebView2Controller(ByVal parentWindow As LongPtr)
    Call StartControllerCreation
    'm_fakeObj_ctrl = CreateWebView2Handler("Invoke_ControllerCompleted")
    Debug.Print "Target HWND: &H" & Hex(targetHWnd)
    Dim hr As Long
    hr = dcf(m_pEnv, 3, "CreateCoreWebView2Controller", targetHWnd, m_fakeObj_ctrl)
    If hr = 0 Then
        Debug.Print "Controller creation requested..."
    Else
        Debug.Print "Failed to call CreateCoreWebView2Controller. hr: &H" & Hex(hr)
    End If
End Function
'4
'    virtual HRESULT STDMETHODCALLTYPE CreateWebResourceResponse(
'        /* [in] */ IStream *content,
'        /* [in] */ int statusCode,
'        /* [in] */ LPCWSTR reasonPhrase,
'        /* [in] */ LPCWSTR headers,
'        /* [retval][out] */ ICoreWebView2WebResourceResponse **response) = 0;
Public Function Environment_CreateWebResourceResponse(ByVal pContent As LongPtr, ByVal statuscode As Long, ByVal resasonPhrase As String, ByVal headers As String)
    Dim pResponse As LongPtr
    Dim hr As Long
    hr = dcf(m_pEnv, 4, "CreateWebResourceResponse", pContent, statuscode, StrPtr(resasonPhrase), StrPtr(headers), VarPtr(pResponse))
    
End Function

'5
'    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_BrowserVersionString(
'        /* [retval][out] */ LPWSTR *versionInfo) = 0;
Public Property Get Environment_BrowserVersionString() As String
    Dim versionInfo As LongPtr
    Dim hr As Long
    hr = dcf(m_pEnv, 5, "Environment_get_BrowserVersionString", VarPtr(versionInfo))
    Environment_BrowserVersionString = PtrToStrW(versionInfo)
End Property
'6
'    virtual HRESULT STDMETHODCALLTYPE add_NewBrowserVersionAvailable(
'        /* [in] */ ICoreWebView2NewBrowserVersionAvailableEventHandler *eventHandler,
'        /* [out] */ EventRegistrationToken *token) = 0;
Public Function Environment_add_NewBrowserVersionAvailable() As Long
    Environment_add_NewBrowserVersionAvailable = RegisterHandlerEx(m_pEnv, 6, "Invoke_Environment_NewBrowserVersionAvailable")
End Function
Public Function Invoke_Environment_NewBrowserVersionAvailable(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent EnvironmentNewBrowserVersionAvailable
    Invoke_Environment_NewBrowserVersionAvailable = 0
End Function
'7
'    virtual HRESULT STDMETHODCALLTYPE remove_NewBrowserVersionAvailable(
'        /* [in] */ EventRegistrationToken token) = 0;
'
Public Function remove_Environment_NewBrowserVersionAvailable(token As Long) As Long
    remove_Environment_NewBrowserVersionAvailable = remove_Handler(Me, 7, m_tokens(6), "remove_Environment_NewBrowserVersionAvailable")
End Function
'};





Public Sub Invoke_EnvironmentCompleted( _
    ByVal pHandlerThis As LongPtr, _
    ByVal errorCode As LongPtr, _
    ByVal pEnvironment As LongPtr)

    If errorCode <> 0 Then
        Debug.Print "Environment Failed. HRESULT=&H" & Hex(errorCode)
        Exit Sub
    End If

    m_pEnv = pEnvironment
    ' AddRef して VBA 側で保持
    dcf m_pEnv, 1, "Env.AddRef"
    Debug.Print "Environment Ready: &H" & Hex(m_pEnv)

    Call StartControllerCreation
End Sub

Private Sub StartControllerCreation()
    Dim fakeObj As LongPtr
    fakeObj = EnsureCallbackHandler("__ctrl_completed", "Invoke_ControllerCompleted")
    If fakeObj = 0 Then Exit Sub

    Debug.Print "Target HWND: &H" & Hex(m_targetHWnd)

    Dim hr As Long
    hr = dcf(m_pEnv, 3, "CreateCoreWebView2Controller", m_targetHWnd, fakeObj)

    If hr = 0 Then
        Debug.Print "Controller creation requested..."
    Else
        Debug.Print "Controller creation failed. hr=&H" & Hex(hr)
    End If
End Sub

Public Sub Invoke_ControllerCompleted( _
    ByVal pHandlerThis As LongPtr, _
    ByVal errorCode As LongPtr, _
    ByVal pController As LongPtr)

    If errorCode <> 0 Or pController = 0 Then
        Debug.Print "Controller Error: &H" & Hex(errorCode)
        Exit Sub
    End If

    ' --- 1. Controller を保持 ---
    m_pController = pController
    dcf m_pController, 1, "Controller.AddRef"

    ' --- 2. 可視化 (put_IsVisible: Index 4) ---
    dcf m_pController, 4, "put_IsVisible", 1&

    ' --- 3. サイズ設定 (put_Bounds: Index 6) ---
    Dim l(3) As Long
    l(0) = 0: l(1) = 0
    l(2) = PtsToPx(m_targetFrame.InsideWidth, False)
    l(3) = PtsToPx(m_targetFrame.InsideHeight, True)
    dcf m_pController, 6, "put_Bounds", VarPtr(l(0))

    ' --- 4. WebView2 本体を取得 ---
    Dim pWV As LongPtr
    dcf m_pController, 25, "get_CoreWebView2", VarPtr(pWV)
    m_pWebView = pWV
    If m_pWebView = 0 Then
        Debug.Print "ICoreWebView2 取得失敗"
        Exit Sub
    End If
    ' get_ は AddRef 済みポインタを返すので追加の AddRef は不要

    ' --- 5. Settings を取得 ---
    dcf m_pWebView, 3, "get_Settings", VarPtr(m_pSettings)

    ' --- 6. イベントハンドラ一括登録 ---
    Call RegisterDefaultHandlers

    ' --- 7. テスト表示 ---
    Debug.Print "Navigating to Google..."
    dcf m_pWebView, 5, "Navigate", StrPtr("https://www.google.com")

    DoEvents
End Sub

























' ???????????????????????????????????????????????
' Section 10: ICoreWebView2Controller
' ???????????????????????????????????????????????

' MIDL_INTERFACE("4d00c0d1-9434-4eb6-8078-8697a560334f")
' VTable Layout (IUnknown=0,1,2):
'  3: get_IsVisible        4: put_IsVisible
'  5: get_Bounds            6: put_Bounds
'  7: get_ZoomFactor        8: put_ZoomFactor
'  9: add_ZoomFactorChanged 10: remove_ZoomFactorChanged
' 11: SetBoundsAndZoomFactor
' 12: MoveFocus
' 13: add_MoveFocusRequested 14: remove_MoveFocusRequested
' 15: add_GotFocus           16: remove_GotFocus
' 17: add_LostFocus          18: remove_LostFocus
' 19: add_AcceleratorKeyPressed 20: remove_AcceleratorKeyPressed
' 21: get_ParentWindow       22: put_ParentWindow
' 23: NotifyParentWindowPositionChanged
' 24: Close
' 25: get_CoreWebView2

' --- Properties ---

Public Property Get Controller_IsVisible() As Boolean
    Controller_IsVisible = GetBoolProperty(m_pController, 3)
End Property
Public Property Let Controller_IsVisible(ByVal value As Boolean)
    LetBoolProperty m_pController, 4, value
End Property

Public Function Controller_get_Bounds(l() As Long) As Long
    Controller_get_Bounds = dcf(m_pController, 5, "get_Bounds", VarPtr(l(0)))
End Function
Public Function Controller_put_Bounds(l() As Long) As Long
    Controller_put_Bounds = dcf(m_pController, 6, "put_Bounds", VarPtr(l(0)))
End Function

Public Property Get Controller_ZoomFactor() As Double
    Dim zf As Double
    dcf m_pController, 7, "get_ZoomFactor", VarPtr(zf)
    Controller_ZoomFactor = zf
End Property
Public Property Let Controller_ZoomFactor(ByVal zf As Double)
    dcf m_pController, 8, "put_ZoomFactor", zf
End Property

Public Property Get Controller_ParentWindow() As LongPtr
    Dim pw As LongPtr
    dcf m_pController, 21, "get_ParentWindow", VarPtr(pw)
    Controller_ParentWindow = pw
End Property
Public Property Let Controller_ParentWindow(ByVal pw As LongPtr)
    dcf m_pController, 22, "put_ParentWindow", pw
End Property

' --- Methods ---

Public Function Controller_SetBoundsAndZoomFactor( _
    ByVal Left As Long, ByVal Top As Long, _
    ByVal Right As Long, ByVal Bottom As Long, _
    ByVal zoomFactor As Double) As Long
    Dim bounds(3) As Long
    bounds(0) = Left: bounds(1) = Top: bounds(2) = Right: bounds(3) = Bottom
    Controller_SetBoundsAndZoomFactor = dcf(m_pController, 11, _
        "SetBoundsAndZoomFactor", VarPtr(bounds(0)), zoomFactor)
End Function

Public Function Controller_MoveFocus(ByVal reason As COREWEBVIEW2_MOVE_FOCUS_REASON) As Long
    Controller_MoveFocus = dcf(m_pController, 12, "MoveFocus", reason)
End Function

Public Sub Controller_NotifyParentWindowPositionChanged()
    dcf m_pController, 23, "NotifyParentWindowPositionChanged"
End Sub

Public Sub Controller_Close()
    dcf m_pController, 24, "Close"
End Sub

Public Function Controller_get_CoreWebView2() As LongPtr
    Dim p As LongPtr
    dcf m_pController, 25, "get_CoreWebView2", VarPtr(p)
    Controller_get_CoreWebView2 = p
End Function

' --- Controller Events ---

Public Function Controller_add_ZoomFactorChanged( _
    Optional ByVal Key As String = "Ctrl.ZoomFactorChanged") As String
    Controller_add_ZoomFactorChanged = _
        AddEventHandler(m_pController, 9, 10, "Invoke_Controller_ZoomFactorChanged", Key)
End Function
Public Function Invoke_Controller_ZoomFactorChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ControllerZoomFactorChanged
    Invoke_Controller_ZoomFactorChanged = 0
End Function

Public Function Controller_add_MoveFocusRequested( _
    Optional ByVal Key As String = "Ctrl.MoveFocusRequested") As String
    Controller_add_MoveFocusRequested = _
        AddEventHandler(m_pController, 13, 14, "Invoke_Controller_MoveFocusRequested", Key)
End Function
Public Function Invoke_Controller_MoveFocusRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ControllerMoveFocusRequested(sender, args)
    Invoke_Controller_MoveFocusRequested = 0
End Function

Public Function Controller_add_GotFocus( _
    Optional ByVal Key As String = "Ctrl.GotFocus") As String
    Controller_add_GotFocus = _
        AddEventHandler(m_pController, 15, 16, "Invoke_Controller_GotFocus", Key)
End Function
Public Function Invoke_Controller_GotFocus( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ControllerGotFocus
    Invoke_Controller_GotFocus = 0
End Function

Public Function Controller_add_LostFocus( _
    Optional ByVal Key As String = "Ctrl.LostFocus") As String
    Controller_add_LostFocus = _
        AddEventHandler(m_pController, 17, 18, "Invoke_Controller_LostFocus", Key)
End Function
Public Function Invoke_Controller_LostFocus( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ControllerLostFocus
    Invoke_Controller_LostFocus = 0
End Function

' ★ 修正: インデックス 19/20 (旧コードは誤って 17 を使用)
Public Function Controller_add_AcceleratorKeyPressed( _
    Optional ByVal Key As String = "Ctrl.AcceleratorKeyPressed") As String
    Controller_add_AcceleratorKeyPressed = _
        AddEventHandler(m_pController, 19, 20, "Invoke_Controller_AcceleratorKeyPressed", Key)
End Function
Public Function Invoke_Controller_AcceleratorKeyPressed( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ControllerAcceleratorKeyPressed(sender, args)
    Invoke_Controller_AcceleratorKeyPressed = 0
End Function

' --- 個別 remove ラッパー (任意のキーで解除可能) ---
Public Function Controller_remove_ZoomFactorChanged( _
    Optional ByVal Key As String = "Ctrl.ZoomFactorChanged") As Long
    Controller_remove_ZoomFactorChanged = RemoveEventHandler(Key)
End Function
Public Function Controller_remove_MoveFocusRequested( _
    Optional ByVal Key As String = "Ctrl.MoveFocusRequested") As Long
    Controller_remove_MoveFocusRequested = RemoveEventHandler(Key)
End Function
Public Function Controller_remove_GotFocus( _
    Optional ByVal Key As String = "Ctrl.GotFocus") As Long
    Controller_remove_GotFocus = RemoveEventHandler(Key)
End Function
Public Function Controller_remove_LostFocus( _
    Optional ByVal Key As String = "Ctrl.LostFocus") As Long
    Controller_remove_LostFocus = RemoveEventHandler(Key)
End Function
Public Function Controller_remove_AcceleratorKeyPressed( _
    Optional ByVal Key As String = "Ctrl.AcceleratorKeyPressed") As Long
    Controller_remove_AcceleratorKeyPressed = RemoveEventHandler(Key)
End Function


' ???????????????????????????????????????????????
' Section 11: ICoreWebView2 Settings
' ???????????????????????????????????????????????

' MIDL_INTERFACE("e562e4f0-d7fa-43ac-8d71-c05150499f00")
' ICoreWebView2Settings : IUnknown
'  3: get_IsScriptEnabled              4: put_IsScriptEnabled
'  5: get_IsWebMessageEnabled          6: put_IsWebMessageEnabled
'  7: get_AreDefaultScriptDialogsEnabled  8: put_AreDefaultScriptDialogsEnabled
'  9: get_IsStatusBarEnabled          10: put_IsStatusBarEnabled
' 11: get_AreDevToolsEnabled          12: put_AreDevToolsEnabled
' 13: get_AreDefaultContextMenusEnabled 14: put_AreDefaultContextMenusEnabled
' 15: get_AreHostObjectsAllowed       16: put_AreHostObjectsAllowed
' 17: get_IsZoomControlEnabled        18: put_IsZoomControlEnabled
' 19: get_IsBuiltInErrorPageEnabled   20: put_IsBuiltInErrorPageEnabled

Public Property Get IsScriptEnabled() As Boolean
    IsScriptEnabled = GetBoolProperty(m_pSettings, 3)
End Property
Public Property Let IsScriptEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 4, value
End Property

Public Property Get IsWebMessageEnabled() As Boolean
    IsWebMessageEnabled = GetBoolProperty(m_pSettings, 5)
End Property
Public Property Let IsWebMessageEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 6, value
End Property

Public Property Get AreDefaultScriptDialogsEnabled() As Boolean
    AreDefaultScriptDialogsEnabled = GetBoolProperty(m_pSettings, 7)
End Property
Public Property Let AreDefaultScriptDialogsEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 8, value
End Property

Public Property Get IsStatusBarEnabled() As Boolean
    IsStatusBarEnabled = GetBoolProperty(m_pSettings, 9)
End Property
Public Property Let IsStatusBarEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 10, value
End Property

Public Property Get AreDevToolsEnabled() As Boolean
    AreDevToolsEnabled = GetBoolProperty(m_pSettings, 11)
End Property
Public Property Let AreDevToolsEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 12, value
End Property

Public Property Get AreDefaultContextMenusEnabled() As Boolean
    AreDefaultContextMenusEnabled = GetBoolProperty(m_pSettings, 13)
End Property
Public Property Let AreDefaultContextMenusEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 14, value
End Property

Public Property Get AreHostObjectsAllowed() As Boolean
    AreHostObjectsAllowed = GetBoolProperty(m_pSettings, 15)
End Property
Public Property Let AreHostObjectsAllowed(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 16, value
End Property

Public Property Get IsZoomControlEnabled() As Boolean
    IsZoomControlEnabled = GetBoolProperty(m_pSettings, 17)
End Property
Public Property Let IsZoomControlEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 18, value
End Property

Public Property Get IsBuiltInErrorPageEnabled() As Boolean
    IsBuiltInErrorPageEnabled = GetBoolProperty(m_pSettings, 19)
End Property
Public Property Let IsBuiltInErrorPageEnabled(ByVal value As Boolean)
    LetBoolProperty m_pSettings, 20, value
End Property


' ???????????????????????????????????????????????
' Section 12: ICoreWebView2 Properties & Methods
' ???????????????????????????????????????????????

' MIDL_INTERFACE("76eceacb-0462-4d94-ac83-423a6793775e")
' ICoreWebView2 : IUnknown
'  3: get_Settings
'  4: get_Source
'  5: Navigate
'  6: NavigateToString
'  7/8:   add/remove_NavigationStarting
'  9/10:  add/remove_ContentLoading
' 11/12:  add/remove_SourceChanged
' 13/14:  add/remove_HistoryChanged
' 15/16:  add/remove_NavigationCompleted
' 17/18:  add/remove_FrameNavigationStarting
' 19/20:  add/remove_FrameNavigationCompleted
' 21/22:  add/remove_ScriptDialogOpening
' 23/24:  add/remove_PermissionRequested
' 25/26:  add/remove_ProcessFailed
' 27:     AddScriptToExecuteOnDocumentCreated
' 28:     RemoveScriptToExecuteOnDocumentCreated
' 29:     ExecuteScript
' 30:     CapturePreview
' 31:     Reload
' 32:     PostWebMessageAsJson
' 33:     PostWebMessageAsString
' 34/35:  add/remove_WebMessageReceived
' 36:     CallDevToolsProtocolMethod
' 37:     get_BrowserProcessId
' 38:     get_CanGoBack
' 39:     get_CanGoForward
' 40:     GoBack
' 41:     GoForward
' 42:     GetDevToolsProtocolEventReceiver
' 43:     Stop
' 44/45:  add/remove_NewWindowRequested
' 46/47:  add/remove_DocumentTitleChanged
' 48:     get_DocumentTitle
' 49:     AddHostObjectToScript
' 50:     RemoveHostObjectFromScript
' 51:     OpenDevToolsWindow
' 52/53:  add/remove_ContainsFullScreenElementChanged
' 54:     get_ContainsFullScreenElement
' 55/56:  add/remove_WebResourceRequested
' 57:     AddWebResourceRequestedFilter
' 58:     RemoveWebResourceRequestedFilter
' 59/60:  add/remove_WindowCloseRequested

' --- Properties ---

Public Property Get Source() As String
    Dim pStr As LongPtr
    dcf m_pWebView, 4, "get_Source", VarPtr(pStr)
    If pStr <> 0 Then
        Source = PtrToStrW(pStr)
        CoTaskMemFree pStr
    End If
End Property

Public Property Get BrowserProcessId() As Long
    Dim value As Long
    dcf m_pWebView, 37, "get_BrowserProcessId", VarPtr(value)
    BrowserProcessId = value
End Property

Public Property Get CanGoBack() As Boolean
    Dim v As Long
    dcf m_pWebView, 38, "get_CanGoBack", VarPtr(v)
    CanGoBack = (v <> 0)
End Property

Public Property Get CanGoForward() As Boolean
    Dim v As Long
    dcf m_pWebView, 39, "get_CanGoForward", VarPtr(v)
    CanGoForward = (v <> 0)
End Property

Public Property Get DocumentTitle() As String
    Dim pStr As LongPtr
    dcf m_pWebView, 48, "get_DocumentTitle", VarPtr(pStr)
    If pStr <> 0 Then
        DocumentTitle = PtrToStrW(pStr)
        CoTaskMemFree pStr
    End If
End Property

Public Property Get ContainsFullScreenElement() As Boolean
    Dim v As Long
    dcf m_pWebView, 54, "get_ContainsFullScreenElement", VarPtr(v)
    ContainsFullScreenElement = (v <> 0)
End Property


' --- Navigation ---

Public Function NavigateAsync(ByVal uri As String) As Long
    NavigateAsync = dcf(m_pWebView, 5, "Navigate", StrPtr(uri))
End Function

Public Sub NavigateSync(ByVal url As String, Optional ByVal TimeoutSec As Double = 10)
    m_IsNavigating = True
    If NavigateAsync(url) = 0 Then
        WaitNavigation TimeoutSec
    End If
End Sub

Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim start As Double: start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - start > TimeoutSec Then
            Debug.Print "WaitNavigation: Timeout!"
            Exit Do
        End If
    Loop
End Sub

Public Function NavigateToString(ByVal htmlContent As String) As Long
    NavigateToString = dcf(m_pWebView, 6, "NavigateToString", StrPtr(htmlContent))
End Function

Public Function Reload() As Long
    Reload = dcf(m_pWebView, 31, "Reload")
End Function

Public Function GoBack() As Long
    GoBack = dcf(m_pWebView, 40, "GoBack")
End Function

Public Function GoForward() As Long
    GoForward = dcf(m_pWebView, 41, "GoForward")
End Function

Public Function Stop_() As Long
    Stop_ = dcf(m_pWebView, 43, "Stop")
End Function


' --- Web Message ---

Public Function PostWebMessageAsJson(ByVal jsonString As String) As Long
    PostWebMessageAsJson = dcf(m_pWebView, 32, "PostWebMessageAsJson", StrPtr(jsonString))
End Function

Public Function PostWebMessageAsString(ByVal webMessage As String) As Long
    PostWebMessageAsString = dcf(m_pWebView, 33, "PostWebMessageAsString", StrPtr(webMessage))
End Function


' --- Host Object ---

Public Function AddHostObjectToScript(ByVal Name As String, ByVal Instance As Object) As Long
    Dim vObj As Variant
    Set vObj = Instance
    AddHostObjectToScript = dcf(m_pWebView, 49, "AddHostObjectToScript", StrPtr(Name), VarPtr(vObj))
End Function

Public Function RemoveHostObjectFromScript(ByVal ObjName As String) As Long
    RemoveHostObjectFromScript = dcf(m_pWebView, 50, "RemoveHostObjectFromScript", StrPtr(ObjName))
End Function


' --- DevTools ---

Public Sub OpenDevToolsWindow()
    dcf m_pWebView, 51, "OpenDevToolsWindow"
End Sub


' --- Filter ---

Public Function AddWebResourceRequestedFilter( _
    ByVal uri As String, ByVal resourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT) As Long
    AddWebResourceRequestedFilter = dcf(m_pWebView, 57, _
        "AddWebResourceRequestedFilter", StrPtr(uri), resourceContext)
End Function

Public Function RemoveWebResourceRequestedFilter( _
    ByVal uri As String, ByVal resourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT) As Long
    RemoveWebResourceRequestedFilter = dcf(m_pWebView, 58, _
        "RemoveWebResourceRequestedFilter", StrPtr(uri), resourceContext)
End Function


' ???????????????????????????????????????????????
' Section 13: ICoreWebView2 イベントハンドラ登録
' ???????????????????????????????????????????????

''' <summary>
''' 初期化時に全デフォルトハンドラを登録する。
''' </summary>
Private Sub RegisterDefaultHandlers()
    ' --- ICoreWebView2 Events ---
    AddEventHandler m_pWebView, 7, 8, "Invoke_NavigationStarting", "WV.NavigationStarting"
    AddEventHandler m_pWebView, 9, 10, "Invoke_ContentLoading", "WV.ContentLoading"
    AddEventHandler m_pWebView, 11, 12, "Invoke_SourceChanged", "WV.SourceChanged"
    AddEventHandler m_pWebView, 13, 14, "Invoke_HistoryChanged", "WV.HistoryChanged"
    AddEventHandler m_pWebView, 15, 16, "Invoke_NavigationCompleted", "WV.NavigationCompleted"
    AddEventHandler m_pWebView, 17, 18, "Invoke_FrameNavigationStarting", "WV.FrameNavigationStarting"
    AddEventHandler m_pWebView, 19, 20, "Invoke_FrameNavigationCompleted", "WV.FrameNavigationCompleted"
    AddEventHandler m_pWebView, 21, 22, "Invoke_ScriptDialogOpening", "WV.ScriptDialogOpening"
    AddEventHandler m_pWebView, 23, 24, "Invoke_PermissionRequested", "WV.PermissionRequested"
    AddEventHandler m_pWebView, 25, 26, "Invoke_ProcessFailed", "WV.ProcessFailed"
    AddEventHandler m_pWebView, 34, 35, "Invoke_WebMessageReceived", "WV.WebMessageReceived"
    AddEventHandler m_pWebView, 44, 45, "Invoke_NewWindowRequested", "WV.NewWindowRequested"
    AddEventHandler m_pWebView, 46, 47, "Invoke_DocumentTitleChanged", "WV.DocumentTitleChanged"
    AddEventHandler m_pWebView, 52, 53, "Invoke_ContainsFullScreenElementChanged", "WV.ContainsFullScreenElementChanged"
    AddEventHandler m_pWebView, 55, 56, "Invoke_WebResourceRequested", "WV.WebResourceRequested"
    AddEventHandler m_pWebView, 59, 60, "Invoke_WindowCloseRequested", "WV.WindowCloseRequested"

    ' --- Filter & Script ---
    AddWebResourceRequestedFilter "*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL

    Dim script1 As String
    script1 = "console.log('script on document created!');"
    AddScriptToExecuteOnDocumentCreated script1
End Sub

' --- 個別 add ラッパー (任意のキーで登録可能) ---

Public Function add_NavigationStarting( _
    Optional ByVal Key As String = "WV.NavigationStarting") As String
    add_NavigationStarting = AddEventHandler(m_pWebView, 7, 8, "Invoke_NavigationStarting", Key)
End Function
Public Function add_ContentLoading( _
    Optional ByVal Key As String = "WV.ContentLoading") As String
    add_ContentLoading = AddEventHandler(m_pWebView, 9, 10, "Invoke_ContentLoading", Key)
End Function
Public Function add_SourceChanged( _
    Optional ByVal Key As String = "WV.SourceChanged") As String
    add_SourceChanged = AddEventHandler(m_pWebView, 11, 12, "Invoke_SourceChanged", Key)
End Function
Public Function add_HistoryChanged( _
    Optional ByVal Key As String = "WV.HistoryChanged") As String
    add_HistoryChanged = AddEventHandler(m_pWebView, 13, 14, "Invoke_HistoryChanged", Key)
End Function
Public Function add_NavigationCompleted( _
    Optional ByVal Key As String = "WV.NavigationCompleted") As String
    add_NavigationCompleted = AddEventHandler(m_pWebView, 15, 16, "Invoke_NavigationCompleted", Key)
End Function
Public Function add_FrameNavigationStarting( _
    Optional ByVal Key As String = "WV.FrameNavigationStarting") As String
    add_FrameNavigationStarting = AddEventHandler(m_pWebView, 17, 18, "Invoke_FrameNavigationStarting", Key)
End Function
Public Function add_FrameNavigationCompleted( _
    Optional ByVal Key As String = "WV.FrameNavigationCompleted") As String
    add_FrameNavigationCompleted = AddEventHandler(m_pWebView, 19, 20, "Invoke_FrameNavigationCompleted", Key)
End Function
Public Function add_ScriptDialogOpening( _
    Optional ByVal Key As String = "WV.ScriptDialogOpening") As String
    add_ScriptDialogOpening = AddEventHandler(m_pWebView, 21, 22, "Invoke_ScriptDialogOpening", Key)
End Function
Public Function add_PermissionRequested( _
    Optional ByVal Key As String = "WV.PermissionRequested") As String
    add_PermissionRequested = AddEventHandler(m_pWebView, 23, 24, "Invoke_PermissionRequested", Key)
End Function
Public Function add_ProcessFailed( _
    Optional ByVal Key As String = "WV.ProcessFailed") As String
    add_ProcessFailed = AddEventHandler(m_pWebView, 25, 26, "Invoke_ProcessFailed", Key)
End Function
Public Function add_WebMessageReceived( _
    Optional ByVal Key As String = "WV.WebMessageReceived") As String
    add_WebMessageReceived = AddEventHandler(m_pWebView, 34, 35, "Invoke_WebMessageReceived", Key)
End Function
Public Function add_NewWindowRequested( _
    Optional ByVal Key As String = "WV.NewWindowRequested") As String
    add_NewWindowRequested = AddEventHandler(m_pWebView, 44, 45, "Invoke_NewWindowRequested", Key)
End Function
Public Function add_DocumentTitleChanged( _
    Optional ByVal Key As String = "WV.DocumentTitleChanged") As String
    add_DocumentTitleChanged = AddEventHandler(m_pWebView, 46, 47, "Invoke_DocumentTitleChanged", Key)
End Function
Public Function add_ContainsFullScreenElementChanged( _
    Optional ByVal Key As String = "WV.ContainsFullScreenElementChanged") As String
    add_ContainsFullScreenElementChanged = AddEventHandler(m_pWebView, 52, 53, _
        "Invoke_ContainsFullScreenElementChanged", Key)
End Function
Public Function add_WebResourceRequested( _
    Optional ByVal Key As String = "WV.WebResourceRequested") As String
    add_WebResourceRequested = AddEventHandler(m_pWebView, 55, 56, "Invoke_WebResourceRequested", Key)
End Function
Public Function add_WindowCloseRequested( _
    Optional ByVal Key As String = "WV.WindowCloseRequested") As String
    add_WindowCloseRequested = AddEventHandler(m_pWebView, 59, 60, "Invoke_WindowCloseRequested", Key)
End Function

' --- 個別 remove ラッパー ---

Public Function remove_NavigationStarting( _
    Optional ByVal Key As String = "WV.NavigationStarting") As Long
    remove_NavigationStarting = RemoveEventHandler(Key)
End Function
Public Function remove_ContentLoading( _
    Optional ByVal Key As String = "WV.ContentLoading") As Long
    remove_ContentLoading = RemoveEventHandler(Key)
End Function
Public Function remove_SourceChanged( _
    Optional ByVal Key As String = "WV.SourceChanged") As Long
    remove_SourceChanged = RemoveEventHandler(Key)
End Function
Public Function remove_HistoryChanged( _
    Optional ByVal Key As String = "WV.HistoryChanged") As Long
    remove_HistoryChanged = RemoveEventHandler(Key)
End Function
Public Function remove_NavigationCompleted( _
    Optional ByVal Key As String = "WV.NavigationCompleted") As Long
    remove_NavigationCompleted = RemoveEventHandler(Key)
End Function
Public Function remove_FrameNavigationStarting( _
    Optional ByVal Key As String = "WV.FrameNavigationStarting") As Long
    remove_FrameNavigationStarting = RemoveEventHandler(Key)
End Function
Public Function remove_FrameNavigationCompleted( _
    Optional ByVal Key As String = "WV.FrameNavigationCompleted") As Long
    remove_FrameNavigationCompleted = RemoveEventHandler(Key)
End Function
Public Function remove_ScriptDialogOpening( _
    Optional ByVal Key As String = "WV.ScriptDialogOpening") As Long
    remove_ScriptDialogOpening = RemoveEventHandler(Key)
End Function
Public Function remove_PermissionRequested( _
    Optional ByVal Key As String = "WV.PermissionRequested") As Long
    remove_PermissionRequested = RemoveEventHandler(Key)
End Function
Public Function remove_ProcessFailed( _
    Optional ByVal Key As String = "WV.ProcessFailed") As Long
    remove_ProcessFailed = RemoveEventHandler(Key)
End Function
Public Function remove_WebMessageReceived( _
    Optional ByVal Key As String = "WV.WebMessageReceived") As Long
    remove_WebMessageReceived = RemoveEventHandler(Key)
End Function
Public Function remove_NewWindowRequested( _
    Optional ByVal Key As String = "WV.NewWindowRequested") As Long
    remove_NewWindowRequested = RemoveEventHandler(Key)
End Function
Public Function remove_DocumentTitleChanged( _
    Optional ByVal Key As String = "WV.DocumentTitleChanged") As Long
    remove_DocumentTitleChanged = RemoveEventHandler(Key)
End Function
Public Function remove_ContainsFullScreenElementChanged( _
    Optional ByVal Key As String = "WV.ContainsFullScreenElementChanged") As Long
    remove_ContainsFullScreenElementChanged = RemoveEventHandler(Key)
End Function
Public Function remove_WebResourceRequested( _
    Optional ByVal Key As String = "WV.WebResourceRequested") As Long
    remove_WebResourceRequested = RemoveEventHandler(Key)
End Function
Public Function remove_WindowCloseRequested( _
    Optional ByVal Key As String = "WV.WindowCloseRequested") As Long
    remove_WindowCloseRequested = RemoveEventHandler(Key)
End Function


' ???????????????????????????????????????????????
' Section 14: Invoke コールバック (ICoreWebView2 Events)
' ???????????????????????????????????????????????

Public Function Invoke_NavigationStarting( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent NavigationStarting(sender, args)
    Invoke_NavigationStarting = 0
End Function

Public Function Invoke_ContentLoading( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ContentLoading(sender, args)
    Invoke_ContentLoading = 0
End Function

Public Function Invoke_SourceChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent SourceChanged(sender, args)
    Invoke_SourceChanged = 0
End Function

Public Function Invoke_HistoryChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent HistoryChanged
    Invoke_HistoryChanged = 0
End Function

Public Function Invoke_NavigationCompleted( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_IsNavigating = False
    RaiseEvent NavigationCompleted(sender, args)
    Invoke_NavigationCompleted = 0
End Function

Public Function Invoke_FrameNavigationStarting( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent FrameNavigationStarting(sender, args)
    Invoke_FrameNavigationStarting = 0
End Function

Public Function Invoke_FrameNavigationCompleted( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent FrameNavigationCompleted(sender, args)
    Invoke_FrameNavigationCompleted = 0
End Function

Public Function Invoke_ScriptDialogOpening( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ScriptDialogOpening(sender, args)
    Invoke_ScriptDialogOpening = 0
End Function

Public Function Invoke_PermissionRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent PermissionRequested(sender, args)
    Invoke_PermissionRequested = 0
End Function

Public Function Invoke_ProcessFailed( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ProcessFailed(sender, args)
    Invoke_ProcessFailed = 0
End Function

Public Function Invoke_WebMessageReceived( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long

    ' get_Source (index 3)
    Dim strSource As String, pStr As LongPtr
    dcf args, 3, "get_Source", VarPtr(pStr)
    If pStr <> 0 Then strSource = PtrToStrW(pStr): CoTaskMemFree pStr: pStr = 0

    ' get_WebMessageAsJson (index 4)
    Dim strJson As String
    dcf args, 4, "get_WebMessageAsJson", VarPtr(pStr)
    If pStr <> 0 Then strJson = PtrToStrW(pStr): CoTaskMemFree pStr: pStr = 0

    ' TryGetWebMessageAsString (index 5)
    Dim strMsg As String
    dcf args, 5, "TryGetWebMessageAsString", VarPtr(pStr)
    If pStr <> 0 Then strMsg = PtrToStrW(pStr): CoTaskMemFree pStr: pStr = 0

    RaiseEvent WebMessageReceived(strSource, strJson, strMsg)
    Invoke_WebMessageReceived = 0
End Function

Public Function Invoke_NewWindowRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent NewWindowRequested(sender, args)
    Invoke_NewWindowRequested = 0
End Function

Public Function Invoke_DocumentTitleChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent DocumentTitleChanged
    Invoke_DocumentTitleChanged = 0
End Function

Public Function Invoke_ContainsFullScreenElementChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ContainsFullScreenElementChanged
    Invoke_ContainsFullScreenElementChanged = 0
End Function

Public Function Invoke_WebResourceRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent WebResourceRequested(sender, args)
    Invoke_WebResourceRequested = 0
End Function

Public Function Invoke_WindowCloseRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent WindowCloseRequested(sender, args)
    Invoke_WindowCloseRequested = 0
End Function


' ???????????????????????????????????????????????
' Section 15: AddScriptToExecuteOnDocumentCreated (Index 27/28)
' ???????????????????????????????????????????????

Public Function AddScriptToExecuteOnDocumentCreated(ByVal javascript As String) As Long

    If m_IsScriptPending Then
        Debug.Print "AddScriptToExecuteOnDocumentCreated: 前回が未完了です"
        AddScriptToExecuteOnDocumentCreated = E_PENDING
        Exit Function
    End If

    m_IsScriptPending = True
    m_PendingScript = javascript

    Dim fakeObj As LongPtr
    fakeObj = EnsureCallbackHandler("__add_script", "Invoke_AddScriptCompleted")
    If fakeObj = 0 Then
        m_IsScriptPending = False
        Exit Function
    End If

    AddScriptToExecuteOnDocumentCreated = dcf(m_pWebView, 27, _
        "AddScriptToExecuteOnDocumentCreated", StrPtr(javascript), fakeObj)
End Function

Public Function Invoke_AddScriptCompleted( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal resultIdPtr As LongPtr) As Long

    Dim scriptId As String
    Dim javascript As String
    scriptId = PtrToStrW(resultIdPtr)
    javascript = m_PendingScript

    If errorCode = 0 Then
        m_ScriptCount = m_ScriptCount + 1
        ReDim Preserve m_Scripts(0 To m_ScriptCount - 1)
        m_Scripts(m_ScriptCount - 1).scriptId = scriptId
        m_Scripts(m_ScriptCount - 1).javascript = javascript
    End If

    m_PendingScript = ""
    m_IsScriptPending = False

    If errorCode = 0 Then
        RaiseEvent AddScriptCompleted(scriptId, javascript)
    Else
        RaiseEvent AddScriptFailed(javascript, errorCode)
    End If

    Invoke_AddScriptCompleted = 0
End Function

Public Function RemoveScriptToExecuteOnDocumentCreated(ByVal scriptId As String) As Long
    RemoveScriptToExecuteOnDocumentCreated = dcf(m_pWebView, 28, _
        "RemoveScriptToExecuteOnDocumentCreated", StrPtr(scriptId))

    ' 内部リストからも削除
    Dim i As Long, j As Long
    For i = 0 To m_ScriptCount - 1
        If m_Scripts(i).scriptId = scriptId Then
            For j = i To m_ScriptCount - 2
                m_Scripts(j) = m_Scripts(j + 1)
            Next j
            m_ScriptCount = m_ScriptCount - 1
            If m_ScriptCount = 0 Then Erase m_Scripts _
            Else ReDim Preserve m_Scripts(0 To m_ScriptCount - 1)
            Exit For
        End If
    Next i
End Function


' ???????????????????????????????????????????????
' Section 16: ExecuteScript (Index 29)
' ???????????????????????????????????????????????

Public Function ExecuteScriptAsync(ByVal script As String) As Long

    If m_IsExecuteScriptPending Then
        Debug.Print "ExecuteScript: 前回が未完了です"
        ExecuteScriptAsync = E_PENDING
        Exit Function
    End If

    m_IsExecuteScriptPending = True

    Dim fakeObj As LongPtr
    fakeObj = EnsureCallbackHandler("__exec_script", "Invoke_ExecuteScript")
    If fakeObj = 0 Then
        m_IsExecuteScriptPending = False
        Exit Function
    End If

    ExecuteScriptAsync = dcf(m_pWebView, 29, "ExecuteScript", StrPtr(script), fakeObj)
End Function

Public Function Invoke_ExecuteScript( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal resultJsonPtr As LongPtr) As Long

    Dim jsonResult As String
    If resultJsonPtr <> 0 Then jsonResult = PtrToStrW(resultJsonPtr)
    LastScriptResult = jsonResult

    m_IsExecuteScriptPending = False
    m_IsScriptRunning = False

    RaiseEvent ReceiveScriptResult(jsonResult)
    Invoke_ExecuteScript = 0
End Function

Public Function ExecuteScriptSync( _
    ByVal script As String, Optional ByVal TimeoutSec As Double = 5) As String

    m_IsScriptRunning = True
    ExecuteScriptAsync script

    Dim start As Double: start = Timer
    Do While m_IsScriptRunning
        DoEvents
        If Timer - start > TimeoutSec Then
            Debug.Print "ExecuteScriptSync: Timeout!"
            m_IsScriptRunning = False
            Exit Do
        End If
    Loop

    ExecuteScriptSync = LastScriptResult
End Function


' ???????????????????????????????????????????????
' Section 17: CapturePreview (Index 30)
' ???????????????????????????????????????????????

Public Function CapturePreview( _
    ByVal format As COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, _
    ByVal pStream As LongPtr) As Long

    Dim fakeObj As LongPtr
    fakeObj = EnsureCallbackHandler("__capture_preview", "Invoke_CapturePreviewCompleted")
    If fakeObj = 0 Then Exit Function

    CapturePreview = dcf(m_pWebView, 30, "CapturePreview", CLng(format), pStream, fakeObj)
End Function

Public Function CapturePreviewToFile( _
    ByVal folderPath As String, ByVal filePath As String, _
    Optional ByVal isJpeg As Boolean = False) As Long

    Dim fullPath As String: fullPath = folderPath & filePath
    CreateDeepFolder folderPath

    Dim pStream As LongPtr
    Dim hr As Long
    hr = SHCreateStreamOnFileW(StrPtr(fullPath), _
        STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_DENY_NONE, pStream)
    If hr <> 0 Or pStream = 0 Then
        CapturePreviewToFile = hr
        Exit Function
    End If

    m_pCaptureStream = pStream
    CapturePreviewToFile = CapturePreview(IIf(isJpeg, 1, 0), pStream)
End Function

Public Function Invoke_CapturePreviewCompleted( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal dummy As LongPtr) As Long

    Debug.Print "CapturePreviewCompleted: hr=&H" & Hex(errorCode)

    If m_pCaptureStream <> 0 Then
        dcf m_pCaptureStream, 2, "IStream.Release"
        m_pCaptureStream = 0
    End If

    RaiseEvent CapturePreviewCompleted(errorCode)
    Invoke_CapturePreviewCompleted = 0
End Function


' ???????????????????????????????????????????????
' Section 18: CallDevToolsProtocolMethod (Index 36)
' ???????????????????????????????????????????????

''' <summary>
''' CDP メソッドを非同期で呼び出す。
''' 完了時に CDPMethodCompleted イベントが発火する。
''' 戻り値: 成功時は正のリクエストID、失敗時は負のHRESULT。
''' </summary>
Public Function CallDevToolsProtocolMethod( _
    ByVal methodName As String, ByVal parameterAsJson As String) As Long

    m_CDPNextId = m_CDPNextId + 1
    Dim requestId As Long: requestId = m_CDPNextId

    ' リクエストごとにハンドラを生成 (再利用不可)
    Dim vtbl As LongPtr: vtbl = CreateHandlerVTable("Invoke_CDPMethodCompleted")
    If vtbl = 0 Then
        CallDevToolsProtocolMethod = &H80004005  ' E_FAIL
        Exit Function
    End If
    Dim fakeObj As LongPtr: fakeObj = AllocFakeObj(vtbl)
    If fakeObj = 0 Then
        VirtualFree vtbl, 0, MEM_RELEASE
        CallDevToolsProtocolMethod = &H80004005
        Exit Function
    End If

    ' リクエスト追跡
    m_CDPRequestCount = m_CDPRequestCount + 1
    ReDim Preserve m_CDPRequests(0 To m_CDPRequestCount - 1)
    With m_CDPRequests(m_CDPRequestCount - 1)
        .requestId = requestId
        .methodName = methodName
        .vtblBlock = vtbl
        .fakeObj = fakeObj
    End With

    ' fakeObj→Key 逆引き登録
    m_fakeObjToKey(CStr(fakeObj)) = "__cdp_" & requestId

    Dim hr As Long
    hr = dcf(m_pWebView, 36, "CallDevToolsProtocolMethod", _
        StrPtr(methodName), StrPtr(parameterAsJson), fakeObj)

    If hr = 0 Then
        CallDevToolsProtocolMethod = requestId
    Else
        ' 失敗時クリーンアップ
        FreeHandlerMemory vtbl, fakeObj
        m_fakeObjToKey.Remove CStr(fakeObj)
        m_CDPRequestCount = m_CDPRequestCount - 1
        If m_CDPRequestCount = 0 Then Erase m_CDPRequests _
        Else ReDim Preserve m_CDPRequests(0 To m_CDPRequestCount - 1)
        CallDevToolsProtocolMethod = hr
    End If
End Function

Public Function Invoke_CDPMethodCompleted( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal resultPtr As LongPtr) As Long

    ' pThis (= fakeObj) から対応リクエストを検索
    Dim i As Long, j As Long
    Dim methodName As String
    Dim requestId As Long
    Dim found As Boolean

    For i = 0 To m_CDPRequestCount - 1
        If m_CDPRequests(i).fakeObj = pThis Then
            methodName = m_CDPRequests(i).methodName
            requestId = m_CDPRequests(i).requestId

            ' メモリ解放
            FreeHandlerMemory m_CDPRequests(i).vtblBlock, m_CDPRequests(i).fakeObj
            m_fakeObjToKey.Remove CStr(pThis)

            ' 配列を詰める
            For j = i To m_CDPRequestCount - 2
                m_CDPRequests(j) = m_CDPRequests(j + 1)
            Next j
            m_CDPRequestCount = m_CDPRequestCount - 1
            If m_CDPRequestCount = 0 Then
                Erase m_CDPRequests
            Else
                ReDim Preserve m_CDPRequests(0 To m_CDPRequestCount - 1)
            End If
            found = True
            Exit For
        End If
    Next i

    Dim strResult As String
    If resultPtr <> 0 Then strResult = PtrToStrW(resultPtr)

    If found Then
        RaiseEvent CDPMethodCompleted(requestId, methodName, errorCode, strResult)
    Else
        Debug.Print "CDPMethodCompleted: 不明なリクエスト pThis=&H" & Hex(pThis)
    End If

    Invoke_CDPMethodCompleted = 0
End Function


' ???????????????????????????????????????????????
' Section 19: DevToolsProtocolEventReceiver (Index 42)
' ???????????????????????????????????????????????

''' <summary>
''' CDP イベントを購読する。
''' eventName がそのままキーとなる (プレフィックス "DTP." 付き)。
''' </summary>
Public Function add_DevToolsProtocolEventReceived(ByVal eventName As String) As String
    Dim Key As String: Key = "DTP." & eventName

    If m_handlers.Exists(Key) Then
        Debug.Print "add_DevToolsProtocolEventReceived: 既に登録済み - " & eventName
        Exit Function
    End If

    ' 1. GetDevToolsProtocolEventReceiver でレシーバー取得
    Dim receiverPtr As LongPtr
    Dim hr As Long
    hr = dcf(m_pWebView, 42, "GetDevToolsProtocolEventReceiver", _
        StrPtr(eventName), VarPtr(receiverPtr))
    If hr <> 0 Or receiverPtr = 0 Then Exit Function

    ' 2. ハンドラ生成
    Dim vtbl As LongPtr: vtbl = CreateHandlerVTable("Invoke_DTPEventReceived")
    If vtbl = 0 Then
        dcf receiverPtr, 2, "Receiver.Release"
        Exit Function
    End If
    Dim fakeObj As LongPtr: fakeObj = AllocFakeObj(vtbl)
    If fakeObj = 0 Then
        VirtualFree vtbl, 0, MEM_RELEASE
        dcf receiverPtr, 2, "Receiver.Release"
        Exit Function
    End If

    ' 3. レシーバーに対して add_DevToolsProtocolEventReceived (index 3)
    '    remove は index 4
    Dim token As LongPtr
    hr = dcf(receiverPtr, 3, "add_DevToolsProtocolEventReceived", fakeObj, VarPtr(token))
    If hr <> 0 Then
        FreeHandlerMemory vtbl, fakeObj
        dcf receiverPtr, 2, "Receiver.Release"
        Exit Function
    End If

    ' 4. エントリ保存
    m_handlerCount = m_handlerCount + 1
    ReDim Preserve m_handlerData(0 To m_handlerCount - 1)
    Dim idx As Long: idx = m_handlerCount - 1

    With m_handlerData(idx)
        .Key = Key
        .vtblBlock = vtbl
        .fakeObj = fakeObj
        .isEvent = True
        .pInterface = receiverPtr  ' remove_ はレシーバーに対して呼ぶ
        .removeIndex = 4           ' remove_DevToolsProtocolEventReceived
        .token = token
        .pExtraRelease = receiverPtr  ' Finalize 時に Release する
        .active = True
    End With

    m_handlers(Key) = idx
    m_fakeObjToKey(CStr(fakeObj)) = Key

    Debug.Print "DTP Event Registered: " & eventName
    add_DevToolsProtocolEventReceived = Key
End Function

Public Function Invoke_DTPEventReceived( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long

    ' pThis (= fakeObj) からキーを逆引き → eventName 取得
    Dim Key As String: Key = LookupKeyByFakeObj(pThis)
    Dim eventName As String
    If Len(Key) > 4 Then eventName = Mid$(Key, 5)  ' "DTP." を除去

    ' get_ParameterObjectAsJson (args の index 3)
    Dim paramJson As String
    Dim pStr As LongPtr
    dcf args, 3, "get_ParameterObjectAsJson", VarPtr(pStr)
    If pStr <> 0 Then
        paramJson = PtrToStrW(pStr)
        CoTaskMemFree pStr
    End If

    RaiseEvent DevToolsProtocolEventReceived(eventName, paramJson)
    Invoke_DTPEventReceived = 0
End Function

Public Function remove_DevToolsProtocolEventReceived(ByVal eventName As String) As Long
    remove_DevToolsProtocolEventReceived = RemoveEventHandler("DTP." & eventName)
End Function


' ???????????????????????????????????????????????
' Section 20: リサイズ
' ???????????????????????????????????????????????

Public Sub ResizeWebView()
    If m_pController = 0 Or m_targetFrame Is Nothing Then Exit Sub

    Dim l(3) As Long
    l(0) = 0: l(1) = 0
    l(2) = PtsToPx(m_targetFrame.InsideWidth, False)
    l(3) = PtsToPx(m_targetFrame.InsideHeight, True)
    dcf m_pController, 6, "put_Bounds", VarPtr(l(0))
End Sub


' ???????????????????????????????????????????????
' Section 21: 全体クリーンアップ (Finalize)
' ???????????????????????????????????????????????

Public Sub Finalize()

    If Not m_Initialized Then Exit Sub

    ' --- 1. 全ハンドラ解除 ---
    Dim keys As Variant
    If m_handlers.Count > 0 Then
        keys = m_handlers.keys
        Dim i As Long
        For i = UBound(keys) To LBound(keys) Step -1
            Dim Key As String: Key = CStr(keys(i))
            ' __xxx は完了コールバック (add_/remove_ なし) → メモリだけ解放
            Dim idx As Long
            If m_handlers.Exists(Key) Then
                idx = m_handlers(Key)
                If m_handlerData(idx).isEvent Then
                    RemoveEventHandler Key
                Else
                    ' 非イベントハンドラのメモリ解放
                    Dim fakeObjKey As String
                    fakeObjKey = CStr(m_handlerData(idx).fakeObj)
                    If m_fakeObjToKey.Exists(fakeObjKey) Then m_fakeObjToKey.Remove fakeObjKey
                    FreeHandlerMemory m_handlerData(idx).vtblBlock, m_handlerData(idx).fakeObj
                    m_handlerData(idx).active = False
                    m_handlers.Remove Key
                End If
            End If
        Next i
    End If

    ' --- 2. CDP リクエスト残存分を解放 ---
    For i = 0 To m_CDPRequestCount - 1
        FreeHandlerMemory m_CDPRequests(i).vtblBlock, m_CDPRequests(i).fakeObj
    Next i
    m_CDPRequestCount = 0
    Erase m_CDPRequests

    ' --- 3. COM インターフェースを Release ---
    If m_pSettings <> 0 Then dcf m_pSettings, 2, "Settings.Release": m_pSettings = 0
    If m_pWebView <> 0 Then dcf m_pWebView, 2, "WebView.Release": m_pWebView = 0
    If m_pController <> 0 Then
        dcf m_pController, 24, "Controller.Close"
        dcf m_pController, 2, "Controller.Release"
        m_pController = 0
    End If
    If m_pEnv <> 0 Then dcf m_pEnv, 2, "Env.Release": m_pEnv = 0

    ' --- 4. CaptureStream ---
    If m_pCaptureStream <> 0 Then
        dcf m_pCaptureStream, 2, "CaptureStream.Release"
        m_pCaptureStream = 0
    End If

    ' --- 5. 共有スタブ解放 ---
    If m_pStubQI <> 0 Then VirtualFree m_pStubQI, 0, MEM_RELEASE: m_pStubQI = 0
    If m_pStubAddRefRelease <> 0 Then VirtualFree m_pStubAddRefRelease, 0, MEM_RELEASE: m_pStubAddRefRelease = 0

    ' --- 6. 内部データクリア ---
    m_handlerCount = 0
    Erase m_handlerData
    m_handlers.RemoveAll
    m_fakeObjToKey.RemoveAll
    m_ScriptCount = 0
    Erase m_Scripts
    m_CDPNextId = 0

    Set m_targetFrame = Nothing
    m_targetHWnd = 0

    m_Initialized = False
    Debug.Print "WebView2 Finalized."
End Sub


' ???????????????????????????????????????????????
' Section 22: ユーティリティ
' ???????????????????????????????????????????????

''' <summary>LPWSTR を VBA の String に変換する</summary>
Private Function PtrToStrW(ByVal p As LongPtr) As String
    If p = 0 Then Exit Function
    Dim cch As Long: cch = lstrlenW(p)
    If cch = 0 Then Exit Function
    PtrToStrW = String$(cch, vbNullChar)
    RtlMoveMemory StrPtr(PtrToStrW), p, CLngPtr(cch * 2)
End Function

''' <summary>Points → Pixels 変換 (DPI対応)</summary>
Private Function PtsToPx(ByVal pts As Single, ByVal isVertical As Boolean) As Long
    Dim hdc As LongPtr: hdc = GetDC(0)
    Dim dpi As Long
    If isVertical Then dpi = GetDeviceCaps(hdc, LOGPIXELSY) _
    Else dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
    PtsToPx = CLng(pts * dpi / 72)
End Function

''' <summary>フォルダを再帰的に作成する</summary>
Private Sub CreateDeepFolder(ByVal folderPath As String)
    If folderPath = "" Then Exit Sub
    If Right$(folderPath, 1) = "\" Then folderPath = Left$(folderPath, Len(folderPath) - 1)
    If Dir(folderPath, vbDirectory) <> "" Then Exit Sub

    Dim parent As String
    Dim pos As Long: pos = InStrRev(folderPath, "\")
    If pos > 0 Then
        parent = Left$(folderPath, pos - 1)
        CreateDeepFolder parent
    End If
    MkDir folderPath
End Sub

''' <summary>登録済みハンドラの一覧をデバッグ出力する</summary>
Public Sub DumpHandlers()
    Debug.Print "=== Registered Handlers ==="
    Debug.Print "Total entries: " & m_handlerCount

    Dim keys As Variant
    If m_handlers.Count = 0 Then
        Debug.Print "(none)"
        Exit Sub
    End If

    keys = m_handlers.keys
    Dim i As Long
    For i = LBound(keys) To UBound(keys)
        Dim k As String: k = CStr(keys(i))
        Dim idx2 As Long: idx2 = m_handlers(k)
        With m_handlerData(idx2)
            Debug.Print "  [" & idx2 & "] Key=" & .Key & _
                " isEvent=" & .isEvent & _
                " active=" & .active & _
                " iface=&H" & Hex(.pInterface) & _
                " rmIdx=" & .removeIndex
        End With
    Next i
    Debug.Print "==========================="
End Sub

''' <summary>登録済み CDP リクエストの一覧をデバッグ出力する</summary>
Public Sub DumpCDPRequests()
    Debug.Print "=== Pending CDP Requests ==="
    Debug.Print "Total: " & m_CDPRequestCount
    Dim i As Long
    For i = 0 To m_CDPRequestCount - 1
        Debug.Print "  [" & i & "] id=" & m_CDPRequests(i).requestId & _
            " method=" & m_CDPRequests(i).methodName
    Next i
    Debug.Print "============================="
End Sub


' ???????????????????????????????????????????????
' Section 23: COM 呼び出しラッパー (dcf)
' ???????????????????????????????????????????????

''' <summary>
''' DispCallFunc を使って COM vtable の任意メソッドを呼び出す汎用ラッパー。
''' 可変引数 (ParamArray) で最大16個のパラメータを受け取る。
'''
''' ? 全パラメータは LongPtr として扱う。
'''   BOOL 等は事前に Long に変換してから渡すこと。
'''   Double は 8 bytes だが LongPtr (x64) も 8 bytes なので互換。
''' </summary>
Private Function dcf( _
    ByVal pInterface As LongPtr, _
    ByVal vtblIndex As Long, _
    ByVal funcName As String, _
    ParamArray args() As Variant) As Long

    If pInterface = 0 Then
        Debug.Print "dcf: null interface - " & funcName
        dcf = &H80004003  ' E_POINTER
        Exit Function
    End If

    ' vtable からメソッドアドレスを取得
    Dim pVTable As LongPtr
    CopyMemory pVTable, ByVal pInterface, LenB(pVTable)
    Dim pMethod As LongPtr
    CopyMemory pMethod, ByVal (pVTable + vtblIndex * LenB(pVTable)), LenB(pMethod)

    ' 引数を配列に展開
    Dim argc As Long
    argc = UBound(args) - LBound(args) + 1
    If argc < 0 Then argc = 0  ' ParamArray が空の場合

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
                Case vbLong:    vt(i) = vbLong
                Case vbLongPtr: vt(i) = vbLongPtr
                Case vbDouble:  vt(i) = vbDouble
                Case Else:      vt(i) = vbLongPtr
                    If VarType(vals(i)) = vbLong Then
                        vals(i) = CLngPtr(CLng(vals(i)))
                    End If
            End Select
            vp(i) = VarPtr(vals(i))
        Next i
    End If

    Dim res As Variant
    Dim hr As Long

    If argc > 0 Then
        hr = DispCallFunc(pInterface, CLngPtr(vtblIndex) * LenB(pVTable), _
            CC_STDCALL, vbLong, argc, vt(0), vp(0), res)
    Else
        hr = DispCallFunc(pInterface, CLngPtr(vtblIndex) * LenB(pVTable), _
            CC_STDCALL, vbLong, 0, 0, 0, res)
    End If

    If hr <> 0 Then
        If Len(funcName) > 0 Then
            Debug.Print "dcf CALL failed: " & funcName & " hr=&H" & Hex(hr)
        End If
        dcf = hr
    Else
        Dim lres As Long: lres = CLng(res)
        If lres <> 0 And Len(funcName) > 0 Then
            Debug.Print "dcf METHOD failed: " & funcName & " result=&H" & Hex(lres)
        End If
        dcf = lres
    End If
End Function



