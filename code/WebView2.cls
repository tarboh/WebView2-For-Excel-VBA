VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Public Class Class1 （このクラス単独でローダー実行からICoreWebView2のイベントハンドラ登録、Invoke受信まで全部やりたいけど・・・）
Option Explicit

'Public WithEvents wv2 As c3_WebView2

Private m_targetFrame As MSForms.Control

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

Private m_pSettings As LongPtr ' ICoreWebView2Settings ポインタ

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

Public Enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL = 0
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_DOCUMENT = 1
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_STYLESHEET = 2
    COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE = 4
End Enum

Public Enum COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT
    PNG = 0
    JPEG = 1
End Enum

'Public Settings As c6_WebView2Settings
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
Private m_IsNavigating As Boolean

Private m_IsScriptPending As Boolean  ' True の間は追加登録を拒否
Private Const E_PENDING As Long = &H8000000A   ' → -2147483638 :COMの標準HRESULT
Private m_PendingScript As String  ' Invoke で受け取るまでの一時保管
Private m_ScriptCount As Long
Private m_Scripts() As ScriptEntry
Private Type ScriptEntry
    scriptId As String
    javascript As String
End Type
Private m_IsExcecuteScriptPending As Boolean
Public LastScriptResult As String
Private m_IsScriptRunning As Boolean
Private Const STGM_CREATE As Long = &H1000
Private Const STGM_READWRITE As Long = &H2
Private Const STGM_SHARE_DENY_NONE As Long = &H40
' API to create an IStream from a file path (Unicode version)
Private Declare PtrSafe Function SHCreateStreamOnFileW Lib "shlwapi.dll" ( _
    ByVal pszFile As LongPtr, _
    ByVal grfMode As Long, _
    ByRef ppStm As LongPtr) As Long
Private m_pCaptureStream As LongPtr  ' 解放用に保持
Private m_CDPNextId As Long
Private m_CDPRequestCount As Long
Private Type CDPEntry
    requestId As Long
    methodName As String
    fakeObj   As LongPtr
    vtbl      As LongPtr
End Type
Private m_CDPRequests() As CDPEntry
Private m_DTPEventCount As Long
Private Type DTPEventEntry
    eventName As String
    receiverPtr As LongPtr  ' ICoreWebView2DevToolsProtocolEventReceiver*
    fakeObj As LongPtr
    vtbl As LongPtr
    token As LongPtr
End Type
Private m_DTPEvents() As DTPEventEntry
Private Const CRYPT_STRING_BASE64 As Long = &H1
Private Declare PtrSafe Function CryptStringToBinary Lib "crypt32.dll" Alias "CryptStringToBinaryW" ( _
    ByVal pszString As LongPtr, _
    ByVal cchString As Long, _
    ByVal dwFlags As Long, _
    ByVal pbBinary As LongPtr, _
    ByRef pcbBinary As Long, _
    ByRef pdwSkip As Long, _
    ByRef pdwFlags As Long) As Long

' Use in a Standard Module
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As LongPtr, ByVal hdc As LongPtr) As Long

' Device capabilities indices for retrieving screen DPI (Used with GetDeviceCaps)
Private Const LOGPIXELSX As Long = 88 ' Logical pixels per inch (X-axis)
Private Const LOGPIXELSY As Long = 90 ' Logical pixels per inch (Y-axis)


'Events
Public Event NavigationStarting()       ' 7
Public Event ContentLoading()           ' 9
Public Event SourceChanged()            '11
Public Event HistoryChanged()           '13
Public Event NavigationCompleted()      '15
Public Event FrameNavigationStarting()  '17
Public Event FrameNavigationCompleted() '19
Public Event ScriptDialogOpening()      '21
Public Event PermissionRequested()      '23
Public Event ProcessFailed()            '25
Public Event AddScriptToExecuteOnDocumentCreatedCompleted(ByVal scriptId As String, ByVal javascript As String) '27-1
Public Event AddScriptToExecuteOnDocumentCreatedFailed(ByVal javascript As String, ByVal errorCode As Long) '27-2
Public Event ReceiveScriptResult(ByVal result As String) '29
Public Event CapturePreviewCompleted(ByVal errorCode As Long) '30
Public Event WebMessageReceived(ByVal Source As String, ByVal messageJson As String, ByVal messageString As String) '34
Public Event CallDevToolsProtocolMethodCompleted(ByVal requestId As Long, ByVal methodName As String, ByVal errorCode As String, ByVal result As String) '36
Public Event DevToolsProtocolEventReceived(ByVal eventName As String, ByVal parameterJson As String) '42
Public Event NewWindowRequested() '44
Public Event DocumentTitleChanged() '46
Public Event ContainsFullScreenElementChanged() '52
Public Event WebResourceRequested() '55
Public Event WindowCloseRequested(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) '59


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

Private Function GetBoolProperty(ByVal Index As Long) As Boolean
    If m_pSettings = 0 Then Exit Function
    
    Dim hr As Long, res As Variant
    Dim outVal As Long ' C++ BOOL is 4 bytes
    
    Dim types(0) As Integer: types(0) = vbLongPtr
    Dim args(0) As Variant:  args(0) = VarPtr(outVal)
    Dim ptrs(0) As LongPtr: ptrs(0) = VarPtr(args(0))
    
    hr = DispCallFunc(m_pSettings, Index * LenB(m_pSettings), CC_STDCALL, vbLong, 1, types(0), ptrs(0), res)
    
    If hr = 0 Then
        GetBoolProperty = (outVal <> 0)
    End If
End Function

Private Sub LetBoolProperty(ByVal Index As Long, ByVal value As Boolean)
    If m_pSettings = 0 Then Exit Sub
    
    Dim hr As Long, res As Variant
    Dim inVal As Long
    inVal = IIf(value, 1, 0) ' C++ TRUE is 1 (VBA True is -1)
    
    Dim types(0) As Integer: types(0) = vbLong
    Dim args(0) As Variant:  args(0) = inVal
    Dim ptrs(0) As LongPtr: ptrs(0) = VarPtr(args(0))
    
    hr = DispCallFunc(m_pSettings, Index * LenB(m_pSettings), CC_STDCALL, vbLong, 1, types(0), ptrs(0), res)
    If hr = 0 And res = 0 Then
        Debug.Print "LetBoolProperty Success."
    End If
End Sub

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


Public Sub CreateWebView2Environment(ByVal TargetFrame As MSForms.Frame)
    
    Set m_targetFrame = TargetFrame
    Dim TargetCtrl As MSForms.Control
    Set TargetCtrl = m_targetFrame
    targetHWnd = TargetCtrl.[_GethWnd]
    
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

    ResizeWebView2Force targetHWnd, PtsToPx(m_targetFrame.InsideWidth, False), PtsToPx(m_targetFrame.InsideHeight, True)

    ' --- 4. WebView2本体を取得 ---
    Call dcf(m_pController, 25, "get_CoreWebView2", VarPtr(m_pWebView))
    'これでモジュールレベル変数のm_pWebViewにICoreWebView2のポインタが入る
    
    If m_pWebView = 0 Then
        Debug.Print "IcoreWebView2インターフェースポインタの取得に失敗しました。"
        Exit Sub
    End If
    
    ' --- 5. ICoreWebView2Settingsインターフェースポインターを取得 ---
    Call get_Settings
    
    'Set wv2 = New c3_WebView2
    'wv2.m_pWebView = m_pWebView
    
    ' --- 5. イベントハンドラの登録 ---
    Call Registar_IcoreWebView2_Handlers
    
    'wv2.NavigateAsync "https://www.google.com"
    
'    ' --- 6. 初期ページ表示テスト ---
    If m_pWebView <> 0 Then
        Debug.Print "Navigating to Google..."
        Call dcf(m_pWebView, 5, "Navigate", StrPtr("https://www.google.com"))
    End If
    
    DoEvents
End Sub

' Converts VBA Points (twips/points) to Pixel coordinates based on monitor DPI scaling
Private Function PtsToPx(ByVal pts As Single, ByVal isVertical As Boolean) As Long
    Dim hdc As LongPtr
    Dim dpi As Long
    
    hdc = GetDC(0) ' Get DC for the primary desktop monitor
    If isVertical Then
        dpi = GetDeviceCaps(hdc, LOGPIXELSY)
    Else
        dpi = GetDeviceCaps(hdc, LOGPIXELSX)
    End If
    ReleaseDC 0, hdc
    
    ' Formula: Points * (Current DPI / 72)
    ' e.g., 150% scaling will yield dpi = 144, resulting in points * 2
    PtsToPx = pts * (dpi / 72)
End Function

' Recursively traverses Chromium's internal Win32 hierarchy to forcefully resize drawing canvases
Public Sub ResizeWebView2Force(ByVal parentHwnd As LongPtr, ByVal w As Long, ByVal h As Long)
    Dim child0 As LongPtr
    Dim child1 As LongPtr
    Dim childD3D As LongPtr
    
    ' 1. Chrome_WidgetWin_0 (Outer Shell)
    child0 = FindWindowEx(parentHwnd, 0, "Chrome_WidgetWin_0", vbNullString)
    If child0 = 0 Then Exit Sub
    MoveWindow child0, 0, 0, w, h, 1
    
    ' 2. Chrome_WidgetWin_1 (Renderer Canvas)
    child1 = FindWindowEx(child0, 0, "Chrome_WidgetWin_1", vbNullString)
    If child1 <> 0 Then
        MoveWindow child1, 0, 0, w, h, 1
        
        ' 3. Intermediate D3D Window (Direct3D Drawing Surface)
        childD3D = FindWindowEx(child1, 0, "Intermediate D3D Window", vbNullString)
        If childD3D <> 0 Then
            MoveWindow childD3D, 0, 0, w, h, 1
        End If
    End If
    
    Debug.Print "Win32 Recursive Resize Completed."
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
Private Function remove_Handler(pWebView2 As LongPtr, vtblIndex As Long, token As Long, HandlerName As String) As Long
    Dim hr As Long, res As Variant
    Dim vtoken As Variant: vtoken = token
    Dim args(0) As Variant: args(0) = vtoken
    Dim argsPtr(0) As Variant: argsPtr(0) = VarPtr(args(0))
    Dim argsType(0) As Integer: argsType(0) = vbLong
    hr = DispCallFunc(pWebView2, vtblIndex * LenB(pWebView2), CC_STDCALL, vbLong, 1, argsType(0), argsPtr(0), res)
    If hr = 0 Then
        If res = 0 Then
            remove_Handler = res
        Else
            Debug.Print HandlerName & " Failed. res:" & res
            remove_Handler = res
        End If
    Else
        Debug.Print HandlerName & " Failed. hr:" & hr
        remove_Handler = hr
    End If
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

Private Sub WV2_NavigationStarting()
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

    'Set wv2 = Nothing
    m_pController = 0
End Sub

Private Function PtrToStrW(ByVal pWStr As LongPtr) As String
    Dim Length As Long
    Dim buf As String

    If pWStr = 0 Then
        PtrToStrW = ""
        Exit Function
    End If

    ' 1. Get length of the string (Unicode characters count)
    Length = lstrlenW(pWStr)

    If Length > 0 Then
        ' 2. Allocate VBA string buffer (1 character = 2 bytes)
        buf = Space$(Length)

        ' 3. Copy from memory to the buffer
        ' Since VBA String is internally Unicode, it can be copied directly
        CopyMemory ByVal StrPtr(buf), ByVal pWStr, Length * 2

        PtrToStrW = buf
    Else
        PtrToStrW = ""
    End If
End Function

''' <summary>
''' Parses a JSON string into a Scripting.Object using JScript runtime.
''' </summary>
Public Function ParseJSON(ByVal jsonString As String) As Object
    Dim html As Object
    Set html = CreateObject("htmlfile")
    
    ' Leverage htmlfile's JavaScript space to parse JSON
    html.parentWindow.execScript "function parse(json) { return JSON.parse(json); }", "JScript"
    
    ' Returns as a VBA Object (usable like a Dictionary)
    Set ParseJSON = html.parentWindow.Parse(jsonString)
End Function

''' <summary>
''' Decodes a Base64 string and converts it into a binary Byte array.
''' </summary>
Public Function Base64Decode(ByVal base64Str As String) As Byte()
    Dim byteLen As Long
    Dim skip As Long, outFlags As Long
    
    ' Calculate the required buffer size
    If CryptStringToBinary(StrPtr(base64Str), Len(base64Str), CRYPT_STRING_BASE64, 0, byteLen, skip, outFlags) = 0 Then
        Err.Raise vbObjectError + 1001, , "Failed to calculate Base64 buffer size."
    End If
    
    Dim bytes() As Byte
    ReDim bytes(byteLen - 1)
    
    ' Perform the actual Base64 decoding
    If CryptStringToBinary(StrPtr(base64Str), Len(base64Str), CRYPT_STRING_BASE64, VarPtr(bytes(0)), byteLen, skip, outFlags) = 0 Then
        Err.Raise vbObjectError + 1002, , "Failed to decode Base64 string."
    End If
    
    Base64Decode = bytes
End Function

''' <summary>
''' Saves a Byte array to a physical file (No ADODB.Stream reference required).
''' </summary>
Public Sub SaveBytesToFile(ByRef bytes() As Byte, ByVal filePath As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 1 ' adTypeBinary
        .Open
        .Write bytes
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
End Sub

''' <summary>
''' Recursively creates directories if they do not exist, including deep subfolders.
''' </summary>
''' <param name="folderPath">The absolute path of the directory to create.</param>
Public Sub CreateDeepFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Do nothing if the folder already exists
    If fso.FolderExists(folderPath) Then Exit Sub
    
    ' Get the parent folder path
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    
    ' If the parent folder does not exist, recursively call itself to create the parent first
    If Not fso.FolderExists(parentPath) Then
        CreateDeepFolder parentPath
    End If
    
    ' Finally, create the target folder
    fso.CreateFolder folderPath
    Debug.Print "Deep folder created: " & folderPath
    
    Set fso = Nothing
End Sub

Public Sub Registar_IcoreWebView2_Handlers()
    Call add_NavigationStarting
    Call add_ContentLoading                           ' 9
    Call add_SourceChanged                            '11
    Call add_HistoryChanged                           '13

    Call add_NavigationCompleted                      '15
    Call add_FrameNavigationStarting                  '17
    Call add_FrameNavigationCompleted                 '19
    Call add_ScriptDialogOpening                      '21
    Call add_PermissionRequested                      '23
    Call add_ProcessFailed                            '25
    Dim script1 As String
    script1 = "console.log('script on document created!');"
    Call AddScriptToExecuteOnDocumentCreated(script1) '27
    Call add_WebMessageReceived '34
    Call add_NewWindowRequested '44
    Call add_DocumentTitleChanged '46
    Call add_ContainsFullScreenElementChanged '52
    Call add_WebResourceRequested '55
    Call AddWebResourceRequestedFilter("*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_ALL) '57
    Call add_WindowCloseRequested '59
End Sub

' --- Cleaning up ---
Public Sub Finalize()
    
    If m_pWebView = 0 Then Exit Sub
    
    ' 全常駐ハンドラを解除
    remove_Handler m_pWebView, 8, m_tokens(7), "remove_NavigationStarting"
    remove_Handler m_pWebView, 10, m_tokens(9), "remove_ContentLoading"
    remove_Handler m_pWebView, 12, m_tokens(11), "remove_SourceChanged"
    remove_Handler m_pWebView, 14, m_tokens(13), "remove_HistoryChanged"
    remove_Handler m_pWebView, 16, m_tokens(15), "remove_NavigationCompleted"
    remove_Handler m_pWebView, 18, m_tokens(17), "remove_FrameNavigationStarting"
    remove_Handler m_pWebView, 20, m_tokens(19), "remove_FrameNavigationCompleted"
    remove_Handler m_pWebView, 22, m_tokens(21), "remove_ScriptDialogOpening"
    remove_Handler m_pWebView, 24, m_tokens(23), "remove_PermissionRequested"
    remove_Handler m_pWebView, 26, m_tokens(25), "remove_ProcessFailed"
    remove_Handler m_pWebView, 35, m_tokens(34), "remove_WebMessageReceived"
    remove_Handler m_pWebView, 45, m_tokens(44), "remove_NewWindowRequested"
    remove_Handler m_pWebView, 47, m_tokens(46), "remove_DocumentTitleChanged"
    remove_Handler m_pWebView, 53, m_tokens(52), "remove_ContainsFullScreenElementChanged"
    remove_Handler m_pWebView, 56, m_tokens(55), "remove_WebResourceRequested"
    remove_Handler m_pWebView, 60, m_tokens(59), "remove_WindowCloseRequested"

    ' DTP イベント全解除
    Dim i As Long
    For i = 0 To m_DTPEventCount - 1
        dcf m_DTPEvents(i).receiverPtr, 4, "remove_DevToolsProtocolEventReceived", m_DTPEvents(i).token
        VirtualFree m_DTPEvents(i).vtbl, 0, MEM_RELEASE
        VirtualFree m_DTPEvents(i).fakeObj, 0, MEM_RELEASE
    Next i
    m_DTPEventCount = 0
    Erase m_DTPEvents

    ' vtble / fakeObj のメモリ解放
    For i = LBound(m_vtbls) To UBound(m_vtbls)
        If m_vtbls(i) <> 0 Then
            VirtualFree m_vtbls(i), 0, MEM_RELEASE
            m_vtbls(i) = 0
        End If
        If m_fakeObjs(i) <> 0 Then
            VirtualFree m_fakeObjs(i), 0, MEM_RELEASE
            m_fakeObjs(i) = 0
        End If
    Next i

    ' 共有スタブ解放
    If m_pStubQI <> 0 Then
        VirtualFree m_pStubQI, 0, MEM_RELEASE
        m_pStubQI = 0
    End If
    If m_pStubAddRefRelease <> 0 Then
        VirtualFree m_pStubAddRefRelease, 0, MEM_RELEASE
        m_pStubAddRefRelease = 0
    End If

    m_pWebView = 0
    
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MIDL_INTERFACE ("e562e4f0-d7fa-43ac-8d71-c05150499f00")
'ICoreWebView2Settings:      Public IUnknown
'{
'public:
'3
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsScriptEnabled(
'    /* [retval][out] */ BOOL *isScriptEnabled) = 0;
Public Property Get IsScriptEnabled() As Boolean
    IsScriptEnabled = GetBoolProperty(3)
End Property
'4
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsScriptEnabled(
'    /* [in] */ BOOL isScriptEnabled) = 0;
Public Property Let IsScriptEnabled(ByVal value As Boolean)
    LetBoolProperty 4, value
End Property
'5
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsWebMessageEnabled(
'    /* [retval][out] */ BOOL *isWebMessageEnabled) = 0;
Public Property Get IsWebMessageEnabled() As Boolean
    IsWebMessageEnabled = GetBoolProperty(5)
End Property
'6
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsWebMessageEnabled(
'    /* [in] */ BOOL isWebMessageEnabled) = 0;
Public Property Let IsWebMessageEnabled(ByVal value As Boolean)
    LetBoolProperty 6, value
End Property
'7
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDefaultScriptDialogsEnabled(
'    /* [retval][out] */ BOOL *areDefaultScriptDialogsEnabled) = 0;
Public Property Get AreDefaultScriptDialogsEnabled() As Boolean
    AreDefaultScriptDialogsEnabled = GetBoolProperty(7)
End Property
'8
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDefaultScriptDialogsEnabled(
'    /* [in] */ BOOL areDefaultScriptDialogsEnabled) = 0;
Public Property Let AreDefaultScriptDialogsEnabled(ByVal value As Boolean)
    LetBoolProperty 8, value
End Property
'9
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsStatusBarEnabled(
'    /* [retval][out] */ BOOL *isStatusBarEnabled) = 0;
Public Property Get IsStatusBarEnabled() As Boolean
    IsStatusBarEnabled = GetBoolProperty(9)
End Property
'10
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsStatusBarEnabled(
'    /* [in] */ BOOL isStatusBarEnabled) = 0;
Public Property Let IsStatusBarEnabled(ByVal value As Boolean)
    LetBoolProperty 10, value
End Property
'11
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDevToolsEnabled(
'    /* [retval][out] */ BOOL *areDevToolsEnabled) = 0;
Public Property Get AreDevToolsEnabled() As Boolean
    AreDevToolsEnabled = GetBoolProperty(11)
End Property
'12
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDevToolsEnabled(
'    /* [in] */ BOOL areDevToolsEnabled) = 0;
Public Property Let AreDevToolsEnabled(ByVal value As Boolean)
    LetBoolProperty 12, value
End Property
'13
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreDefaultContextMenusEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'14
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreDefaultContextMenusEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'15
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_AreHostObjectsAllowed(
'    /* [retval][out] */ BOOL *allowed) = 0;
'
'16
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_AreHostObjectsAllowed(
'    /* [in] */ BOOL allowed) = 0;
'
'17
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsZoomControlEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'18
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsZoomControlEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'19
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsBuiltInErrorPageEnabled(
'    /* [retval][out] */ BOOL *enabled) = 0;
'
'20
'virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_IsBuiltInErrorPageEnabled(
'    /* [in] */ BOOL enabled) = 0;
'
'};

''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''

'MIDL_INTERFACE ("76eceacb-0462-4d94-ac83-423a6793775e")
'ICoreWebView2:          Public IUnknown
'{
'public:

'3
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Settings(
'    /* [retval][out] */ ICoreWebView2Settings **settings) = 0;
Public Function get_Settings() As Long
    'Dim psettings As LongPtr
    Dim hr As Long
    hr = dcf(m_pWebView, 3, "get_Settings", VarPtr(m_pSettings))
    If hr = 0 Then
        Debug.Print "m_psettings:" & m_pSettings
        'Set Settings = New c6_WebView2Settings
        'Settings.Init psettings
    End If
End Function

'4
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Source(
'    /* [retval][out] */ LPWSTR *uri) = 0;
Public Property Get Source() As String
    'Get current URL
    
    Dim strsource As String
    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(m_pWebView, 4, "get_Source", VarPtr(pStr))
    If pStr <> 0 Then
        strsource = PtrToStrW(pStr)
        CoTaskMemFree pStr
        pStr = 0
        Source = strsource
    End If

End Property

'5
'virtual HRESULT STDMETHODCALLTYPE Navigate(
'    /* [in] */ LPCWSTR uri) = 0;
Public Function NavigateAsync(ByVal uri As String) As Long
    'Navigate to the specified URI.
    NavigateAsync = dcf(m_pWebView, 5, "Navigate", StrPtr(uri))
End Function
Public Sub NavigateSync(ByVal url As String, Optional ByVal TimeoutSec As Double = 10)
    ' Synchronized version: Wait until loading is complete.
    m_IsNavigating = True
    If NavigateAsync(url) = 0 Then
        WaitNavigation TimeoutSec
    End If
End Sub
Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim start As Double: start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - start > TimeoutSec Then Exit Do
    Loop
End Sub

'6
'virtual HRESULT STDMETHODCALLTYPE NavigateToString(
'    /* [in] */ LPCWSTR htmlContent) = 0;
Public Function NavigateToString(htmlContent As String) As Long
    'Renders the specified HTML code.
    NavigateToString = dcf(m_pWebView, 6, "NavigateToString", StrPtr(htmlContent))
End Function

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
'8
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationStarting(token As Long) As Long
    'remove_NavigationStarting = remove_Handler(Me, 8, token, "remove_NavigationStarting")
    remove_NavigationStarting = remove_Handler(Me, 8, m_tokens(7), "remove_NavigationStarting")
End Function

'9
'virtual HRESULT STDMETHODCALLTYPE add_ContentLoading(
'    /* [in] */ ICoreWebView2ContentLoadingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ContentLoading() As Long
    'add_ContentLoading = DCF_RegisterHandler(Me, 9, "add_ContentLoading", AddressOf ContentLoading_Invoke, "ContentLoadingHandler")
    add_ContentLoading = RegisterHandler(9, "Invoke_ContentLoading")
End Function
Public Function Invoke_ContentLoading(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ContentLoading
    Invoke_ContentLoading = 0
End Function
'10
'virtual HRESULT STDMETHODCALLTYPE remove_ContentLoading(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ContentLoading(token As Long) As Long
    'remove_ContentLoading = remove_Handler(Me, 10, token, "remove_ContentLoading")
    remove_ContentLoading = remove_Handler(Me, 10, m_tokens(9), "remove_ContentLoading")
End Function

'11
'virtual HRESULT STDMETHODCALLTYPE add_SourceChanged(
'    /* [in] */ ICoreWebView2SourceChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_SourceChanged() As Long
    'add_SourceChanged = DCF_RegisterHandler(Me, 11, "add_SourceChanged", AddressOf SourceChanged_Invoke, "SourceChangedHandler")
    add_SourceChanged = RegisterHandler(11, "Invoke_SourceChanged")
End Function
Public Function Invoke_SourceChanged(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent SourceChanged
    Invoke_SourceChanged = 0
End Function
'12
'virtual HRESULT STDMETHODCALLTYPE remove_SourceChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_SourceChanged(token As Long) As Long
    'remove_SourceChanged = remove_Handler(Me, 12, token, "remove_SourceChanged")
    remove_SourceChanged = remove_Handler(Me, 12, m_tokens(11), "remove_SourceChanged")
End Function

'13
'virtual HRESULT STDMETHODCALLTYPE add_HistoryChanged(
'    /* [in] */ ICoreWebView2HistoryChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_HistoryChanged() As Long
    'add_HistoryChanged = DCF_RegisterHandler(Me, 13, "add_HistoryChanged", AddressOf HistoryChanged_Invoke, "HistoryChangedHandler")
    add_HistoryChanged = RegisterHandler(13, "Invoke_HistoryChanged")
End Function
Public Function Invoke_HistoryChanged(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent HistoryChanged
    Invoke_HistoryChanged = 0
End Function
'14
'virtual HRESULT STDMETHODCALLTYPE remove_HistoryChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_HistoryChanged(token As Long) As Long
    'remove_HistoryChanged = remove_Handler(Me, 14, token, "remove_HistoryChanged")
    remove_HistoryChanged = remove_Handler(Me, 14, m_tokens(13), "remove_HistoryChanged")
End Function

'15
'virtual HRESULT STDMETHODCALLTYPE add_NavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'add_NavigationCompleted
Public Function add_NavigationCompleted() As Long
    'add_NavigationCompleted = mHelpers.DCF_RegisterHandler(Me, 15, "add_NavigationCompleted", AddressOf NavCompleted_Invoke, "NavigationCompleted")
    add_NavigationCompleted = RegisterHandler(15, "Invoke_NavigationCompleted")
End Function
Public Function Invoke_NavigationCompleted(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent NavigationCompleted
    Invoke_NavigationCompleted = 0
End Function
'16
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationCompleted(token As Long) As Long
    'remove_NavigationCompleted = remove_Handler(Me, 16, token, "remove_NavigationCompleted")
    remove_NavigationCompleted = remove_Handler(Me, 16, m_tokens(15), "remove_NavigationCompleted")
End Function

'17
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationStarting() As Long
    'add_FrameNavigationStarting = mHelpers.DCF_RegisterHandler(Me, 17, "add_FrameNavigationStarting", AddressOf FrameNavigationStarting_Invoke, "FrameNavigationStarting")
    add_FrameNavigationStarting = RegisterHandler(17, "Invoke_FrameNavigationStarting")
End Function
Public Function Invoke_FrameNavigationStarting(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent FrameNavigationStarting
    Invoke_FrameNavigationStarting = 0
End Function
'18
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationStarting(token As Long) As Long
    'remove_FrameNavigationStarting = remove_Handler(Me, 18, token, "remove_FrameNavigationStarting")
    remove_FrameNavigationStarting = remove_Handler(Me, 18, m_tokens(17), "remove_FrameNavigationStarting")
End Function

'19
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationCompleted() As Long
    'add_FrameNavigationCompleted = mHelpers.DCF_RegisterHandler(Me, 19, "add_FrameNavigationCompleted", AddressOf FrameNavigationCompleted_Invoke, "FrameNavigationCompleted")
    add_FrameNavigationCompleted = RegisterHandler(19, "Invoke_FrameNavigationCompleted")
End Function
Public Function Invoke_FrameNavigationCompleted(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent FrameNavigationCompleted
    Invoke_FrameNavigationCompleted = 0
End Function
'20
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationCompleted(token As Long) As Long
    'remove_FrameNavigationCompleted = remove_Handler(Me, 20, token, "remove_FrameNavigationCompleted")
    remove_FrameNavigationCompleted = remove_Handler(Me, 20, m_tokens(19), "remove_FrameNavigationCompleted")
End Function

'21
'virtual HRESULT STDMETHODCALLTYPE add_ScriptDialogOpening(
'    /* [in] */ ICoreWebView2ScriptDialogOpeningEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'   ** This event is triggered only if the ICoreWebView2Settings::AreDefaultScriptDialogsEnabled property is set to FALSE.
Public Function add_ScriptDialogOpening() As Long
    'add_ScriptDialogOpening = mHelpers.DCF_RegisterHandler(Me, 21, "add_ScriptDialogOpening", AddressOf ScriptDialogOpening_Invoke, "ScriptDialogOpening")
    add_ScriptDialogOpening = RegisterHandler(21, "Invoke_ScriptDialogOpening")
End Function
Public Function Invoke_ScriptDialogOpening(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ScriptDialogOpening
    Invoke_ScriptDialogOpening = 0
End Function
'22
'virtual HRESULT STDMETHODCALLTYPE remove_ScriptDialogOpening(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ScriptDialogOpening(token As Long) As Long
    'remove_ScriptDialogOpening = remove_Handler(Me, 22, token, "remove_ScriptDialogOpening")
    remove_ScriptDialogOpening = remove_Handler(Me, 22, m_tokens(21), "remove_ScriptDialogOpening")
End Function

'23
'virtual HRESULT STDMETHODCALLTYPE add_PermissionRequested(
'    /* [in] */ ICoreWebView2PermissionRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_PermissionRequested() As Long
    'add_PermissionRequested = mHelpers.DCF_RegisterHandler(Me, 23, "add_PermissionRequested", AddressOf PermissionRequested_Invoke, "PermissionRequested")
    add_PermissionRequested = RegisterHandler(23, "Invoke_PermissionRequested")
End Function
Public Function Invoke_PermissionRequested(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent PermissionRequested
    Invoke_PermissionRequested = 0
End Function
'24
'virtual HRESULT STDMETHODCALLTYPE remove_PermissionRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_PermissionRequested(token As Long) As Long
    'remove_PermissionRequested = remove_Handler(Me, 24, token, "remove_PermissionRequested")
    remove_PermissionRequested = remove_Handler(Me, 24, m_tokens(23), "remove_PermissionRequested")
End Function

'25
'virtual HRESULT STDMETHODCALLTYPE add_ProcessFailed(
'    /* [in] */ ICoreWebView2ProcessFailedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ProcessFailed() As Long
    'add_ProcessFailed = mHelpers.DCF_RegisterHandler(Me, 25, "add_ProcessFailed", AddressOf ProcessFailed_Invoke, "ProcessFailed")
    add_ProcessFailed = RegisterHandler(25, "Invoke_ProcessFailed")
End Function
Public Function Invoke_ProcessFailed(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ProcessFailed
    Invoke_ProcessFailed = 0
End Function
'26
'virtual HRESULT STDMETHODCALLTYPE remove_ProcessFailed(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ProcessFailed(token As Long) As Long
    'remove_ProcessFailed = remove_Handler(Me, 26, token, "remove_ProcessFailed")
    remove_ProcessFailed = remove_Handler(Me, 26, m_tokens(25), "remove_ProcessFailed")
End Function

'** Note that handlers 27 and 28 are handled differently from the other handlers.
' Instead of receiving a callback when an event occurs,
' this function passes the script as a string in order to execute JavaScript inside WebView2, so
' both Add and Remove have different signatures.
'27
'virtual HRESULT STDMETHODCALLTYPE AddScriptToExecuteOnDocumentCreated(
'    /* [in] */ LPCWSTR javaScript,
'    /* [in] */ ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler *handler) = 0;
Public Function AddScriptToExecuteOnDocumentCreated(ByVal javascript As String) As Long
    
    If m_IsScriptPending Then
        Debug.Print "AddScriptToExecuteOnDocumentCreated: 前の登録が完了していません"
        AddScriptToExecuteOnDocumentCreated = E_PENDING
        Exit Function
    Else
        m_IsScriptPending = True
    End If
    
    m_PendingScript = javascript
    
    '初回のみ生成
    If m_vtbls(27) = 0 Then
        ' vtble 生成
        m_vtbls(27) = CreateHandlerVTable("Invoke_AddScriptToExecuteOnDocumentCreated")
        If m_vtbls(27) = 0 Then Exit Function
    
        ' fakeObj 生成
        m_fakeObjs(27) = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        If m_fakeObjs(27) = 0 Then Exit Function
        RtlMoveMemory m_fakeObjs(27), VarPtr(m_vtbls(27)), 8
    End If
    
    ' WebView2 に登録
    AddScriptToExecuteOnDocumentCreated = dcf(m_pWebView, 27, "AddScriptToExecuteOnDocumentCreated", _
         StrPtr(javascript), m_fakeObjs(27))
End Function
Public Function Invoke_AddScriptToExecuteOnDocumentCreated( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal result As LongPtr) As Long

    ' ローカルに退避
    Dim scriptId As String
    Dim javascript As String
    scriptId = PtrToStrW(result)
    javascript = m_PendingScript
    
    ' 登録情報を保存
    If errorCode = 0 Then
        m_ScriptCount = m_ScriptCount + 1
        ReDim Preserve m_Scripts(0 To m_ScriptCount - 1)
        m_Scripts(m_ScriptCount - 1).scriptId = scriptId
        m_Scripts(m_ScriptCount - 1).javascript = javascript
    End If
    
    ' ロック解放（RaiseEvent の前）
    m_PendingScript = ""
    m_IsScriptPending = False
    
    ' ローカル変数からイベント発火
    If errorCode = 0 Then
        RaiseEvent AddScriptToExecuteOnDocumentCreatedCompleted(scriptId, javascript)
    Else
        RaiseEvent AddScriptToExecuteOnDocumentCreatedFailed(javascript, errorCode)
    End If
    
    Invoke_AddScriptToExecuteOnDocumentCreated = 0
End Function
'28
'virtual HRESULT STDMETHODCALLTYPE RemoveScriptToExecuteOnDocumentCreated(
'    /* [in] */ LPCWSTR id) = 0;
Public Function RemoveScriptToExecuteOnDocumentCreated(scriptId As String) As Long
    Dim hr As Long
    hr = dcf(m_pWebView, 28, "RemoveScriptToExecuteOnDocumentCreated", StrPtr(scriptId))
    If hr = 0 Then
        Debug.Print "RemoveScriptToExecuteOnDocumentCreated Success."
    Else
        Debug.Print "RemoveScriptToExecuteOnDocumentCreated Failed. Hr:" & hr
    End If
    RemoveScriptToExecuteOnDocumentCreated = hr
End Function
'
'Public Property Get scriptId() As String
'    scriptId = m_ScriptId
'End Property

'29
'virtual HRESULT STDMETHODCALLTYPE ExecuteScript(
'    /* [in] */ LPCWSTR javaScript,
'    /* [in] */ ICoreWebView2ExecuteScriptCompletedHandler *handler) = 0;
Public Function ExecuteScriptAsync(ByVal script As String) As Long
' Asynchronous version: Simply request execution and return. Receive results via events.

    If m_IsExcecuteScriptPending Then
        Debug.Print "ExecuteScript: 前の登録が完了していません"
        ExecuteScriptAsync = E_PENDING
        Exit Function
    Else
        m_IsExcecuteScriptPending = True
    End If

    '初回のみ生成
    If m_vtbls(29) = 0 Then
        ' vtble 生成
        m_vtbls(29) = CreateHandlerVTable("Invoke_ExecuteScript")
        If m_vtbls(29) = 0 Then Exit Function
    
        ' fakeObj 生成
        m_fakeObjs(29) = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        If m_fakeObjs(29) = 0 Then Exit Function
        RtlMoveMemory m_fakeObjs(29), VarPtr(m_vtbls(29)), 8
    End If

    ' WebView2 に登録
    ExecuteScriptAsync = dcf(m_pWebView, 29, "ExecuteScriptAsync", _
         StrPtr(script), m_fakeObjs(29))
End Function
Public Function Invoke_ExecuteScript(ByVal this As LongPtr, ByVal errorCode As Long, ByVal resultJsonPtr As LongPtr) As Long
    Dim jsonResult As String
    jsonResult = PtrToStrW(resultJsonPtr)
    LastScriptResult = jsonResult
    
    m_IsExcecuteScriptPending = False
    m_IsScriptRunning = False
    
    RaiseEvent ReceiveScriptResult(jsonResult)
    
    Invoke_ExecuteScript = 0
End Function

' Synchronized version: Receives the execution result directly as a return value.
Public Function ExecuteScriptSync(ByVal script As String, Optional ByVal TimeoutSec As Double = 5) As String
    m_IsScriptRunning = True
    ExecuteScriptAsync script
    
    ' Dedicated wait loop
    Dim start As Double: start = Timer
    Do While m_IsScriptRunning
        Debug.Print "waiting ExecuteScriptSync"
        DoEvents
        If Timer - start > TimeoutSec Then
            Debug.Print "ExecuteScriptSync: Timeout!"
            Exit Do
        End If
    Loop
    
    ExecuteScriptSync = LastScriptResult
End Function


'30
'virtual HRESULT STDMETHODCALLTYPE CapturePreview(
'    /* [in] */ COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT imageFormat,
'    /* [in] */ IStream *imageStream,
'    /* [in] */ ICoreWebView2CapturePreviewCompletedHandler *handler) = 0;
Public Function CapturePreview(ByVal format As COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT, ByVal pStream As LongPtr) As Long
    ' ※ format: 0=PNG, 1=JPEG
    ' ※ pStream: Pointer to the pre-created IStream
    
    '初回のみ生成
    If m_vtbls(30) = 0 Then
        ' vtble 生成
        m_vtbls(30) = CreateHandlerVTable("Invoke_CapturePreviewCompleted")
        If m_vtbls(30) = 0 Then Exit Function
    
        ' fakeObj 生成
        m_fakeObjs(30) = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
        If m_fakeObjs(30) = 0 Then Exit Function
        RtlMoveMemory m_fakeObjs(30), VarPtr(m_vtbls(30)), 8
    End If
    
    Dim thunkAddr As LongPtr
    RtlMoveMemory VarPtr(thunkAddr), m_vtbls(30) + (3 * 8), 8

    ' WebView2 に登録
    Dim hr As Long
    hr = dcf(m_pWebView, 30, "CapturePreview", format, pStream, m_fakeObjs(30))
    CapturePreview = hr

End Function
''' <summary>
''' Saves a screenshot of the WebView2 to the specified file path (Asynchronous).
''' </summary>
''' <param name="folderPath">The absolute path of the directory to save the file (e.g., "C:\temp\")</param>
''' <param name="filePath">The file name with extension (e.g., "capture.png")</param>
''' <param name="isJpeg">True to save as JPEG, False to save as PNG</param>
Public Function CapturePreviewToFile(ByVal folderPath As String, ByVal filePath As String, Optional ByVal isJpeg As Boolean = False) As Long
    
    Dim hr As Long
    Dim pStream As LongPtr
    
    ' Combine paths to prevent file naming collisions
    Dim fullPath As String
    fullPath = folderPath & filePath
    
    ' Automatically create the directory if it does not exist
    CreateDeepFolder folderPath
    
    ' 1. Create an IStream bound to the file (New creation & Read/Write mode)
    hr = SHCreateStreamOnFileW(StrPtr(fullPath), STGM_CREATE Or STGM_READWRITE Or STGM_SHARE_DENY_NONE, pStream)
    If hr <> 0 Or pStream = 0 Then
        CapturePreviewToFile = hr
        Exit Function
    End If

    m_pCaptureStream = pStream  ' 保持しておく

    Dim format As Long: format = IIf(isJpeg, 1, 0)
    
    hr = CapturePreview(format, pStream)
    CapturePreviewToFile = hr
End Function

Public Function Invoke_CapturePreviewCompleted( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal dummy As LongPtr) As Long

    Debug.Print "Invoke_CapturePreviewCompleted: " & errorCode

    If m_pCaptureStream <> 0 Then
        Call dcf(m_pCaptureStream, 2, "IUnknown::Release")
        m_pCaptureStream = 0
    End If

    RaiseEvent CapturePreviewCompleted(errorCode)
    Invoke_CapturePreviewCompleted = 0
End Function


'31
'virtual HRESULT STDMETHODCALLTYPE Reload( void) = 0;
Public Function Reload() As Long
    Reload = dcf(m_pWebView, 31, "Reload")
End Function

'32
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsJson(
'    /* [in] */ LPCWSTR webMessageAsJson) = 0;
Public Function PostWebMessageAsJson(ByVal jsonString As String) As String
    PostWebMessageAsJson = dcf(m_pWebView, 32, "PostWebMessageAsJson", StrPtr(jsonString))
End Function

'33
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsString(
'    /* [in] */ LPCWSTR webMessageAsString) = 0;
Public Function PostWebMessageAsString(ByVal webMessage As String) As Long
    PostWebMessageAsString = dcf(m_pWebView, 33, "PostWebMessageAsString", StrPtr(webMessage))
End Function

'34
'virtual HRESULT STDMETHODCALLTYPE add_WebMessageReceived(
'    /* [in] */ ICoreWebView2WebMessageReceivedEventHandler *handler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_WebMessageReceived() As Long
    'add_WebMessageReceived = mHelpers.DCF_RegisterHandler(Me, 34, "add_WebMessageReceived", AddressOf WebMessageReceived_Invoke, "WebMessageReceived")
    add_WebMessageReceived = RegisterHandler(34, "Invoke_WebMessageReceived")
End Function
Public Function Invoke_WebMessageReceived(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    
    'MIDL_INTERFACE ("0f99a40c-e962-4207-9e92-e3d542eff849")
    'ICoreWebView2WebMessageReceivedEventArgs: public IUnknown
    '{
    'public:
    '    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Source(
    '        /* [retval][out] */ LPWSTR *value) = 0;
    '
    '    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_WebMessageAsJson(
    '        /* [retval][out] */ LPWSTR *value) = 0;
    '
    '    virtual HRESULT STDMETHODCALLTYPE TryGetWebMessageAsString(
    '        /* [retval][out] */ LPWSTR *value) = 0;
    '
    '};
    
    'get_Source
    Dim strsource As String
    Dim pStr As LongPtr
    Call dcf(args, 3, "get_Source", VarPtr(pStr))
    If pStr <> 0 Then
        strsource = PtrToStrW(pStr)
        CoTaskMemFree pStr
        pStr = 0
    End If
    
    'get_WebMessageAsJson
    Dim strJson As String
    Call dcf(args, 4, "get_WebMessageAsJson", VarPtr(pStr))
    If pStr <> 0 Then
        strJson = PtrToStrW(pStr)
        CoTaskMemFree pStr
        pStr = 0
    End If
    
    'TryGetWebMessageAsString
    Dim strWebMessage As String
    Call dcf(args, 5, "TryGetWebMessageAsString", VarPtr(pStr))
    If pStr <> 0 Then
        strWebMessage = PtrToStrW(pStr)
        CoTaskMemFree pStr
        pStr = 0
    End If
    
    RaiseEvent WebMessageReceived(strsource, strJson, strWebMessage)
    Invoke_WebMessageReceived = 0
End Function
'35
'virtual HRESULT STDMETHODCALLTYPE remove_WebMessageReceived(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WebMessageReceived(token As Long) As Long
    'remove_WebMessageReceived = remove_Handler(Me, 35, token, "remove_WebMessageReceived")
    remove_WebMessageReceived = remove_Handler(Me, 35, m_tokens(34), "remove_WebMessageReceived")
End Function

'36
'virtual HRESULT STDMETHODCALLTYPE CallDevToolsProtocolMethod(
'    /* [in] */ LPCWSTR methodName,
'    /* [in] */ LPCWSTR parametersAsJson,
'    /* [in] */ ICoreWebView2CallDevToolsProtocolMethodCompletedHandler *handler) = 0;
Public Function CallDevToolsProtocolMethod(ByVal methodName As String, ByVal parameterAsJson As String) As Long

    m_CDPNextId = m_CDPNextId + 1
    Dim requestId As Long: requestId = m_CDPNextId

    ' リクエストを記録
    m_CDPRequestCount = m_CDPRequestCount + 1
    ReDim Preserve m_CDPRequests(0 To m_CDPRequestCount - 1)
    m_CDPRequests(m_CDPRequestCount - 1).requestId = requestId
    m_CDPRequests(m_CDPRequestCount - 1).methodName = methodName

    ' リクエストごとに vtble/fakeObj を作成
    Dim vtbl As LongPtr: vtbl = CreateHandlerVTable("Invoke_CallDevToolsProtocolMethod")
    Dim fakeObj As LongPtr
    fakeObj = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    RtlMoveMemory fakeObj, VarPtr(vtbl), 8

    ' fakeObj と requestId を対応させて保存
    m_CDPRequests(m_CDPRequestCount - 1).fakeObj = fakeObj
    m_CDPRequests(m_CDPRequestCount - 1).vtbl = vtbl

    ' HRESULTではなくリクエストIDを返す
    Dim hr As Long
    hr = dcf(m_pWebView, 36, "CallDevToolsProtocolMethod", _
        StrPtr(methodName), StrPtr(parameterAsJson), fakeObj)
    
    If hr = 0 Then
        CallDevToolsProtocolMethod = requestId  ' 成功時はリクエストID
    Else
        ' 失敗時はクリーンアップしてHRESULTを返す
        VirtualFree vtbl, 0, MEM_RELEASE
        VirtualFree fakeObj, 0, MEM_RELEASE
        m_CDPRequestCount = m_CDPRequestCount - 1
        CallDevToolsProtocolMethod = hr  ' 負の値（エラー）
    End If

End Function

Public Function Invoke_CallDevToolsProtocolMethod( _
    ByVal pThis As LongPtr, _
    ByVal errorCode As Long, _
    ByVal result As LongPtr) As Long

    ' pThis から対応するリクエストを検索
    Dim i As Long
    Dim methodName As String
    Dim requestId As Long
    For i = 0 To m_CDPRequestCount - 1
        If m_CDPRequests(i).fakeObj = pThis Then
            methodName = m_CDPRequests(i).methodName
            requestId = m_CDPRequests(i).requestId
            ' 使用済みエントリを解放
            VirtualFree m_CDPRequests(i).vtbl, 0, MEM_RELEASE
            VirtualFree m_CDPRequests(i).fakeObj, 0, MEM_RELEASE
            ' 配列から削除する処理
            Exit For
        End If
    Next i

    Dim strResult As String
    If result <> 0 Then strResult = PtrToStrW(result)

    RaiseEvent CallDevToolsProtocolMethodCompleted(requestId, methodName, errorCode, strResult)
    Invoke_CallDevToolsProtocolMethod = 0
End Function

'37
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_BrowserProcessId(
'    /* [retval][out] */ UINT32 *value) = 0;
Public Property Get BrowserProcessId() As Long
    Dim value As Long: value = 0
    Dim hr As Long
    hr = dcf(m_pWebView, 37, "", VarPtr(value))
    BrowserProcessId = value
End Property

'38
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoBack(
'    /* [retval][out] */ BOOL *canGoBack) = 0;
Public Property Get CanGoBack() As Boolean
    Dim lcanGoBack As Long
    Dim hr As Long
    hr = dcf(m_pWebView, 38, "get_canGoBack", VarPtr(lcanGoBack))
    If lcanGoBack = 1 Then
        CanGoBack = True
    Else
        CanGoBack = False
    End If
End Property

'39
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoForward(
'    /* [retval][out] */ BOOL *canGoForward) = 0;
Public Property Get CanGoForward() As Boolean
    Dim lCanGoForward As Long
    Dim hr As Long
    hr = dcf(m_pWebView, 39, "get_CanGoForward", VarPtr(lCanGoForward))
    If lCanGoForward = 1 Then
        CanGoForward = True
    Else
        CanGoForward = False
    End If
End Property

'40
'virtual HRESULT STDMETHODCALLTYPE GoBack( void) = 0;
Public Function GoBack() As Long
    GoBack = dcf(m_pWebView, 40, "GoBack")
End Function

'41
'virtual HRESULT STDMETHODCALLTYPE GoForward( void) = 0;
Public Function GoForward() As Long
    GoForward = dcf(m_pWebView, 41, "GoForward")
End Function

'42
'virtual HRESULT STDMETHODCALLTYPE GetDevToolsProtocolEventReceiver(
'    /* [in] */ LPCWSTR eventName,
'    /* [retval][out] */ ICoreWebView2DevToolsProtocolEventReceiver **receiver) = 0;
'
'    MIDL_INTERFACE ("b32ca51a-8371-45e9-9317-af021d080367")
'    ICoreWebView2DevToolsProtocolEventReceiver: Public IUnknown
'    {
'    public:
'        virtual HRESULT STDMETHODCALLTYPE add_DevToolsProtocolEventReceived(
'            /* [in] */ ICoreWebView2DevToolsProtocolEventReceivedEventHandler *eventHandler,
'            /* [out] */ EventRegistrationToken *token) = 0;
'
'        virtual HRESULT STDMETHODCALLTYPE remove_DevToolsProtocolEventReceived(
'            /* [in] */ EventRegistrationToken token) = 0;
'
'    };
Public Function add_DevToolsProtocolEventReceived(ByVal eventName As String) As Long
    
    ' 既に登録済みの場合は何もしない
    Dim i As Long
    For i = 0 To m_DTPEventCount - 1
        If m_DTPEvents(i).eventName = eventName Then
            Debug.Print "add_DevToolsProtocolEventReceived: 既に登録済み - " & eventName
            Exit Function
        End If
    Next i

    ' 1. GetDevToolsProtocolEventReceiver でレシーバー取得
    Dim receiverPtr As LongPtr
    Dim hr As Long
    hr = dcf(m_pWebView, 42, "GetDevToolsProtocolEventReceiver", StrPtr(eventName), VarPtr(receiverPtr))
    If hr <> 0 Or receiverPtr = 0 Then
        add_DevToolsProtocolEventReceived = hr
        Exit Function
    End If

    ' 2. エントリ追加
    m_DTPEventCount = m_DTPEventCount + 1
    ReDim Preserve m_DTPEvents(0 To m_DTPEventCount - 1)
    Dim idx As Long: idx = m_DTPEventCount - 1
    m_DTPEvents(idx).eventName = eventName
    m_DTPEvents(idx).receiverPtr = receiverPtr

    ' 3. ハンドラ生成
    m_DTPEvents(idx).vtbl = CreateHandlerVTable("Invoke_DevToolsProtocolEventReceived")
    m_DTPEvents(idx).fakeObj = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    RtlMoveMemory m_DTPEvents(idx).fakeObj, VarPtr(m_DTPEvents(idx).vtbl), 8

    ' 4. レシーバーに対して add_DevToolsProtocolEventReceived (vtbl index 3)
    add_DevToolsProtocolEventReceived = dcf(receiverPtr, 3, "add_DevToolsProtocolEventReceived", _
        m_DTPEvents(idx).fakeObj, VarPtr(m_DTPEvents(idx).token))

End Function
Public Function Invoke_DevToolsProtocolEventReceived( _
    ByVal pThis As LongPtr, _
    ByVal sender As LongPtr, _
    ByVal args As LongPtr) As Long

    ' pThis から eventName を逆引き
    Dim eventName As String
    Dim i As Long
    For i = 0 To m_DTPEventCount - 1
        If m_DTPEvents(i).fakeObj = pThis Then
            eventName = m_DTPEvents(i).eventName
            Exit For
        End If
    Next i

    ' parameterObjectAsJson を取得
    Dim paramJson As String
    Dim pStr As LongPtr: pStr = 0
    Dim hr As Long
    hr = dcf(args, 3, "get_ParameterObjectAsJson", VarPtr(pStr))
    If hr = 0 And pStr <> 0 Then
        paramJson = PtrToStrW(pStr)
        CoTaskMemFree pStr
    End If

    RaiseEvent DevToolsProtocolEventReceived(eventName, paramJson)
    Invoke_DevToolsProtocolEventReceived = 0
End Function
Public Function remove_DevToolsProtocolEventReceived(ByVal eventName As String) As Long

    Dim i As Long
    For i = 0 To m_DTPEventCount - 1
        If m_DTPEvents(i).eventName = eventName Then

            ' ハンドラ解除
            dcf m_DTPEvents(i).receiverPtr, 4, "remove_DevToolsProtocolEventReceived", m_DTPEvents(i).token

            ' メモリ解放
            VirtualFree m_DTPEvents(i).vtbl, 0, MEM_RELEASE
            VirtualFree m_DTPEvents(i).fakeObj, 0, MEM_RELEASE

            ' 配列を詰め直す
            Dim j As Long
            For j = i To m_DTPEventCount - 2
                m_DTPEvents(j) = m_DTPEvents(j + 1)
            Next j
            m_DTPEventCount = m_DTPEventCount - 1

            If m_DTPEventCount = 0 Then
                Erase m_DTPEvents
            Else
                ReDim Preserve m_DTPEvents(0 To m_DTPEventCount - 1)
            End If

            Exit Function
        End If
    Next i

    Debug.Print "remove_DevToolsProtocolEventReceived: 未登録 - " & eventName
End Function

'43
'virtual HRESULT STDMETHODCALLTYPE Stop( void) = 0;
Public Function Stop_() As Long
    Stop_ = dcf(m_pWebView, 43, "Stop_")
End Function

'44
'virtual HRESULT STDMETHODCALLTYPE add_NewWindowRequested(
'    /* [in] */ ICoreWebView2NewWindowRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_NewWindowRequested() As Long
    'add_NewWindowRequested = mHelpers.DCF_RegisterHandler(Me, 44, "add_NewWindowRequested", AddressOf NewWindowRequested_Invoke, "NewWindowRequested")
    add_NewWindowRequested = RegisterHandler(44, "Invoke_NewWindowRequested")
End Function
Public Function Invoke_NewWindowRequested(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent NewWindowRequested
    Invoke_NewWindowRequested = 0
End Function
'45
'virtual HRESULT STDMETHODCALLTYPE remove_NewWindowRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NewWindowRequested(token As Long) As Long
    'remove_NewWindowRequested = remove_Handler(Me, 45, token, "remove_NewWindowRequested")
    remove_NewWindowRequested = remove_Handler(Me, 45, m_tokens(44), "remove_NewWindowRequested")
End Function

'46
'virtual HRESULT STDMETHODCALLTYPE add_DocumentTitleChanged(
'    /* [in] */ ICoreWebView2DocumentTitleChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_DocumentTitleChanged() As Long
    'add_DocumentTitleChanged = mHelpers.DCF_RegisterHandler(Me, 46, "add_DocumentTitleChanged", AddressOf DocumentTitleChanged_Invoke, "DocumentTitleChanged")
    add_DocumentTitleChanged = RegisterHandler(46, "Invoke_DocumentTitleChanged")
End Function
Public Function Invoke_DocumentTitleChanged(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent DocumentTitleChanged
    Invoke_DocumentTitleChanged = 0
End Function
'47
'virtual HRESULT STDMETHODCALLTYPE remove_DocumentTitleChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_DocumentTitleChanged(token As Long) As Long
    'remove_DocumentTitleChanged = remove_Handler(Me, 47, token, "remove_DocumentTitleChanged")
    remove_DocumentTitleChanged = remove_Handler(Me, 47, m_tokens(46), "remove_DocumentTitleChanged")
End Function

'48
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_DocumentTitle(
'    /* [retval][out] */ LPWSTR *title) = 0;
Public Property Get DocumentTitle() As String
    Dim strTitle As String
    Dim pStr As LongPtr
    Dim hr As Long
    hr = dcf(m_pWebView, 48, "get_DocumentTitle", VarPtr(pStr))
    If pStr <> 0 Then
        strTitle = PtrToStrW(pStr)
        CoTaskMemFree pStr
        pStr = 0
        DocumentTitle = strTitle
    End If
End Property

'49
'virtual HRESULT STDMETHODCALLTYPE AddHostObjectToScript(
'    /* [in] */ LPCWSTR name,
'    /* [in] */ VARIANT *object) = 0;
Public Function AddHostObjectToScript(ByVal Name As String, ByVal Instance As Object) As Long
    Dim vObj As Variant
    Set vObj = Instance  ' Object を Variant に包む
    AddHostObjectToScript = dcf(m_pWebView, 49, "AddHostObjectToScript", StrPtr(Name), VarPtr(vObj))
    'AddHostObjectToScript = DCF_TwoArgsStringAndObject(m_pWebView, 49, "AddHostObjectToScript", Name, Instance)
End Function

'50
'virtual HRESULT STDMETHODCALLTYPE RemoveHostObjectFromScript(
'    /* [in] */ LPCWSTR name) = 0;
Public Function RemoveHostObjectFromScript(ByVal ObjName As String) As Long
    'RemoveHostObjectFromScript = DCF_OneArgString(m_pWebView, 50, "RemoveHostObjectFromScript", ObjName)
    RemoveHostObjectFromScript = dcf(m_pWebView, 50, "RemoveHostObjectFromScript", StrPtr(ObjName))
End Function

'51
'virtual HRESULT STDMETHODCALLTYPE OpenDevToolsWindow( void) = 0;
Public Function OpenDevToolsWindow()
    'Call DispCallFunc(m_pWebView, 51 * LenB(m_pWebView), CC_STDCALL, vbLong, 0, 0, 0, OpenDevToolsWindow)
    Call dcf(m_pWebView, 51, "OpenDevToolsWindow")
End Function

'52
'virtual HRESULT STDMETHODCALLTYPE add_ContainsFullScreenElementChanged(
'    /* [in] */ ICoreWebView2ContainsFullScreenElementChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ContainsFullScreenElementChanged() As Long
    'add_ContainsFullScreenElementChanged = mHelpers.DCF_RegisterHandler(Me, 52, "add_ContainsFullScreenElementChanged", AddressOf ContainsFullScreenElementChanged_Invoke, "ContainsFullScreenElementChanged")
    add_ContainsFullScreenElementChanged = RegisterHandler(52, "Invoke_ContainsFullScreenElementChanged")
End Function
Public Function Invoke_ContainsFullScreenElementChanged(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent ContainsFullScreenElementChanged
    Invoke_ContainsFullScreenElementChanged = 0
End Function
'53
'virtual HRESULT STDMETHODCALLTYPE remove_ContainsFullScreenElementChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ContainsFullScreenElementChanged(token As Long) As Long
    'remove_ContainsFullScreenElementChanged = remove_Handler(Me, 53, token, "remove_ContainsFullScreenElementChanged")
    remove_ContainsFullScreenElementChanged = remove_Handler(Me, 53, m_tokens(52), "remove_ContainsFullScreenElementChanged")
End Function


'54
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ContainsFullScreenElement(
'    /* [retval][out] */ BOOL *containsFullScreenElement) = 0;
Public Property Get ContainsFullScreenElement() As Boolean
    Dim lng_ContainsFullScreenElement As Long
    Dim hr As Long
    hr = dcf(m_pWebView, 54, "ContainsFullScreenElement", VarPtr(lng_ContainsFullScreenElement))
    If lng_ContainsFullScreenElement = 1 Then
        ContainsFullScreenElement = True
    Else
        ContainsFullScreenElement = False
    End If
End Property

'55
' virtual HRESULT STDMETHODCALLTYPE add_WebResourceRequested(
'     /* [in] */ ICoreWebView2WebResourceRequestedEventHandler *eventHandler,
'     /* [out] */ EventRegistrationToken *token) = 0;

' NOTE: To enable WebResourceRequested, simply registering the handler is not enough.
' You MUST call AddWebResourceRequestedFilter (VTable Index 57) first to tell the browser
' which URLs (e.g., "*" for all, or a specific domain) and resource types to intercept.
' If this filter is not registered, the event will never fire.
Public Function add_WebResourceRequested() As Long
    'Dim hr As Long
    'hr = mHelpers.DCF_RegisterHandler(Me, 55, "add_WebResourceRequested", AddressOf WebResourceRequested_Invoke, "WebResourceRequested")
    'Debug.Print "add_WebResourceRequested :" & hr
    'add_WebResourceRequested = hr
    add_WebResourceRequested = RegisterHandler(55, "Invoke_WebResourceRequested")
End Function
Public Function Invoke_WebResourceRequested(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent WebResourceRequested
    Invoke_WebResourceRequested = 0
End Function
'56
'virtual HRESULT STDMETHODCALLTYPE remove_WebResourceRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WebResourceRequested(token As Long) As Long
    'remove_WebResourceRequested = remove_Handler(Me, 56, token, "remove_WebResourceRequested")
    remove_WebResourceRequested = remove_Handler(Me, 56, m_tokens(55), "remove_WebResourceRequested")
End Function

'57
'virtual HRESULT STDMETHODCALLTYPE AddWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;
''' <summary>
''' Adds a filter to intercept specific web resources (Index 57).
''' </summary>
Public Function AddWebResourceRequestedFilter(ByVal uri As String, ByVal resourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT) As Long
    AddWebResourceRequestedFilter = dcf(m_pWebView, 57, "AddWebResourceRequestedFilter", StrPtr(uri), resourceContext)
End Function

'58
'virtual HRESULT STDMETHODCALLTYPE RemoveWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;
''' <summary>
''' Removes an intercept filter (Index 58).
''' </summary>
Public Function RemoveWebResourceRequestedFilter(ByVal uri As String, ByVal resourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT) As Long
    RemoveWebResourceRequestedFilter = dcf(m_pWebView, 58, "RemoveWebResourceRequestedFilter", StrPtr(uri), resourceContext)
End Function

'59
'virtual HRESULT STDMETHODCALLTYPE add_WindowCloseRequested(
'    /* [in] */ ICoreWebView2WindowCloseRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
''' <summary>
''' Registers a handler for when the WebView window requests to close (Index 59).
''' </summary>
Public Function add_WindowCloseRequested() As Long
    'add_WindowCloseRequested = mHelpers.DCF_RegisterHandler(Me, 59, "add_WindowCloseRequested", AddressOf WindowCloseRequested_Invoke, "WindowCloseRequested")
    add_WindowCloseRequested = RegisterHandler(59, "Invoke_WindowCloseRequested")
End Function
Public Function Invoke_WindowCloseRequested(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    RaiseEvent WindowCloseRequested(pThis, sender, args)
    Invoke_WindowCloseRequested = 0
End Function
'60
'virtual HRESULT STDMETHODCALLTYPE remove_WindowCloseRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WindowCloseRequested(token As Long) As Long
    'remove_WindowCloseRequested = remove_Handler(Me, 60, token, "remove_WindowCloseRequested")
    remove_WindowCloseRequested = remove_Handler(Me, 60, m_tokens(59), "remove_WindowCloseRequested")
End Function
'
'};
