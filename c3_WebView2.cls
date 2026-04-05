VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c3_WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c3_WebView2

Option Explicit

Private m_captureHandlerRef As Long

Private m_funcPtrs As Object  ' Scripting.Dictionary: 関数名(小文字) → LongPtr
Private m_vtbls(7 To 60)    As LongPtr  ' vTblIndex → vtble+サンク ブロック
Private m_fakeObjs(7 To 60) As LongPtr  ' vTblIndex → fakeObj ポインタ
Private m_tokens(7 To 60)   As Long  ' vTblIndex → 登録解除用トークン

Private Const E_PENDING As Long = &H8000000A   ' → -2147483638 :COMの標準HRESULT

' vTbleIndex27のAddScriptToExecuteOnDocumetCreatedで取得できるScriptIdを保持しておく場所。
Private m_PendingScript As String  ' Invoke で受け取るまでの一時保管
Private m_IsScriptPending As Boolean  ' True の間は追加登録を拒否
Private Type ScriptEntry
    scriptId As String
    javascript As String
End Type
Private m_Scripts() As ScriptEntry
Private m_ScriptCount As Long

Private m_IsExcecuteScriptPending As Boolean

Private m_pCaptureStream As LongPtr  ' 解放用に保持

Private Type CDPEntry
    requestId As Long
    methodName As String
    fakeObj   As LongPtr
    vtbl      As LongPtr
End Type

Private m_CDPRequests() As CDPEntry
Private m_CDPRequestCount As Long
Private m_CDPNextId As Long


Private Type DTPEventEntry
    eventName As String
    receiverPtr As LongPtr  ' ICoreWebView2DevToolsProtocolEventReceiver*
    fakeObj As LongPtr
    vtbl As LongPtr
    token As LongPtr
End Type

Private m_DTPEvents() As DTPEventEntry
Private m_DTPEventCount As Long


Private m_pStubAddRefRelease As LongPtr  ' AddRef と Release で共用
Private m_pStubQI            As LongPtr
Private Const MEM_COMMIT  As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const MEM_RELEASE As Long = &H8000
' --- API Declarations (for x64) ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

Private m_vtbl_NavigationCompleted  As LongPtr
Private m_fakeObj_NavigationCompleted As LongPtr
Private m_token_NavigationCompleted As LongPtr

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)
' File access mode constants for STGM (Storage Management)
Private Const STGM_READWRITE As Long = &H2
Private Const STGM_CREATE As Long = &H1000
Private Const STGM_SHARE_DENY_NONE As Long = &H40
    ' API to create an IStream from a file path (Unicode version)
    Private Declare PtrSafe Function SHCreateStreamOnFileW Lib "shlwapi.dll" ( _
        ByVal pszFile As LongPtr, _
        ByVal grfMode As Long, _
        ByRef ppStm As LongPtr) As Long

Private Const CC_STDCALL As Long = 4 ' Calling convention for the 3rd argument of DispCallFunc (__stdcall)
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
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long


Public ppWebView2 As LongPtr

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

Public Event WebResourceRequested()

Public LastScriptResult As String

Public Event WindowCloseRequested(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)

Public Col_Handler As Collection 'New Collection
Private m_IsNavigating As Boolean
Private m_IsScriptRunning As Boolean

#If Win64 Then
'Public cHandler As c4_Handler2
#End If

Public Settings As c6_WebView2Settings

Private m_ScriptId As String

Public col_DTPEventReceiver As New Collection

Private erPtr As LongPtr
Private vTblePtrs(4) As LongPtr

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


' --- For internal notifications ---
'Public Sub NotifyNavigationStarting()
'    RaiseEvent NavigationStarting
'End Sub
'Public Sub NotifyContentLoading()
'    RaiseEvent ContentLoading
'End Sub
'Public Sub NotifySourceChanged()
'    RaiseEvent SourceChanged
'End Sub
'Public Sub NotifyHistoryChanged()
'    RaiseEvent HistoryChanged
'End Sub
'Public Sub NotifyNavigationCompleted()
'    m_IsNavigating = False
'    RaiseEvent NavigationCompleted
'End Sub
'Public Sub NotifyFrameNavigationStarting()
'    RaiseEvent FrameNavigationStarting
'End Sub
'Public Sub NotifyFrameNavigationCompleted()
'    RaiseEvent FrameNavigationCompleted
'End Sub
'Public Sub NotifyScriptDialogOpening()
'    RaiseEvent ScriptDialogOpening
'End Sub
'Public Sub NotifyPermissionRequested()
'    RaiseEvent PermissionRequested
'End Sub
'Public Sub NotifyProcessFailed()
'    RaiseEvent ProcessFailed
'End Sub
'Public Sub NotifyWebMessageReceived()
'    RaiseEvent WebMessageReceived("", "", "")
'End Sub
'Public Sub NotifyExecuteScriptCompleted(ByVal jsonResult As String)
'    LastScriptResult = jsonResult
'    m_IsScriptRunning = False
'    RaiseEvent ReceiveScriptResult(jsonResult)
'End Sub
'Public Sub NotifyNewWindowRequested()
'    RaiseEvent NewWindowRequested
'End Sub
'Public Sub NotifyDocumentTitleChanged()
'    RaiseEvent DocumentTitleChanged
'End Sub
'Public Sub NotifyContainsFullScreenElementChanged()
'    RaiseEvent ContainsFullScreenElementChanged
'End Sub
'Public Sub NotifyWebResourceRequested()
'    RaiseEvent WebResourceRequested
'End Sub
'Public Sub NotifyAddScriptToExecuteOnDocumentCreatedCompleted(ByVal errorCode As Long, result As LongPtr)
'    Dim scriptId As String
'    scriptId = PtrToStrW(result)
'    m_ScriptId = scriptId
'    'RaiseEvent AddScriptToExecuteOnDocumentCreatedCompleted
'End Sub
'Public Sub NotifyCapturePreviewCompleted(ByVal errorCode As Long)
'    If errorCode = 0 Then
'        Debug.Print "CapturePreview: Success!"
'    Else
'        Debug.Print "CapturePreview: Failed. ErrorCode: &H" & Hex(errorCode)
'    End If
'
'    RaiseEvent CapturePreviewCompleted(errorCode)
'
'End Sub
'Public Sub NotifyCallDevToolsProtocolMethodCompleted(ByVal errorCode As Long, result As LongPtr)
'    Dim strResult As String
'    strResult = PtrToStrW(result)
'    RaiseEvent CallDevToolsProtocolMethodCompleted(0, "", errorCode, strResult)
'End Sub
'Public Sub NotifyDevToolsProtocolEventReceived(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)
'
'    ' this   : DevToolsProctocoleventReceiverHandler *
'    ' sender : ICoreWebView2 *
'    ' args   : ICoreWebView2DevToolsProtocolEventReceivedEventArgs *
'
''    The parameter object of the corresponding DevToolsProtocol event represented as a JSON string.
''
''    public HRESULT get_ParameterObjectAsJson(LPWSTR * parameterObjectAsJson)
''
''    The caller must free the returned string with CoTaskMemFree. See API Conventions.
'
'    Dim parameterObjectAsJson As String, paramObjPtr As LongPtr
'    parameterObjectAsJson = DCF_GetStringNoArgs(args, 3, "get_ParameterObjectAsJson")
'
'    RaiseEvent DevToolsProtocolEventReceived("", parameterObjectAsJson)
'End Sub
'
'Public Sub NotifyWindowCloseRequested(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)
'    RaiseEvent WindowCloseRequested(this, sender, args)
'End Sub

' --- Cleaning up ---
Public Sub Finalize()
    
    If ppWebView2 = 0 Then Exit Sub
    
    ' 全常駐ハンドラを解除
    remove_Handler Me, 8, m_tokens(7), "remove_NavigationStarting"
    remove_Handler Me, 10, m_tokens(9), "remove_ContentLoading"
    remove_Handler Me, 12, m_tokens(11), "remove_SourceChanged"
    remove_Handler Me, 14, m_tokens(13), "remove_HistoryChanged"
    remove_Handler Me, 16, m_tokens(15), "remove_NavigationCompleted"
    remove_Handler Me, 18, m_tokens(17), "remove_FrameNavigationStarting"
    remove_Handler Me, 20, m_tokens(19), "remove_FrameNavigationCompleted"
    remove_Handler Me, 22, m_tokens(21), "remove_ScriptDialogOpening"
    remove_Handler Me, 24, m_tokens(23), "remove_PermissionRequested"
    remove_Handler Me, 26, m_tokens(25), "remove_ProcessFailed"
    remove_Handler Me, 35, m_tokens(34), "remove_WebMessageReceived"
    remove_Handler Me, 45, m_tokens(44), "remove_NewWindowRequested"
    remove_Handler Me, 47, m_tokens(46), "remove_DocumentTitleChanged"
    remove_Handler Me, 53, m_tokens(52), "remove_ContainsFullScreenElementChanged"
    remove_Handler Me, 56, m_tokens(55), "remove_WebResourceRequested"
    remove_Handler Me, 60, m_tokens(59), "remove_WindowCloseRequested"

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

    ppWebView2 = 0
End Sub

Private Sub Class_Initialize()
    m_pStubQI = CreateStubThunk_QI()
    m_pStubAddRefRelease = CreateStubThunk_AddRefRelease()

    'BuildFuncPtrCache

End Sub

Private Sub Class_Terminate()
    'Finalize
    Debug.Print "c3_WebView2 Terminated."
End Sub

' --- Standby helper ---
Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim start As Double: start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - start > TimeoutSec Then Exit Do
    Loop
End Sub


'MIDL_INTERFACE ("76eceacb-0462-4d94-ac83-423a6793775e")
'ICoreWebView2:          Public IUnknown
'{
'public:

'3
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Settings(
'    /* [retval][out] */ ICoreWebView2Settings **settings) = 0;
Public Function get_Settings() As Long
    Dim psettings As LongPtr
    Dim hr As Long
    hr = dcf(ppWebView2, 3, "get_Settings", VarPtr(psettings))
    If hr = 0 Then
        Debug.Print "psettings:" & psettings
        Set Settings = New c6_WebView2Settings
        Settings.Init psettings
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
    hr = dcf(ppWebView2, 4, "get_Source", VarPtr(pStr))
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
    NavigateAsync = dcf(ppWebView2, 5, "Navigate", StrPtr(uri))
End Function
Public Sub NavigateSync(ByVal url As String, Optional ByVal TimeoutSec As Double = 10)
    ' Synchronized version: Wait until loading is complete.
    m_IsNavigating = True
    If NavigateAsync(url) = 0 Then
        WaitNavigation TimeoutSec
    End If
End Sub

'6
'virtual HRESULT STDMETHODCALLTYPE NavigateToString(
'    /* [in] */ LPCWSTR htmlContent) = 0;
Public Function NavigateToString(htmlContent As String) As Long
    'Renders the specified HTML code.
    NavigateToString = dcf(ppWebView2, 6, "NavigateToString", StrPtr(htmlContent))
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
    AddScriptToExecuteOnDocumentCreated = dcf(ppWebView2, 27, "AddScriptToExecuteOnDocumentCreated", _
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
    hr = dcf(ppWebView2, 28, "RemoveScriptToExecuteOnDocumentCreated", StrPtr(scriptId))
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
    ExecuteScriptAsync = dcf(ppWebView2, 29, "ExecuteScriptAsync", _
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
    hr = dcf(ppWebView2, 30, "CapturePreview", format, pStream, m_fakeObjs(30))
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
    Reload = dcf(ppWebView2, 31, "Reload")
End Function

'32
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsJson(
'    /* [in] */ LPCWSTR webMessageAsJson) = 0;
Public Function PostWebMessageAsJson(ByVal jsonString As String) As String
    PostWebMessageAsJson = dcf(ppWebView2, 32, "PostWebMessageAsJson", StrPtr(jsonString))
End Function

'33
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsString(
'    /* [in] */ LPCWSTR webMessageAsString) = 0;
Public Function PostWebMessageAsString(ByVal webMessage As String) As Long
    PostWebMessageAsString = dcf(ppWebView2, 33, "PostWebMessageAsString", StrPtr(webMessage))
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
    hr = dcf(ppWebView2, 36, "CallDevToolsProtocolMethod", _
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
    hr = dcf(ppWebView2, 37, "", VarPtr(value))
    BrowserProcessId = value
End Property

'38
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoBack(
'    /* [retval][out] */ BOOL *canGoBack) = 0;
Public Property Get CanGoBack() As Boolean
    Dim lcanGoBack As Long
    Dim hr As Long
    hr = dcf(ppWebView2, 38, "get_canGoBack", VarPtr(lcanGoBack))
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
    hr = dcf(ppWebView2, 39, "get_CanGoForward", VarPtr(lCanGoForward))
    If lCanGoForward = 1 Then
        CanGoForward = True
    Else
        CanGoForward = False
    End If
End Property

'40
'virtual HRESULT STDMETHODCALLTYPE GoBack( void) = 0;
Public Function GoBack() As Long
    GoBack = dcf(ppWebView2, 40, "GoBack")
End Function

'41
'virtual HRESULT STDMETHODCALLTYPE GoForward( void) = 0;
Public Function GoForward() As Long
    GoForward = dcf(ppWebView2, 41, "GoForward")
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
    hr = dcf(ppWebView2, 42, "GetDevToolsProtocolEventReceiver", StrPtr(eventName), VarPtr(receiverPtr))
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
    Stop_ = dcf(ppWebView2, 43, "Stop_")
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
    hr = dcf(ppWebView2, 48, "get_DocumentTitle", VarPtr(pStr))
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
    AddHostObjectToScript = dcf(ppWebView2, 49, "AddHostObjectToScript", StrPtr(Name), VarPtr(vObj))
    'AddHostObjectToScript = DCF_TwoArgsStringAndObject(ppWebView2, 49, "AddHostObjectToScript", Name, Instance)
End Function

'50
'virtual HRESULT STDMETHODCALLTYPE RemoveHostObjectFromScript(
'    /* [in] */ LPCWSTR name) = 0;
Public Function RemoveHostObjectFromScript(ByVal ObjName As String) As Long
    'RemoveHostObjectFromScript = DCF_OneArgString(ppWebView2, 50, "RemoveHostObjectFromScript", ObjName)
    RemoveHostObjectFromScript = dcf(ppWebView2, 50, "RemoveHostObjectFromScript", StrPtr(ObjName))
End Function

'51
'virtual HRESULT STDMETHODCALLTYPE OpenDevToolsWindow( void) = 0;
Public Function OpenDevToolsWindow()
    'Call DispCallFunc(ppWebView2, 51 * LenB(ppWebView2), CC_STDCALL, vbLong, 0, 0, 0, OpenDevToolsWindow)
    Call dcf(ppWebView2, 51, "OpenDevToolsWindow")
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
    hr = dcf(ppWebView2, 54, "ContainsFullScreenElement", VarPtr(lng_ContainsFullScreenElement))
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
    AddWebResourceRequestedFilter = dcf(ppWebView2, 57, "AddWebResourceRequestedFilter", StrPtr(uri), resourceContext)
End Function

'58
'virtual HRESULT STDMETHODCALLTYPE RemoveWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;
''' <summary>
''' Removes an intercept filter (Index 58).
''' </summary>
Public Function RemoveWebResourceRequestedFilter(ByVal uri As String, ByVal resourceContext As COREWEBVIEW2_WEB_RESOURCE_CONTEXT) As Long
    RemoveWebResourceRequestedFilter = dcf(ppWebView2, 58, "RemoveWebResourceRequestedFilter", StrPtr(uri), resourceContext)
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

''' <summary>
''' Registers the NavigationCompleted event handler
''' </summary>
#If Win64 Then
'Public Function AddNavigationCompletedHandler(ByVal Handler As c4_Handler2) As Long
'
'    If Col_Handler Is Nothing Then Set Col_Handler = New Collection
'    Handler.ppWebView2 = ppWebView2
'
'    Dim tkn As LongPtr ' Variable to receive the EventRegistrationToken for unregistering
'    Dim hr As Long
'    hr = dcf(ppWebView2, 15, "AddNavigationCompletedHandler", Handler.Pointer, VarPtr(tkn))
'    If hr = 0 Then
'        Handler.token = tkn
'        Col_Handler.Add Handler, CStr(tkn)
'        Debug.Print "--- Event Handler Registered Successfully ---"
'        Debug.Print "Token:", tkn
'    End If
'    AddNavigationCompletedHandler = hr
'End Function
#End If














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
    RegisterHandler = dcf(ppWebView2, vtblIndex, methodName, _
        m_fakeObjs(vtblIndex), VarPtr(m_tokens(vtblIndex)))

End Function

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

Private Function CreateStubThunk_AddRefRelease() As LongPtr
    ' AddRef / Release 共用
    ' rcx=pThis のみ、rax=1 を返すだけ
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function

    Dim b(7) As Byte
    ' mov eax, 1
    b(0) = &HB8: b(1) = &H1: b(2) = &H0: b(3) = &H0: b(4) = &H0
    ' ret
    b(5) = &HC3

    RtlMoveMemory addr, VarPtr(b(0)), 8
    CreateStubThunk_AddRefRelease = addr
End Function

Private Function CreateStubThunk_QI() As LongPtr
    ' QI: rcx=pThis, rdx=riid, r8=ppv
    ' *ppv = pThis, return S_OK(0)
    Dim addr As LongPtr
    addr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    If addr = 0 Then Exit Function

    Dim b(7) As Byte
    ' mov [r8], rcx   ; *ppv = pThis
    b(0) = &H49: b(1) = &H89: b(2) = &H8
    ' xor eax, eax    ; rax = 0 (S_OK)
    b(3) = &H31: b(4) = &HC0
    ' ret
    b(5) = &HC3

    RtlMoveMemory addr, VarPtr(b(0)), 8
    CreateStubThunk_QI = addr
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


'Private Sub FreeHandlerVTable(ByVal vtblAddr As LongPtr)
'    If vtblAddr = 0 Then Exit Sub
'    VirtualFree vtblAddr, 0, MEM_RELEASE
'End Sub

'Private Function CreateThunk(ByVal targetObjPtr As LongPtr, _
'                             ByVal funcPtr As LongPtr) As LongPtr
'
'    Dim thunkAddr As LongPtr
'    thunkAddr = VirtualAlloc(0, 32, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
'    If thunkAddr = 0 Then Exit Function
'
'    ' x64 呼び出し規約では WebView2 → Invoke のとき
'    '   rcx = pThis(fakeObj),  rdx = sender,  r8 = args
'    ' これをクラスメソッドの呼び出し規約に合わせて並べ替える：
'    '   rcx = ObjPtr(Me),  rdx = pThis,  r8 = sender,  r9 = args
'
'    Dim b(31) As Byte
'    Dim i As Long
'
'    ' mov r9, r8        ; args   → 4番目の引数へ退避
'    b(i) = &H4D: b(i + 1) = &H89: b(i + 2) = &HC1: i = i + 3
'    ' mov r8, rdx       ; sender → 3番目の引数へ退避
'    b(i) = &H49: b(i + 1) = &H89: b(i + 2) = &HD0: i = i + 3
'    ' mov rdx, rcx      ; pThis  → 2番目の引数へ退避
'    b(i) = &H48: b(i + 1) = &H89: b(i + 2) = &HCA: i = i + 3
'    ' mov rcx, imm64    ; ObjPtr(Me) → 1番目の引数へ
'    b(i) = &H48: b(i + 1) = &HB9: i = i + 2
'    RtlMoveMemory VarPtr(b(i)), VarPtr(targetObjPtr), 8: i = i + 8
'    ' mov rax, imm64    ; ジャンプ先アドレス
'    b(i) = &H48: b(i + 1) = &HB8: i = i + 2
'    RtlMoveMemory VarPtr(b(i)), VarPtr(funcPtr), 8: i = i + 8
'    ' jmp rax
'    b(i) = &HFF: b(i + 1) = &HE0
'
'    RtlMoveMemory thunkAddr, VarPtr(b(0)), 32
'
'    CreateThunk = thunkAddr
'
'End Function

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
            Debug.Print "登録", i, bstrName
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

'Public Sub RebuildFuncPtrCache()
'    BuildFuncPtrCache
'End Sub

' Helper: PtrToStrW (Converts Unicode pointer to VBA String)
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

Private Function remove_Handler(WebView2 As c3_WebView2, vtblIndex As Long, token As Long, HandlerName As String) As Long
    Dim ppWebView2 As LongPtr: ppWebView2 = WebView2.ppWebView2
    Dim hr As Long, res As Variant
    Dim vtoken As Variant: vtoken = token
    Dim args(0) As Variant: args(0) = vtoken
    Dim argsPtr(0) As Variant: argsPtr(0) = VarPtr(args(0))
    Dim argsType(0) As Integer: argsType(0) = vbLong
    hr = DispCallFunc(ppWebView2, vtblIndex * LenB(ppWebView2), CC_STDCALL, vbLong, 1, argsType(0), argsPtr(0), res)
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

''' <summary>
''' Recursively creates directories if they do not exist, including deep subfolders.
''' </summary>
''' <param name="folderPath">The absolute path of the directory to create.</param>
Private Sub CreateDeepFolder(ByVal folderPath As String)
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

