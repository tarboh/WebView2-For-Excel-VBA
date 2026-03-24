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

Public ppWebView2 As LongPtr
Public Event NavigationStarting()
Public Event ContentLoading()
Public Event SourceChanged()
Public Event HistoryChanged()
Public Event NavigationCompleted()
Public Event FrameNavigationStarting()
Public Event FrameNavigationCompleted()
Public Event ScriptDialogOpening()
Public Event PermissionRequested()
Public Event ProcessFailed()
Public Event WebMessageReceived()
Public Event NewWindowRequested()
Public Event DocumentTitleChanged()
Public Event ContainsFullScreenElementChanged()
Public Event WebResourceRequested()
Public Event AddScriptToExecuteOnDocumentCreatedCompleted()
Public Event CapturePreviewCompleted(ByVal errorCode As Long)

Public Event ReceiveScriptResult(ByVal result As String)
Public LastScriptResult As String


Public Col_Handler As New Collection
Private m_IsNavigating As Boolean
Private m_IsScriptRunning As Boolean

#If Win64 Then
Public cHandler As c4_Handler2
#End If

Public Settings As c6_WebView2Settings

Private m_ScriptId As String

' --- For internal notifications ---
Public Sub NotifyNavigationStarting()
    RaiseEvent NavigationStarting
End Sub
Public Sub NotifyContentLoading()
    RaiseEvent ContentLoading
End Sub
Public Sub NotifySourceChanged()
    RaiseEvent SourceChanged
End Sub
Public Sub NotifyHistoryChanged()
    RaiseEvent HistoryChanged
End Sub
Public Sub NotifyNavigationCompleted()
    m_IsNavigating = False
    RaiseEvent NavigationCompleted
End Sub
Public Sub NotifyFrameNavigationStarting()
    RaiseEvent FrameNavigationStarting
End Sub
Public Sub NotifyFrameNavigationCompleted()
    RaiseEvent FrameNavigationCompleted
End Sub
Public Sub NotifyScriptDialogOpening()
    RaiseEvent ScriptDialogOpening
End Sub
Public Sub NotifyPermissionRequested()
    RaiseEvent PermissionRequested
End Sub
Public Sub NotifyProcessFailed()
    RaiseEvent ProcessFailed
End Sub
Public Sub NotifyWebMessageReceived()
    RaiseEvent WebMessageReceived
End Sub
Public Sub NotifyExecuteScriptCompleted(ByVal jsonResult As String)
    LastScriptResult = jsonResult
    m_IsScriptRunning = False
    RaiseEvent ReceiveScriptResult(jsonResult)
End Sub
Public Sub NotifyNewWindowRequested()
    RaiseEvent NewWindowRequested
End Sub
Public Sub NotifyDocumentTitleChanged()
    RaiseEvent DocumentTitleChanged
End Sub
Public Sub NotifyContainsFullScreenElementChanged()
    RaiseEvent ContainsFullScreenElementChanged
End Sub
Public Sub NotifyWebResourceRequested()
    RaiseEvent WebResourceRequested
End Sub
Public Sub NotifyAddScriptToExecuteOnDocumentCreatedCompleted(ByVal errorCode As Long, result As LongPtr)
    Dim ScriptId As String
    ScriptId = PtrToStrW(result)
    m_ScriptId = ScriptId
    RaiseEvent AddScriptToExecuteOnDocumentCreatedCompleted
End Sub
Public Sub NotifyCapturePreviewCompleted(ByVal errorCode As Long)
    If errorCode = 0 Then
        Debug.Print "CapturePreview: Success!"
    Else
        Debug.Print "CapturePreview: Failed. ErrorCode: &H" & Hex(errorCode)
    End If
    
    RaiseEvent CapturePreviewCompleted(errorCode)
End Sub

' --- Cleaning up ---
Public Sub Finalize()
    Dim h As c4_Handler
    On Error Resume Next
    For Each h In Col_Handler
        UnregisterInstance h.Pointer
    Next h
    Set Col_Handler = New Collection
End Sub

Private Sub Class_Terminate()
    Finalize
    Debug.Print "c3_WebView2 Terminated."
End Sub

' --- Standby helper ---
Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim Start As Double: Start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - Start > TimeoutSec Then Exit Do
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
    Dim args(0) As Variant: args(0) = VarPtr(psettings)
    Dim args_Ptr(0) As LongPtr: args_Ptr(0) = VarPtr(args(0))
    Dim args_Type(0) As Integer: args_Type(0) = vbLongPtr
    Dim hr As Long, res As Variant
    
    hr = DispCallFunc(ppWebView2, 3 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        Debug.Print "get_Settings Result:" & res
        get_Settings = res
        If res = 0 Then
            Debug.Print "psettings:" & psettings
            Set Settings = New c6_WebView2Settings
            Settings.Init psettings
        End If
    Else
        Debug.Print "DispCallFunc(get_Settings) Failed. HR:" & hr
    End If
End Function

'4
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Source(
'    /* [retval][out] */ LPWSTR *uri) = 0;
Public Property Get Source() As String
    'Get current URL
    Source = DCF_GetStringNoArgs(ppWebView2, 4, "get_Source")
End Property

'5
'virtual HRESULT STDMETHODCALLTYPE Navigate(
'    /* [in] */ LPCWSTR uri) = 0;
Public Function NavigateAsync(ByVal uri As String) As Long
    'Navigate to the specified URI.
    NavigateAsync = DCF_OneArgString(ppWebView2, 5, "Navigate", uri)
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
    NavigateToString = DCF_OneArgString(ppWebView2, 6, "NavigateToString", htmlContent)
End Function

'7
'virtual HRESULT STDMETHODCALLTYPE add_NavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_NavigationStarting() As Long
    add_NavigationStarting = DCF_RegisterHandler(Me, 7, "add_NavigationStarting", AddressOf NavigationStarting_Invoke, "NavigationStartingHandler")
End Function
Public Function add_NavigationStarting_1(Handler As c4_Handler) As Long
    'I want to modify it so that it receives a token by passing a vTble pointer to the event handler.
End Function

'8
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationStarting(token As Long) As Long
    remove_NavigationStarting = remove_Handler(Me, 8, token, "remove_NavigationStarting")
End Function

'9
'virtual HRESULT STDMETHODCALLTYPE add_ContentLoading(
'    /* [in] */ ICoreWebView2ContentLoadingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ContentLoading() As Long
    add_ContentLoading = DCF_RegisterHandler(Me, 9, "add_ContentLoading", AddressOf ContentLoading_Invoke, "ContentLoadingHandler")
End Function

'10
'virtual HRESULT STDMETHODCALLTYPE remove_ContentLoading(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ContentLoading(token As Long) As Long
    remove_ContentLoading = remove_Handler(Me, 10, token, "remove_ContentLoading")
End Function

'11
'virtual HRESULT STDMETHODCALLTYPE add_SourceChanged(
'    /* [in] */ ICoreWebView2SourceChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_SourceChanged() As Long
    add_SourceChanged = DCF_RegisterHandler(Me, 11, "add_SourceChanged", AddressOf SourceChanged_Invoke, "SourceChangedHandler")
End Function

'12
'virtual HRESULT STDMETHODCALLTYPE remove_SourceChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_SourceChanged(token As Long) As Long
    remove_SourceChanged = remove_Handler(Me, 12, token, "remove_SourceChanged")
End Function

'13
'virtual HRESULT STDMETHODCALLTYPE add_HistoryChanged(
'    /* [in] */ ICoreWebView2HistoryChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_HistoryChanged() As Long
    add_HistoryChanged = DCF_RegisterHandler(Me, 13, "add_HistoryChanged", AddressOf HistoryChanged_Invoke, "HistoryChangedHandler")
End Function

'14
'virtual HRESULT STDMETHODCALLTYPE remove_HistoryChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_HistoryChanged(token As Long) As Long
    remove_HistoryChanged = remove_Handler(Me, 14, token, "remove_HistoryChanged")
End Function

'15
'virtual HRESULT STDMETHODCALLTYPE add_NavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'add_NavigationCompleted
Public Function add_NavigationCompleted() As Long
    add_NavigationCompleted = mHelpers.DCF_RegisterHandler(Me, 15, "add_NavigationCompleted", AddressOf NavCompleted_Invoke, "NavigationCompleted")
End Function

'16
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationCompleted(token As Long) As Long
    remove_NavigationCompleted = remove_Handler(Me, 16, token, "remove_NavigationCompleted")
End Function

'17
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationStarting() As Long
    add_FrameNavigationStarting = mHelpers.DCF_RegisterHandler(Me, 17, "add_FrameNavigationStarting", AddressOf FrameNavigationStarting_Invoke, "FrameNavigationStarting")
End Function

'18
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationStarting(token As Long) As Long
    remove_FrameNavigationStarting = remove_Handler(Me, 18, token, "remove_FrameNavigationStarting")
End Function

'19
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationCompleted() As Long
    add_FrameNavigationCompleted = mHelpers.DCF_RegisterHandler(Me, 19, "add_FrameNavigationCompleted", AddressOf FrameNavigationCompleted_Invoke, "FrameNavigationCompleted")
End Function

'20
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationCompleted(token As Long) As Long
    remove_FrameNavigationCompleted = remove_Handler(Me, 20, token, "remove_FrameNavigationCompleted")
End Function

'21
'virtual HRESULT STDMETHODCALLTYPE add_ScriptDialogOpening(
'    /* [in] */ ICoreWebView2ScriptDialogOpeningEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'   ** This event is triggered only if the ICoreWebView2Settings::AreDefaultScriptDialogsEnabled property is set to FALSE.
Public Function add_ScriptDialogOpening() As Long
    add_ScriptDialogOpening = mHelpers.DCF_RegisterHandler(Me, 21, "add_ScriptDialogOpening", AddressOf ScriptDialogOpening_Invoke, "ScriptDialogOpening")
End Function

'22
'virtual HRESULT STDMETHODCALLTYPE remove_ScriptDialogOpening(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ScriptDialogOpening(token As Long) As Long
    remove_ScriptDialogOpening = remove_Handler(Me, 22, token, "remove_ScriptDialogOpening")
End Function

'23
'virtual HRESULT STDMETHODCALLTYPE add_PermissionRequested(
'    /* [in] */ ICoreWebView2PermissionRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_PermissionRequested() As Long
    add_PermissionRequested = mHelpers.DCF_RegisterHandler(Me, 23, "add_PermissionRequested", AddressOf PermissionRequested_Invoke, "PermissionRequested")
End Function

'24
'virtual HRESULT STDMETHODCALLTYPE remove_PermissionRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_PermissionRequested(token As Long) As Long
    remove_PermissionRequested = remove_Handler(Me, 24, token, "remove_PermissionRequested")
End Function

'25
'virtual HRESULT STDMETHODCALLTYPE add_ProcessFailed(
'    /* [in] */ ICoreWebView2ProcessFailedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ProcessFailed() As Long
    add_ProcessFailed = mHelpers.DCF_RegisterHandler(Me, 25, "add_ProcessFailed", AddressOf ProcessFailed_Invoke, "ProcessFailed")
End Function

'26
'virtual HRESULT STDMETHODCALLTYPE remove_ProcessFailed(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ProcessFailed(token As Long) As Long
    remove_ProcessFailed = remove_Handler(Me, 26, token, "remove_ProcessFailed")
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
    Dim pStr As LongPtr: pStr = StrPtr(javascript)
    Dim args(1) As Variant, args_Ptr(1) As LongPtr, args_Type(1) As Integer
    
    args(0) = pStr
    args_Ptr(0) = VarPtr(args(0))
    args_Type(0) = vbLongPtr
    
    Dim Handler As c4_Handler
    Set Handler = New c4_Handler
    Col_Handler.Add Handler
    Call Handler.CreateVTble(AddressOf AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke, ppWebView2)
    Handler.Namae = "AddScriptToExecuteOnDocumentCreated"
    
    args(1) = Handler.Pointer
    args_Ptr(1) = VarPtr(args(1))
    args_Type(1) = vbLongPtr
    
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 27 * LenB(ppWebView2), CC_STDCALL, vbLong, 2, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        If res = 0 Then
            Debug.Print "AddScriptToExecuteOnDocumentCreated Success."
            RegisterInstance Handler.Pointer, Me
        Else
            Debug.Print "AddScriptToExecuteOnDocumentCreated Failed. Res:" & res
        End If
    Else
        Debug.Print "AddScriptToExecuteOnDocumentCreated Failed. Hr:" & hr
    End If
    
End Function

'28
'virtual HRESULT STDMETHODCALLTYPE RemoveScriptToExecuteOnDocumentCreated(
'    /* [in] */ LPCWSTR id) = 0;
Public Function RemoveScriptToExecuteOnDocumentCreated(ScriptId As String) As Long
    Dim args(0) As Variant:      args(0) = StrPtr(ScriptId)
    Dim args_Ptr(0)  As LongPtr: args_Ptr(0) = VarPtr(args(0))
    Dim args_Type(0) As Integer: args_Type(0) = vbLongPtr
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 28 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        If res = 0 Then
            Debug.Print "RemoveScriptToExecuteOnDocumentCreated Success."
        Else
            Debug.Print "RemoveScriptToExecuteOnDocumentCreated Failed. Res:" & res
        End If
    Else
        Debug.Print "RemoveScriptToExecuteOnDocumentCreated Failed. Hr:" & hr
    End If
End Function

Public Property Get ScriptId() As String
    ScriptId = m_ScriptId
End Property

'29
'virtual HRESULT STDMETHODCALLTYPE ExecuteScript(
'    /* [in] */ LPCWSTR javaScript,
'    /* [in] */ ICoreWebView2ExecuteScriptCompletedHandler *handler) = 0;
Public Function ExecuteScriptAsync(ByVal script As String) As Long
' Asynchronous version: Simply request execution and return. Receive results via events.
    ExecuteScriptAsync = DCF_RegisterHandlerWithString(Me, 29, "ExecuteScript", script, AddressOf ExecuteScript_Invoke)
End Function

' Synchronized version: Receives the execution result directly as a return value.
Public Function ExecuteScriptSync(ByVal script As String, Optional ByVal TimeoutSec As Double = 5) As String
    m_IsScriptRunning = True
    ExecuteScriptAsync script
    
    ' Dedicated wait loop
    Dim Start As Double: Start = Timer
    Do While m_IsScriptRunning
        Debug.Print "waiting ExecuteScriptSync"
        DoEvents
        If Timer - Start > TimeoutSec Then
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
Public Function CapturePreview(ByVal format As Long, ByVal pStream As LongPtr) As Long
    ' Å¶ format: 0=PNG, 1=JPEG
    ' Å¶ pStream: Pointer to the pre-created IStream
    
    Dim args(2) As Variant, args_Ptr(2) As LongPtr, args_Type(2) As Integer
    
    ' 1. Fomat (Enum / Long)
    args(0) = format
    args_Ptr(0) = VarPtr(args(0))
    args_Type(0) = vbLong
    
    ' 2. Pointer to the pre-created IStream
    args(1) = pStream
    args_Ptr(1) = VarPtr(args(1))
    args_Type(1) = vbLongPtr
    
    ' 3. Completion handler (c4_Handler)
    Dim Handler As c4_Handler
    Set Handler = New c4_Handler
    Col_Handler.Add Handler, "CapturePreviewCompletedHandler"
    Call Handler.CreateVTble(AddressOf CapturePreviewCompletedHandler_Invoke, ppWebView2)
    
    args(2) = Handler.Pointer
    args_Ptr(2) = VarPtr(args(2))
    args_Type(2) = vbLongPtr
    
    ' Call ICoreWebView2::CapturePreview (VTable Index 30)
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 30 * LenB(ppWebView2), CC_STDCALL, vbLong, 3, args_Type(0), args_Ptr(0), res)
    
    If hr = 0 Then
        RegisterInstance Handler.Pointer, Me
        CapturePreview = res
    End If
End Function


'31
'virtual HRESULT STDMETHODCALLTYPE Reload( void) = 0;
Public Function Reload() As Long
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 31 * LenB(ppWebView2), CC_STDCALL, vbLong, 0, 0, 0, res)
    Reload = res
End Function

'32
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsJson(
'    /* [in] */ LPCWSTR webMessageAsJson) = 0;
Public Function PostWebMessageAsJson(ByVal jsonString As String) As String
    Dim args(0) As Variant, args_Type(0) As Integer, args_Ptr(0) As LongPtr
    args(0) = StrPtr(jsonString)
    args_Type(0) = vbLongPtr
    args_Ptr(0) = VarPtr(args(0))
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 32 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        If res = 0 Then
            Debug.Print "PostWebMessageAsJson success."
        Else
            Debug.Print "PostWebMessageAsJson failed. res:" & res
        End If
        PostWebMessageAsJson = res
    Else
        Debug.Print "PostWebMessageAsJson failed. hr:" & hr
        PostWebMessageAsJson = hr
    End If
End Function

'33
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsString(
'    /* [in] */ LPCWSTR webMessageAsString) = 0;
Public Function PostWebMessageAsString(ByVal webMessage As String) As Long
    Dim args(0) As Variant, args_Type(0) As Integer, args_Ptr(0) As LongPtr
    args(0) = StrPtr(webMessage)
    args_Type(0) = vbLongPtr
    args_Ptr(0) = VarPtr(args(0))
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 33 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        If res = 0 Then
            Debug.Print "PostWebMessageAsString success."
        Else
            Debug.Print "PostWebMessageAsString failed. res:" & res
        End If
        PostWebMessageAsString = res
    Else
        Debug.Print "PostWebMessageAsJson failed. hr:" & hr
        PostWebMessageAsString = hr
    End If
End Function

'34
'virtual HRESULT STDMETHODCALLTYPE add_WebMessageReceived(
'    /* [in] */ ICoreWebView2WebMessageReceivedEventHandler *handler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_WebMessageReceived() As Long
    add_WebMessageReceived = mHelpers.DCF_RegisterHandler(Me, 34, "add_WebMessageReceived", AddressOf WebMessageReceived_Invoke, "WebMessageReceived")
End Function

'35
'virtual HRESULT STDMETHODCALLTYPE remove_WebMessageReceived(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WebMessageReceived(token As Long) As Long
    remove_WebMessageReceived = remove_Handler(Me, 35, token, "remove_WebMessageReceived")
End Function

'36
'virtual HRESULT STDMETHODCALLTYPE CallDevToolsProtocolMethod(
'    /* [in] */ LPCWSTR methodName,
'    /* [in] */ LPCWSTR parametersAsJson,
'    /* [in] */ ICoreWebView2CallDevToolsProtocolMethodCompletedHandler *handler) = 0;

'37
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_BrowserProcessId(
'    /* [retval][out] */ UINT32 *value) = 0;

'38
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoBack(
'    /* [retval][out] */ BOOL *canGoBack) = 0;

'39
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoForward(
'    /* [retval][out] */ BOOL *canGoForward) = 0;

'40
'virtual HRESULT STDMETHODCALLTYPE GoBack( void) = 0;

'41
'virtual HRESULT STDMETHODCALLTYPE GoForward( void) = 0;

'42
'virtual HRESULT STDMETHODCALLTYPE GetDevToolsProtocolEventReceiver(
'    /* [in] */ LPCWSTR eventName,
'    /* [retval][out] */ ICoreWebView2DevToolsProtocolEventReceiver **receiver) = 0;

'43
'virtual HRESULT STDMETHODCALLTYPE Stop( void) = 0;

'44
'virtual HRESULT STDMETHODCALLTYPE add_NewWindowRequested(
'    /* [in] */ ICoreWebView2NewWindowRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_NewWindowRequested() As Long
    add_NewWindowRequested = mHelpers.DCF_RegisterHandler(Me, 44, "add_NewWindowRequested", AddressOf NewWindowRequested_Invoke, "NewWindowRequested")
End Function
'45
'virtual HRESULT STDMETHODCALLTYPE remove_NewWindowRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NewWindowRequested(token As Long) As Long
    remove_NewWindowRequested = remove_Handler(Me, 45, token, "remove_NewWindowRequested")
End Function

'46
'virtual HRESULT STDMETHODCALLTYPE add_DocumentTitleChanged(
'    /* [in] */ ICoreWebView2DocumentTitleChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_DocumentTitleChanged() As Long
    add_DocumentTitleChanged = mHelpers.DCF_RegisterHandler(Me, 46, "add_DocumentTitleChanged", AddressOf DocumentTitleChanged_Invoke, "DocumentTitleChanged")
End Function

'47
'virtual HRESULT STDMETHODCALLTYPE remove_DocumentTitleChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_DocumentTitleChanged(token As Long) As Long
    remove_DocumentTitleChanged = remove_Handler(Me, 47, token, "remove_DocumentTitleChanged")
End Function

'48
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_DocumentTitle(
'    /* [retval][out] */ LPWSTR *title) = 0;
Public Property Get DocumentTitle() As String
    DocumentTitle = DCF_GetStringNoArgs(ppWebView2, 48, "get_DocumentTitle")
End Property

'49
'virtual HRESULT STDMETHODCALLTYPE AddHostObjectToScript(
'    /* [in] */ LPCWSTR name,
'    /* [in] */ VARIANT *object) = 0;
Public Function AddHostObjectToScript(ByVal Name As String, ByVal Instance As Object) As Long
    AddHostObjectToScript = DCF_TwoArgsStringAndObject(ppWebView2, 49, "AddHostObjectToScript", Name, Instance)
End Function

'50
'virtual HRESULT STDMETHODCALLTYPE RemoveHostObjectFromScript(
'    /* [in] */ LPCWSTR name) = 0;
Public Function RemoveHostObjectFromScript(ByVal ObjName As String) As Long
    RemoveHostObjectFromScript = DCF_OneArgString(ppWebView2, 50, "RemoveHostObjectFromScript", ObjName)
End Function

'51
'virtual HRESULT STDMETHODCALLTYPE OpenDevToolsWindow( void) = 0;
Public Function OpenDevToolsWindow()
    Call DispCallFunc(ppWebView2, 51 * LenB(ppWebView2), CC_STDCALL, vbLong, 0, 0, 0, OpenDevToolsWindow)
End Function

'52
'virtual HRESULT STDMETHODCALLTYPE add_ContainsFullScreenElementChanged(
'    /* [in] */ ICoreWebView2ContainsFullScreenElementChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ContainsFullScreenElementChanged() As Long
    add_ContainsFullScreenElementChanged = mHelpers.DCF_RegisterHandler(Me, 52, "add_ContainsFullScreenElementChanged", AddressOf ContainsFullScreenElementChanged_Invoke, "ContainsFullScreenElementChanged")
End Function

'53
'virtual HRESULT STDMETHODCALLTYPE remove_ContainsFullScreenElementChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ContainsFullScreenElementChanged(token As Long) As Long
    remove_ContainsFullScreenElementChanged = remove_Handler(Me, 53, token, "remove_ContainsFullScreenElementChanged")
End Function


'54
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ContainsFullScreenElement(
'    /* [retval][out] */ BOOL *containsFullScreenElement) = 0;


'55
' virtual HRESULT STDMETHODCALLTYPE add_WebResourceRequested(
'     /* [in] */ ICoreWebView2WebResourceRequestedEventHandler *eventHandler,
'     /* [out] */ EventRegistrationToken *token) = 0;

' NOTE: To enable WebResourceRequested, simply registering the handler is not enough.
' You MUST call AddWebResourceRequestedFilter (VTable Index 57) first to tell the browser
' which URLs (e.g., "*" for all, or a specific domain) and resource types to intercept.
' If this filter is not registered, the event will never fire.
Public Function add_WebResourceRequested() As Long
    add_WebResourceRequested = mHelpers.DCF_RegisterHandler(Me, 55, "add_WebResourceRequested", AddressOf WebResourceRequested_Invoke, "WebResourceRequested")
End Function

'56
'virtual HRESULT STDMETHODCALLTYPE remove_WebResourceRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WebResourceRequested(token As Long) As Long
    remove_WebResourceRequested = remove_Handler(Me, 56, token, "remove_WebResourceRequested")
End Function

'57
'virtual HRESULT STDMETHODCALLTYPE AddWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;

'58
'virtual HRESULT STDMETHODCALLTYPE RemoveWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;

'59
'virtual HRESULT STDMETHODCALLTYPE add_WindowCloseRequested(
'    /* [in] */ ICoreWebView2WindowCloseRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'60
'virtual HRESULT STDMETHODCALLTYPE remove_WindowCloseRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_WindowCloseRequested(token As Long) As Long
    remove_WindowCloseRequested = remove_Handler(Me, 60, token, "remove_WindowCloseRequested")
End Function
'
'};


''' <summary>
''' Registers the NavigationCompleted event handler
''' </summary>
#If Win64 Then
Public Function AddNavigationCompletedHandler(ByVal Handler As c4_Handler2) As Long
    On Error GoTo ErrorHandler

    ' 1. Initialize the handler collection (First time only)
    If Col_Handler Is Nothing Then Set Col_Handler = New Collection

    Handler.ppWebView2 = ppWebView2

    ' 2. Prepare the registration for WebView2 (COM Interface)
    ' ICoreWebView2::add_NavigationCompleted VTable Index is 15 (x64)
    Const IDX_ADD_NAV_COMPLETED As Long = 15

    Dim pWebView2 As LongPtr: pWebView2 = Me.ppWebView2 ' Cached ICoreWebView2 pointer
    Dim pHandlerObj As LongPtr: pHandlerObj = Handler.Pointer ' Emulated/Fake COM Object Pointer
    Dim tkn As LongPtr ' Variable to receive the EventRegistrationToken for unregistering

    ' DispCallFunc argument preparation
    ' HRESULT add_NavigationCompleted([in] ICoreWebView2NavigationCompletedEventHandler* eventHandler, [out] EventRegistrationToken* token)
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
    Dim res As Variant
    Dim hr As Long

    ' Arg 1: [in] Pointer to our fake COM Object
    args(0) = pHandlerObj
    argTypes(0) = vbLongPtr
    argPtrs(0) = VarPtr(args(0))

    ' Arg 2: [out] Pointer to the Token variable
    args(1) = VarPtr(tkn)
    argTypes(1) = vbLongPtr
    argPtrs(1) = VarPtr(args(1))

    ' 3. Execute via DispCallFunc
    hr = DispCallFunc(pWebView2, _
                      IDX_ADD_NAV_COMPLETED * LenB(pWebView2), _
                      CC_STDCALL, _
                      vbLong, _
                      2, _
                      argTypes(0), _
                      argPtrs(0), _
                      res)

    ' 4. Evaluate results
    If hr = 0 Then
        If res = 0 Then
            ' Success: Save the token and keep the handler instance alive in the collection
            Handler.token = tkn
            Col_Handler.Add Handler, CStr(tkn) ' Using the token as a unique key to keep reference

            Debug.Print "--- Event Handler Registered Successfully ---"
            Debug.Print "Token:", tkn
            AddNavigationCompletedHandler = 0 ' S_OK
        Else
            Debug.Print "WebView2 error occurred. HRESULT: &H" & Hex(res)
            AddNavigationCompletedHandler = res
        End If
    Else
        Debug.Print "DispCallFunc itself failed. hr: &H" & Hex(hr)
        AddNavigationCompletedHandler = hr
    End If

    Exit Function

ErrorHandler:
    Debug.Print "Unexpected error in AddNavigationCompletedHandler: " & Err.Description
    AddNavigationCompletedHandler = -1
End Function
#End If


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

    Dim format As Long: format = IIf(isJpeg, 1, 0)
    Dim args(2) As Variant, args_Ptr(2) As LongPtr, args_Type(2) As Integer
    Dim res As Variant
    
    ' Arg 0: Image format (0: PNG, 1: JPEG)
    args(0) = format
    args_Ptr(0) = VarPtr(args(0))
    args_Type(0) = vbLong

    ' Arg 1: Pointer to the pre-created IStream
    args(1) = pStream
    args_Ptr(1) = VarPtr(args(1))
    args_Type(1) = vbLongPtr

    Dim Handler As c4_Handler
    Set Handler = New c4_Handler
    
    ' Use ObjPtr (numeric address) instead of a literal string key to avoid collection collisions
    Handler.Namae = "Cap_" & ObjPtr(Handler)
    Col_Handler.Add Handler, Handler.Namae
    
    Call Handler.CreateVTble(AddressOf CapturePreviewCompletedHandler_Invoke, ppWebView2)

    ' Arg 2: Completion handler (c4_Handler)
    args(2) = Handler.Pointer
    args_Ptr(2) = VarPtr(args(2))
    args_Type(2) = vbLongPtr

    ' Call ICoreWebView2::CapturePreview (VTable Index 30)
    hr = DispCallFunc(ppWebView2, 30 * LenB(ppWebView2), CC_STDCALL, vbLong, 3, args_Type(0), args_Ptr(0), res)

    If hr = 0 Then
        RegisterInstance Handler.Pointer, Me
        CapturePreviewToFile = res
    End If

    ' Release the IStream pointer to unlock the file (Decrements the reference count)
    If pStream <> 0 Then
        Dim resRelease As Variant
        ' IUnknown::Release is VTable Index 2
        Call DispCallFunc(pStream, 2 * LenB(pStream), CC_STDCALL, vbLong, 0, 0, 0, resRelease)
    End If
    
End Function
