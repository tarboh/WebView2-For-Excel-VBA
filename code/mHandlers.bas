Attribute VB_Name = "mHandlers"
' --- Standard Module: mHandlers ---

Option Explicit

' Helper function to receive AddressOf as LongPtr
Public Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Public Function Handler_QueryInterface(ByVal this As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    OutputDebugString StrPtr("QI Called from WebView2! " & this)
    Debug.Print "QueryInterface called!"
    ppvObject = this
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Public Function Handler_AddRef(ByVal this As LongPtr) As Long:
    OutputDebugString StrPtr("AddRef Called from WebView2! " & this)
    Handler_AddRef = 1
End Function
Public Function Handler_Release(ByVal this As LongPtr) As Long
    OutputDebugString StrPtr("Release Called from WebView2! " & this)
    Handler_Release = 1
End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Public Function Handler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    Debug.Print "WebView2 Environment Created. ErrorCode: " & errorCode

    If errorCode = 0 Then
        Call UserForm1.WV2Environment.CreateWebView2Controller(pEnvironment)
    End If

    Handler_Invoke = 0
End Function

' Callback called by WebView2 when Controller creation is completed
Public Function ControllerHandler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal pController As LongPtr) As Long
    
    Debug.Print "ControllerHandler_Invoke called. pController: " & pController
    
    If errorCode <> 0 Then Exit Function
    
    ' --- CRITICAL: Prevent WebView2 from being destroyed ---
    CallAddRef pController
    
    Set UserForm1.WV2Controller = New c2_WebView2Controller
    
    ' Register pointer
    UserForm1.WV2Controller.pController = pController
    
    ' Make it visible
    UserForm1.WV2Controller.IsVisible = True

    ' Retrieve WebView2 object
    Call UserForm1.WV2Controller.GetWebView2
    Set UserForm1.wv2 = UserForm1.WV2Controller.WebView2
    
    ' Get Settings
    Call UserForm1.WV2Controller.WebView2.get_Settings
    
    ' Set ScriptDialogsEnabled Property
    UserForm1.WV2Controller.WebView2.Settings.AreDefaultScriptDialogsEnabled = True
    
    ' Register Navigation/Event handlers
    Call UserForm1.WV2Controller.WebView2.add_NavigationStarting
    Call UserForm1.WV2Controller.WebView2.add_ContentLoading
    Call UserForm1.WV2Controller.WebView2.add_SourceChanged
    Call UserForm1.WV2Controller.WebView2.add_HistoryChanged
    Call UserForm1.WV2Controller.WebView2.add_NavigationCompleted
    Call UserForm1.WV2Controller.WebView2.add_FrameNavigationStarting
    Call UserForm1.WV2Controller.WebView2.add_FrameNavigationCompleted
    Call UserForm1.WV2Controller.WebView2.add_ScriptDialogOpening
    Call UserForm1.WV2Controller.WebView2.add_PermissionRequested
    Call UserForm1.WV2Controller.WebView2.add_ProcessFailed
    Call UserForm1.WV2Controller.WebView2.add_WebMessageReceived
    Call UserForm1.WV2Controller.WebView2.add_NewWindowRequested
    Call UserForm1.WV2Controller.WebView2.add_DocumentTitleChanged
    Call UserForm1.WV2Controller.WebView2.add_ContainsFullScreenElementChanged
    Call UserForm1.WV2Controller.WebView2.add_WebResourceRequested
    Call UserForm1.WV2Controller.WebView2.AddScriptToExecuteOnDocumentCreated("console.log('script on document created!');")

    
    ' Register events via Handler2 approach
    #If Win64 Then
    Call UserForm1.WV2Controller.WebView2.AddNavigationCompletedHandler(UserForm1.NavigationCompletedHandler)
    #End If
    
    Debug.Print "ppWebView2:", UserForm1.WV2Controller.WebView2.ppWebView2
    
    ' 4. Force visibility using Win32 API
    DoEvents
    Dim childHwnd As LongPtr
    ' Directly manipulate "Chrome_WidgetWin_0" discovered in previous inspection
    childHwnd = FindWindowEx(TargetHwnd, 0, "Chrome_WidgetWin_0", vbNullString)

    If childHwnd <> 0 Then
        ' Even if put_Bounds fails inside WebView2,
        ' we can force the size at the OS level if we have the window handle.
        MoveWindow childHwnd, 0, 0, 800, 600, 1
        Debug.Print "Final Sync via Win32 API. childHwnd: " & childHwnd & "(" & Hex(childHwnd) & ")"
    End If

    Call Resize ' (Change to english if function name is defined, e.g., AdjustLayout)
    
    Call UserForm1.WV2Controller.ReadyCompleted

    ControllerHandler_Invoke = 0
End Function

Public Function NavigationStarting_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        ' Call the method on the class side
        target.NotifyNavigationStarting
    Else
        ' ÅyCRITICALÅzIf target is not found (after class is destroyed),
        ' it might be a "ghost handler" left on the WebView2 side.
        ' Clean up this pointer from the dictionary just in case.
        UnregisterInstance this
    End If
    
    NavigationStarting_Invoke = 0
End Function

Public Function ContentLoading_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyContentLoading
    Else
        UnregisterInstance this
    End If
    
    ContentLoading_Invoke = 0
End Function

Public Function SourceChanged_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifySourceChanged
    Else
        UnregisterInstance this
    End If
    
    SourceChanged_Invoke = 0
End Function

Public Function HistoryChanged_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyHistoryChanged
    Else
        UnregisterInstance this
    End If
    
    HistoryChanged_Invoke = 0
End Function

Public Function NavCompleted_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyNavigationCompleted
    Else
        UnregisterInstance this
    End If
    
    NavCompleted_Invoke = 0
End Function

Public Function FrameNavigationStarting_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyFrameNavigationStarting
    Else
        UnregisterInstance this
    End If
    
    FrameNavigationStarting_Invoke = 0
End Function

Public Function FrameNavigationCompleted_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyFrameNavigationCompleted
    Else
        UnregisterInstance this
    End If
    
    FrameNavigationCompleted_Invoke = 0
End Function

Public Function ScriptDialogOpening_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyScriptDialogOpening
    Else
        UnregisterInstance this
    End If
    
    ScriptDialogOpening_Invoke = 0
End Function

Public Function PermissionRequested_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyPermissionRequested
    Else
        UnregisterInstance this
    End If
    
    PermissionRequested_Invoke = 0
End Function

Public Function ProcessFailed_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyProcessFailed
    Else
        UnregisterInstance this
    End If
    
    ProcessFailed_Invoke = 0
End Function

Public Function WebMessageReceived_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyWebMessageReceived
    Else
        UnregisterInstance this
    End If
    
    WebMessageReceived_Invoke = 0
End Function

' Callback upon completion of ExecuteScript
' Index 3: Invoke(HRESULT errorCode, LPCWSTR resultObjectAsJson)
Public Function ExecuteScript_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal resultJsonPtr As LongPtr) As Long
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        ' Extract string from pointer and pass to target
        target.NotifyExecuteScriptCompleted PtrToStrW(resultJsonPtr)
        
        ' Remove completed handler from registry (one-time use)
        UnregisterInstance this
    End If
    ExecuteScript_Invoke = 0
End Function

Public Function NewWindowRequested_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyNewWindowRequested
    Else
        UnregisterInstance this
    End If
    
    NewWindowRequested_Invoke = 0
End Function

Public Function DocumentTitleChanged_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyDocumentTitleChanged
    Else
        UnregisterInstance this
    End If
    
    DocumentTitleChanged_Invoke = 0
End Function

Public Function ContainsFullScreenElementChanged_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyContainsFullScreenElementChanged
    Else
        UnregisterInstance this
    End If
    
    ContainsFullScreenElementChanged_Invoke = 0
End Function

Public Function WebResourceRequested_Invoke(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        target.NotifyWebResourceRequested
    Else
        UnregisterInstance this
    End If
    
    WebResourceRequested_Invoke = 0
End Function

'public HRESULT Invoke(HRESULT errorCode, LPCWSTR result)
Public Function AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal result As LongPtr) As Long
    
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(this)
    
    If Not target Is Nothing Then
        Call target.NotifyAddScriptToExecuteOnDocumentCreatedCompleted(errorCode, result)
    Else
        UnregisterInstance this
    End If
    
    AddScriptToExecuteOnDocumentCreatedCompletedHandler_Invoke = 0
    
End Function

Public Function CapturePreviewCompletedHandler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long) As Long
    CapturePreviewCompletedHandler_Invoke = 0
    
    Dim wv2 As c3_WebView2
    Set wv2 = GetInstance(this)
    
    If Not wv2 Is Nothing Then
        wv2.NotifyCapturePreviewCompleted errorCode
        
        ' For stability, do not remove the handler from the collection here.
        ' Let the Class_Terminate (destruction of c3_WebView2) handle the cleanup to prevent crashes.
    End If
    
    UnregisterInstance this
End Function

'MIDL_INTERFACE ("5c4889f0-5ef6-4c5a-952c-d8f1b92d0574")
'ICoreWebView2CallDevToolsProtocolMethodCompletedHandler:      Public IUnknown
'{
'public:
'    virtual HRESULT STDMETHODCALLTYPE Invoke(
'        /* [in] */ HRESULT errorCode,
'        /* [in] */ LPCWSTR result) = 0;
'
'};
Public Function CallDevToolsProtocolMethodCompletedHandler_Invoke(ByVal this As LongPtr, ByVal errorCode As Long, ByVal result As LongPtr) As Long
    
    Dim wv2 As c3_WebView2
    Set wv2 = GetInstance(this)
    
    If Not wv2 Is Nothing Then
        wv2.NotifyCallDevToolsProtocolMethodCompleted errorCode, result
        'wv2.Col_Handler.Remove "CallDevToolsProtocolMethodCompletedHandler"
        ' For stability, do not remove the handler from the collection here.
        ' Let the Class_Terminate (destruction of c3_WebView2) handle the cleanup to prevent crashes.
    End If
    
    UnregisterInstance this
    
    CallDevToolsProtocolMethodCompletedHandler_Invoke = 0
    
End Function

'public HRESULT Invoke(ICoreWebView2 * sender, ICoreWebView2DevToolsProtocolEventReceivedEventArgs * args)
''' <summary>
''' Event handler for ICoreWebView2DevToolsProtocolEventReceivedEventHandler::Invoke.
''' This intercepts real-time CDP events (Console logs, Network interception, etc.).
''' </summary>
Public Function DevToolsProtocolEventReceivedHandler_Invoke( _
    ByVal this As LongPtr, _
    ByVal sender As LongPtr, _
    ByVal args As LongPtr) As Long

    Debug.Print "DTPEventHandler_Invoke. this: " & this & " sender: " & sender

    ' S_OK (Success) by default
    DevToolsProtocolEventReceivedHandler_Invoke = 0
    
    'OutputDebugString StrPtr("DTPEventHandler_Invoke. this: " & this & " sender: " & sender)
    
    ' Fail-safe: Ensure args pointer is valid
    If args = 0 Then
        Debug.Print "Exit"
        Exit Function
    End If
    ' To prevent crashes, use standard error handling for low-level COM operations
    On Error GoTo ErrorHandler
    
    Dim wv2 As c3_WebView2
    Set wv2 = GetInstance(this)
    
    If Not wv2 Is Nothing Then
        wv2.NotifyDevToolsProtocolEventReceived
        'wv2.Col_Handler.Remove "NotifyDevToolsProtocolEventReceivedHandler"
        ' For stability, do not remove the handler from the collection here.
        ' Let the Class_Terminate (destruction of c3_WebView2) handle the cleanup to prevent crashes.
    End If
    
    'UnregisterInstance this

    ' 1. Capture the specific event parameters (e.g., Get the JSON data payload)
    ' (Assuming you have a helper class c3_WebView2 or similar to wrap ICoreWebView2DevToolsProtocolEventReceivedEventArgs)
    ' Example:
    ' Dim eventArgs As New c6_DevToolsEventArgs
    ' eventArgs.Initialize args
    ' Debug.Print eventArgs.ParameterObjectAsJson

    Exit Function

ErrorHandler:
    ' Return HRESULT E_FAIL on crash/error to notify the WebView2 runtime
    Debug.Print "Error!"
    DevToolsProtocolEventReceivedHandler_Invoke = &H80004005
End Function

' Helper: PtrToStrW (Converts Unicode pointer to VBA String)
Public Function PtrToStrW(ByVal pWStr As LongPtr) As String
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
