Attribute VB_Name = "mHandlers"
' --- Standard Module: mHandlers ---

Option Explicit

' Helper function to receive AddressOf as LongPtr
Public Function GetAddr(ByVal addr As LongPtr) As LongPtr
    GetAddr = addr
End Function

' IUnknown::QueryInterface
Public Function Handler_QueryInterface(ByVal This As LongPtr, ByVal riid As LongPtr, ByRef ppvObject As LongPtr) As Long
    ' Normally used to check GUID, but for now it returns itself
    Debug.Print "QueryInterface called!"
    ppvObject = This
    Handler_QueryInterface = S_OK
End Function

' IUnknown::AddRef / Release (Returns 1 as a stub/dummy)
Public Function Handler_AddRef(ByVal This As LongPtr) As Long: Handler_AddRef = 1: End Function
Public Function Handler_Release(ByVal This As LongPtr) As Long: Handler_Release = 1: End Function

' ICoreWebView2CreateCoreWebView2EnvironmentCompletedHandler::Invoke
' Receives the initialization result from WebView2
Public Function Handler_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal pEnvironment As LongPtr) As Long
    Debug.Print "WebView2 Environment Created. ErrorCode: " & errorCode

    If errorCode = 0 Then
        Call UserForm1.WV2Environment.CreateWebView2Controller(pEnvironment)
    End If

    Handler_Invoke = 0
End Function

' Callback called by WebView2 when Controller creation is completed
Public Function ControllerHandler_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal pController As LongPtr) As Long
    
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
    Set UserForm1.WV2 = UserForm1.WV2Controller.WebView2
    
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

Public Function NavigationStarting_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' Call the method on the class side
        target.NotifyNavigationStarting
    Else
        ' yCRITICALzIf target is not found (after class is destroyed),
        ' it might be a "ghost handler" left on the WebView2 side.
        ' Clean up this pointer from the dictionary just in case.
        UnregisterInstance This
    End If
    
    NavigationStarting_Invoke = 0
End Function

Public Function ContentLoading_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyContentLoading
    Else
        UnregisterInstance This
    End If
    
    ContentLoading_Invoke = 0
End Function

Public Function SourceChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifySourceChanged
    Else
        UnregisterInstance This
    End If
    
    SourceChanged_Invoke = 0
End Function

Public Function HistoryChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyHistoryChanged
    Else
        UnregisterInstance This
    End If
    
    HistoryChanged_Invoke = 0
End Function

Public Function NavCompleted_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyNavigationCompleted
    Else
        UnregisterInstance This
    End If
    
    NavCompleted_Invoke = 0
End Function

Public Function FrameNavigationStarting_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyFrameNavigationStarting
    Else
        UnregisterInstance This
    End If
    
    FrameNavigationStarting_Invoke = 0
End Function

Public Function FrameNavigationCompleted_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyFrameNavigationCompleted
    Else
        UnregisterInstance This
    End If
    
    FrameNavigationCompleted_Invoke = 0
End Function

Public Function ScriptDialogOpening_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyScriptDialogOpening
    Else
        UnregisterInstance This
    End If
    
    ScriptDialogOpening_Invoke = 0
End Function

Public Function PermissionRequested_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyPermissionRequested
    Else
        UnregisterInstance This
    End If
    
    PermissionRequested_Invoke = 0
End Function

Public Function ProcessFailed_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyProcessFailed
    Else
        UnregisterInstance This
    End If
    
    ProcessFailed_Invoke = 0
End Function

Public Function WebMessageReceived_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyWebMessageReceived
    Else
        UnregisterInstance This
    End If
    
    WebMessageReceived_Invoke = 0
End Function

' Callback upon completion of ExecuteScript
' Index 3: Invoke(HRESULT errorCode, LPCWSTR resultObjectAsJson)
Public Function ExecuteScript_Invoke(ByVal This As LongPtr, ByVal errorCode As Long, ByVal resultJsonPtr As LongPtr) As Long
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        ' Extract string from pointer and pass to target
        target.NotifyExecuteScriptCompleted PtrToStrW(resultJsonPtr)
        
        ' Remove completed handler from registry (one-time use)
        UnregisterInstance This
    End If
    ExecuteScript_Invoke = 0
End Function

Public Function NewWindowRequested_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyNewWindowRequested
    Else
        UnregisterInstance This
    End If
    
    NewWindowRequested_Invoke = 0
End Function

Public Function DocumentTitleChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyDocumentTitleChanged
    Else
        UnregisterInstance This
    End If
    
    DocumentTitleChanged_Invoke = 0
End Function

Public Function ContainsFullScreenElementChanged_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyContainsFullScreenElementChanged
    Else
        UnregisterInstance This
    End If
    
    ContainsFullScreenElementChanged_Invoke = 0
End Function

Public Function WebResourceRequested_Invoke(ByVal This As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next
    Dim target As c3_WebView2
    Set target = GetInstance(This)
    
    If Not target Is Nothing Then
        target.NotifyWebResourceRequested
    Else
        UnregisterInstance This
    End If
    
    WebResourceRequested_Invoke = 0
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
