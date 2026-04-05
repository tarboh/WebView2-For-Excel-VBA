VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20760
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Public UserForm UserFrom1

Option Explicit


'Public WV2Loader As New c0_WebView2Loader
'Public WV2Loader As New Class1
'Public WV2Environment As New c1_WebView2Environment
'Public WithEvents WV2Controller As c2_WebView2Controller
'Public WithEvents WV2 As c3_WebView2
Public c5 As New ObjectForJS

Private WithEvents Console As fm_Console
Attribute Console.VB_VarHelpID = -1
Private WithEvents WV2 As WebView2
Attribute WV2.VB_VarHelpID = -1

Public m_InstanceMap As Object

Private Sub CheckBox_Attach_c5ToJS_Click()
    If CheckBox_Attach_c5ToJS.value = True Then
        If WV2.AddHostObjectToScript("VBAObj", c5) = 0 Then
            Debug.Print "c5 attached as 'VBAObj'"
        Else
            Debug.Print "c5 attache failed"
        End If
    Else
        If WV2.RemoveHostObjectFromScript("VBAObj") = 0 Then
            Debug.Print "c5 remove success"
        Else
            Debug.Print "c5 remove failed"
        End If
    End If
End Sub

Private Sub CheckBox_BuiltInErrorPageEnabled_Click()
    WV2.IsBuiltInErrorPageEnabled = CheckBox_BuiltInErrorPageEnabled.value
    Debug.Print "IsBuiltInErrorPageEnabled:" & WV2.IsBuiltInErrorPageEnabled
End Sub

Private Sub CheckBox_Controller_IsVisible_Click()
    WV2.Controller_IsVisible = CheckBox_Controller_IsVisible.value
    Debug.Print "Controller_IsVisible:" & WV2.Controller_IsVisible
End Sub

Private Sub CheckBox_DefaultContextMenusEnabled_Change()
    WV2.AreDefaultContextMenusEnabled = CheckBox_DefaultContextMenusEnabled.value
    Debug.Print "AreDefaultContextMenusEnabled:" & WV2.AreDefaultContextMenusEnabled
End Sub

Private Sub CheckBox_DevToolsEnabled_Change()
    WV2.AreDevToolsEnabled = CheckBox_DevToolsEnabled.value
    Debug.Print "AreDevToolsEnabled:" & WV2.AreDevToolsEnabled
End Sub

Private Sub CheckBox_HostObjectsAllowed_Change()
    WV2.AreHostObjectsAllowed = CheckBox_HostObjectsAllowed.value
    Debug.Print "AreHostObjectsAllowed:" & WV2.AreHostObjectsAllowed
End Sub

Private Sub CheckBox_InterceptDialogs_Change()
    If CheckBox_InterceptDialogs.value = True Then
        WV2.AreDefaultScriptDialogsEnabled = False
    Else
        WV2.AreDefaultScriptDialogsEnabled = True
    End If
    Debug.Print WV2.AreDefaultScriptDialogsEnabled
    WV2.Reload
End Sub

Private Sub CheckBox_ScriptEnabled_Change()
    WV2.IsScriptEnabled = CheckBox_ScriptEnabled.value
End Sub

Private Sub CheckBox_StatusBarEnabled_change()
    WV2.IsStatusBarEnabled = CheckBox_StatusBarEnabled.value
End Sub

Private Sub CheckBox_WebMessageEnabled_Change()
    WV2.IsWebMessageEnabled = CheckBox_WebMessageEnabled.value
End Sub

Private Sub CheckBox_ZoomControlEnabled_change()
    WV2.IsZoomControlEnabled = CheckBox_ZoomControlEnabled.value
    Debug.Print "IsZoomControlEnabled:" & WV2.IsZoomControlEnabled
End Sub

Private Sub CmdBtn_SetBoundsAndZoomFactor_Click()
    Call WV2.Controller_SetBoundsAndZoomFactor(TextBox_Bounds_Left, TextBox_Bounds_Top, TextBox_Bounds_Right, TextBox_Bounds_Bottom, TextBox_Controller_ZoomFactor)
End Sub

Private Sub CommandButton_CallDevToolsProtocolMethod_Click()
    Dim strMethodName As String
    strMethodName = "Page.printToPDF"
    Dim strParametersAsJson As String
    strParametersAsJson = "{" & _
        """paperWidth"": 8.27," & _
        """paperHeight"": 11.69," & _
        """marginTop"": 0," & _
        """marginBottom"": 0," & _
        """marginLeft"": 0," & _
        """marginRight"": 0," & _
        """printBackground"": true," & _
        """landscape"": false," & _
        """displayHeaderFooter"": false" & _
    "}"
    Call WV2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
End Sub

Private Sub CommandButton_CapturePreviewToFile_Click()
    
    Dim folderPath As String
    folderPath = "C:\temp\VBA_WebView2\ScreenShot\"
    
    Dim uniquePath As String
    uniquePath = "cap_" & format(Now, "yyyymmdd_hhnnss") & "_" & Right("000" & Int(Timer * 1000) Mod 1000, 3) & ".png"
    
    WV2.CapturePreviewToFile folderPath, uniquePath
End Sub

Private Sub CommandButton_Console_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
End Sub



Private Sub CommandButton_Controller_Close_Click()
    WV2.Controller_Close
End Sub

Private Sub CommandButton_Controller_Get_ParentWindow_Click()
    Dim hwnd As LongPtr
    hwnd = WV2.Controller_ParentWindow
    TextBox_Controller_ParentWindow.Text = hwnd
End Sub

Private Sub CommandButton_Controller_get_ZoomFactor_Click()
    TextBox_Controller_ZoomFactor.Text = WV2.Controller_ZoomFactor
End Sub

Private Sub CommandButton_Controller_MoveFocus_Click()
    Dim reason As COREWEBVIEW2_MOVE_FOCUS_REASON
    reason = ComboBox_MOVE_FOCUS_REASON.ListIndex
    Call WV2.Controller_MoveFocus(reason)
End Sub

Private Sub CommandButton_Controller_Put_ZoomFactor_Click()
    WV2.Controller_ZoomFactor = TextBox_Controller_ZoomFactor.Text
End Sub

Private Sub CommandButton_ExeCuteVBAInJavaScript_Click()
    Call WV2.ExecuteScriptAsync("window.chrome.webview.hostObjects.sync.VBAObj.Func1(15);")
End Sub

Private Sub CommandButton_GoBack_Click()
    WV2.GoBack
End Sub

Private Sub CommandButton_GoForward_Click()
    WV2.GoForward
End Sub

Private Sub CommandButton_Navigate_Click()
    
    Dim url As String
    url = TextBox_URL.Text
        
    If Left(url, 11) = "javascript:" Then
        Call WV2.ExecuteScriptAsync(url)
    ElseIf Left(url, 4) = "http" Then
        Call WV2.NavigateAsync(url)
    Else
        Call WV2.NavigateToString(url)
    End If

End Sub


Private Sub CommandButton_NavToStr_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
    
    Dim uri As String
    uri = Console.TextBox_Console.Text
    Debug.Print uri
    Call WV2.NavigateToString(uri)
End Sub

Private Sub CommandButton_Open_Click()
    Call Create_WebView2
End Sub

Private Sub CommandButton_OpenDevTools_Click()
    WV2.OpenDevToolsWindow
End Sub

Private Sub CommandButton_PostWebMessageAsJson_Click()
    Dim strJson As String
    strJson = "{""funcName"": ""calculateAndDisplay"", ""args"": [""Sum Result"", 123, 456]}"
    Debug.Print WV2.PostWebMessageAsJson(strJson)
End Sub

Private Sub CommandButton_PostWebMessageAsString_Click()
    Dim webMessage As String
    webMessage = "System Check Complete"
    Debug.Print WV2.PostWebMessageAsString(webMessage)
End Sub

Private Sub CommandButton_Reload_Click()
    Call WV2.Reload
End Sub

Private Sub CommandButton_RunScript_Click()
    Dim script As String
    script = TextBox_Script.Text
    Call WV2.ExecuteScriptAsync(script)
End Sub

Private Sub CommandButton_Stop_Click()
    WV2.Stop_
End Sub

Private Sub CommandButton_StopAutoJS_Click()
'    WV2.RemoveScriptToExecuteOnDocumentCreated ( _
'        WV2.scriptId)
End Sub

Private Sub CommandButton4_Click()
    
    Call WV2.add_DevToolsProtocolEventReceived("Network.responseReceived")

    ' ネットワーク監視機能を有効化する（これを投げないとイベントが来ない）
    Dim strMethodName As String
    strMethodName = "Network.enable"
    
    Dim strParametersAsJson As String
    strParametersAsJson = "{}" ' パラメータは空のJSONオブジェクトでOK
    
    Dim hr As Long
    hr = WV2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
    Debug.Print "登録結果：" & hr
    
End Sub

Private Sub CommandButton5_Click()
    Call WV2.AddWebResourceRequestedFilter("*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE)
    Call WV2.add_WebResourceRequested
End Sub


Private Sub CommandButton_Controller_get_Bounds_Click()
    Dim hr As Long
    Dim l(3) As Long
    hr = WV2.Controller_get_Bounds(l)
    Debug.Print "get_Bounds hr:" & hr
    TextBox_Bounds_Left.Text = l(0)
    TextBox_Bounds_Top.Text = l(1)
    TextBox_Bounds_Right.Text = l(2)
    TextBox_Bounds_Bottom.Text = l(3)
End Sub

Private Sub CommandButton_Controller_put_Bounds_Click()
    Dim hr As Long
    Dim l(3) As Long
    l(0) = TextBox_Bounds_Left.Text
    l(1) = TextBox_Bounds_Top.Text
    l(2) = TextBox_Bounds_Right.Text
    l(3) = TextBox_Bounds_Bottom.Text
    hr = WV2.Controller_put_Bounds(l)
    Debug.Print "put_Bounds hr:" & hr
End Sub

Private Sub CommandButton7_Click()
    WV2.Controller_NotifyParentWindowPositionChanged
End Sub

Private Sub CommandButtonController_Put_ParentWindow_Click()
    WV2.Controller_ParentWindow = TextBox_Controller_ParentWindow.Text
End Sub

Private Sub Console_QueryClose()
    Set Console = Nothing
End Sub

#If Win64 Then
Private Sub NavigationCompletedHandler_Invoked(ByVal pThis As LongLong, ByVal sender As LongLong, ByVal args As LongLong)
    Debug.Print "C4_Handler2_NavigationCompleted!"
End Sub
#End If





Private Sub WV2_AddScriptToExecuteOnDocumentCreatedCompleted(ByVal scriptId As String, ByVal javascript As String)
    Debug.Print "AddScriptToExecuteOnDocumentCreatedCompleted"
End Sub

Private Sub wv2_AddScriptToExecuteOnDocumentCreatedFailed(ByVal javascript As String, ByVal errorCode As Long)
    Debug.Print "AddScriptToExecuteOnDocumentCreatedFailed"
End Sub

Private Sub wv2_CallDevToolsProtocolMethodCompleted(ByVal requestId As Long, ByVal methodName As String, ByVal errorCode As String, ByVal result As String)
    
    'Debug.Print "CallDevToolsProtocolMethodCompleted result:" & result
    
    Debug.Print methodName
    Select Case methodName
        Case "Page.printToPDF"
            ' VBA can directly access JavaScript properties (e.g., .data) retrieved from JScript!
            Dim jsonObject As Object
            Set jsonObject = WV2.ParseJSON(result)
            
            ' Safely retrieve the Base64 PDF string directly via Dot Notation
            Dim base64PDF As String
            base64PDF = CallByName(jsonObject, "data", VbGet)
            
            If Len(base64PDF) > 0 Then
                Dim pdfBytes() As Byte
                pdfBytes = WV2.Base64Decode(base64PDF)
                
                Dim folderPath As String
                folderPath = "C:\temp\VBA_WebView2\PDF\"
                
                WV2.CreateDeepFolder folderPath
                
                Dim uniquePath As String
                uniquePath = format(Now, "yyyymmdd_hhnnss") & "_" & Right("000" & Int(Timer * 1000) Mod 1000, 3) & ".pdf"
                
                WV2.SaveBytesToFile pdfBytes, folderPath & uniquePath
                Debug.Print "PDF saved successfully to Desktop!"
            End If
        Case 2
    End Select
    
End Sub

Private Sub wv2_CapturePreviewCompleted(ByVal errorCode As Long)
    Debug.Print "EventColled : CapturePreviewCompleted"
End Sub

Private Sub WV2_ContainsFullScreenElementChanged()
    Debug.Print "ContainsFullScreenElementChanged"
    'Dim Source As String
    'Source = WV2Controller.WebView2.Source
    
    Dim Title As String
    Title = WV2.DocumentTitle
    
    Debug.Print "NavigationCompleted(From Standard Module) "
    'Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title
    
    'TextBox_URL.text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & WV2.ContainsFullScreenElement
End Sub

Private Sub WV2_ControllerAcceleratorKeyPressed()
    Debug.Print "ControllerAcceleratorKeyPressed"
End Sub

Private Sub WV2_ControllerGotFocus()
    Debug.Print "ControllerGotFocus"
End Sub

Private Sub WV2_ControllerLostFocus()
    Debug.Print "ControllerLostFocus"
End Sub

Private Sub WV2_ControllerMoveFocusRequested()
    Debug.Print "ControllerMoveFocuceRequested"
End Sub

Private Sub WV2_ControllerZoomFactorChanged()
    Debug.Print "ControllerZoomChanged ZoomFactor:" & WV2.Controller_ZoomFactor
    TextBox_Controller_ZoomFactor.Text = WV2.Controller_ZoomFactor
End Sub

Private Sub wv2_DevToolsProtocolEventReceived(ByVal eventName As String, ByVal parameterJson As String)
    Debug.Print "DevToolsProtocolEventReceived. JSON:" & parameterJson
End Sub

Private Sub WV2_DocumentTitleChanged()
    Debug.Print "DocumentTitleChanged"
End Sub

Private Sub WV2_NewWindowRequested()
    Debug.Print "NewWindowRequested"
End Sub

Private Sub WV2_PermissionRequested()
    Debug.Print "PermissionRequested"
End Sub

Private Sub WV2_ProcessFailed()
    Debug.Print "ProcessFailed"
End Sub

Private Sub WV2_ReceiveScriptResult(ByVal result As String)
    Debug.Print "ReceiveScriptResult result : " & result
End Sub

Private Sub WV2_ScriptDialogOpening()
    Debug.Print "ScriptDialogOpening"
    MsgBox "Dialog On VBA!"
End Sub

Private Sub CommandButton3_Click()
    Call WV2.ExecuteScriptAsync("alert('Dialog On WebView2 !');")
End Sub

Private Sub UserForm_Initialize()
'    Me.width = 1920 * 0.75
'    Me.Height = 1080 * 0.75
'    Frame1.width = 1800 * 0.75
'    Frame1.Height = 1000 * 0.75
    
'    Set wv2 = New c3_WebView2
'    Call wv2.BuildFuncPtrCache
    
    Set m_InstanceMap = CreateObject("Scripting.Dictionary")
    
    Set c5 = New ObjectForJS
    
    ComboBox_MOVE_FOCUS_REASON.AddItem "PROGRAMMATIC"
    ComboBox_MOVE_FOCUS_REASON.AddItem "NEXT"
    ComboBox_MOVE_FOCUS_REASON.AddItem "PREVIOUS"
    ComboBox_MOVE_FOCUS_REASON.ListIndex = 0
    
    Call Create_WebView2
End Sub


'Create WebView2 In Frame
Public Sub Create_WebView2()

    'Use Hidden Property
    'Notified by KallunWillock via GitHub Issue. Thank you!
    Dim targetHWnd As LongPtr
    targetHWnd = Frame1.[_GethWnd]
    Debug.Print targetHWnd
    
    Set WV2 = New WebView2
    
    'Call WV2Loader.CreateWebView2Environment
    Call WV2.CreateWebView2Environment(Frame1) 'targetHWnd)
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 1. まずハンドラを全解除
        If Not WV2 Is Nothing Then
            WV2.Finalize
        End If
    
    ' 2. WebView2プロセスをシャットダウン
    
        WV2.CloseWebView2
    
    ' 3. 参照を解放
    Set m_InstanceMap = Nothing
    
    Debug.Print "Console解放前"
    Set Console = Nothing
    Debug.Print "Console解放後"
    
    Debug.Print "QueryClose完了"
End Sub

Private Sub WV2_ContentLoading()
    Debug.Print "ContentLoading"
End Sub

Private Sub WV2_FrameNavigationCompleted()
    Debug.Print "FrameNavigationCompleted"
End Sub

Private Sub WV2_FrameNavigationStarting()
    Debug.Print "FrameNavigationStarting"
End Sub

Private Sub WV2_HistoryChanged()
    Debug.Print "HistoryChanged"
End Sub

Private Sub WV2_NavigationCompleted()
     
'    Debug.Print "NavigationCompleted"
    Dim Source As String
    Source = WV2.Source

    Dim Title As String
    Title = WV2.DocumentTitle

    Debug.Print "NavigationCompleted(From Standard Module) "
    Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title

    TextBox_URL.Text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & WV2.ContainsFullScreenElement
    
End Sub

Private Sub WV2_NavigationStarting()
    
    Debug.Print "NavigationStarting"
End Sub

Private Sub WV2_SourceChanged()
    CommandButton_GoBack.Enabled = WV2.CanGoBack
    CommandButton_GoForward.Enabled = WV2.CanGoForward
    Debug.Print "SourceChanged"
End Sub

Private Sub WV2_WebMessageReceived(ByVal Source As String, ByVal messageJson As String, ByVal messageString As String)
    Debug.Print "WebMessageReceived"
    Debug.Print "    source        :" & Source
    Debug.Print "    mssage(json)  :" & messageJson
    Debug.Print "    mssage(string):" & messageString
End Sub

Private Sub WV2_WebResourceRequested()
    Debug.Print "WebResourceRequested"
End Sub

Private Sub wv2_WindowCloseRequested(ByVal this As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "WindowCloseRequested"
End Sub

'Private Sub WV2Controller_ScriptResultReceived(ByVal result As String)
'
'    Debug.Print "ScriptResultReceived:", result
'    TextBox_URL.text = WV2Controller.WebView2.Source
'
'End Sub

'Private Sub WV2Controller_WebView2ReadyCompleted()
'
'    Debug.Print "WV2Controller_WebView2ReadyCompleted proccessid:" & WV2Controller.WebView2.BrowserProcessId
'    Call WV2Controller.WebView2.NavigateAsync("https://www.google.com/")
'
'End Sub



