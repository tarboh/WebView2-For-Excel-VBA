VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   10560
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15990
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

Public WV2Loader As New c0_WebView2Loader
Public WV2Environment As New c1_WebView2Environment
Public WithEvents WV2Controller As c2_WebView2Controller
Attribute WV2Controller.VB_VarHelpID = -1
Public WithEvents wv2 As c3_WebView2
Attribute wv2.VB_VarHelpID = -1
Public c5 As New c5_ObjectForJS

Private WithEvents Console As fm_Console
Attribute Console.VB_VarHelpID = -1

#If Win64 Then
Public WithEvents NavigationCompletedHandler As c4_Handler2
Attribute NavigationCompletedHandler.VB_VarHelpID = -1
#End If

Private Sub CheckBox_Attach_c5ToJS_Click()
    If CheckBox_Attach_c5ToJS.value = True Then
        Call wv2.AddHostObjectToScript("VBAObj", c5)
    Else
        Call wv2.RemoveHostObjectFromScript("VBAObj")
    End If
End Sub

Private Sub CheckBox_InterceptDialogs_Change()
    If CheckBox_InterceptDialogs.value = True Then
        wv2.Settings.AreDefaultScriptDialogsEnabled = False
    Else
        wv2.Settings.AreDefaultScriptDialogsEnabled = True
    End If
    Debug.Print wv2.Settings.AreDefaultScriptDialogsEnabled
    wv2.Reload
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
    Call WV2Controller.WebView2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
End Sub

Private Sub CommandButton_CapturePreviewToFile_Click()
    
    Dim folderPath As String
    folderPath = "C:\temp\VBA_WebView2\ScreenShot\"
    
    Dim uniquePath As String
    uniquePath = "cap_" & format(Now, "yyyymmdd_hhnnss") & "_" & Right("000" & Int(Timer * 1000) Mod 1000, 3) & ".png"
    
    WV2Controller.WebView2.CapturePreviewToFile folderPath, uniquePath
End Sub

Private Sub CommandButton_Console_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
End Sub



Private Sub CommandButton_ExeCuteVBAInJavaScript_Click()
    Call WV2Controller.WebView2.ExecuteScriptAsync("window.chrome.webview.hostObjects.sync.VBAObj.Func1(15);")
End Sub

Private Sub CommandButton_GoBack_Click()
    WV2Controller.WebView2.GoBack
End Sub

Private Sub CommandButton_GoForward_Click()
    WV2Controller.WebView2.GoForward
End Sub

Private Sub CommandButton_Navigate_Click()
    
    Dim url As String
    url = TextBox_URL.text
        
    If Left(url, 11) = "javascript:" Then
        Call WV2Controller.WebView2.ExecuteScriptAsync(url)
    ElseIf Left(url, 4) = "http" Then
        Call WV2Controller.WebView2.NavigateAsync(url)
    Else
        Call WV2Controller.WebView2.NavigateToString(url)
    End If

End Sub


Private Sub CommandButton_NavToStr_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
    
    Dim uri As String
    uri = Console.TextBox_Console.text
    Debug.Print uri
    Call WV2Controller.WebView2.NavigateToString(uri)
End Sub

Private Sub CommandButton_OpenDevTools_Click()
    wv2.OpenDevToolsWindow
End Sub

Private Sub CommandButton_PostWebMessageAsJson_Click()
    Dim strjson As String
    strjson = "{""funcName"": ""calculateAndDisplay"", ""args"": [""Sum Result"", 123, 456]}"
    Debug.Print WV2Controller.WebView2.PostWebMessageAsJson(strjson)
End Sub

Private Sub CommandButton_PostWebMessageAsString_Click()
    Dim webMessage As String
    webMessage = "System Check Complete"
    Debug.Print WV2Controller.WebView2.PostWebMessageAsString(webMessage)
End Sub

Private Sub CommandButton_Reload_Click()
    Call WV2Controller.WebView2.Reload
End Sub

Private Sub CommandButton_RunScript_Click()
    Dim script As String
    script = TextBox_Script.text
    Call WV2Controller.WebView2.ExecuteScriptAsync(script)
End Sub

Private Sub CommandButton_Stop_Click()
    WV2Controller.WebView2.Stop_
End Sub

Private Sub CommandButton_StopAutoJS_Click()
    WV2Controller.WebView2.RemoveScriptToExecuteOnDocumentCreated ( _
        WV2Controller.WebView2.ScriptId)
End Sub

Private Sub CommandButton4_Click()
    
    Call WV2Controller.WebView2.GetDevToolsProtocolEventReceiver("Network.responseReceived")

    ' ネットワーク監視機能を有効化する（これを投げないとイベントが来ない）
    Dim strMethodName As String
    strMethodName = "Network.enable"
    
    Dim strParametersAsJson As String
    strParametersAsJson = "{}" ' パラメータは空のJSONオブジェクトでOK
    
    Dim hr As Long
    hr = WV2Controller.WebView2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
    Debug.Print "登録結果：" & hr
    
End Sub

Private Sub CommandButton5_Click()
    Call WV2Controller.WebView2.AddWebResourceRequestedFilter("*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE)
    Call WV2Controller.WebView2.add_WebResourceRequested
End Sub

Private Sub Console_QueryClose()
    Set Console = Nothing
End Sub

#If Win64 Then
Private Sub NavigationCompletedHandler_Invoked(ByVal pThis As LongLong, ByVal sender As LongLong, ByVal args As LongLong)
    Debug.Print "C4_Handler2_NavigationCompleted!"
End Sub
#End If

Private Sub WV2_AddScriptToExecuteOnDocumentCreatedCompleted()
    Debug.Print "AddScriptToExecuteOnDocumentCreatedCompleted"
End Sub

Private Sub wv2_CallDevToolsProtocolMethodCompleted(ByVal errorCode As String, ByVal result As String)
    
    'Debug.Print "CallDevToolsProtocolMethodCompleted result:" & result
    
    ' VBA can directly access JavaScript properties (e.g., .data) retrieved from JScript!
    Dim jsonObject As Object
    Set jsonObject = ParseJSON(result)
    
    ' Safely retrieve the Base64 PDF string directly via Dot Notation
    Dim base64PDF As String
    base64PDF = CallByName(jsonObject, "data", VbGet)
    
    If Len(base64PDF) > 0 Then
        Dim pdfBytes() As Byte
        pdfBytes = Base64Decode(base64PDF)
        
        Dim folderPath As String
        folderPath = "C:\temp\VBA_WebView2\PDF\"
        
        CreateDeepFolder folderPath
        
        Dim uniquePath As String
        uniquePath = format(Now, "yyyymmdd_hhnnss") & "_" & Right("000" & Int(Timer * 1000) Mod 1000, 3) & ".pdf"
        
        SaveBytesToFile pdfBytes, folderPath & uniquePath
        Debug.Print "PDF saved successfully to Desktop!"
    End If
    
End Sub

Private Sub wv2_CapturePreviewCompleted(ByVal errorCode As Long)
    Debug.Print "CapturePreviewCompleted"
End Sub

Private Sub WV2_ContainsFullScreenElementChanged()
    Debug.Print "ContainsFullScreenElementChanged"
    'Dim Source As String
    'Source = WV2Controller.WebView2.Source
    
    Dim Title As String
    Title = WV2Controller.WebView2.DocumentTitle
    
    Debug.Print "NavigationCompleted(From Standard Module) "
    'Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title
    
    'TextBox_URL.text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & WV2Controller.WebView2.ContainsFullScreenElement
End Sub

Private Sub wv2_DevToolsProtocolEventReceived(ByRef parameterObjectAsJson As String)
    Debug.Print "DevToolsProtocolEventReceived. JSON:" & parameterObjectAsJson
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
    Call WV2Controller.WebView2.ExecuteScriptAsync("alert('Dialog On WebView2 !');")
End Sub

Private Sub UserForm_Initialize()
'    Me.width = 1920 * 0.75
'    Me.Height = 1080 * 0.75
'    Frame1.width = 1800 * 0.75
'    Frame1.Height = 1000 * 0.75
    
    #If Win64 Then
    Set NavigationCompletedHandler = New c4_Handler2
    #End If
    Call Create_WebView2
End Sub


'Create WebView2 In Frame
Public Sub Create_WebView2()

    'Use Hidden Property
    'Notified by KallunWillock via GitHub Issue. Thank you!
    TargetHwnd = Frame1.[_GethWnd]
    Debug.Print TargetHwnd
    
    Call WV2Loader.CreateWebView2Environment
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 1. Shut down the WebView2 process first
    Call WV2Controller.CloseWebView2
    
    ' 2. CRITICAL: Explicitly release references like Dictionaries
    If Not WV2Controller.WebView2 Is Nothing Then
        WV2Controller.WebView2.Finalize
    End If
    
    ' Currently, Class_Terminate won't fire unless the dictionary is released
    Set m_InstanceMap = Nothing
    
    ' 3. Finally, release the main controller reference
    Set WV2Controller = Nothing
    
    #If Win64 Then
    ' Release the handler that allocates/holds the Thunk memory area
    Set NavigationCompletedHandler = Nothing
    #End If
    
    Set Console = Nothing
    
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
        
    Dim Source As String
    Source = WV2Controller.WebView2.Source
    
    Dim Title As String
    Title = WV2Controller.WebView2.DocumentTitle
    
    Debug.Print "NavigationCompleted(From Standard Module) "
    Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title
    
    TextBox_URL.text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & WV2Controller.WebView2.ContainsFullScreenElement
    
End Sub

Private Sub WV2_NavigationStarting()
    
    Debug.Print "NavigationStarting"
End Sub

Private Sub WV2_SourceChanged()
    CommandButton_GoBack.Enabled = WV2Controller.WebView2.CanGoBack
    CommandButton_GoForward.Enabled = WV2Controller.WebView2.CanGoForward
    Debug.Print "SourceChanged"
End Sub

Private Sub WV2_WebMessageReceived()
    Debug.Print "WebMessageReceived"
End Sub

Private Sub WV2_WebResourceRequested()
    Debug.Print "WebResourceRequested"
End Sub

Private Sub wv2_WindowCloseRequested(ByVal this As LongLong, ByVal sender As LongLong, ByVal args As LongLong)
    Debug.Print "WindowCloseRequested"
End Sub

Private Sub WV2Controller_ScriptResultReceived(ByVal result As String)

    Debug.Print "ScriptResultReceived:", result
    TextBox_URL.text = WV2Controller.WebView2.Source

End Sub

Private Sub WV2Controller_WebView2ReadyCompleted()

    Debug.Print "WV2Controller_WebView2ReadyCompleted proccessid:" & WV2Controller.WebView2.BrowserProcessId
    Call WV2Controller.WebView2.NavigateAsync("https://www.google.com/")

End Sub

