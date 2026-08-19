VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   11400
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



Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public c5 As New ObjectForJS

Private WithEvents Console As fm_Console
Attribute Console.VB_VarHelpID = -1
Private WithEvents wv2 As WebView2
Attribute wv2.VB_VarHelpID = -1

Private m_step As Long

Private rng As Range
Private v(1 To 10, 1 To 7) As String
#If メモ Then

    SolveRPAChallenge

    WV2_ReceiveScriptResult
    
#End If

Private WithEvents Header As WebView2
Attribute Header.VB_VarHelpID = -1
Private Sub HeaderSetUp()
    
    Set Header = New WebView2
    'header.
    
End Sub

Private Sub RPA()
    
    Dim script As String
    
    Select Case m_step
        Case 0
            
            'データ準備
            Set rng = Workbooks("challenge_ja.xlsx").Sheets("Sheet1").Range("A2:G11")
            
            Dim r As Long, c As Long
            For r = 1 To 10
                For c = 1 To 7
                    v(r, c) = rng.Cells(r, c).value
                Next
            Next
    
            'まずはスタートボタンをクリック
            wv2.ExecuteScriptAsync "document.querySelector('button.uiColorButton').click();"
            m_step = m_step + 1
            
        Case 1 To 10
    
            script = GetRPAScript("苗字", v(m_step, 1)) & _
                     GetRPAScript("名前", v(m_step, 2)) & _
                     GetRPAScript("会社名", v(m_step, 3)) & _
                     GetRPAScript("部署", v(m_step, 4)) & _
                     GetRPAScript("住所", v(m_step, 5)) & _
                     GetRPAScript("メールアドレス", v(m_step, 6)) & _
                     GetRPAScript("電話番号", v(m_step, 7)) & _
                     "document.querySelector('input.btn').click();" ' Submitボタン
                     
            wv2.ExecuteScriptAsync script
            'DoEvents
            m_step = m_step + 1

    End Select
    
End Sub


Private Sub CheckBox_Attach_c5ToJS_Click()
    If CheckBox_Attach_c5ToJS.value = True Then
        If wv2.AddHostObjectToScript("VBAObj", c5) = 0 Then
            Debug.Print "c5 attached as 'VBAObj'"
        Else
            Debug.Print "c5 attache failed"
        End If
    Else
        If wv2.RemoveHostObjectFromScript("VBAObj") = 0 Then
            Debug.Print "c5 remove success"
        Else
            Debug.Print "c5 remove failed"
        End If
    End If
End Sub

Private Sub CheckBox_BuiltInErrorPageEnabled_Click()
    wv2.IsBuiltInErrorPageEnabled = CheckBox_BuiltInErrorPageEnabled.value
    Debug.Print "IsBuiltInErrorPageEnabled:" & wv2.IsBuiltInErrorPageEnabled
End Sub

Private Sub CheckBox_Controller_IsVisible_Click()
    wv2.Controller_IsVisible = CheckBox_Controller_IsVisible.value
    Debug.Print "Controller_IsVisible:" & wv2.Controller_IsVisible
End Sub

Private Sub CheckBox_DefaultContextMenusEnabled_Change()
    wv2.AreDefaultContextMenusEnabled = CheckBox_DefaultContextMenusEnabled.value
    Debug.Print "AreDefaultContextMenusEnabled:" & wv2.AreDefaultContextMenusEnabled
End Sub

Private Sub CheckBox_DevToolsEnabled_Change()
    wv2.AreDevToolsEnabled = CheckBox_DevToolsEnabled.value
    Debug.Print "AreDevToolsEnabled:" & wv2.AreDevToolsEnabled
End Sub

Private Sub CheckBox_HostObjectsAllowed_Change()
    wv2.AreHostObjectsAllowed = CheckBox_HostObjectsAllowed.value
    Debug.Print "AreHostObjectsAllowed:" & wv2.AreHostObjectsAllowed
End Sub

Private Sub CheckBox_InterceptDialogs_Change()
    If CheckBox_InterceptDialogs.value = True Then
        wv2.AreDefaultScriptDialogsEnabled = False
    Else
        wv2.AreDefaultScriptDialogsEnabled = True
    End If
    Debug.Print wv2.AreDefaultScriptDialogsEnabled
    wv2.Reload
End Sub

Private Sub CheckBox_ScriptEnabled_Change()
    wv2.IsScriptEnabled = CheckBox_ScriptEnabled.value
End Sub

Private Sub CheckBox_StatusBarEnabled_change()
    wv2.IsStatusBarEnabled = CheckBox_StatusBarEnabled.value
End Sub

Private Sub CheckBox_WebMessageEnabled_Change()
    wv2.IsWebMessageEnabled = CheckBox_WebMessageEnabled.value
End Sub

Private Sub CheckBox_ZoomControlEnabled_change()
    wv2.IsZoomControlEnabled = CheckBox_ZoomControlEnabled.value
    Debug.Print "IsZoomControlEnabled:" & wv2.IsZoomControlEnabled
End Sub

Private Sub CmdBtn_SetBoundsAndZoomFactor_Click()
    Call wv2.Controller_SetBoundsAndZoomFactor(TextBox_Bounds_Left, TextBox_Bounds_Top, TextBox_Bounds_Right, TextBox_Bounds_Bottom, TextBox_Controller_ZoomFactor)
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
    Call wv2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
End Sub

Private Sub CommandButton_CapturePreviewToFile_Click()
    
    Dim folderPath As String
    folderPath = "C:\temp\VBA_WebView2\ScreenShot\"
    
    Dim uniquePath As String
    uniquePath = "cap_" & format(Now, "yyyymmdd_hhnnss") & "_" & Right("000" & Int(Timer * 1000) Mod 1000, 3) & ".png"
    
    wv2.CapturePreviewToFile folderPath, uniquePath
End Sub

Private Sub CommandButton_Console_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
End Sub



Private Sub CommandButton_Controller_Close_Click()
    wv2.Controller_Close
End Sub

Private Sub CommandButton_Controller_Get_ParentWindow_Click()
    Dim hwnd As LongPtr
    hwnd = wv2.Controller_ParentWindow
    TextBox_Controller_ParentWindow.Text = hwnd
End Sub

Private Sub CommandButton_Controller_get_ZoomFactor_Click()
    TextBox_Controller_ZoomFactor.Text = wv2.Controller_ZoomFactor
End Sub

Private Sub CommandButton_Controller_MoveFocus_Click()
    Dim reason As COREWEBVIEW2_MOVE_FOCUS_REASON
    reason = ComboBox_MOVE_FOCUS_REASON.ListIndex
    Call wv2.Controller_MoveFocus(reason)
End Sub

Private Sub CommandButton_Controller_Put_ZoomFactor_Click()
    wv2.Controller_ZoomFactor = TextBox_Controller_ZoomFactor.Text
End Sub

Private Sub CommandButton_ExeCuteVBAInJavaScript_Click()
    Call wv2.ExecuteScriptAsync("window.chrome.webview.hostObjects.sync.VBAObj.Func1(15);")
End Sub

Private Sub CommandButton_GoBack_Click()
    wv2.GoBack
End Sub

Private Sub CommandButton_GoForward_Click()
    wv2.GoForward
End Sub

Private Sub CommandButton_Navigate_Click()
    
    Dim url As String
    url = TextBox_URL.Text
        
    If Left(url, 11) = "javascript:" Then
        Call wv2.ExecuteScriptAsync(url)
    ElseIf Left(url, 4) = "http" Then
        Call wv2.NavigateAsync(url)
    Else
        Call wv2.NavigateToString(url)
    End If

End Sub


Private Sub CommandButton_NavToStr_Click()
    If Console Is Nothing Then Set Console = New fm_Console
    Console.Show
    
    Dim uri As String
    uri = Console.TextBox_Console.Text
    Debug.Print uri
    Call wv2.NavigateToString(uri)
End Sub

Private Sub CommandButton_Open_Click()
    Call Create_WebView2
End Sub

Private Sub CommandButton_OpenDevTools_Click()
    wv2.OpenDevToolsWindow
End Sub

Private Sub CommandButton_PostWebMessageAsJson_Click()
    Dim strJson As String
    strJson = "{""funcName"": ""calculateAndDisplay"", ""args"": [""Sum Result"", 123, 456]}"
    Debug.Print wv2.PostWebMessageAsJson(strJson)
End Sub

Private Sub CommandButton_PostWebMessageAsString_Click()
    Dim webMessage As String
    webMessage = "System Check Complete"
    Debug.Print wv2.PostWebMessageAsString(webMessage)
End Sub

Private Sub CommandButton_Reload_Click()
    Call wv2.Reload
End Sub

Private Sub CommandButton_RunScript_Click()
    Dim script As String
    script = TextBox_Script.Text
    Call wv2.ExecuteScriptAsync(script)
End Sub

Private Sub CommandButton_Stop_Click()
    wv2.Stop_
End Sub

Private Sub CommandButton_StopAutoJS_Click()
'    WV2.RemoveScriptToExecuteOnDocumentCreated ( _
'        WV2.scriptId)
End Sub

Private Sub CommandButton10_Click()

    'Call SolveRPAChallenge

    Call RPA

End Sub

Public Sub SolveRPAChallenge()
    ' 1. まずはスタートボタンをクリック
    'wv2.ExecuteScriptAsync "document.querySelector('button.uiColorButton').click();"
    
    ' 2. 各フィールドに値を入力（Excelシートからデータを取る想定）
    ' ※ここでは例として固定値を入力するJSを投げます
    
    Dim script As String
    script = GetRPAScript("名前", "Taro") & _
             GetRPAScript("苗字", "Tanaka") & _
             GetRPAScript("会社名", "VBA-Hacker Inc.") & _
             GetRPAScript("部署", "Architect") & _
             GetRPAScript("住所", "Tokyo, Japan") & _
             GetRPAScript("メールアドレス", "taro@example.com") & _
             GetRPAScript("電話番号", "0123456789") & _
             "document.querySelector('input.btn').click();" ' Submitボタン
             
    wv2.ExecuteScriptAsync script
End Sub

' 特定のラベルに対応するinputに値をセットするJSコードを生成
Private Function GetRPAScript(labelName As String, value As String) As String
    GetRPAScript = "document.evaluate(""//label[contains(., '" & labelName & "')]/following-sibling::input"", " & _
                   "document, null, XPathResult.FIRST_ORDERED_NODE_TYPE, null).singleNodeValue.value = '" & value & "';"
End Function


Private Sub CommandButton4_Click()
    
    Call wv2.add_DevToolsProtocolEventReceived("Network.responseReceived")

    ' ネットワーク監視機能を有効化する（これを投げないとイベントが来ない）
    Dim strMethodName As String
    strMethodName = "Network.enable"
    
    Dim strParametersAsJson As String
    strParametersAsJson = "{}" ' パラメータは空のJSONオブジェクトでOK
    
    Dim hr As Long
    hr = wv2.CallDevToolsProtocolMethod(strMethodName, strParametersAsJson)
    Debug.Print "登録結果：" & hr
    
End Sub

Private Sub CommandButton5_Click()
    Call wv2.AddWebResourceRequestedFilter("*", COREWEBVIEW2_WEB_RESOURCE_CONTEXT_IMAGE)
    Call wv2.add_WebResourceRequested
End Sub


Private Sub CommandButton_Controller_get_Bounds_Click()
    Dim hr As Long
    Dim l(3) As Long
    hr = wv2.Controller_get_Bounds(l)
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
    hr = wv2.Controller_put_Bounds(l)
    Debug.Print "put_Bounds hr:" & hr
End Sub



Private Sub CommandButton7_Click()
    wv2.Controller_NotifyParentWindowPositionChanged
End Sub

Private Sub CommandButton8_Click()
        
    
    Dim html As String
    html = html & "<!DOCTYPE html>                                                                              " & vbCrLf
    html = html & "<html>                                                                                       " & vbCrLf
    html = html & "<head>                                                                                       " & vbCrLf
    html = html & "    <meta charset=""UTF-8"">                                                                   " & vbCrLf
    html = html & "    <style>                                                                                  " & vbCrLf
    html = html & "        body {                                                                               " & vbCrLf
    html = html & "            background: #0f172a;                                                             " & vbCrLf
    html = html & "            color: #f8fafc;                                                                  " & vbCrLf
    html = html & "            font-family: 'Segoe UI', sans-serif;                                             " & vbCrLf
    html = html & "            display: flex;                                                                   " & vbCrLf
    html = html & "            flex-direction: column;                                                          " & vbCrLf
    html = html & "            justify-content: center;                                                         " & vbCrLf
    html = html & "            align-items: center;                                                             " & vbCrLf
    html = html & "            height: 100vh;                                                                   " & vbCrLf
    html = html & "            margin: 0;                                                                       " & vbCrLf
    html = html & "            overflow: hidden;                                                                " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        .container {                                                                         " & vbCrLf
    html = html & "            width: 80%;                                                                      " & vbCrLf
    html = html & "            max-width: 400px;                                                                " & vbCrLf
    html = html & "            text-align: center;                                                              " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        .status-text {                                                                       " & vbCrLf
    html = html & "            font-size: 1.2rem;                                                               " & vbCrLf
    html = html & "            margin-bottom: 20px;                                                             " & vbCrLf
    html = html & "            font-weight: 300;                                                                " & vbCrLf
    html = html & "            letter-spacing: 0.1em;                                                           " & vbCrLf
    html = html & "            color: #38bdf8;                                                                  " & vbCrLf
    html = html & "            text-shadow: 0 0 10px rgba(56, 189, 248, 0.5);                                   " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        /* プログレスバーの外枠 */                                                           " & vbCrLf
    html = html & "        .progress-bg {                                                                       " & vbCrLf
    html = html & "            background: rgba(30, 41, 59, 0.8);                                               " & vbCrLf
    html = html & "            height: 8px;                                                                     " & vbCrLf
    html = html & "            border-radius: 4px;                                                              " & vbCrLf
    html = html & "            overflow: hidden;                                                                " & vbCrLf
    html = html & "            position: relative;                                                              " & vbCrLf
    html = html & "            box-shadow: inset 0 2px 4px rgba(0,0,0,0.3);                                     " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        /* 動くバー本体 */                                                                   " & vbCrLf
    html = html & "        .progress-fill {                                                                     " & vbCrLf
    html = html & "            width: 0%;                                                                       " & vbCrLf
    html = html & "            height: 100%;                                                                    " & vbCrLf
    html = html & "            background: linear-gradient(90deg, #0ea5e9, #22d3ee);                            " & vbCrLf
    html = html & "            box-shadow: 0 0 15px #0ea5e9;                                                    " & vbCrLf
    html = html & "            transition: width 0.4s cubic-bezier(0.4, 0, 0.2, 1);                             " & vbCrLf
    html = html & "            position: relative;                                                              " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        /* 流れる光のエフェクト */                                                           " & vbCrLf
    html = html & "        .progress-fill::after {                                                              " & vbCrLf
    html = html & "            content: "";                                                                     " & vbCrLf
    html = html & "            position: absolute;                                                              " & vbCrLf
    html = html & "            top: 0; left: 0; bottom: 0; right: 0;                                            " & vbCrLf
    html = html & "            background: linear-gradient(                                                     " & vbCrLf
    html = html & "                90deg,                                                                       " & vbCrLf
    html = html & "                transparent,                                                                 " & vbCrLf
    html = html & "                rgba(255, 255, 255, 0.4),                                                    " & vbCrLf
    html = html & "                transparent                                                                  " & vbCrLf
    html = html & "            );                                                                               " & vbCrLf
    html = html & "            animation: shine 1.5s infinite;                                                  " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        @keyframes shine {                                                                   " & vbCrLf
    html = html & "            from { transform: translateX(-100%); }                                           " & vbCrLf
    html = html & "            to { transform: translateX(100%); }                                              " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "        .percentage {                                                                        " & vbCrLf
    html = html & "            margin-top: 10px;                                                                " & vbCrLf
    html = html & "            font-family: 'Consolas', monospace;                                              " & vbCrLf
    html = html & "            font-size: 0.9rem;                                                               " & vbCrLf
    html = html & "            color: #94a3b8;                                                                  " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "    </style>                                                                                 " & vbCrLf
    html = html & "</head>                                                                                      " & vbCrLf
    html = html & "<body>                                                                                       " & vbCrLf
    html = html & "    <div class=""container"">                                                                  " & vbCrLf
    html = html & "        <div id=""status"" class=""status-text"">SYSTEM INITIALIZING...</div>                    " & vbCrLf
    html = html & "        <div class=""progress-bg"">                                                            " & vbCrLf
    html = html & "            <div id=""bar"" class=""progress-fill""></div>                                       " & vbCrLf
    html = html & "        </div>                                                                               " & vbCrLf
    html = html & "        <div id=""percent"" class=""percentage"">0%</div>                                        " & vbCrLf
    html = html & "    </div>                                                                                   " & vbCrLf
    html = html & "                                                                                             " & vbCrLf
    html = html & "    <script>                                                                                 " & vbCrLf
    html = html & "        // VBA側から ExecuteScript(""updateProgress(50, 'DATA EXTRACTION...')"") のように叩く用" & vbCrLf
    html = html & "        function updateProgress(val, text) {                                                 " & vbCrLf
    html = html & "            const bar = document.getElementById('bar');                                      " & vbCrLf
    html = html & "            const status = document.getElementById('status');                                " & vbCrLf
    html = html & "            const percent = document.getElementById('percent');                              " & vbCrLf
    html = html & "                                                                                             " & vbCrLf
    html = html & "            val = Math.min(Math.max(val, 0), 100);                                           " & vbCrLf
    html = html & "            bar.style.width = val + '%';                                                     " & vbCrLf
    html = html & "            percent.innerText = val + '%';                                                   " & vbCrLf
    html = html & "            if(text) status.innerText = text;                                                " & vbCrLf
    html = html & "        }                                                                                    " & vbCrLf
    html = html & "    </script>                                                                                " & vbCrLf
    html = html & "</body>                                                                                      " & vbCrLf
    html = html & "</html>                                                                                      " & vbCrLf
    
    wv2.NavigateToString html
    
    Dim i As Long
    For i = 0 To 100
        UpdateUI i, "Loading..."
        Sleep 100
        DoEvents
    Next
    
    UpdateUI i, "Completed!!"
    
End Sub

Public Sub UpdateUI(ByVal percent As Integer, ByVal msg As String)
    wv2.ExecuteScriptAsync "updateProgress(" & percent & ", '" & msg & "');"
End Sub


Private Sub CommandButton_VersionInfo_Click()
    MsgBox "WebView2 Version : " & wv2.Environment_BrowserVersionString
End Sub

Private Sub CommandButton9_Click()

    wv2.NavigateAsync "https://www.rpachallenge.com/?lang=ja"
    m_step = 0
    'wv2.ExecuteScriptAsync "document.getElementById(""t3eXO"").value = ""a"";"

End Sub

Private Sub CommandButtonController_Put_ParentWindow_Click()
    wv2.Controller_ParentWindow = TextBox_Controller_ParentWindow.Text
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
    Title = wv2.DocumentTitle
    
    Debug.Print "NavigationCompleted(From Standard Module) "
    'Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title
    
    'TextBox_URL.text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & wv2.ContainsFullScreenElement
End Sub

Private Sub WV2_ControllerAcceleratorKeyPressed(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "ControllerAcceleratorKeyPressed"
End Sub

Private Sub WV2_ControllerGotFocus()
    Debug.Print "ControllerGotFocus"
End Sub

Private Sub WV2_ControllerLostFocus()
    Debug.Print "ControllerLostFocus"
End Sub

Private Sub WV2_ControllerMoveFocusRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "ControllerMoveFocuceRequested"
End Sub

Private Sub WV2_ControllerZoomFactorChanged()
    Debug.Print "ControllerZoomChanged ZoomFactor:" & wv2.Controller_ZoomFactor
    TextBox_Controller_ZoomFactor.Text = wv2.Controller_ZoomFactor
End Sub

Private Sub wv2_DevToolsProtocolEventReceived(ByVal eventName As String, ByVal parameterJson As String)
    Debug.Print "DevToolsProtocolEventReceived. JSON:" & parameterJson
End Sub

Private Sub WV2_DocumentTitleChanged()
    Debug.Print "DocumentTitleChanged"
End Sub

Private Sub WV2_EnvironmentNewBrowserVersionAvailable()
    Debug.Print "NewBrowserVersionAvailable"
End Sub

Private Sub WV2_NewWindowRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "NewWindowRequested"
    
    'argsを解析
    'MIDL_INTERFACE ("34acb11c-fc37-4418-9132-f9c21d1eafb9")
    'ICoreWebView2NewWindowRequestedEventArgs:  Public IUnknown
    '{
    'public:
    '3    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Uri(
    '        /* [retval][out] */ LPWSTR *uri) = 0;
    Dim puri As LongPtr
    Dim hr As Long
    hr = dcf(args, 3, "get_Uri", VarPtr(puri))
    Dim uri As String
    uri = PtrToString(puri)
    CoTaskMemFree puri
    Debug.Print uri
    '
    '4    virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_NewWindow(
    '        /* [in] */ ICoreWebView2 *newWindow) = 0;
    '
    '5    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_NewWindow(
    '        /* [retval][out] */ ICoreWebView2 **newWindow) = 0;
'    Dim pnewWindow As LongPtr
'    hr = dcf(args, 5, "get_NewWindow", VarPtr(pnewWindow))
'    Debug.Print pnewWindow
'    Stop
    '
    '6    virtual /* [propput] */ HRESULT STDMETHODCALLTYPE put_Handled(
    '        /* [in] */ BOOL handled) = 0;
    hr = dcf(args, 6, "put_Handled", 1)
    '
    '7    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Handled(
    '        /* [retval][out] */ BOOL *handled) = 0;
    '
    '8    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_IsUserInitiated(
    '        /* [retval][out] */ BOOL *isUserInitiated) = 0;
    '
    '9    virtual HRESULT STDMETHODCALLTYPE GetDeferral(
    '        /* [retval][out] */ ICoreWebView2Deferral **deferral) = 0;
    '
    '10    virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_WindowFeatures(
    '        /* [retval][out] */ ICoreWebView2WindowFeatures **value) = 0;
    '
    '};
    
    Dim NewWindow As New UserForm1
    NewWindow.Show
    Call NewWindow.Create_WebView2(wv2.EnvPtr, uri)
    
End Sub

Public Sub ShowPopupWindow(url As String)
    'Environmentだけ引き継いで、CreateCorewWebView2Controllerのシーケンスから実行させたい
    Dim fm As New UserForm1
    
End Sub

Public Sub ShowBrowserWindow(wv2 As WebView2, url As String)
    Set wv2 = New WebView2
    Call wv2.CreateWebView2Environment(Frame1, url)
End Sub

Private Sub WV2_PermissionRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "PermissionRequested"
End Sub

Private Sub WV2_ProcessFailed(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "ProcessFailed"
End Sub

Private Sub WV2_ReceiveScriptResult(ByVal result As String)
    'Debug.Print "ReceiveScriptResult result : " & result
    Call RPA
'    If m_step = 11 Then
'        m_step = 0
'    End If
End Sub

Private Sub WV2_ScriptDialogOpening(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "ScriptDialogOpening"
    MsgBox "Dialog On VBA!"
End Sub

Private Sub CommandButton3_Click()
    Call wv2.ExecuteScriptAsync("alert('Dialog On WebView2 !');")
End Sub

Private Sub UserForm_Initialize()

'    Me.width = 1920 * 0.75
'    Me.Height = 1080 * 0.75
'    Frame1.width = 1800 * 0.75
'    Frame1.Height = 1000 * 0.75
    
'    Set wv2 = New c3_WebView2
'    Call wv2.BuildFuncPtrCache
    
    Set c5 = New ObjectForJS
    
    ComboBox_MOVE_FOCUS_REASON.AddItem "PROGRAMMATIC"
    ComboBox_MOVE_FOCUS_REASON.AddItem "NEXT"
    ComboBox_MOVE_FOCUS_REASON.AddItem "PREVIOUS"
    ComboBox_MOVE_FOCUS_REASON.ListIndex = 0
    
    'Call Create_WebView2
End Sub

'Create WebView2 In Frame
Public Sub Create_WebView2(Optional EnvPtr As LongPtr, Optional StartURL As String)

    Set wv2 = New WebView2
    If EnvPtr = 0 Then
        Call wv2.CreateWebView2Environment(Frame1, StartURL) 'targetHWnd)
    Else
        wv2.EnvPtr = EnvPtr
        Call wv2.StartControllerCreation(Frame1, StartURL)
    End If
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 1. まずハンドラを全解除
        If Not wv2 Is Nothing Then
            wv2.Finalize
        End If
    
    ' 2. WebView2プロセスをシャットダウン
    
        wv2.Controller_Close
    
    ' 3. 参照を解放
    
    Debug.Print "Console解放前"
    Set Console = Nothing
    Debug.Print "Console解放後"
    
    Debug.Print "QueryClose完了"
End Sub

Private Sub WV2_ContentLoading(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "ContentLoading"
End Sub

Private Sub WV2_FrameNavigationCompleted(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "FrameNavigationCompleted"
End Sub

Private Sub WV2_FrameNavigationStarting(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "FrameNavigationStarting"
End Sub

Private Sub WV2_HistoryChanged()
    Debug.Print "HistoryChanged"
End Sub

Private Sub WV2_NavigationCompleted(ByVal sender As LongPtr, ByVal args As LongPtr)
     
'    Debug.Print "NavigationCompleted"
    Dim Source As String
    Source = wv2.Source

    Dim Title As String
    Title = wv2.DocumentTitle

    Debug.Print "NavigationCompleted(From Standard Module) "
    Debug.Print "    Source : " & Source
    Debug.Print "    Title  : " & Title

    TextBox_URL.Text = Source
    Me.Caption = Title & " ContainsFullScreenElement:" & wv2.ContainsFullScreenElement
    
End Sub

Private Sub WV2_NavigationStarting(ByVal sender As LongPtr, ByVal args As LongPtr)
    
    Debug.Print "NavigationStarting"
End Sub

Private Sub WV2_SourceChanged(ByVal sender As LongPtr, ByVal args As LongPtr)
    CommandButton_GoBack.Enabled = wv2.CanGoBack
    CommandButton_GoForward.Enabled = wv2.CanGoForward
    Debug.Print "SourceChanged"
End Sub

Private Sub WV2_WebMessageReceived(ByVal Source As String, ByVal messageJson As String, ByVal messageString As String)
    Debug.Print "WebMessageReceived"
    Debug.Print "    source        :" & Source
    Debug.Print "    mssage(json)  :" & messageJson
    Debug.Print "    mssage(string):" & messageString
End Sub

Private Sub WV2_WebResourceRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "WebResourceRequested"
End Sub

Private Sub wv2_WindowCloseRequested(ByVal sender As LongPtr, ByVal args As LongPtr)
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

Public Sub Navigate(url As String)
    
    wv2.NavigateAsync url
End Sub
