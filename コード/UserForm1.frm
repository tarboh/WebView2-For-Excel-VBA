VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   9120.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12360
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
Public WithEvents WV2 As c3_WebView2
Attribute WV2.VB_VarHelpID = -1
Public c5 As New c5_ObjectForJS

#If Win64 Then
Public WithEvents NavigationCompletedHandler As c4_Handler2
Attribute NavigationCompletedHandler.VB_VarHelpID = -1
#End If

Private Sub CommandButton1_Click()
    
    Dim url As String
    url = TextBox1.Text
        
    If Left(url, 11) = "javascript:" Then
        Call WV2Controller.WebView2.ExecuteScriptAsync(url)
    ElseIf Left(url, 4) = "http" Then
        Call WV2Controller.WebView2.NavigateAsync(url)
    Else
        Call WV2Controller.WebView2.NavigateToString(url)
    End If

End Sub


Private Sub CommandButton2_Click()
    Debug.Print WV2Controller.WebView2.Source
End Sub

#If Win64 Then
Private Sub NavigationCompletedHandler_Invoked(ByVal pThis As LongLong, ByVal sender As LongLong, ByVal args As LongLong)
    Debug.Print "C4_Handler2_NavigationCompleted!"
End Sub
#End If

Private Sub ToggleButton1_Change()

    If ToggleButton1.Value = True Then
        WV2.Settings.AreDefaultScriptDialogsEnabled = False
    Else
        WV2.Settings.AreDefaultScriptDialogsEnabled = True
    End If
    Debug.Print WV2.Settings.AreDefaultScriptDialogsEnabled
    WV2.Reload

End Sub

Private Sub WV2_ContainsFullScreenElementChanged()
    Debug.Print "ContainsFullScreenElementChanged"
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
    Debug.Print result
End Sub

Private Sub WV2_ScriptDialogOpening()
    Debug.Print "ScriptDialogOpening"
    MsgBox "Dialog On VBA!"
End Sub

Private Sub CommandButton3_Click()
    Call WV2Controller.WebView2.ExecuteScriptAsync("alert('Dialog On WebView2 !');")
End Sub

Private Sub UserForm_Activate()
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
    Debug.Print "NavigationCompleted(From Standard Module) Source:" & Source
    TextBox1.Text = Source
End Sub

Private Sub WV2_NavigationStarting()
    Debug.Print "NavigationStarting"
End Sub

Private Sub WV2_SourceChanged()
    Debug.Print "SourceChanged"
End Sub

Private Sub WV2_WebMessageReceived()
    Debug.Print "WebMessageReceived"
End Sub

Private Sub WV2_WebResourceRequested()
    Debug.Print "WebResourceRequested"
End Sub

Private Sub WV2Controller_ScriptResultReceived(ByVal result As String)

    Debug.Print "ScriptResultReceived:", result
    TextBox1.Text = WV2Controller.WebView2.Source

End Sub

Private Sub WV2Controller_WebView2ReadyCompleted()

    Call WV2.AddHostObjectToScript("VBAObj", c5)

    Call WV2Controller.WebView2.NavigateAsync("https://www.google.com/")

End Sub

