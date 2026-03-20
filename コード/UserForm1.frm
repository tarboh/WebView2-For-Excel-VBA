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
Public WithEvents NavigationCompletedHandler As c4_Handler2
Attribute NavigationCompletedHandler.VB_VarHelpID = -1
Public c5 As New c5_ObjectForJS

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

Private Sub NavigationCompletedHandler_Invoked(ByVal pThis As LongLong, ByVal sender As LongLong, ByVal args As LongLong)
    Debug.Print "C4_Handler2_NavigationCompleted!"
End Sub

Private Sub UserForm_Activate()
    Set NavigationCompletedHandler = New c4_Handler2
    Call WebView2錬成
End Sub

'フォームにWebView2を生成する処理
Public Sub WebView2錬成()

    '隠しプロパティを使えば直接ウィンドウハンドルが取得できる
    '※KallunWillockさんからのIssueで教えてもらいました。ありがとう。
    TargetHwnd = Frame1.[_GethWnd]
    Debug.Print TargetHwnd
    
    Call WV2Loader.CreateWebView2Environment
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' 1. まず WebView2 本体のプロセスを止める
    Call WV2Controller.CloseWebView2
    
    ' 2. 重要：Dictionary 等の参照を明示的に外す
    If Not WV2Controller.WebView2 Is Nothing Then
        WV2Controller.WebView2.Finalize
    End If
    
    '現状、辞書を解放しないとTerminateが発動しない
    Set m_InstanceMap = Nothing
    
    ' 3. 最後に参照を切る
    Set WV2Controller = Nothing
    
    'サンクを領域展開しているハンドラを消す
    Set NavigationCompletedHandler = Nothing
End Sub

Private Sub WV2_NavigationCompleted()
    Dim Source As String
    Source = WV2Controller.WebView2.Source
    Debug.Print "（標準モジュール由来）NavigationCompleted Source:" & Source
    TextBox1.Text = Source
End Sub

Private Sub WV2_NavigationStarting()
    Debug.Print "NavigationStarting"
End Sub

Private Sub WV2Controller_ScriptResultReceived(result As String)

    Debug.Print "ScriptResultReceived:", result
    TextBox1.Text = WV2Controller.WebView2.Source

End Sub

Private Sub WV2Controller_WebVeiw2ReadyCompleted()

    Call WV2.AddHostObjectToScript("VBAObj", c5)

    Call WV2Controller.WebView2.NavigateAsync("https://www.google.co.jp")

End Sub

