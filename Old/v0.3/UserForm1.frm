VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   19575
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' =============================================================================
' UserForm1 コード例
' =============================================================================
' フォーム構成:
'   Frame1 (UIBarFrame) : Top=0, Height=76pt, 幅=フォーム幅
'   Frame2 (ContentFrame): Top=76, Height=残り, 幅=フォーム幅
'
' ※ Frame の BorderStyle = 0 (fmBorderStyleNone) 推奨
' =============================================================================
Option Explicit

Private WithEvents m_Manager As WebView2Manager
Attribute m_Manager.VB_VarHelpID = -1
Private WithEvents wv2 As WebView2
Attribute wv2.VB_VarHelpID = -1

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

Public Sub TestWithOldClass()
    ' v0.3 の Manager で Environment を作り、
    ' v0.2 の WebView2 クラスで Controller を作る
    Dim wv As New WebView2
    wv.EnvPtr = m_Manager.EnvPtr  ' v0.3 の Environment を渡す
    wv.StartControllerCreation Me.Frame2, "https://www.google.com"
End Sub

Private Sub m_Manager_EnvironmentReady()
    Debug.Print "環境準備完了 - testing controller..."
    'm_Manager.TestCreateController
    TestWithOldClass
End Sub
Private Sub UserForm_Initialize()
    'Set m_Manager = New WebView2Manager
'    m_Manager.CreateEnvironment Me.Frame1, Me.Frame2
    
End Sub

Public Sub StartBrowser()
    Set m_Manager = New WebView2Manager
    m_Manager.CreateEnvironment Me.Frame1, Me.Frame2
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Debug.Print "=== QueryClose ==="
    If Not m_Manager Is Nothing Then
        m_Manager.Finalize
        Debug.Print "=== Manager.Finalize done ==="
    End If
End Sub

'Private Sub UserForm_Terminate()
'    If Not m_Manager Is Nothing Then
'        m_Manager.Finalize
'        Set m_Manager = Nothing
'    End If
'End Sub
' UserForm 側
'Private Sub UserForm_Terminate()
'    Debug.Print "=== Form Terminate start ==="
'    If Not m_Manager Is Nothing Then
'        m_Manager.Finalize
'        Debug.Print "=== Manager.Finalize done ==="
'        Set m_Manager = Nothing
'        Debug.Print "=== Manager released ==="
'    End If
'    Debug.Print "=== Form Terminate end ==="
'End Sub
Private Sub UserForm_Terminate()
    Debug.Print "=== Form Terminate ==="
    Set m_Manager = Nothing
End Sub

' --- リサイズ対応 ---
Private Sub UserForm_Resize()
    If Not m_Manager Is Nothing Then
        m_Manager.ResizeAll
    End If
End Sub

' =============================================================================
' Manager イベントハンドラ
' =============================================================================

'Private Sub m_Manager_EnvironmentReady()
'    Debug.Print "環境準備完了"
'        m_Manager.CreateUIBarNow
'    m_Manager.CreateTab "https://www.google.com"
'
'    ' UIバーの準備ができた後、最初のタブを作成
'    ' (UIバーの ControllerCompleted 後に呼ぶのが安全)
'End Sub

Private Sub m_Manager_TabCreated(ByVal tabId As Long)
    Debug.Print "タブ作成: " & tabId
    ' 初回タブがなければ作成
    If m_Manager.TabCount = 0 Then
        m_Manager.CreateTab "https://www.google.com"
    End If
End Sub

Private Sub m_Manager_TabClosed(ByVal tabId As Long)
    Debug.Print "タブ閉じた: " & tabId
End Sub

Private Sub m_Manager_ActiveTabChanged(ByVal tabId As Long)
    Debug.Print "アクティブタブ変更: " & tabId
End Sub

Private Sub m_Manager_TabNavigationStarting( _
    ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "Tab[" & tabId & "] NavigationStarting"
End Sub

Private Sub m_Manager_TabNavigationCompleted( _
    ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "Tab[" & tabId & "] NavigationCompleted"
End Sub

Private Sub m_Manager_TabDocumentTitleChanged(ByVal tabId As Long, ByVal title As String)
    Debug.Print "Tab[" & tabId & "] Title: " & title
End Sub

Private Sub m_Manager_TabWebMessageReceived(ByVal tabId As Long, ByVal jsonMessage As String)
    Debug.Print "Tab[" & tabId & "] WebMessage: " & jsonMessage
End Sub

Private Sub m_Manager_TabNewWindowRequested(ByVal tabId As Long, ByVal uri As String)
    Debug.Print "Tab[" & tabId & "] NewWindow -> " & uri
End Sub

Private Sub m_Manager_TabProcessFailed( _
    ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
    Debug.Print "Tab[" & tabId & "] ProcessFailed!"
End Sub

Private Sub m_Manager_TabWindowCloseRequested(ByVal tabId As Long)
    Debug.Print "Tab[" & tabId & "] WindowCloseRequested"
End Sub

Private Sub m_Manager_TabScriptResult(ByVal tabId As Long, ByVal jsonResult As String)
    Debug.Print "Tab[" & tabId & "] Script Result: " & jsonResult
End Sub


