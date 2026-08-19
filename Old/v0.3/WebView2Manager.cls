VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' WebView2Manager.cls  (Core集約版)
' =============================================================================
' タブブラウザのメインコントローラー。
' Environment 保持、Tab 配列管理、UIバー、フォームへの RaiseEvent 中継。
' ハンドラ管理は Core に委譲。プレフィックスは "MGR."。
' =============================================================================

#If Win64 Then
#Else
    #Error "64 ビット VBA 専用"
#End If

Private Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader" ( _
    ByVal browserExecutableFolder As LongPtr, ByVal userDataFolder As LongPtr, _
    ByVal environmentOptions As LongPtr, ByVal environmentCreatedHandler As LongPtr) As Long

Private Const S_OK As Long = 0
Private Const MGR_PREFIX As String = "MGR."

' =============================================================================
' フォームへのイベント
' =============================================================================

Public Event TabNavigationStarting(ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event TabContentLoading(ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event TabSourceChanged(ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event TabHistoryChanged(ByVal tabId As Long)
Public Event TabNavigationCompleted(ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event TabDocumentTitleChanged(ByVal tabId As Long, ByVal title As String)
Public Event TabWebMessageReceived(ByVal tabId As Long, ByVal jsonMessage As String)
Public Event TabNewWindowRequested(ByVal tabId As Long, ByVal uri As String)
Public Event TabProcessFailed(ByVal tabId As Long, ByVal sender As LongPtr, ByVal args As LongPtr)
Public Event TabWindowCloseRequested(ByVal tabId As Long)
Public Event TabScriptResult(ByVal tabId As Long, ByVal jsonResult As String)
Public Event EnvironmentReady()
Public Event TabCreated(ByVal tabId As Long)
Public Event TabClosed(ByVal tabId As Long)
Public Event ActiveTabChanged(ByVal tabId As Long)

' =============================================================================
' メンバ
' =============================================================================

Private m_Core As WebView2Core
Private m_funcPtrs As Object       ' Manager 自身のメソッドアドレス

Private m_pEnv As LongPtr

Private m_Tabs As Object           ' Dictionary: tabId -> WebView2Tab
Private m_TabOrder() As Long
Private m_TabCount As Long
Private m_NextTabId As Long
Private m_ActiveTabId As Long

Private m_UIBar As WebView2Tab
Private m_UIBarFrame As Object
Private m_ContentFrame As MSForms.Control

' 保留タブ
Private Type PendingTab
    tabId As Long
    StartURL As String
End Type
Private m_Pending() As PendingTab
Private m_PendingCount As Long

Private m_EnvReady As Boolean
Private m_Initialized As Boolean


' =============================================================================
' ライフサイクル
' =============================================================================

Private Sub Class_Initialize()
    Set m_Core = New WebView2Core
    Set m_Tabs = CreateObject("Scripting.Dictionary")
    Set m_funcPtrs = CreateObject("Scripting.Dictionary")
    m_funcPtrs.CompareMode = 1
    m_NextTabId = 1: m_ActiveTabId = -1
    m_Core.BuildFuncPtrCache ObjPtr(Me), m_funcPtrs
    m_Initialized = True
' Manager.Class_Initialize の最後に追加
Debug.Print "Manager funcPtrs count: " & m_funcPtrs.Count
Debug.Print "Has Invoke_EnvironmentCompleted: " & m_funcPtrs.Exists("Invoke_EnvironmentCompleted")
If m_funcPtrs.Exists("Invoke_EnvironmentCompleted") Then
    Dim pp As LongPtr: pp = m_funcPtrs("Invoke_EnvironmentCompleted")
    Debug.Print "  pFunc = &H" & Hex(pp)
End If
End Sub

Private Sub Class_Terminate()
    ' Finalize は UserForm から明示的に呼ばれる
    Debug.Print "Manager.Class_Terminate called"
    'Finalize
End Sub

Public Property Get EnvPtr() As LongPtr
    EnvPtr = m_pEnv
End Property

Public Function Invoke_TabControllerCompleted( _
    ByVal pThis As LongPtr, ByVal errorCode As LongPtr, ByVal pCtrl As LongPtr) As Long

    Debug.Print "Mgr.Invoke_TabControllerCompleted entered err=&H" & Hex(errorCode)
    Invoke_TabControllerCompleted = 0
End Function

' Manager に追加（元のコードのやり方をそのまま再現）
'Public Sub TestCreateController()
'    Dim fo As LongPtr
'    fo = m_Core.EnsureCallback(MGR_PREFIX, ObjPtr(Me), _
'        "__ctrl_test", "Invoke_TestCtrlCompleted", m_funcPtrs)
'    If fo = 0 Then
'        Debug.Print "TestCreateController: fakeObj failed"
'        Exit Sub
'    End If
'
'    Dim hwnd As LongPtr
'    hwnd = m_ContentFrame.[_GethWnd] 'CallByName(m_ContentFrame, "_GethWnd", VbGet)
'    Debug.Print "TestCreateController hwnd=&H" & Hex(hwnd) & " fo=&H" & Hex(fo)
'
'    Dim hr As Long
'    hr = m_Core.dcf(m_pEnv, 3, "CreateCoreWebView2Controller", hwnd, fo)
'    Debug.Print "TestCreateController hr=&H" & Hex(hr)
'End Sub
Public Sub TestCreateController()
    Dim fo As LongPtr
    fo = m_Core.EnsureCallback(MGR_PREFIX, ObjPtr(Me), _
        "__ctrl_test", "Invoke_TestCtrlCompleted", m_funcPtrs)
    If fo = 0 Then Exit Sub
    
    Dim hwnd As LongPtr
    hwnd = m_ContentFrame.[_GethWnd] 'CallByName(m_ContentFrame, "_GethWnd", VbGet)
    Debug.Print "TestCreateController hwnd=&H" & Hex(hwnd) & " fo=&H" & Hex(fo)
    
    ' Core.dcf ではなく元の標準モジュールの dcf を使う
    Dim hr As Long
    hr = dcf(m_pEnv, 3, "CreateCoreWebView2Controller", hwnd, fo)
    Debug.Print "TestCreateController hr=&H" & Hex(hr)
End Sub


Public Function Invoke_TestCtrlCompleted( _
    ByVal pThis As LongPtr, ByVal errCode As LongPtr, ByVal pCtrl As LongPtr) As Long
    Debug.Print "=== TestCtrlCompleted! err=&H" & Hex(errCode) & " pCtrl=&H" & Hex(pCtrl)
    Invoke_TestCtrlCompleted = 0
End Function

' =============================================================================
' 公開 API
' =============================================================================

Public Sub CreateEnvironment(ByVal UIBarFrame As Object, ByVal ContentFrame As Object)
    Set m_UIBarFrame = UIBarFrame
    Set m_ContentFrame = ContentFrame

    Dim fo As LongPtr
    fo = m_Core.EnsureCallback(MGR_PREFIX, ObjPtr(Me), _
        "__env_completed", "Invoke_EnvironmentCompleted", m_funcPtrs)
    If fo = 0 Then Exit Sub

    Dim udp As String: udp = "C:\Temp\VBA_WebView2_v03"
    m_Core.CreateDeepFolder "C:\Temp\"

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(udp), 0, fo)
    If hr <> S_OK Then Debug.Print "Mgr: env failed &H" & Hex(hr)
End Sub

'Public Function Invoke_EnvironmentCompleted( _
'    ByVal pThis As LongPtr, ByVal errCode As LongPtr, ByVal pEnv As LongPtr) As Long
'
'    If errCode <> 0 Then
'        Debug.Print "Mgr: env error &H" & Hex(errCode)
'        Exit Function
'    End If
'
'    m_pEnv = pEnv
'
''    m_Core.dcf m_pEnv, 1, "Env.AddRef"
'    ' AddRef を直接 DispCallFunc で呼ぶ (dcf の ParamArray を経由しない)
''    Dim rc As Long
''    rc = m_Core.AddRefCOM(m_pEnv)
''    Debug.Print "Env.AddRef refCount=" & rc
''
''    ' AddRef/Release は行わない - COM 側が参照を管理
''    Debug.Print "Mgr: Environment Ready &H" & Hex(m_pEnv)
'
'    ' AddRef して VBA 側で保持（Release は呼ばない）
'    Dim rc As Long
'    rc = m_Core.AddRefCOM(m_pEnv)
'    Debug.Print "Mgr: Environment Ready, refCount=" & rc
'
'    m_EnvReady = True
'
'
'    RaiseEvent EnvironmentReady
'
'
'    ' UIバー・タブ作成は COM コールバックから戻った後に実行
'    ' CreateUIBar をここでは呼ばない
'
'    'Exit Function
'
''    ' UIバー作成
''    Debug.Print "Mgr: About to CreateUIBar..."
''    CreateUIBar
''    Debug.Print "Mgr: CreateUIBar done."
''
''    ' 保留タブ
''    ' 保留タブはまだ止めておく
'''    Dim i As Long
'''    For i = 0 To m_PendingCount - 1
'''        CreateTabInternal m_Pending(i).tabId, m_Pending(i).StartURL, False
'''    Next i
''    m_PendingCount = 0: Erase m_Pending
'
'    Invoke_EnvironmentCompleted = 0
'End Function

Public Function Invoke_EnvironmentCompleted( _
    ByVal pThis As LongPtr, ByVal errCode As LongPtr, ByVal pEnv As LongPtr) As Long

    If errCode <> 0 Then Exit Function

    m_pEnv = pEnv
    Dim rc As Long
    rc = m_Core.AddRefCOM(m_pEnv)
    m_EnvReady = True
    Debug.Print "Mgr: Environment Ready, refCount=" & rc

    RaiseEvent EnvironmentReady

    ' ここで直接呼ぶ
    Debug.Print "Mgr: calling TestCreateController directly..."
    TestCreateController

    Invoke_EnvironmentCompleted = 0
End Function

''' <summary>新タブ作成。戻り値 = tabId。</summary>
Public Function CreateTab(Optional ByVal StartURL As String = "about:blank") As Long
    Dim tid As Long: tid = m_NextTabId: m_NextTabId = m_NextTabId + 1

    If m_EnvReady Then
        CreateTabInternal tid, StartURL, False
    Else
        m_PendingCount = m_PendingCount + 1
        ReDim Preserve m_Pending(0 To m_PendingCount - 1)
        m_Pending(m_PendingCount - 1).tabId = tid
        m_Pending(m_PendingCount - 1).StartURL = StartURL
    End If
    CreateTab = tid
End Function

Public Sub CreateUIBarNow()
    If m_EnvReady Then CreateUIBar
End Sub

Public Sub SwitchTab(ByVal tabId As Long)
    If Not m_Tabs.Exists(tabId) Then Exit Sub
    If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then
        Dim old As WebView2Tab: Set old = m_Tabs(m_ActiveTabId)
        old.SetVisible False
    End If
    m_ActiveTabId = tabId
    Dim nw As WebView2Tab: Set nw = m_Tabs(tabId)
    nw.SetVisible True: nw.ResizeWebView
    RaiseEvent ActiveTabChanged(tabId)
    UpdateUIBar
End Sub

Public Sub CloseTab(ByVal tabId As Long)
    If Not m_Tabs.Exists(tabId) Then Exit Sub
    Dim t As WebView2Tab: Set t = m_Tabs(tabId)
    t.Finalize: m_Tabs.Remove tabId
    RemoveFromOrder tabId
    RaiseEvent TabClosed(tabId)
    If m_ActiveTabId = tabId Then
        If m_TabCount > 0 Then SwitchTab m_TabOrder(m_TabCount - 1) _
        Else m_ActiveTabId = -1
    End If
    UpdateUIBar
End Sub

Public Property Get ActiveTabId() As Long: ActiveTabId = m_ActiveTabId: End Property
Public Property Get ActiveTab() As WebView2Tab
    If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then Set ActiveTab = m_Tabs(m_ActiveTabId)
End Property
Public Property Get TabCount() As Long: TabCount = m_TabCount: End Property
Public Function GetTab(ByVal tabId As Long) As WebView2Tab
    If m_Tabs.Exists(tabId) Then Set GetTab = m_Tabs(tabId)
End Function
Public Property Get CoreInstance() As WebView2Core: Set CoreInstance = m_Core: End Property


' =============================================================================
' 内部: タブ作成
' =============================================================================

Private Sub CreateTabInternal(ByVal tabId As Long, ByVal StartURL As String, ByVal IsUIBar As Boolean)
    Dim frm As MSForms.Control
    If IsUIBar Then
        Set frm = m_UIBarFrame
    Else
        Set frm = m_ContentFrame
    End If

    Dim hwnd As LongPtr
    hwnd = frm.[_GethWnd]  'CallByName(frm, "_GethWnd", VbGet)

    Dim tab_ As New WebView2Tab
    tab_.Initialize m_Core, Me, tabId, frm, hwnd, StartURL, IsUIBar
    tab_.RequestCreateController m_pEnv

    If IsUIBar Then
        Set m_UIBar = tab_
    Else
        Set m_Tabs(tabId) = tab_
        m_TabCount = m_TabCount + 1
        ReDim Preserve m_TabOrder(0 To m_TabCount - 1)
        m_TabOrder(m_TabCount - 1) = tabId
    End If
End Sub

Private Sub CreateUIBar()
    CreateTabInternal 0, "", True
End Sub
'Private Sub CreateUIBar()
'    ' Controller作成なしで Tab オブジェクトだけ作る
'    Dim frm As MSForms.Control
'    Set frm = m_UIBarFrame
'    Dim hwnd As LongPtr
'    hwnd = frm.[_GethWnd]
'
'    Dim tab_ As New WebView2Tab
'    tab_.Initialize m_Core, Me, 0, frm, hwnd, "", True
'    Set m_UIBar = tab_
'
'    Debug.Print "UIBar: initialized (no controller yet)"
'    ' RequestCreateController は呼ばない
'End Sub


Private Sub RemoveFromOrder(ByVal tabId As Long)
    Dim i As Long, j As Long
    For i = 0 To m_TabCount - 1
        If m_TabOrder(i) = tabId Then
            For j = i To m_TabCount - 2: m_TabOrder(j) = m_TabOrder(j + 1): Next j
            m_TabCount = m_TabCount - 1
            If m_TabCount = 0 Then Erase m_TabOrder _
            Else ReDim Preserve m_TabOrder(0 To m_TabCount - 1)
            Exit For
        End If
    Next i
End Sub

Public Function GetControllerCallback() As LongPtr
    GetControllerCallback = m_Core.EnsureCallback( _
        MGR_PREFIX, ObjPtr(Me), _
        "__ctrl_completed", "Invoke_TabControllerCompleted", m_funcPtrs)
End Function


' =============================================================================
' UIバー更新
' =============================================================================

Private Sub UpdateUIBar()
    If m_UIBar Is Nothing Then Exit Sub
    If m_UIBar.WebViewPtr = 0 Then Exit Sub

    Dim msg As String
    msg = "{""type"":""updateState""," & _
          """activeTabId"":" & m_ActiveTabId & "," & _
          """tabs"":" & BuildTabListJson()

    If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then
        Dim at As WebView2Tab: Set at = m_Tabs(m_ActiveTabId)
        msg = msg & ",""url"":""" & EscJ(at.Source) & """"
        msg = msg & ",""canGoBack"":" & LCase(CStr(at.CanGoBack))
        msg = msg & ",""canGoForward"":" & LCase(CStr(at.CanGoForward))
    End If
    msg = msg & "}"

    m_UIBar.PostWebMessageAsJson msg
End Sub

Private Function BuildTabListJson() As String
    Dim s As String: s = "["
    Dim i As Long
    For i = 0 To m_TabCount - 1
        Dim tid As Long: tid = m_TabOrder(i)
        If m_Tabs.Exists(tid) Then
            Dim t As WebView2Tab: Set t = m_Tabs(tid)
            If i > 0 Then s = s & ","
            s = s & "{""id"":" & tid & ",""title"":""" & EscJ(t.TabTitle) & """" & _
                ",""active"":" & LCase(CStr(tid = m_ActiveTabId)) & "}"
        End If
    Next i
    BuildTabListJson = s & "]"
End Function

Private Function EscJ(ByVal s As String) As String
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r"): s = Replace(s, vbLf, "\n")
    EscJ = s
End Function


' =============================================================================
' UIバーコマンド処理
' =============================================================================

Public Sub OnUIBarMessage(ByVal json As String)
    Dim cmd As String: cmd = ExtJ(json, "command")
    Dim url As String: url = ExtJ(json, "url")
    Dim tid As String: tid = ExtJ(json, "tabId")

    Select Case cmd
        Case "navigate"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then
                m_Tabs(m_ActiveTabId).Navigate url
            End If
        Case "back"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then m_Tabs(m_ActiveTabId).GoBack
        Case "forward"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then m_Tabs(m_ActiveTabId).GoForward
        Case "reload"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then m_Tabs(m_ActiveTabId).Reload
        Case "stop"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then m_Tabs(m_ActiveTabId).StopNavigation
        Case "newtab"
            Dim nid As Long: nid = CreateTab(IIf(url <> "", url, "about:blank"))
            SwitchTab nid
        Case "closetab"
            If tid <> "" Then
                CloseTab CLng(tid)
            ElseIf m_ActiveTabId > 0 Then
                CloseTab m_ActiveTabId
            End If
        Case "switchtab"
            If tid <> "" Then SwitchTab CLng(tid)
        Case "devtools"
            If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then _
                m_Tabs(m_ActiveTabId).OpenDevToolsWindow
    End Select
End Sub

Private Function ExtJ(ByVal json As String, ByVal Key As String) As String
    Dim sr As String: sr = """" & Key & """:"
    Dim p As Long: p = InStr(1, json, sr, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(sr)
    Do While p <= Len(json) And Mid$(json, p, 1) = " ": p = p + 1: Loop
    If p > Len(json) Then Exit Function
    Dim ch As String: ch = Mid$(json, p, 1)
    If ch = """" Then
        Dim ep As Long: ep = InStr(p + 1, json, """")
        If ep = 0 Then Exit Function
        ExtJ = Mid$(json, p + 1, ep - p - 1)
    Else
        Dim ep2 As Long: ep2 = p
        Do While ep2 <= Len(json)
            ch = Mid$(json, ep2, 1)
            If ch = "," Or ch = "}" Or ch = " " Then Exit Do
            ep2 = ep2 + 1
        Loop
        ExtJ = Mid$(json, p, ep2 - p)
    End If
End Function


' =============================================================================
' Tab からのイベント受信
' =============================================================================

Public Sub OnTabEvent( _
    ByVal tabId As Long, ByVal evType As String, _
    ByVal sender As LongPtr, ByVal args As LongPtr, _
    Optional ByVal extra As String = "")

    Select Case evType
        Case "ControllerCompleted"
            RaiseEvent TabCreated(tabId)
            If m_ActiveTabId <= 0 And tabId > 0 Then SwitchTab tabId
            UpdateUIBar

        Case "NavigationStarting"
            RaiseEvent TabNavigationStarting(tabId, sender, args)
            UpdateUIBar

        Case "ContentLoading"
            RaiseEvent TabContentLoading(tabId, sender, args)

        Case "SourceChanged"
            RaiseEvent TabSourceChanged(tabId, sender, args)
            UpdateUIBar

        Case "HistoryChanged"
            RaiseEvent TabHistoryChanged(tabId)
            UpdateUIBar

        Case "NavigationCompleted"
            RaiseEvent TabNavigationCompleted(tabId, sender, args)
            UpdateUIBar

        Case "DocumentTitleChanged"
            Dim title As String
            If m_Tabs.Exists(tabId) Then title = m_Tabs(tabId).TabTitle
            RaiseEvent TabDocumentTitleChanged(tabId, title)
            UpdateUIBar

        Case "WebMessageReceived"
            RaiseEvent TabWebMessageReceived(tabId, extra)

        Case "NewWindowRequested"
            Dim nid2 As Long: nid2 = CreateTab(extra)
            SwitchTab nid2
            RaiseEvent TabNewWindowRequested(tabId, extra)

        Case "ProcessFailed"
            RaiseEvent TabProcessFailed(tabId, sender, args)

        Case "WindowCloseRequested"
            RaiseEvent TabWindowCloseRequested(tabId)
            CloseTab tabId

        Case "ScriptResult"
            RaiseEvent TabScriptResult(tabId, extra)
    End Select
End Sub


' =============================================================================
' リサイズ
' =============================================================================

Public Sub ResizeAll()
    If Not m_UIBar Is Nothing Then m_UIBar.ResizeWebView
    If m_ActiveTabId > 0 And m_Tabs.Exists(m_ActiveTabId) Then
        m_Tabs(m_ActiveTabId).ResizeWebView
    End If
End Sub


' =============================================================================
' クリーンアップ
' =============================================================================

Public Sub Finalize()
    Debug.Print "Manager.Finalize called, m_Initialized=" & m_Initialized
    If Not m_Initialized Then Exit Sub
    

    ' 全タブ
    Debug.Print "Mgr: releasing tabs..."
    If m_Tabs.Count > 0 Then
        Dim keys As Variant: keys = m_Tabs.keys
        Dim i As Long
        For i = UBound(keys) To LBound(keys) Step -1
            Dim t As WebView2Tab: Set t = m_Tabs(keys(i))
            t.Finalize
        Next i
        m_Tabs.RemoveAll
    End If
    Debug.Print "Mgr: tabs done"

    ' UIバー
    Debug.Print "Mgr: releasing UIBar..."
    If Not m_UIBar Is Nothing Then
        m_UIBar.Finalize
        Debug.Print "Mgr: UIBar.Finalize done"
        Set m_UIBar = Nothing
        Debug.Print "Mgr: UIBar set Nothing done"
    End If
    Debug.Print "Mgr: UIBar done"
    
    ' Manager ハンドラ
    Debug.Print "Mgr: unregister prefix..."
    m_Core.UnregisterByPrefix MGR_PREFIX
    Debug.Print "Mgr: prefix done"

    ' Environment
    Debug.Print "Mgr: releasing env..."
    If m_pEnv <> 0 Then
'        'Debug.Print "Mgr: env ptr = &H" & Hex(m_pEnv) & " (skipping Release for test)"
'        'm_Core.dcf m_pEnv, 2, "Env.Release"
'        Dim rc As Long
'        rc = m_Core.ReleaseCOM(m_pEnv)
'        Debug.Print "Env.Release refCount=" & rc
'        m_pEnv = 0
    
'        Dim rc1 As Long
'        rc1 = m_Core.AddRefCOM(m_pEnv)
'        Debug.Print "Env check: AddRef returned " & rc1
'
'        Dim rc2 As Long
'        rc2 = m_Core.ReleaseCOM(m_pEnv)
'        Debug.Print "Env: Release(1) returned " & rc2
'
'        Dim rc3 As Long
'        rc3 = m_Core.ReleaseCOM(m_pEnv)
'        Debug.Print "Env: Release(2) returned " & rc3
'        m_pEnv = 0
    
'        Dim rc As Long
'        rc = m_Core.ReleaseCOM(m_pEnv)
'        Debug.Print "Env.Release refCount=" & rc
'        m_pEnv = 0
    
    ' Environment は Release しない
    ' (ブラウザプロセスの終了と連動するため、明示的 Release はクラッシュの原因になる)
    If m_pEnv <> 0 Then
        Debug.Print "Mgr: env ptr cleared (no Release)"
        m_pEnv = 0
    End If
    
    End If
    
    Debug.Print "Mgr: env done"
    
    m_TabCount = 0: Erase m_TabOrder
    Set m_funcPtrs = Nothing
    Set m_Core = Nothing
    m_Initialized = False
    Debug.Print "Mgr: Finalized."
End Sub


