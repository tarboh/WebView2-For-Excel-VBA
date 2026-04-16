VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WebView2Tab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' =============================================================================
' WebView2Tab.cls  (Core集約版)
' =============================================================================
' 単一 WebView2 Controller + ICoreWebView2。
' ハンドラ管理は全て Core に委譲。自身は Invoke コールバックと
' Manager.OnTabEvent() 通知のみ担当。
' =============================================================================

#If Win64 Then
#Else
    #Error "64 ビット VBA 専用"
#End If

Private Declare PtrSafe Sub CoTaskMemFree Lib "ole32" (ByVal pv As LongPtr)

Private Const E_PENDING As Long = &H8000000A

Private Declare PtrSafe Function DispCallFunc Lib "oleaut32" ( _
    ByVal pvInstance As LongPtr, ByVal oVft As LongPtr, _
    ByVal cc As Long, ByVal vtReturn As Integer, _
    ByVal cActuals As Long, ByRef prgvt As Any, _
    ByRef prgpvarg As Any, ByRef pvargResult As Any) As Long
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    ByRef dest As Any, ByRef src As Any, ByVal Length As LongPtr)

' --- スクリプト管理 ---
Private Type ScriptEntry
    scriptId As String
    javascript As String
End Type

Private Type CDPRequestEntry
    requestId As Long
    methodName As String
    vtblBlock As LongPtr
    fakeObj As LongPtr
End Type

' =============================================================================
' メンバ
' =============================================================================

Private m_Core As WebView2Core
Private m_Manager As WebView2Manager
Private m_funcPtrs As Object       ' 自身のメソッドアドレス Dictionary

Private m_TabId As Long
Private m_TabTitle As String
Private m_IsUIBar As Boolean
Private m_Prefix As String         ' "T1." or "T0." (UIBar)

Private m_pController As LongPtr
Private m_pWebView As LongPtr
Private m_pSettings As LongPtr

Private m_targetFrame As Object
Private m_targetHWnd As LongPtr
Private m_StartURL As String

' --- スクリプト ---
Private m_Scripts() As ScriptEntry
Private m_ScriptCount As Long
Private m_PendingScript As String
Private m_IsScriptPending As Boolean
Private m_IsExecuteScriptPending As Boolean
Private m_IsScriptRunning As Boolean
Public LastScriptResult As String

' --- ナビゲーション ---
Private m_IsNavigating As Boolean

' --- CDP ---
Private m_CDPRequests() As CDPRequestEntry
Private m_CDPRequestCount As Long
Private m_CDPNextId As Long

' --- CapturePreview ---
Private m_pCaptureStream As LongPtr

Private m_Initialized As Boolean


' =============================================================================
' 初期化
' =============================================================================

Public Sub Initialize( _
    ByVal Core As WebView2Core, _
    ByVal Manager As WebView2Manager, _
    ByVal tabId As Long, _
    ByVal TargetFrame As Object, _
    ByVal targetHWnd As LongPtr, _
    Optional ByVal StartURL As String = "", _
    Optional ByVal IsUIBar As Boolean = False)

    Set m_Core = Core
    Set m_Manager = Manager
    m_TabId = tabId
    Set m_targetFrame = TargetFrame
    m_targetHWnd = targetHWnd
    m_StartURL = StartURL
    m_IsUIBar = IsUIBar
    m_Prefix = "T" & tabId & "."

    Set m_funcPtrs = CreateObject("Scripting.Dictionary")
    m_funcPtrs.CompareMode = 1
    m_Core.BuildFuncPtrCache ObjPtr(Me), m_funcPtrs


    Debug.Print "Tab[" & m_TabId & "] funcPtrs count=" & m_funcPtrs.Count
    Debug.Print "Tab[" & m_TabId & "] has Invoke_ControllerCompleted=" & _
        m_funcPtrs.Exists("Invoke_ControllerCompleted")
    If m_funcPtrs.Exists("Invoke_ControllerCompleted") Then
        Dim pp As LongPtr: pp = m_funcPtrs("Invoke_ControllerCompleted")
        Debug.Print "Tab[" & m_TabId & "] pFunc=&H" & Hex(pp)
    End If

    m_Initialized = True
End Sub

Public Property Get tabId() As Long: tabId = m_TabId: End Property
Public Property Get TabTitle() As String: TabTitle = m_TabTitle: End Property
Public Property Get IsUIBar() As Boolean: IsUIBar = m_IsUIBar: End Property
Public Property Get WebViewPtr() As LongPtr: WebViewPtr = m_pWebView: End Property
Public Property Get ControllerPtr() As LongPtr: ControllerPtr = m_pController: End Property


' =============================================================================
' Controller 作成要求
' =============================================================================

Public Sub RequestCreateController(ByVal pEnv As LongPtr)
    Debug.Print "Tab[" & m_TabId & "].RequestCreateController"
    
    Dim fo As LongPtr
    fo = m_Core.EnsureCallback(m_Prefix, ObjPtr(Me), _
        "__ctrl_completed", "Invoke_ControllerCompleted", m_funcPtrs)
    Debug.Print "Tab[" & m_TabId & "] fakeObj=&H" & Hex(fo)
    If fo = 0 Then Exit Sub

    ' サンクのダンプ
    Dim fakeObjAddr As LongPtr
    fakeObjAddr = fo
    
    ' fakeObj → vtable pointer
    Dim pVtbl As LongPtr
    CopyMemory pVtbl, ByVal fakeObjAddr, 8
    Debug.Print "fakeObj=&H" & Hex(fakeObjAddr) & " vtbl=&H" & Hex(pVtbl)
    
    ' vtable slots
    Dim slot(3) As LongPtr
    CopyMemory slot(0), ByVal pVtbl, 32
    Debug.Print "  [0]QI=&H" & Hex(slot(0))
    Debug.Print "  [1]AR=&H" & Hex(slot(1))
    Debug.Print "  [2]AR=&H" & Hex(slot(2))
    Debug.Print "  [3]Invoke=&H" & Hex(slot(3))
    
    ' thunk bytes (先頭32バイト)
    Dim thkBytes(31) As Byte
    CopyMemory thkBytes(0), ByVal slot(3), 32
    Dim s As String, j As Long
    For j = 0 To 31
        s = s & Right$("0" & Hex(thkBytes(j)), 2) & " "
    Next j
    Debug.Print "  thunk: " & s


    Dim hr As Long
    hr = m_Core.dcf(pEnv, 3, "CreateCoreWebView2Controller", m_targetHWnd, fo)
    Debug.Print "Tab[" & m_TabId & "] CreateController hr=&H" & Hex(hr)
End Sub

Public Function Invoke_ControllerCompleted( _
    ByVal pThis As LongPtr, ByVal errorCode As LongPtr, ByVal pCtrl As LongPtr) As Long

    Debug.Print "Tab[" & m_TabId & "].ControllerCompleted entered, err=&H" & Hex(errorCode) & " pCtrl=&H" & Hex(pCtrl)

    If errorCode <> 0 Or pCtrl = 0 Then
        Debug.Print "Tab[" & m_TabId & "] ctrl error &H" & Hex(errorCode)
        Exit Function
    End If

    m_pController = pCtrl
    m_Core.dcf m_pController, 1, "Controller.AddRef"
    m_Core.dcf m_pController, 4, "put_IsVisible", 1&

    ' サイズ
    Dim l(3) As Long
    l(0) = 0: l(1) = 0
    l(2) = m_Core.PtsToPx(CSng(CallByName(m_targetFrame, "InsideWidth", VbGet)), False)
    l(3) = m_Core.PtsToPx(CSng(CallByName(m_targetFrame, "InsideHeight", VbGet)), True)
    m_Core.dcf m_pController, 6, "put_Bounds", VarPtr(l(0))

    ' WebView2
    Dim pWV As LongPtr
    m_Core.dcf m_pController, 25, "get_CoreWebView2", VarPtr(pWV)
    m_pWebView = pWV
    If m_pWebView = 0 Then Exit Function

    ' Settings
    m_Core.dcf m_pWebView, 3, "get_Settings", VarPtr(m_pSettings)

    ' イベント登録
    RegisterDefaultHandlers

    ' 初期ナビゲーション
    If m_StartURL <> "" Then m_Core.dcf m_pWebView, 5, "Navigate", StrPtr(m_StartURL)

    m_Manager.OnTabEvent m_TabId, "ControllerCompleted", 0, 0
    DoEvents
    Invoke_ControllerCompleted = 0
End Function


' =============================================================================
' デフォルトハンドラ登録
' =============================================================================

Private Sub RegisterDefaultHandlers()
    Dim p As String: p = m_Prefix
    Dim me_ As LongPtr: me_ = ObjPtr(Me)
    Dim fp As Object: Set fp = m_funcPtrs
    Dim wv As LongPtr: wv = m_pWebView

    m_Core.RegisterEvent p, me_, wv, 7, 8, "Invoke_NavigationStarting", "WV.NavStart", fp
    m_Core.RegisterEvent p, me_, wv, 9, 10, "Invoke_ContentLoading", "WV.ContentLoad", fp
    m_Core.RegisterEvent p, me_, wv, 11, 12, "Invoke_SourceChanged", "WV.SrcChanged", fp
    m_Core.RegisterEvent p, me_, wv, 13, 14, "Invoke_HistoryChanged", "WV.History", fp
    m_Core.RegisterEvent p, me_, wv, 15, 16, "Invoke_NavigationCompleted", "WV.NavComplete", fp
    m_Core.RegisterEvent p, me_, wv, 17, 18, "Invoke_FrameNavStarting", "WV.FrmNavStart", fp
    m_Core.RegisterEvent p, me_, wv, 19, 20, "Invoke_FrameNavCompleted", "WV.FrmNavComp", fp
    m_Core.RegisterEvent p, me_, wv, 21, 22, "Invoke_ScriptDialogOpening", "WV.ScriptDlg", fp
    m_Core.RegisterEvent p, me_, wv, 23, 24, "Invoke_PermissionRequested", "WV.Permission", fp
    m_Core.RegisterEvent p, me_, wv, 25, 26, "Invoke_ProcessFailed", "WV.ProcFail", fp
    m_Core.RegisterEvent p, me_, wv, 34, 35, "Invoke_WebMessageReceived", "WV.WebMsg", fp
    m_Core.RegisterEvent p, me_, wv, 44, 45, "Invoke_NewWindowRequested", "WV.NewWin", fp
    m_Core.RegisterEvent p, me_, wv, 46, 47, "Invoke_DocTitleChanged", "WV.DocTitle", fp
    m_Core.RegisterEvent p, me_, wv, 52, 53, "Invoke_FullScreenChanged", "WV.FullScr", fp
    m_Core.RegisterEvent p, me_, wv, 55, 56, "Invoke_WebResourceRequested", "WV.WebRes", fp
    m_Core.RegisterEvent p, me_, wv, 59, 60, "Invoke_WindowCloseRequested", "WV.WinClose", fp

    ' 全リソースフィルタ
    m_Core.dcf m_pWebView, 57, "AddWebResourceRequestedFilter", StrPtr("*"), CLng(0)
End Sub


' =============================================================================
' プロパティ・メソッド
' =============================================================================

Public Property Get Source() As String
    Dim pStr As LongPtr
    m_Core.dcf m_pWebView, 4, "get_Source", VarPtr(pStr)
    If pStr <> 0 Then Source = m_Core.PtrToStrW(pStr): CoTaskMemFree pStr
End Property

Public Property Get CanGoBack() As Boolean
    Dim v As Long: m_Core.dcf m_pWebView, 38, "get_CanGoBack", VarPtr(v)
    CanGoBack = (v <> 0)
End Property

Public Property Get CanGoForward() As Boolean
    Dim v As Long: m_Core.dcf m_pWebView, 39, "get_CanGoForward", VarPtr(v)
    CanGoForward = (v <> 0)
End Property

Public Property Get DocumentTitle() As String
    Dim pStr As LongPtr
    m_Core.dcf m_pWebView, 48, "get_DocumentTitle", VarPtr(pStr)
    If pStr <> 0 Then DocumentTitle = m_Core.PtrToStrW(pStr): CoTaskMemFree pStr
End Property

Public Function Navigate(ByVal uri As String) As Long
    Navigate = m_Core.dcf(m_pWebView, 5, "Navigate", StrPtr(uri))
End Function

Public Function NavigateToString(ByVal html As String) As Long
    NavigateToString = m_Core.dcf(m_pWebView, 6, "NavigateToString", StrPtr(html))
End Function

Public Function Reload() As Long
    Reload = m_Core.dcf(m_pWebView, 31, "Reload")
End Function

Public Function GoBack() As Long
    GoBack = m_Core.dcf(m_pWebView, 40, "GoBack")
End Function

Public Function GoForward() As Long
    GoForward = m_Core.dcf(m_pWebView, 41, "GoForward")
End Function

Public Function StopNavigation() As Long
    StopNavigation = m_Core.dcf(m_pWebView, 43, "Stop")
End Function

Public Function PostWebMessageAsJson(ByVal json As String) As Long
    PostWebMessageAsJson = m_Core.dcf(m_pWebView, 32, "PostWebMessageAsJson", StrPtr(json))
End Function

Public Function PostWebMessageAsString(ByVal msg As String) As Long
    PostWebMessageAsString = m_Core.dcf(m_pWebView, 33, "PostWebMessageAsString", StrPtr(msg))
End Function

Public Sub OpenDevToolsWindow()
    m_Core.dcf m_pWebView, 51, "OpenDevToolsWindow"
End Sub

Public Sub SetVisible(ByVal v As Boolean)
    If m_pController = 0 Then Exit Sub
    m_Core.dcf m_pController, 4, "put_IsVisible", IIf(v, 1&, 0&)
End Sub

Public Sub ResizeWebView()
    If m_pController = 0 Or m_targetFrame Is Nothing Then Exit Sub
    Dim l(3) As Long
    l(0) = 0: l(1) = 0
    l(2) = m_Core.PtsToPx(CSng(CallByName(m_targetFrame, "InsideWidth", VbGet)), False)
    l(3) = m_Core.PtsToPx(CSng(CallByName(m_targetFrame, "InsideHeight", VbGet)), True)
    m_Core.dcf m_pController, 6, "put_Bounds", VarPtr(l(0))
End Sub


' =============================================================================
' Settings
' =============================================================================

Public Sub SetScriptEnabled(ByVal v As Boolean)
    m_Core.dcf m_pSettings, 4, "put_IsScriptEnabled", IIf(v, 1&, 0&)
End Sub
Public Sub SetWebMessageEnabled(ByVal v As Boolean)
    m_Core.dcf m_pSettings, 6, "put_IsWebMessageEnabled", IIf(v, 1&, 0&)
End Sub
Public Sub SetDevToolsEnabled(ByVal v As Boolean)
    m_Core.dcf m_pSettings, 12, "put_AreDevToolsEnabled", IIf(v, 1&, 0&)
End Sub
Public Sub SetContextMenuEnabled(ByVal v As Boolean)
    m_Core.dcf m_pSettings, 14, "put_AreDefaultContextMenusEnabled", IIf(v, 1&, 0&)
End Sub
Public Sub SetStatusBarEnabled(ByVal v As Boolean)
    m_Core.dcf m_pSettings, 10, "put_IsStatusBarEnabled", IIf(v, 1&, 0&)
End Sub


' =============================================================================
' ExecuteScript
' =============================================================================

Public Function ExecuteScriptAsync(ByVal script As String) As Long
    If m_IsExecuteScriptPending Then ExecuteScriptAsync = E_PENDING: Exit Function
    m_IsExecuteScriptPending = True
    Dim fo As LongPtr
    fo = m_Core.EnsureCallback(m_Prefix, ObjPtr(Me), _
        "__exec_script", "Invoke_ExecuteScript", m_funcPtrs)
    If fo = 0 Then m_IsExecuteScriptPending = False: Exit Function
    ExecuteScriptAsync = m_Core.dcf(m_pWebView, 29, "ExecuteScript", StrPtr(script), fo)
End Function

Public Function Invoke_ExecuteScript( _
    ByVal pThis As LongPtr, ByVal errCode As Long, ByVal resPtr As LongPtr) As Long
    If resPtr <> 0 Then LastScriptResult = m_Core.PtrToStrW(resPtr)
    m_IsExecuteScriptPending = False: m_IsScriptRunning = False
    m_Manager.OnTabEvent m_TabId, "ScriptResult", 0, 0, LastScriptResult
    Invoke_ExecuteScript = 0
End Function

Public Function ExecuteScriptSync( _
    ByVal script As String, Optional ByVal tmo As Double = 5) As String
    m_IsScriptRunning = True
    ExecuteScriptAsync script
    Dim st As Double: st = Timer
    Do While m_IsScriptRunning
        DoEvents
        If Timer - st > tmo Then Exit Do
    Loop
    ExecuteScriptSync = LastScriptResult
End Function


' =============================================================================
' AddScriptToExecuteOnDocumentCreated
' =============================================================================

Public Function AddScriptOnDocCreated(ByVal js As String) As Long
    If m_IsScriptPending Then AddScriptOnDocCreated = E_PENDING: Exit Function
    m_IsScriptPending = True: m_PendingScript = js
    Dim fo As LongPtr
    fo = m_Core.EnsureCallback(m_Prefix, ObjPtr(Me), _
        "__add_script", "Invoke_AddScriptCompleted", m_funcPtrs)
    If fo = 0 Then m_IsScriptPending = False: Exit Function
    AddScriptOnDocCreated = m_Core.dcf(m_pWebView, 27, _
        "AddScriptToExecuteOnDocumentCreated", StrPtr(js), fo)
End Function

Public Function Invoke_AddScriptCompleted( _
    ByVal pThis As LongPtr, ByVal errCode As Long, ByVal idPtr As LongPtr) As Long
    If errCode = 0 Then
        Dim sid As String: sid = m_Core.PtrToStrW(idPtr)
        m_ScriptCount = m_ScriptCount + 1
        ReDim Preserve m_Scripts(0 To m_ScriptCount - 1)
        m_Scripts(m_ScriptCount - 1).scriptId = sid
        m_Scripts(m_ScriptCount - 1).javascript = m_PendingScript
    End If
    m_PendingScript = "": m_IsScriptPending = False
    Invoke_AddScriptCompleted = 0
End Function


' =============================================================================
' Invoke コールバック群
' =============================================================================

Public Function Invoke_NavigationStarting( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_IsNavigating = True
    m_Manager.OnTabEvent m_TabId, "NavigationStarting", sender, args
    Invoke_NavigationStarting = 0
End Function

Public Function Invoke_ContentLoading( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "ContentLoading", sender, args
    Invoke_ContentLoading = 0
End Function

Public Function Invoke_SourceChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "SourceChanged", sender, args
    Invoke_SourceChanged = 0
End Function

Public Function Invoke_HistoryChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "HistoryChanged", sender, args
    Invoke_HistoryChanged = 0
End Function

Public Function Invoke_NavigationCompleted( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_IsNavigating = False
    m_TabTitle = Me.DocumentTitle
    m_Manager.OnTabEvent m_TabId, "NavigationCompleted", sender, args
    Invoke_NavigationCompleted = 0
End Function

Public Function Invoke_FrameNavStarting( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "FrameNavigationStarting", sender, args
    Invoke_FrameNavStarting = 0
End Function

Public Function Invoke_FrameNavCompleted( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "FrameNavigationCompleted", sender, args
    Invoke_FrameNavCompleted = 0
End Function

Public Function Invoke_ScriptDialogOpening( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "ScriptDialogOpening", sender, args
    Invoke_ScriptDialogOpening = 0
End Function

Public Function Invoke_PermissionRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "PermissionRequested", sender, args
    Invoke_PermissionRequested = 0
End Function

Public Function Invoke_ProcessFailed( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "ProcessFailed", sender, args
    Invoke_ProcessFailed = 0
End Function

Public Function Invoke_WebMessageReceived( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long

    Dim pStr As LongPtr, sJson As String

    m_Core.dcf args, 4, "get_WebMessageAsJson", VarPtr(pStr)
    If pStr <> 0 Then sJson = m_Core.PtrToStrW(pStr): CoTaskMemFree pStr: pStr = 0

    If m_IsUIBar Then
        m_Manager.OnUIBarMessage sJson
    Else
        m_Manager.OnTabEvent m_TabId, "WebMessageReceived", sender, args, sJson
    End If
    Invoke_WebMessageReceived = 0
End Function

Public Function Invoke_NewWindowRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    Dim pStr As LongPtr, uri As String
    m_Core.dcf args, 3, "get_Uri", VarPtr(pStr)
    If pStr <> 0 Then uri = m_Core.PtrToStrW(pStr): CoTaskMemFree pStr
    m_Core.dcf args, 6, "put_Handled", 1&
    m_Manager.OnTabEvent m_TabId, "NewWindowRequested", sender, args, uri
    Invoke_NewWindowRequested = 0
End Function

Public Function Invoke_DocTitleChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_TabTitle = Me.DocumentTitle
    m_Manager.OnTabEvent m_TabId, "DocumentTitleChanged", sender, args
    Invoke_DocTitleChanged = 0
End Function

Public Function Invoke_FullScreenChanged( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "FullScreenChanged", sender, args
    Invoke_FullScreenChanged = 0
End Function

Public Function Invoke_WebResourceRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "WebResourceRequested", sender, args
    Invoke_WebResourceRequested = 0
End Function

Public Function Invoke_WindowCloseRequested( _
    ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    m_Manager.OnTabEvent m_TabId, "WindowCloseRequested", sender, args
    Invoke_WindowCloseRequested = 0
End Function


' =============================================================================
' CDP
' =============================================================================

Public Function CallDevToolsProtocolMethod( _
    ByVal methodName As String, ByVal paramJson As String) As Long

    m_CDPNextId = m_CDPNextId + 1
    Dim rid As Long: rid = m_CDPNextId

    Dim vtbl As LongPtr, fo As LongPtr
    If Not m_Core.CreateOneShot(ObjPtr(Me), "Invoke_CDPCompleted", m_funcPtrs, vtbl, fo) Then
        CallDevToolsProtocolMethod = &H80004005: Exit Function
    End If

    m_CDPRequestCount = m_CDPRequestCount + 1
    ReDim Preserve m_CDPRequests(0 To m_CDPRequestCount - 1)
    With m_CDPRequests(m_CDPRequestCount - 1)
        .requestId = rid: .methodName = methodName: .vtblBlock = vtbl: .fakeObj = fo
    End With

    Dim hr As Long
    hr = m_Core.dcf(m_pWebView, 36, "CallDevToolsProtocolMethod", _
        StrPtr(methodName), StrPtr(paramJson), fo)

    If hr = 0 Then
        CallDevToolsProtocolMethod = rid
    Else
        m_Core.FreeOneShot vtbl, fo
        m_CDPRequestCount = m_CDPRequestCount - 1
        CallDevToolsProtocolMethod = hr
    End If
End Function

Public Function Invoke_CDPCompleted( _
    ByVal pThis As LongPtr, ByVal errCode As Long, ByVal resPtr As LongPtr) As Long

    Dim i As Long, j As Long, mn As String, rid As Long, found As Boolean
    For i = 0 To m_CDPRequestCount - 1
        If m_CDPRequests(i).fakeObj = pThis Then
            mn = m_CDPRequests(i).methodName: rid = m_CDPRequests(i).requestId
            m_Core.FreeOneShot m_CDPRequests(i).vtblBlock, m_CDPRequests(i).fakeObj
            For j = i To m_CDPRequestCount - 2
                m_CDPRequests(j) = m_CDPRequests(j + 1)
            Next j
            m_CDPRequestCount = m_CDPRequestCount - 1
            found = True: Exit For
        End If
    Next i

    Dim sr As String
    If resPtr <> 0 Then sr = m_Core.PtrToStrW(resPtr)
    If found Then m_Manager.OnTabEvent m_TabId, "CDPCompleted", 0, 0, sr

    Invoke_CDPCompleted = 0
End Function


' =============================================================================
' クリーンアップ
' =============================================================================

Public Sub Finalize()
    
    Debug.Print "Tab[" & m_TabId & "].Finalize called"
    
    If Not m_Initialized Then Exit Sub

    ' Core からプレフィックス一括解除
    m_Core.UnregisterByPrefix m_Prefix

    ' CDP 残存
    Dim i As Long
    For i = 0 To m_CDPRequestCount - 1
        m_Core.FreeOneShot m_CDPRequests(i).vtblBlock, m_CDPRequests(i).fakeObj
    Next i
    m_CDPRequestCount = 0

    ' COM Release
    If m_pSettings <> 0 Then m_Core.dcf m_pSettings, 2, "Settings.Release": m_pSettings = 0
    If m_pWebView <> 0 Then m_Core.dcf m_pWebView, 2, "WebView.Release": m_pWebView = 0
    If m_pController <> 0 Then
        m_Core.dcf m_pController, 24, "Controller.Close"
        m_Core.dcf m_pController, 2, "Controller.Release"
        m_pController = 0
    End If
    If m_pCaptureStream <> 0 Then
        m_Core.dcf m_pCaptureStream, 2, "Stream.Release": m_pCaptureStream = 0
    End If

    Set m_funcPtrs = Nothing
    Set m_targetFrame = Nothing
    'Set m_Manager = Nothing
    'Set m_Core = Nothing
    m_Initialized = False
    Debug.Print "Tab[" & m_TabId & "].Finalize completed"
End Sub

Private Sub Class_Terminate()
    ' Finalize は Manager から明示的に呼ばれるので、
    ' Class_Terminate では何もしない。
    ' (VBA の破棄順序が不定なため、ここから Core 等にアクセスすると危険)
    
    Debug.Print "Tab[" & m_TabId & "].Class_Terminate called"
    'Finalize
End Sub


