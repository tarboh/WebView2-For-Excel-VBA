VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c3_WebView2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c3_WebView2

Option Explicit

Public ppWebView2 As LongPtr
Public Event NavigationStarting()
Public Event ContentLoading()
Public Event SourceChanged()
Public Event HistoryChanged()
Public Event NavigationCompleted()
Public Event FrameNavigationStarting()
Public Event FrameNavigationCompleted()
Public Event ScriptDialogOpening()

Public Event ReceiveScriptResult(ByVal result As String)
Public LastScriptResult As String


Public Col_Handler As New Collection
Private m_IsNavigating As Boolean
Private m_IsScriptRunning As Boolean

Public cHandler As c4_Handler2

Public Settings As c6_WebView2Settings

' --- 内部通知用 ---
Public Sub NotifyNavigationStarting()
    RaiseEvent NavigationStarting
End Sub
Public Sub NotifyContentLoading()
    RaiseEvent ContentLoading
End Sub
Public Sub NotifySourceChanged()
    RaiseEvent SourceChanged
End Sub
Public Sub NotifyHistoryChanged()
    RaiseEvent HistoryChanged
End Sub
Public Sub NotifyNavigationCompleted()
    m_IsNavigating = False
    RaiseEvent NavigationCompleted
End Sub
Public Sub NotifyFrameNavigationStarting()
    RaiseEvent FrameNavigationStarting
End Sub
Public Sub NotifyFrameNavigationCompleted()
    RaiseEvent FrameNavigationCompleted
End Sub
Public Sub NotifyScriptDialogOpening()
    RaiseEvent ScriptDialogOpening
End Sub

Public Sub NotifyExecuteScriptCompleted(ByVal jsonResult As String)
    LastScriptResult = jsonResult
    m_IsScriptRunning = False
    RaiseEvent ReceiveScriptResult(jsonResult)
End Sub

' --- 後始末 ---
Public Sub Finalize()
    Dim h As c4_Handler
    On Error Resume Next
    For Each h In Col_Handler
        UnregisterInstance h.Pointer
    Next h
    Set Col_Handler = New Collection
End Sub

Private Sub Class_Terminate()
    Finalize
    Debug.Print "c3_WebView2 Terminated."
End Sub

' --- 待機用ヘルパー ---
Private Sub WaitNavigation(ByVal TimeoutSec As Double)
    Dim Start As Double: Start = Timer
    Do While m_IsNavigating
        DoEvents
        If Timer - Start > TimeoutSec Then Exit Do
    Loop
End Sub


'MIDL_INTERFACE ("76eceacb-0462-4d94-ac83-423a6793775e")
'ICoreWebView2:          Public IUnknown
'{
'public:

'3
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Settings(
'    /* [retval][out] */ ICoreWebView2Settings **settings) = 0;
Public Function get_Settings() As Long
    Dim psettings As LongPtr
    Dim args(0) As Variant: args(0) = VarPtr(psettings)
    Dim args_Ptr(0) As LongPtr: args_Ptr(0) = VarPtr(args(0))
    Dim args_Type(0) As Integer: args_Type(0) = vbLongPtr
    Dim hr As Long, res As Variant
    
    hr = DispCallFunc(ppWebView2, 3 * LenB(ppWebView2), CC_STDCALL, vbLong, 1, args_Type(0), args_Ptr(0), res)
    If hr = 0 Then
        Debug.Print "get_Settings Result:" & res
        get_Settings = res
        If res = 0 Then
            Debug.Print "psettings:" & psettings
            Set Settings = New c6_WebView2Settings
            Settings.Init psettings
        End If
    Else
        Debug.Print "DispCallFunc(get_Settings) Failed. HR:" & hr
    End If
End Function

'4
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_Source(
'    /* [retval][out] */ LPWSTR *uri) = 0;
Public Property Get Source() As String
    '現在のURLを取得する
    Source = DCF_引数無しで文字列を取得(ppWebView2, 4, "get_Source")
End Property

'5
'virtual HRESULT STDMETHODCALLTYPE Navigate(
'    /* [in] */ LPCWSTR uri) = 0;
Public Function NavigateAsync(ByVal uri As String) As Long
    '指定したuriに遷移する
    NavigateAsync = DCF_引数1つで文字列を渡す(ppWebView2, 5, "Navigate", uri)
End Function
Public Sub NavigateSync(ByVal url As String, Optional ByVal TimeoutSec As Double = 10)
    ' 同期版：読み込み完了まで待機する
    m_IsNavigating = True
    If NavigateAsync(url) = 0 Then
        WaitNavigation TimeoutSec
    End If
End Sub

'6
'virtual HRESULT STDMETHODCALLTYPE NavigateToString(
'    /* [in] */ LPCWSTR htmlContent) = 0;
Public Function NavigateToString(htmlContent As String) As Long
    '指定したHTMLコードを描画する
    NavigateToString = DCF_引数1つで文字列を渡す(ppWebView2, 6, "NavigateToString", htmlContent)
End Function

'7
'virtual HRESULT STDMETHODCALLTYPE add_NavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_NavigationStarting() As Long
    add_NavigationStarting = DCF_ハンドラ登録(Me, 7, "add_NavigationStarting", AddressOf NavigationStarting_Invoke, "NavigationStartingHandler")
End Function
Public Function add_NavigationStarting_1(Handler As c4_Handler) As Long
    'イベントハンドラのvTbleポインタを渡してトークンを受け取る形に修正したい
End Function

'8
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationStarting(token As Long) As Long
    remove_NavigationStarting = remove_Handler(Me, 8, token, "remove_NavigationStarting")
End Function

'9
'virtual HRESULT STDMETHODCALLTYPE add_ContentLoading(
'    /* [in] */ ICoreWebView2ContentLoadingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_ContentLoading() As Long
    add_ContentLoading = DCF_ハンドラ登録(Me, 9, "add_ContentLoading", AddressOf ContentLoading_Invoke, "ContentLoadingHandler")
End Function

'10
'virtual HRESULT STDMETHODCALLTYPE remove_ContentLoading(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ContentLoading(token As Long) As Long
    remove_ContentLoading = remove_Handler(Me, 10, token, "remove_ContentLoading")
End Function

'11
'virtual HRESULT STDMETHODCALLTYPE add_SourceChanged(
'    /* [in] */ ICoreWebView2SourceChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_SourceChanged() As Long
    add_SourceChanged = DCF_ハンドラ登録(Me, 11, "add_SourceChanged", AddressOf SourceChanged_Invoke, "SourceChangedHandler")
End Function

'12
'virtual HRESULT STDMETHODCALLTYPE remove_SourceChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_SourceChanged(token As Long) As Long
    remove_SourceChanged = remove_Handler(Me, 12, token, "remove_SourceChanged")
End Function

'13
'virtual HRESULT STDMETHODCALLTYPE add_HistoryChanged(
'    /* [in] */ ICoreWebView2HistoryChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_HistoryChanged() As Long
    add_HistoryChanged = DCF_ハンドラ登録(Me, 13, "add_HistoryChanged", AddressOf HistoryChanged_Invoke, "HistoryChangedHandler")
End Function

'14
'virtual HRESULT STDMETHODCALLTYPE remove_HistoryChanged(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_HistoryChanged(token As Long) As Long
    remove_HistoryChanged = remove_Handler(Me, 14, token, "remove_HistoryChanged")
End Function

'15
'virtual HRESULT STDMETHODCALLTYPE add_NavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'add_NavigationCompleted
Public Function add_NavigationCompleted() As Long
    add_NavigationCompleted = mHelpers.DCF_ハンドラ登録(Me, 15, "add_NavigationCompleted", AddressOf NavCompleted_Invoke, "NavigationCompleted")
End Function

'16
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_NavigationCompleted(token As Long) As Long
    remove_NavigationCompleted = remove_Handler(Me, 16, token, "remove_NavigationCompleted")
End Function

'17
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationStarting() As Long
    add_FrameNavigationStarting = mHelpers.DCF_ハンドラ登録(Me, 17, "add_FrameNavigationStarting", AddressOf FrameNavigationStarting_Invoke, "FrameNavigationStarting")
End Function

'18
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationStarting(token As Long) As Long
    remove_FrameNavigationStarting = remove_Handler(Me, 18, token, "remove_FrameNavigationStarting")
End Function

'19
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
Public Function add_FrameNavigationCompleted() As Long
    add_FrameNavigationCompleted = mHelpers.DCF_ハンドラ登録(Me, 19, "add_FrameNavigationCompleted", AddressOf FrameNavigationCompleted_Invoke, "FrameNavigationCompleted")
End Function

'20
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_FrameNavigationCompleted(token As Long) As Long
    remove_FrameNavigationCompleted = remove_Handler(Me, 20, token, "remove_FrameNavigationCompleted")
End Function

'21
'virtual HRESULT STDMETHODCALLTYPE add_ScriptDialogOpening(
'    /* [in] */ ICoreWebView2ScriptDialogOpeningEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'   ※ このイベントのみ ICoreWebView2Settings:: AreDefaultScriptDialogsEnabledプロパティがFALSEに設定されている場合にトリガーされます｡
Public Function add_ScriptDialogOpening() As Long
    add_ScriptDialogOpening = mHelpers.DCF_ハンドラ登録(Me, 21, "add_ScriptDialogOpening", AddressOf ScriptDialogOpening_Invoke, "ScriptDialogOpening")
End Function

'22
'virtual HRESULT STDMETHODCALLTYPE remove_ScriptDialogOpening(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ScriptDialogOpening(token As Long) As Long
    remove_ScriptDialogOpening = remove_Handler(Me, 22, token, "remove_ScriptDialogOpening")
End Function


'23
'virtual HRESULT STDMETHODCALLTYPE add_PermissionRequested(
'    /* [in] */ ICoreWebView2PermissionRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'24
'virtual HRESULT STDMETHODCALLTYPE remove_PermissionRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_PermissionRequested(token As Long) As Long
    remove_PermissionRequested = remove_Handler(Me, 24, token, "remove_PermissionRequested")
End Function

'25
'virtual HRESULT STDMETHODCALLTYPE add_ProcessFailed(
'    /* [in] */ ICoreWebView2ProcessFailedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'26
'virtual HRESULT STDMETHODCALLTYPE remove_ProcessFailed(
'    /* [in] */ EventRegistrationToken token) = 0;
Public Function remove_ProcessFailed(token As Long) As Long
    remove_ProcessFailed = remove_Handler(Me, 26, token, "remove_ProcessFailed")
End Function

'※ 27,28は他のハンドラとは扱いが違うので注意。
'   イベント発生時にコールバックを受けるのではなく、
'   WebView2内部でJavaScriptを実行させるために、そのスクリプトを文字列として渡す関数なので
'   Add,Removeどちらもシグネチャが異なる。
'27
'virtual HRESULT STDMETHODCALLTYPE AddScriptToExecuteOnDocumentCreated(
'    /* [in] */ LPCWSTR javaScript,
'    /* [in] */ ICoreWebView2AddScriptToExecuteOnDocumentCreatedCompletedHandler *handler) = 0;

'28
'virtual HRESULT STDMETHODCALLTYPE RemoveScriptToExecuteOnDocumentCreated(
'    /* [in] */ LPCWSTR id) = 0;


'29
'virtual HRESULT STDMETHODCALLTYPE ExecuteScript(
'    /* [in] */ LPCWSTR javaScript,
'    /* [in] */ ICoreWebView2ExecuteScriptCompletedHandler *handler) = 0;
Public Function ExecuteScriptAsync(ByVal script As String) As Long
' 非同期版：実行を依頼するだけで戻る。結果はイベントで受け取る
    ExecuteScriptAsync = DFC_ハンドラ登録_文字列渡しあり(Me, 29, "ExecuteScript", script, AddressOf ExecuteScript_Invoke)
End Function

' 同期版：実行結果を戻り値として直接受け取る
Public Function ExecuteScriptSync(ByVal script As String, Optional ByVal TimeoutSec As Double = 5) As String
    m_IsScriptRunning = True
    ExecuteScriptAsync script
    
    ' 専用の待機ループ
    Dim Start As Double: Start = Timer
    Do While m_IsScriptRunning
        Debug.Print "waiting ExecuteScriptSync"
        DoEvents
        If Timer - Start > TimeoutSec Then
            Debug.Print "ExecuteScriptSync: Timeout!"
            Exit Do
        End If
    Loop
    
    ExecuteScriptSync = LastScriptResult
End Function


'30
'virtual HRESULT STDMETHODCALLTYPE CapturePreview(
'    /* [in] */ COREWEBVIEW2_CAPTURE_PREVIEW_IMAGE_FORMAT imageFormat,
'    /* [in] */ IStream *imageStream,
'    /* [in] */ ICoreWebView2CapturePreviewCompletedHandler *handler) = 0;

'31
'virtual HRESULT STDMETHODCALLTYPE Reload( void) = 0;
Public Function Reload() As Long
    Dim hr As Long, res As Variant
    hr = DispCallFunc(ppWebView2, 31 * LenB(ppWebView2), CC_STDCALL, vbLong, 0, 0, 0, res)
    Reload = res
End Function

'32
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsJson(
'    /* [in] */ LPCWSTR webMessageAsJson) = 0;

'33
'virtual HRESULT STDMETHODCALLTYPE PostWebMessageAsString(
'    /* [in] */ LPCWSTR webMessageAsString) = 0;

'34
'virtual HRESULT STDMETHODCALLTYPE add_WebMessageReceived(
'    /* [in] */ ICoreWebView2WebMessageReceivedEventHandler *handler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'35
'virtual HRESULT STDMETHODCALLTYPE remove_WebMessageReceived(
'    /* [in] */ EventRegistrationToken token) = 0;

'36
'virtual HRESULT STDMETHODCALLTYPE CallDevToolsProtocolMethod(
'    /* [in] */ LPCWSTR methodName,
'    /* [in] */ LPCWSTR parametersAsJson,
'    /* [in] */ ICoreWebView2CallDevToolsProtocolMethodCompletedHandler *handler) = 0;

'37
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_BrowserProcessId(
'    /* [retval][out] */ UINT32 *value) = 0;

'38
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoBack(
'    /* [retval][out] */ BOOL *canGoBack) = 0;

'39
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_CanGoForward(
'    /* [retval][out] */ BOOL *canGoForward) = 0;

'40
'virtual HRESULT STDMETHODCALLTYPE GoBack( void) = 0;

'41
'virtual HRESULT STDMETHODCALLTYPE GoForward( void) = 0;

'42
'virtual HRESULT STDMETHODCALLTYPE GetDevToolsProtocolEventReceiver(
'    /* [in] */ LPCWSTR eventName,
'    /* [retval][out] */ ICoreWebView2DevToolsProtocolEventReceiver **receiver) = 0;

'43
'virtual HRESULT STDMETHODCALLTYPE Stop( void) = 0;

'44
'virtual HRESULT STDMETHODCALLTYPE add_NewWindowRequested(
'    /* [in] */ ICoreWebView2NewWindowRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'45
'virtual HRESULT STDMETHODCALLTYPE remove_NewWindowRequested(
'    /* [in] */ EventRegistrationToken token) = 0;

'46
'virtual HRESULT STDMETHODCALLTYPE add_DocumentTitleChanged(
'    /* [in] */ ICoreWebView2DocumentTitleChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'47
'virtual HRESULT STDMETHODCALLTYPE remove_DocumentTitleChanged(
'    /* [in] */ EventRegistrationToken token) = 0;

'48
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_DocumentTitle(
'    /* [retval][out] */ LPWSTR *title) = 0;

'49
'virtual HRESULT STDMETHODCALLTYPE AddHostObjectToScript(
'    /* [in] */ LPCWSTR name,
'    /* [in] */ VARIANT *object) = 0;
Public Function AddHostObjectToScript(ByVal Name As String, ByVal Instance As Object) As Long
    AddHostObjectToScript = DCF_引数2つ_StringとObject(ppWebView2, 49, "AddHostObjectToScript", Name, Instance)
End Function

'50
'virtual HRESULT STDMETHODCALLTYPE RemoveHostObjectFromScript(
'    /* [in] */ LPCWSTR name) = 0;

'51
'virtual HRESULT STDMETHODCALLTYPE OpenDevToolsWindow( void) = 0;

'52
'virtual HRESULT STDMETHODCALLTYPE add_ContainsFullScreenElementChanged(
'    /* [in] */ ICoreWebView2ContainsFullScreenElementChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'53
'virtual HRESULT STDMETHODCALLTYPE remove_ContainsFullScreenElementChanged(
'    /* [in] */ EventRegistrationToken token) = 0;

'54
'virtual /* [propget] */ HRESULT STDMETHODCALLTYPE get_ContainsFullScreenElement(
'    /* [retval][out] */ BOOL *containsFullScreenElement) = 0;

'55
'virtual HRESULT STDMETHODCALLTYPE add_WebResourceRequested(
'    /* [in] */ ICoreWebView2WebResourceRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'56
'virtual HRESULT STDMETHODCALLTYPE remove_WebResourceRequested(
'    /* [in] */ EventRegistrationToken token) = 0;

'57
'virtual HRESULT STDMETHODCALLTYPE AddWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;

'58
'virtual HRESULT STDMETHODCALLTYPE RemoveWebResourceRequestedFilter(
'    /* [in] */ const LPCWSTR uri,
'    /* [in] */ const COREWEBVIEW2_WEB_RESOURCE_CONTEXT resourceContext) = 0;

'59
'virtual HRESULT STDMETHODCALLTYPE add_WindowCloseRequested(
'    /* [in] */ ICoreWebView2WindowCloseRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'60
'virtual HRESULT STDMETHODCALLTYPE remove_WindowCloseRequested(
'    /* [in] */ EventRegistrationToken token) = 0;
'
'};


''' <summary>
''' NavigationCompleted イベントハンドラを登録する
''' </summary>
Public Function AddNavigationCompletedHandler(Handler As c4_Handler2) As Long
    On Error GoTo ErrorHandler

    ' 1. コレクションの初期化（初回のみ）
    If Col_Handler Is Nothing Then Set Col_Handler = New Collection

    ' 2. ハンドラクラスのインスタンス化
    ' ※ Initialize 内で VTable と マシン語サンクが自動生成されます
    'Set handler = New c4_Handler2
    Handler.ppWebView2 = ppWebView2

    ' 3. WebView2 (COMインターフェース) への登録準備
    ' ICoreWebView2::add_NavigationCompleted の VTable Index は 15 (x64)
    Const IDX_ADD_NAV_COMPLETED As Long = 15

    Dim pWebView2 As LongPtr: pWebView2 = Me.ppWebView2 ' 保持している ICoreWebView2 ポインタ
    Dim pHandlerObj As LongPtr: pHandlerObj = Handler.Pointer ' 捏造された FakeObjPtr
    Dim tkn As LongPtr ' 登録解除用のトークンを受け取る変数

    ' DispCallFunc の引数設定
    ' HRESULT add_NavigationCompleted(handler, [out] token)
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
    Dim res As Variant
    Dim hr As Long

    ' 第1引数: [in] ICoreWebView2NavigationCompletedEventHandler*
    args(0) = pHandlerObj
    argTypes(0) = vbLongPtr
    argPtrs(0) = VarPtr(args(0))

    ' 第2引数: [out] EventRegistrationToken* (ポインタを渡す)
    args(1) = VarPtr(tkn)
    argTypes(1) = vbLongPtr
    argPtrs(1) = VarPtr(args(1))

    ' 4. DispCallFunc で API 実行
    hr = DispCallFunc(pWebView2, _
                      IDX_ADD_NAV_COMPLETED * LenB(pWebView2), _
                      CC_STDCALL, _
                      vbLong, _
                      2, _
                      argTypes(0), _
                      argPtrs(0), _
                      res)

    ' 5. 結果の判定
    If hr = 0 Then
        If res = 0 Then
            ' 成功：トークンを保存し、インスタンスをコレクションで維持する
            Handler.token = tkn
            Col_Handler.Add Handler, CStr(tkn) ' トークンをキーにして保持

            Debug.Print "--- ハンドラ登録成功 ---"
            Debug.Print "Token:", tkn
            AddNavigationCompletedHandler = 0 ' S_OK
        Else
            Debug.Print "WebView2側でエラー発生。HRESULT: " & Hex(res)
            AddNavigationCompletedHandler = res
        End If
    Else
        Debug.Print "DispCallFunc 自体が失敗。hr: " & Hex(hr)
        AddNavigationCompletedHandler = hr
    End If

    Exit Function

ErrorHandler:
    Debug.Print "AddNavigationCompletedHandler で予期せぬエラー: " & Err.Description
    AddNavigationCompletedHandler = -1
End Function

