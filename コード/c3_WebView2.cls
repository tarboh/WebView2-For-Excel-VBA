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
Public Event NavigationCompleted()
Public Event ReceiveScriptResult(ByVal result As String)
Public LastScriptResult As String


Public Col_Handler As New Collection
Private m_IsNavigating As Boolean
Private m_IsScriptRunning As Boolean

Public cHandler As c4_Handler2

' --- 内部通知用 ---
Public Sub NotifyNavigationStarting()
    RaiseEvent NavigationStarting
End Sub
Public Sub NotifyNavigationCompleted()
    m_IsNavigating = False
    RaiseEvent NavigationCompleted
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
    add_NavigationStarting = DCF_ハンドラ登録(Me, 7, "add_NavigationStarting", AddressOf NavigationStarting_Invoke)
End Function


'8
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;

'9
'virtual HRESULT STDMETHODCALLTYPE add_ContentLoading(
'    /* [in] */ ICoreWebView2ContentLoadingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'10
'virtual HRESULT STDMETHODCALLTYPE remove_ContentLoading(
'    /* [in] */ EventRegistrationToken token) = 0;

'11
'virtual HRESULT STDMETHODCALLTYPE add_SourceChanged(
'    /* [in] */ ICoreWebView2SourceChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'12
'virtual HRESULT STDMETHODCALLTYPE remove_SourceChanged(
'    /* [in] */ EventRegistrationToken token) = 0;

'13
'virtual HRESULT STDMETHODCALLTYPE add_HistoryChanged(
'    /* [in] */ ICoreWebView2HistoryChangedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'14
'virtual HRESULT STDMETHODCALLTYPE remove_HistoryChanged(
'    /* [in] */ EventRegistrationToken token) = 0;

'15
'virtual HRESULT STDMETHODCALLTYPE add_NavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;
'add_NavigationCompleted
Public Function add_NavigationCompleted() As Long
    add_NavigationCompleted = mHelpers.DCF_ハンドラ登録(Me, 15, "add_NavigationCompleted", AddressOf NavCompleted_Invoke)
End Function
Public Function add_NavigationCompleted2() As Long
        
    Set cHandler = New c4_Handler2
    Dim funcPtr As LongPtr, vtbloffset As Long
    funcPtr = GetClassMethodPtr(cHandler, "NavCompleted_Invoke", vtbloffset)
    
    add_NavigationCompleted2 = mHelpers.DCF_ハンドラ登録(Me, 15, "add_NavigationCompleted2", funcPtr)

End Function

Public Function Register_NavCompleted(ByVal wb2Ptr As LongPtr, ByVal handler As c4_Handler2) As LongPtr
    ' wb2Ptr: ICoreWebView2 の生ポインタ
    ' handler: 捏造 VTable を持った c4_Handler2 のインスタンス
    
    Dim hr As Long
    Dim pHandler As LongPtr: pHandler = handler.Pointer
    Dim tkn As LongPtr ' トークンを受け取る変数
    Dim res As Variant
    
    ' DispCallFunc 用の準備 (引数2つ)
    ' 1: [in]  eventHandler (IUnknown*) -> pHandler
    ' 2: [out] token (EventRegistrationToken*) -> VarPtr(tkn)
    
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
    Dim args(1) As Variant
    
    ' 第1引数: Handlerのポインタそのものを渡す (ByVal)
    args(0) = pHandler
    argTypes(0) = vbLongPtr
    argPtrs(0) = VarPtr(args(0))
    
    ' 第2引数: Tokenを受け取る変数のアドレスを渡す (ByRef相当)
    args(1) = VarPtr(tkn)
    argTypes(1) = vbLongPtr
    argPtrs(1) = VarPtr(args(1))
    
    ' add_NavigationCompleted の VTable Index は 15 (64bit環境)
    ' ※環境により異なる場合は要調整
    Const IDX_ADD_NAV_COMPLETED As Long = 15
    
    hr = DispCallFunc(wb2Ptr, IDX_ADD_NAV_COMPLETED * LenB(wb2Ptr), _
                      CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
                      
    If hr = 0 And res = 0 Then
        ' 成功したらトークンを保持しておく
        handler.Token = tkn
        Register_NavCompleted = tkn
        Debug.Print "ハンドラ登録成功！ Token: " & tkn
    Else
        Debug.Print "ハンドラ登録失敗。 hr: " & Hex(hr) & " / res: " & res
    End If
End Function


'16
'virtual HRESULT STDMETHODCALLTYPE remove_NavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;

'17
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationStarting(
'    /* [in] */ ICoreWebView2NavigationStartingEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'18
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationStarting(
'    /* [in] */ EventRegistrationToken token) = 0;

'19
'virtual HRESULT STDMETHODCALLTYPE add_FrameNavigationCompleted(
'    /* [in] */ ICoreWebView2NavigationCompletedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'20
'virtual HRESULT STDMETHODCALLTYPE remove_FrameNavigationCompleted(
'    /* [in] */ EventRegistrationToken token) = 0;

'21
'virtual HRESULT STDMETHODCALLTYPE add_ScriptDialogOpening(
'    /* [in] */ ICoreWebView2ScriptDialogOpeningEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'22
'virtual HRESULT STDMETHODCALLTYPE remove_ScriptDialogOpening(
'    /* [in] */ EventRegistrationToken token) = 0;

'23
'virtual HRESULT STDMETHODCALLTYPE add_PermissionRequested(
'    /* [in] */ ICoreWebView2PermissionRequestedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'24
'virtual HRESULT STDMETHODCALLTYPE remove_PermissionRequested(
'    /* [in] */ EventRegistrationToken token) = 0;

'25
'virtual HRESULT STDMETHODCALLTYPE add_ProcessFailed(
'    /* [in] */ ICoreWebView2ProcessFailedEventHandler *eventHandler,
'    /* [out] */ EventRegistrationToken *token) = 0;

'26
'virtual HRESULT STDMETHODCALLTYPE remove_ProcessFailed(
'    /* [in] */ EventRegistrationToken token) = 0;

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
Public Function AddNavigationCompletedHandler(handler As c4_Handler2) As Long
    On Error GoTo ErrorHandler

    ' 1. コレクションの初期化（初回のみ）
    If Col_Handler Is Nothing Then Set Col_Handler = New Collection

    ' 2. ハンドラクラスのインスタンス化
    ' ※ Initialize 内で VTable と マシン語サンクが自動生成されます
    'Set handler = New c4_Handler2
    handler.ppWebView2 = ppWebView2

    ' 3. WebView2 (COMインターフェース) への登録準備
    ' ICoreWebView2::add_NavigationCompleted の VTable Index は 15 (x64)
    Const IDX_ADD_NAV_COMPLETED As Long = 15
    
    Dim pWebView2 As LongPtr: pWebView2 = Me.ppWebView2 ' 保持している ICoreWebView2 ポインタ
    Dim pHandlerObj As LongPtr: pHandlerObj = handler.Pointer ' 捏造された FakeObjPtr
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
            handler.Token = tkn
            Col_Handler.Add handler, CStr(tkn) ' トークンをキーにして保持
            
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

