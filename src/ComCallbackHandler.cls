VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComCallbackHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls N-3 段階 (レスポンスのステータスとヘッダ) ---
'
'   N-3 の変更点:
'     - Handler_Invoke に Case HK_WebResourceResponseReceived を追加。
'       2 引数イベント (sender, args) なので既存と同形。
'       → m_owner.View2_OnWebResourceResponseReceived arg1, arg2
'
'   ★命名は View2_ 接頭辞★ (設計原則31 / 第9.9b 規約)
'     このイベントは ICoreWebView2_2 (派生 IF) のもの。add / remove も
'     m_pView2 で叩く。基底の HK_WebResourceRequested (View_) と混ぜないこと。
'
'   ★この分岐も Debug.Print を撃たない★
'     ★このイベントにはフィルタが無い★ ので、画像も CSS も全部来る。
'     N-1 の WebResourceRequested より発火数が多い。記録だけしてすぐ返す。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls N-1 段階 (WebResourceRequested の配線) ---
'
'   N-1 の変更点:
'     - Handler_Invoke に Case HK_WebResourceRequested を追加。
'       2 引数イベント (sender, args) なので DocumentTitleChanged 分岐と同形。
'       → m_owner.View_OnWebResourceRequested arg1, arg2
'
'   ★この分岐だけ Debug.Print を撃たない (N-1 論点5)★
'     WebResourceRequested は 1 ページで数十～数百件発火し、
'     ★ハンドラが遅いと通信そのものが詰まる★ (要求を止める / 書き換えるための
'     イベントなので、ランタイムは Invoke の復帰を待って初めて通信を進める)。
'     他の Case にある 4 行の Debug.Print は、この種別だけ意図的に外している。
'     記録は Wv2Pane 側のリングバッファが持ち、取り出しは NetLogDrain で
'     後からまとめて行う (論点6)。
'
'   ★ハンドラの中で EvalSync を呼んではいけない★ (仕様事実14 / 20 の領域)。
'     Wv2Pane 側の in-callback ガードが働くので呼んでも固まりはしないが、
'     そもそも呼ばない設計にしてある。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- ComCallbackHandler.cls 第9.9b段階 (旧称: Class7) ---
'
'   第9.9b の変更点 (SPA 対応):
'     - Handler_Invoke に 2 種類の分岐を追加:
'         Case HK_HistoryChanged   → m_owner.View_OnHistoryChanged arg1, arg2
'         Case HK_DOMContentLoaded → m_owner.View2_OnDOMContentLoaded arg1, arg2
'       どちらも 2 引数イベント (sender, args)。DocumentTitleChanged 分岐と同形。
'     - 本体ロジック (Handler_QueryInterface / AddRef / Release) は無変更
'       (kind は動的アクセスなので新 kind=9/10 も自動的に通る)
'
'   第9.9a の変更点: なし (ComCallbackHandler はノータッチだった)
'
'   第9.8c の変更点 (継続):
'     - 新規モジュール変数 m_callbackId As Long 追加
'       (HK_ExecuteScriptCompleted 用、同時並行 ExecuteScript の呼び出しを識別)
'     - 新規 Friend Property: CallbackId (Get/Let、Wv2Pane.View_ExecuteScript が
'       Acquire 直後に Let で設定する)
'     - Handler_Invoke に Case HK_ExecuteScriptCompleted を追加 (1 ショット系、
'       LPCWSTR result を String に PtrToString してから Wv2Pane へ通知)
'     - Wv2Pane.OnExecuteScriptCompleted(callbackId, errorCode, resultStr) に
'       動的バインディングで通知 (案 α 第一実装はログのみ)
'
'   第9.8a/9.8b の継続事項:
'     - クラス名 Class7 → ComCallbackHandler (9.8a)
'     - 9.8b では ComCallbackHandler の変更なし
'
'   役割: COM コールバック (サンク経由) の VBA 側ラッパ。
'     インスタンスは「1 イベントハンドラ = 1 個」の関係で生成され、
'     WebView2 が握る vtable の AddRef/Release/Invoke を VBA の
'     m_refCount + m_owner.動的バインディング呼び出しで実装する。
'
'   実装済み種別:
'     HK_EnvironmentCompleted    (1 ショット、案 D1-α)
'     HK_ControllerCompleted     (1 ショット、案 D1-α)
'     HK_NavigationStarting      (永続、案 P1)
'     HK_NavigationCompleted     (永続、案 P1)
'     HK_WebMessageReceived      (永続、案 P1)
'     HK_DocumentTitleChanged    (永続、案 P1)
'     HK_NewWindowRequested      (永続、案 P1)
'     HK_ExecuteScriptCompleted  (1 ショット、案 a = 同時並行 OK、第9.8c で追加)
'     HK_HistoryChanged          (永続、第9.9b で追加、基底 ICoreWebView2)
'     HK_DOMContentLoaded        (永続、第9.9b で追加、ICoreWebView2_2)
'     HK_WebResourceRequested    (永続、N-1 で追加、基底 ICoreWebView2)
'     HK_WebResourceResponseReceived (永続、N-3 で追加、ICoreWebView2_2)
'
'   引数解釈の法則 (設計原則 11):
'     ComCallbackHandler の Handler_Invoke が受ける (this, arg1, arg2) は
'     常に WebView2 視点の (C++ this, 1st 引数, 2nd 引数) に一致する。
'     2 引数イベント (sender, args) の Invoke では:
'       this = handler obj (使わない)
'       arg1 = sender (= m_pView と同じはず)
'       arg2 = args (各イベント固有の args インターフェース)
'     2 引数 1 ショット系 (Environment/Controller/ExecuteScript Completed) では:
'       this = handler obj (使わない)
'       arg1 = errorCode (HRESULT、LongPtrLowDword で下位 32 bit 抽出)
'       arg2 = 結果ポインタ (pEnvironment/pController/LPCWSTR result)
'
'   命名規約 (設計原則 31):
'     ICoreWebView2 のイベントレシーバーは View_On<EventName> 形式
'     (基底 IF) / View2_On<EventName> 形式 (派生 IF、第9.9b で初導入)。
'     いずれも Wv2Pane 内の Public Sub で、ComCallbackHandler.Handler_Invoke
'     から動的バインディングで呼ばれる。
'
'   重要: m_owner.View_OnXxx は **Public Sub** でなければ動的バインディング
'     経由で呼べない。Wv2Pane の全 View_OnXxx / View2_OnXxx / OnXxxCompleted
'     メソッドは Public 必須。
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "このクラスは 64 ビット VBA (x64) が必要です"
#End If

Private m_kind     As HandlerKind
Private m_owner    As Object
Private m_pSlot    As LongPtr
Private m_refCount As Long

' 第9.8c で追加:
'   HK_ExecuteScriptCompleted は同時並行 OK (案 a)。各呼び出しを識別するため、
'   ComCallbackHandler インスタンスが callback ID を保持する。
'   HK_ExecuteScriptCompleted 以外の kind では 0 のまま使われない。
Private m_callbackId As Long


' ============================================================
' 初期化 (基盤の AcquireHandlerFor / Thunks_AcquireSlot から呼ばれる前提)
'
'   refcount は 0 から開始する (案 X-1 採用):
'     - COM 意味論: ハンドラを WebView2 に渡したら、WebView2 が自分で
'       AddRef して責任を持つ。VBA 側は参照を持たない。
'     - CreateCoreWebView2EnvironmentWithOptions の内部で WebView2 が
'       AddRef を呼ぶことが実機で確認済み (refcount: 0 → 1)。
'     - Invoke 後に WebView2 が Release を呼んで refcount = 0 となり、
'       HandlerReleaseInternal の自動解放経路が走る。
' ============================================================
Friend Sub Init( _
    ByVal kind As HandlerKind, _
    ByVal owner As Object, _
    ByVal pSlot As LongPtr)

    m_kind = kind
    Set m_owner = owner
    m_pSlot = pSlot
    m_refCount = 0
End Sub


' ============================================================
' プロパティ群 (基盤側からの操作用)
' ============================================================
Friend Property Get RefCount() As Long
    RefCount = m_refCount
End Property

Friend Property Let RefCount(ByVal v As Long)
    m_refCount = v
End Property

Friend Property Get Slot() As LongPtr
    Slot = m_pSlot
End Property

Friend Property Get kind() As HandlerKind
    kind = m_kind
End Property


' 第9.8c で追加: CallbackId プロパティ (HK_ExecuteScriptCompleted で使用)
'   Wv2Pane.View_ExecuteScript が呼び出しごとに連番 ID を採番し、
'   AcquireHandlerFor の直後にこの Let で設定する。
'   その後 Handler_Invoke 内で m_callbackId として参照され、
'   Wv2Pane.OnExecuteScriptCompleted に渡される。
Friend Property Get callbackId() As Long
    callbackId = m_callbackId
End Property

Friend Property Let callbackId(ByVal v As Long)
    m_callbackId = v
End Property


' ============================================================
' ClearOwner
'   refcount = 0 になったときに基盤側 (HandlerReleaseInternal) から呼ばれる。
'   循環参照を切るため、m_owner を Nothing に倒す。
' ============================================================
Friend Sub ClearOwner()
    Set m_owner = Nothing
End Sub


' ============================================================
' Handler_Invoke (vTable slot 7 固定)
'
'   WebView2 からのコールバックを受ける唯一のエントリ。
'   サンクが ObjPtr(this) を Me に注入し、引数を 1 個ずつシフトして
'   このメソッドに到達する。
'
'   引数仕様:
'     this : サンク後の rdx (= WebView2 が rcx に乗せていた obj)
'            HK_EnvironmentCompleted では使用しない (識別は m_kind で行う)
'     arg1 : サンク後の r8   (= WebView2 が rdx に乗せていた値)
'     arg2 : サンク後の r9   (= WebView2 が r8 に乗せていた値)
'
'   m_kind ごとの解釈:
'     HK_EnvironmentCompleted: WebView2 視点 Invoke(this, errorCode, pEnv)
'                              → this = handler obj, arg1 = errorCode (下位32bit),
'                                 arg2 = pEnvironment
'     HK_ControllerCompleted:  WebView2 視点 Invoke(this, errorCode, pCtrl)
'                              → this = handler obj, arg1 = errorCode (下位32bit),
'                                 arg2 = pController
'     HK_NavigationStarting:   WebView2 視点 Invoke(this, sender, args)
'                              → this = handler obj, arg1 = sender (=m_pView),
'                                 arg2 = args (= ICoreWebView2NavigationStartingEventArgs)
'     HK_NavigationCompleted:  WebView2 視点 Invoke(this, sender, args)
'                              → 同上、args の型が NavigationCompletedEventArgs
'     HK_WebMessageReceived:   同上、args の型が WebMessageReceivedEventArgs
'     HK_DocumentTitleChanged: 同上、args の型が IUnknown* (情報なし)
'     HK_NewWindowRequested:   同上、args の型が
'                              ICoreWebView2NewWindowRequestedEventArgs
'                              (第9.6a で追加)
'
'   重要 (第9.4a 段階で確定):
'     ComCallbackHandler の this 引数は全種別共通で「handler obj アドレス」を受ける。
'     WebView2 視点での「第 1 引数」(C++ this の次に置かれる引数) は arg1。
'     第9.3 段階で「this = obj」と書いたのは正しいが、HK_EnvironmentCompleted
'     等で「arg1 = errorCode」が偶々 0 で気づかなかっただけで、3 引数ハンドラ
'     でも 2 引数ハンドラでも同じ法則が成り立つ。
' ============================================================
Public Function Handler_Invoke( _
    ByVal this As LongPtr, _
    ByVal arg1 As LongPtr, _
    ByVal arg2 As LongPtr) As Long

    Select Case m_kind
        Case HK_EnvironmentCompleted
            Dim errorCode As Long
            errorCode = LongPtrLowDword(arg1)
            Debug.Print "ComCallbackHandler.Invoke[HK_EnvironmentCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  errorCode   :", errorCode
            Debug.Print "  pEnvironment:", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 上位クラスへ通知 (m_owner が居るときだけ)
            ' 動的バインディング: m_owner は Object 型だが
            ' OnEnvironmentCompleted という名前のメソッドさえ持っていれば呼べる。
            ' 第9.3a 段階で Wv2Environment (旧称 Class8) がこのメソッドを実装している。
            '
            ' 注 (第9.6a): 1 ショット系コールバックは出自接頭辞なしの命名を維持。
            ' 派生 IF (ICoreWebView2Environment2 等) を直接受け取る関数が
            ' 存在しないためバージョン識別の必要がない。
            If Not (m_owner Is Nothing) Then
                m_owner.OnEnvironmentCompleted errorCode, arg2
            End If

        Case HK_ControllerCompleted
            Dim errorCode2 As Long
            errorCode2 = LongPtrLowDword(arg1)
            Debug.Print "ComCallbackHandler.Invoke[HK_ControllerCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  errorCode   :", errorCode2
            Debug.Print "  pController :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 上位クラスへ通知 (m_owner が居るときだけ)
            ' m_owner は Wv2Pane (旧称 Class9、Controller+View 合体クラス) を指す前提
            ' 1 ショット系のため出自接頭辞なし (第9.6a 規約より)
            If Not (m_owner Is Nothing) Then
                m_owner.OnControllerCompleted errorCode2, arg2
            End If

        Case HK_NavigationStarting
            Debug.Print "ComCallbackHandler.Invoke[HK_NavigationStarting]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 上位クラスへ通知 (m_owner が居るときだけ)
            '
            ' サンクの動作 (再確認):
            '   サンクは「引数を 1 個右にシフトして rcx に ObjPtr を注入」
            '   する仕組み。VBA 視点では:
            '     Me      = サンクが注入した ObjPtr (VBA が暗黙に消費)
            '     this    = rdx (元 WebView2 が rcx に乗せた C++ this) = ハンドラ obj
            '     arg1    = r8  (元 WebView2 が rdx に乗せた 1st 引数)
            '     arg2    = r9  (元 WebView2 が r8 に乗せた 2nd 引数)
            '
            ' C++ シグネチャ (Microsoft Learn で確定):
            '   HRESULT Invoke(ICoreWebView2* sender, ICoreWebView2NavigationStartingEventArgs* args)
            ' これは C++ メンバ関数 = 暗黙の this を含めて 3 引数:
            '   rcx = this (handler), rdx = sender, r8 = args
            ' サンク後の VBA 視点:
            '   this (我々の) = handler obj (C++ this を受けただけ、通常は使わない)
            '   arg1          = sender = ICoreWebView2 ポインタ (= m_pView と同じはず)
            '   arg2          = args   = ICoreWebView2NavigationStartingEventArgs ポインタ
            '
            ' 注意 (第9.4a 段階での教訓):
            '   第9.3 で確立した「this = obj」は HK_EnvironmentCompleted/Controller でも
            '   同じ。あの段階で errorCode = 0 が偶々全件成功で arg1 = errorCode と
            '   正しく解釈できていただけ。ComCallbackHandler が受け取る (this, arg1, arg2) は常に
            '   サンク後の (rdx, r8, r9) で、WebView2 視点では (1st, 2nd, 3rd) 引数。
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationStarting arg1, arg2
            End If

        Case HK_NavigationCompleted
            Debug.Print "ComCallbackHandler.Invoke[HK_NavigationCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2NavigationCompletedEventArgs)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNavigationCompleted arg1, arg2
            End If

        Case HK_WebMessageReceived
            Debug.Print "ComCallbackHandler.Invoke[HK_WebMessageReceived]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2WebMessageReceivedEventArgs)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebMessageReceived arg1, arg2
            End If

        Case HK_DocumentTitleChanged
            Debug.Print "ComCallbackHandler.Invoke[HK_DocumentTitleChanged]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= IUnknown* で情報なし、sender 経由で
            '                  ICoreWebView2::get_DocumentTitle を呼ぶ必要)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnDocumentTitleChanged arg1, arg2
            End If

        Case HK_NewWindowRequested
            Debug.Print "ComCallbackHandler.Invoke[HK_NewWindowRequested]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2NewWindowRequestedEventArgs)
            '
            ' args の vtable レイアウト (旧版 WebView2.cls の IDL コメントより):
            '    3 : get_Uri(LPWSTR* uri)
            '    4 : put_NewWindow(ICoreWebView2*)
            '    5 : get_NewWindow(ICoreWebView2**)
            '    6 : put_Handled(BOOL)
            '    7 : get_Handled(BOOL*)
            '    8 : get_IsUserInitiated(BOOL*)
            '    9 : GetDeferral(ICoreWebView2Deferral**)
            '   10 : get_WindowFeatures(ICoreWebView2WindowFeatures**)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnNewWindowRequested arg1, arg2
            End If

        Case HK_HistoryChanged
            ' 第9.9b で追加。永続、基底 ICoreWebView2 (vtable 13/14)。
            Debug.Print "ComCallbackHandler.Invoke[HK_HistoryChanged]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= IUnknown* で情報なし。HistoryChanged の args は
            '                  情報を持たないので、sender 経由で get_Source /
            '                  get_CanGoBack / get_CanGoForward を呼ぶ)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnHistoryChanged arg1, arg2
            End If

        Case HK_DOMContentLoaded
            ' 第9.9b で追加。永続、ICoreWebView2_2 (vtable 64/65)。
            Debug.Print "ComCallbackHandler.Invoke[HK_DOMContentLoaded]"
            Debug.Print "  this        :", this
            Debug.Print "  arg1 (raw)  :", arg1
            Debug.Print "  arg2 (raw)  :", arg2
            Debug.Print "  m_refCount  :", m_refCount

            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView、ただし ICoreWebView2_2 の Invoke は
            '          sender を ICoreWebView2* で渡す。m_pView と同じ値)
            '   arg2 = args   (= ICoreWebView2DOMContentLoadedEventArgs)
            '          vtable 3 = get_NavigationId(UINT64*)
            '
            ' 受け口の命名は View2_OnDOMContentLoaded (派生 IF 由来なので View2_、
            ' 第9.6a/9.9b 命名規約)。
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnDOMContentLoaded arg1, arg2
            End If

        Case HK_WebResourceRequested
            ' N-1 で追加。永続、基底 ICoreWebView2 (vtable 55/56)。
            '
            ' ★ここでは Debug.Print を撃たない (N-1 論点5、ヘッダー参照)★
            '   数十～数百件が連続で飛んでくるうえ、ハンドラが遅いと
            '   通信そのものが詰まる。記録だけしてすぐ返す。
            '
            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView)
            '   arg2 = args   (= ICoreWebView2WebResourceRequestedEventArgs)
            '
            ' args の vtable レイアウト (★tools/vtable.py で WebView2.h から確定★):
            '    3 : get_Request(ICoreWebView2WebResourceRequest**)
            '    4 : get_Response      5 : put_Response      6 : GetDeferral
            '    7 : get_ResourceContext(COREWEBVIEW2_WEB_RESOURCE_CONTEXT*)
            If Not (m_owner Is Nothing) Then
                m_owner.View_OnWebResourceRequested arg1, arg2
            End If

        Case HK_WebResourceResponseReceived
            ' N-3 で追加。永続、ICoreWebView2_2 (vtable 61/62)。
            '
            ' ★ここでも Debug.Print を撃たない★ (N-1 と同じ理由。しかも
            '   このイベントには★フィルタが無い★ので発火数はさらに多い)
            '
            ' 2 引数イベント: Invoke(sender, args)
            '   arg1 = sender (= m_pView。_2 IF でも sender は ICoreWebView2*)
            '   arg2 = args   (= ICoreWebView2WebResourceResponseReceivedEventArgs)
            '
            ' args の vtable (★tools/vtable.py で確定★):
            '    3 : get_Request(ICoreWebView2WebResourceRequest**)
            '    4 : get_Response(ICoreWebView2WebResourceResponseView**)
            If Not (m_owner Is Nothing) Then
                m_owner.View2_OnWebResourceResponseReceived arg1, arg2
            End If

        Case HK_ExecuteScriptCompleted
            ' 第9.8c で追加。1 ショット系、同時並行可 (案 a 採用)。
            '
            ' SDK シグネチャ:
            '   ICoreWebView2ExecuteScriptCompletedHandler::Invoke(
            '       HRESULT errorCode,
            '       LPCWSTR result)
            '
            ' 引数解釈:
            '   arg1 = errorCode  (LongPtr の下位 32 bit、LongPtrLowDword で取り出す)
            '   arg2 = result     (LPCWSTR、JSON エンコード済みの結果文字列)
            '
            ' result の解放責任: LPCWSTR は呼び出し側 (= VBA) が解放する規約ではない
            '   (= 1 ショットの Invoke 引数として渡され、Invoke 戻り後は無効になる)。
            '   ここで PtrToString で String にコピーすれば、その後の参照は安全。
            Dim errorCode3 As Long
            errorCode3 = LongPtrLowDword(arg1)
            Dim resultStr As String
            If arg2 <> 0 Then
                resultStr = PtrToString(arg2)
            End If

            Debug.Print "ComCallbackHandler.Invoke[HK_ExecuteScriptCompleted]"
            Debug.Print "  this        :", this
            Debug.Print "  callbackId  :", m_callbackId
            Debug.Print "  errorCode   :", errorCode3
            Debug.Print "  resultLen   :", Len(resultStr)
            Debug.Print "  m_refCount  :", m_refCount

            ' 上位クラス (Wv2Pane) へ通知。
            ' 案 α 第一実装: Wv2Pane.OnExecuteScriptCompleted がログ出力 + 切り捨て。
            ' 将来 UserForm 経由で受け渡したい場合は Wv2Pane 側を拡張する。
            ' 動的バインディング: m_owner は Wv2Pane を指す前提
            ' 1 ショット系のため出自接頭辞なし (第9.6a 規約より)
            If Not (m_owner Is Nothing) Then
                m_owner.OnExecuteScriptCompleted m_callbackId, errorCode3, resultStr
            End If

            ' 同時並行のため、Wv2Pane は ComCallbackHandler への参照を持っていない
            ' (= AcquireHandlerFor 後すぐに Set h = Nothing して捨てている)。
            ' 完了通知を受けたので、ここで自分の m_owner 参照を切る。
            ' WebView2 ランタイム側の Release が走れば refCount=0 となり、
            ' Handler_Release の自動解放経路から ClearOwner も呼ばれて
            ' 循環参照が綺麗に解消される。
            ' 実際は m_owner = Wv2Pane を即座に切らなくても、ClearOwner は
            ' Handler_Release 内で発火するため、ここで明示的に Set はしない。
            ' (= 既存 HK_EnvironmentCompleted / HK_ControllerCompleted と同じ流れ)

        Case Else
            Debug.Print "ComCallbackHandler.Invoke[未知の kind=" & m_kind & "]"
    End Select

    Handler_Invoke = 0   ' S_OK
End Function


' ============================================================
' LongPtr の下位32bitを符号付き Long として取り出す
'
'   x64 calling convention では、32bit 引数 (HRESULT 等) を渡すとき
'   上位 32bit が未定義のまま rdx 等に乗る可能性がある。
'   サンクは「mov rdx, rcx」のように 64bit シフトしているため、
'   そのまま LongPtr で受けると上位 32bit のゴミが残る。
'
'   下位 32bit を取り出して、0x80000000 以上なら負の Long に折り返す
'   (CLng は Long 範囲外でオーバーフローするため、自前で処理する)。
' ============================================================
Private Function LongPtrLowDword(ByVal v As LongPtr) As Long
    Dim u As LongLong
    u = v And &HFFFFFFFF^
    If u >= &H80000000^ Then
        LongPtrLowDword = CLng(u - &H100000000^)
    Else
        LongPtrLowDword = CLng(u)
    End If
End Function

