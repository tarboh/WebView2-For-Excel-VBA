VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Environment.cls 第9.8a段階 (旧称: Class8) ---
'
'   第9.8a の変更点 (改名のみ、ロジック変更なし):
'     - クラス名 Class8 → Wv2Environment に正式固定
'     - Class7 への言及を ComCallbackHandler に変更
'     - 自己クラスのデバッグログを "Class8.xxx" → "Wv2Environment.xxx" に変更
'     - 第9.3a で確立した責務・寿命管理・状態遷移は不変
'
'   役割: ICoreWebView2Environment の VBA 側ラッパ。
'     - HK_EnvironmentCompleted ハンドラ (ComCallbackHandler) を 1 個生成・保持する
'     - CreateCoreWebView2EnvironmentWithOptions を呼んで WebView2 環境作成を依頼する
'     - サンク経由で ComCallbackHandler.Handler_Invoke が叩かれ、その m_kind
'       分岐から自分の OnEnvironmentCompleted が呼ばれる流れで完了通知を受ける
'     - 受け取った pEnvironment に対して COM 規約通り AddRef を撃ち、
'       Class_Terminate で Release する責任を持つ
'
'   寿命管理 (案 D1-α: 1 ショットハンドラパターン):
'     - HK_EnvironmentCompleted は「1 回呼ばれて終わり」のハンドラ。
'     - OnEnvironmentCompleted の最後で Set m_handler = Nothing して
'       ComCallbackHandler への外部参照を切る。
'     - WebView2 が ComCallbackHandler を Release すれば refcount=0 となり、
'       基盤側の自動解放経路が走る (ComCallbackHandler.ClearOwner で
'       m_owner=Me が切られ、Wv2Environment ←→ ComCallbackHandler の
'       循環参照が解消)。
'     - その後 Wv2Environment (Me) はテストプロシージャのローカル変数のスコープ
'       終了等で解放され、Class_Terminate で m_pEnvironment->Release が走る。
'
'   状態遷移:
'     Es_Created       : Class_Initialize 直後
'     Es_Waiting       : Init で API を呼んだ後、コールバック待ち
'     Es_Ready         : OnEnvironmentCompleted で errorCode=0 を受け取った
'     Es_Failed        : OnEnvironmentCompleted で errorCode<>0 を受け取った
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "このクラスは 64 ビット VBA (x64) が必要です"
#End If


Public Enum EnvironmentState
    Es_Created = 0
    Es_Waiting = 1
    Es_Ready = 2
    Es_Failed = 3
End Enum


' --- 内部状態 ---
' m_handler:      HK_EnvironmentCompleted ハンドラへの強参照
'                 (OnEnvironmentCompleted 後に Nothing にする)
' m_pEnvironment: WebView2 から受け取った ICoreWebView2Environment
'                 (Class_Terminate で Release する責任を持つ)
Private m_handler      As ComCallbackHandler
Private m_pEnvironment As LongPtr
Private m_state        As EnvironmentState
Private m_lastError    As Long


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_state = Es_Created
    m_pEnvironment = 0
    m_lastError = 0
End Sub

Private Sub Class_Terminate()
    ' COM 所有権規約: Init で受け取った pEnvironment に対して
    ' OnEnvironmentCompleted で AddRef を撃って VBA 側の所有権分を確保しているので、
    ' ここで Release を撃って解放する。
    If m_pEnvironment <> 0 Then
        Debug.Print "Wv2Environment.Terminate: Releasing m_pEnvironment = " & m_pEnvironment
        ComRelease m_pEnvironment
        m_pEnvironment = 0
    End If

    ' m_handler は通常 OnEnvironmentCompleted で Nothing になっているはずだが、
    ' 万一 Init で例外が起きた等で残っている場合は、ここで切る。
    ' (ComCallbackHandler 側は基盤の m_handlers(idx) からも参照されているので、
    '  ここで切ることでスロットの自動解放経路が走る。)
    If Not (m_handler Is Nothing) Then
        Debug.Print "Wv2Environment.Terminate: m_handler が残っているので切る (異常終了パス)"
        Set m_handler = Nothing
    End If
End Sub


' ============================================================
' Init
'
'   ICoreWebView2Environment の作成を開始する。
'   このメソッドはコールバックを待たずに即座にリターンする (非同期)。
'   完了通知は OnEnvironmentCompleted で受ける。
'
'   呼び出し側は env.IsReady が True になるまで DoEvents で待つ。
' ============================================================
Public Sub Init(ByVal userDataPath As String)

    If m_state <> Es_Created Then
        MsgBox "Wv2Environment.Init は二重呼び出し不可", vbCritical
        Exit Sub
    End If

    ' --- 1. ハンドラ (ComCallbackHandler) を 1 個 Acquire ---
    '   AcquireHandlerFor の中で:
    '     - ComCallbackHandler を New
    '     - Thunks_AcquireSlot でスロット確保 + サンク書き込み
    '     - ComCallbackHandler.Init(HK_EnvironmentCompleted, owner=Me, pSlot)
    '     - m_handlers(idx) に登録
    '   が一括で行われ、初期化済みの ComCallbackHandler が返ってくる。
    Set m_handler = AcquireHandlerFor(HK_EnvironmentCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Environment.Init: ハンドラの Acquire に失敗", vbCritical
        m_state = Es_Failed
        Exit Sub
    End If

    ' --- 2. WebView2 が握る obj アドレスを計算 ---
    '   pSlot + VTABLE_OBJ_OFFSET が vtableObj の先頭。
    '   このアドレスが obj として WebView2 に渡され、後の AddRef/Release/Invoke の
    '   this として使われる。
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. ユーザデータフォルダを準備 ---
    Dim path As String
    path = userDataPath
    If LenB(path) = 0 Then path = "C:\Temp\VBA_WebView2"
    EnsureFolder path

    ' --- 4. CreateCoreWebView2EnvironmentWithOptions を呼ぶ ---
    '
    '   重要: m_state = Es_Waiting は API 呼び出しの「前」に立てる。
    '   理由は、WebView2Loader が API 内部で同期的にコールバックを発火させる
    '   ケースがあるため (実機ログで確認済み):
    '
    '     CreateCoreWebView2EnvironmentWithOptions 呼び出し開始
    '       └─ 内部で WebView2 が AddRef → Invoke → Release を同期実行
    '          └─ Wv2Environment.OnEnvironmentCompleted が呼ばれて m_state = Es_Ready
    '       ← API が S_OK で返ってくる
    '
    '   このとき、API 呼び出しの「後」に m_state = Es_Waiting を立てると、
    '   コールバックで設定された Es_Ready が Es_Waiting で上書きされてしまい、
    '   呼び出し側の IsReady ループが永遠に True にならない。
    '
    '   先に Es_Waiting を立てておけば、同期コールバックでも非同期コールバックでも
    '   最終状態は OnEnvironmentCompleted が設定した値 (Es_Ready / Es_Failed) に
    '   なって整合する。API 失敗時のみ後で Es_Failed に上書きする。
    m_state = Es_Waiting

    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(path), 0, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2EnvironmentWithOptions failed. HRESULT: &H" & Hex(hr), _
               vbCritical
        m_state = Es_Failed
        m_lastError = hr
        ' 失敗時のリーク対策は第9.3a では未実装。
        ' WebView2 が AddRef を撃たずに失敗復帰した場合、refcount=0 のまま
        ' m_handlers(idx) が残ってスロットがリークする。
        ' 復旧は Thunks_Shutdown を呼ぶことで全スロット強制クリーンアップに頼る。
        ' 第9.3b 以降で Controller 生成失敗パスと併せて統一的なヘルパーを設計する予定。
        Exit Sub
    End If
End Sub


' ============================================================
' OnEnvironmentCompleted
'
'   ComCallbackHandler.Handler_Invoke の HK_EnvironmentCompleted 分岐から
'   m_owner.OnEnvironmentCompleted が動的バインディングで呼ばれる。
'
'   このメソッドは Public Sub 必須:
'     ComCallbackHandler 側は m_owner As Object 型なので呼び出しは遅延バインディング
'     (IDispatch::Invoke 経由) になる。Friend では IDispatch に乗らないため
'     「オブジェクトは、このプロパティまたはメソッドをサポートしていません。」
'     ランタイムエラーが出る。Public にすると IDispatch に乗って呼べるようになる。
'
'   引数:
'     errorCode    : HRESULT (S_OK = 0 で成功)
'     pEnvironment : ICoreWebView2Environment ポインタ
'                    (WebView2 が一時的に AddRef した状態で渡される。
'                     コールバックから戻った後 WebView2 が Release するため、
'                     VBA 側で保持し続けるなら自分で AddRef が必要。)
' ============================================================
Public Sub OnEnvironmentCompleted( _
    ByVal errorCode As Long, _
    ByVal pEnvironment As LongPtr)

    Debug.Print "Wv2Environment.OnEnvironmentCompleted: errorCode=" & errorCode & _
                ", pEnvironment=" & pEnvironment

    m_lastError = errorCode

    If errorCode = 0 And pEnvironment <> 0 Then
        ' VBA 側で保持する分の参照を確保
        m_pEnvironment = pEnvironment
        ComAddRef m_pEnvironment
        m_state = Es_Ready
    Else
        m_pEnvironment = 0
        m_state = Es_Failed
    End If

    ' 1 ショットハンドラパターン (案 D1-α):
    '   ハンドラはもう用途を終えたので参照を切る。
    '   このタイミングでは ComCallbackHandler.Handler_Invoke の中で WebView2 が
    '   まだ refcount=1 を保持している (WebView2 が直後に Release を撃つ)。
    '   ここで m_handler を切ると、Wv2Environment → ComCallbackHandler の参照は消えるが、
    '   基盤側の m_handlers(idx) からはまだ参照されているため ComCallbackHandler は
    '   生存し続ける。WebView2 が Release を撃って refcount=0 になった瞬間、
    '   基盤の HandlerReleaseInternal の自動解放経路が走り、
    '   ComCallbackHandler.ClearOwner で m_owner (= Me Wv2Environment) への参照も切れる。
    Set m_handler = Nothing
End Sub


' ============================================================
' プロパティ群
' ============================================================
Public Property Get state() As EnvironmentState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Es_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Es_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

Public Property Get EnvironmentPtr() As LongPtr
    EnvironmentPtr = m_pEnvironment
End Property




