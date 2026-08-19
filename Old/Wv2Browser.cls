''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  第9.10b 段階 (骨格) ---
'
'   複数タブを束ねる包括クラス。Wv2Environment を 1 個だけ保持し、
'   そこから N 個の Wv2Pane (= タブ) を生成・管理する。
'   UserForm のフィールドが将来 m_env + m_pane から m_browser 1 個に集約される
'   構想の中核 (UserForm 接続そのものは 9.10c)。
'
'   設計方針 (第9.10 段階 確定事項):
'     - クラス名 Wv2Browser (Wv2 接頭辞統一、設計原則 38)
'     - Wv2Pane は温存 (案 D-1)。本クラスは Collection で束ねるだけで、
'       各タブの View_/Ctrl_/Settings_ には ActivePane 経由で到達する。
'     - Environment は 1 個を全タブ共有 (案 E-1)。
'     - 非同期完了待ちは AddTab / Init の内部に閉じ込める (論点① a 案)。
'       → UserForm 側から Do While ... DoEvents の待機ロジックを消せる。
'
'   9.10b のスコープ (骨格のみ):
'     作る  : Init / AddTab / AddTabWithUrl / AddTabWithHtml / TabCount /
'             ActivePane / Class_Terminate
'     作らない: ActivateTab (表示切替) / RemoveTab / UserForm 接続 /
'               マウスリサイズ連動 (すべて 9.10c 以降)
'     → 全タブは同じ m_hwndHost に重なって生成される (視覚的には最後の
'       1 個が手前)。9.10b は「各 Pane が Controller/View まで到達する」ことの
'       確認までが目的 (論点②)。
'
'   Release 順序 (設計原則 43: 子から親へ):
'     Class_Terminate では全 Pane を解放してから Environment を解放する。
'     各 Wv2Pane の Class_Terminate が Settings→View2→View→Ctrl の順で
'     解放するので、本クラスは「Pane 群 → Environment」の粒度で順序を守る。
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' Browser 状態
' ============================================================
Public Enum BrowserState
    Bs_Created = 0     ' New 直後、Init 前
    Bs_EnvReady = 1    ' Environment 生成完了、タブ追加可能
    Bs_Failed = 2      ' Environment 生成失敗
End Enum

' --- 非同期完了待ちのタイムアウト秒 (既存 StartWebView2 と同値) ---
Private Const READY_TIMEOUT As Single = 10#

' ============================================================
' フィールド
' ============================================================
Private m_env      As Wv2Environment   ' 1 個、全タブ共有
Private m_tabs     As Collection       ' Wv2Pane を格納 (1-based)
Private m_hwndHost As LongPtr          ' 親フレーム HWND (全タブ共通)
Private m_state    As BrowserState

' 第9.10c-1 段階で追加: アクティブタブの index (1-based、0 = アクティブなし)。
'   ActivateTab が更新する唯一の状態。ActivePane はこれを参照する。
Private m_activeIndex As Long


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    Set m_tabs = New Collection
    Set m_env = Nothing
    m_hwndHost = 0
    m_state = Bs_Created
    m_activeIndex = 0
End Sub


' ============================================================
' Init ? Environment を 1 個生成し Ready まで待つ
'   hwndHost   : 全タブが載る親フレームの HWND
'   userDataPath: WebView2 の userDataFolder
'   戻り値      : Environment が Ready になれば True
' ============================================================
Public Function Init( _
    ByVal hwndHost As LongPtr, _
    ByVal userDataPath As String) As Boolean

    If m_state <> Bs_Created Then
        Debug.Print "Wv2Browser.Init: 二重呼び出し (state=" & m_state & ")"
        Init = (m_state = Bs_EnvReady)
        Exit Function
    End If

    m_hwndHost = hwndHost

    ' --- Environment 生成 ---
    Set m_env = New Wv2Environment
    m_env.Init userDataPath

    ' --- Ready 待ち (待機ループは本クラス内部に閉じ込め = 論点① a) ---
    If Not WaitEnvReady() Then
        Debug.Print "Wv2Browser.Init: Environment 生成失敗 (state=" & _
                    m_env.state & ", lastError=&H" & Hex(m_env.LastError) & ")"
        m_state = Bs_Failed
        Exit Function
    End If

    Debug.Print "Wv2Browser.Init: Environment 生成 OK (pEnv=" & _
                m_env.EnvironmentPtr & ")"
    m_state = Bs_EnvReady
    Init = True
End Function


' ============================================================
' AddTab ? 器だけのタブを 1 個追加する (Navigate しない)
'   新 Wv2Pane を生成し、共有 Environment から Controller を生成、
'   Ready まで待って Bounds を設定し、Collection に追加する。
'   戻り値: 生成された Wv2Pane (失敗時 Nothing)
' ============================================================
Public Function AddTab() As Wv2Pane
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.AddTab: Environment が未 Ready (state=" & m_state & ")"
        Exit Function
    End If

    ' --- 1. Pane を生成し、共有 Environment から Controller を生成 ---
    Dim pane As Wv2Pane
    Set pane = New Wv2Pane
    pane.Init m_env, m_hwndHost

    ' --- 2. Controller Ready 待ち (内部に閉じ込め) ---
    If Not WaitPaneReady(pane) Then
        If pane.IsFailed Then
            Debug.Print "Wv2Browser.AddTab: Controller 生成失敗 (state=" & _
                        pane.state & ", lastError=&H" & Hex(pane.LastError) & ")"
        Else
            Debug.Print "Wv2Browser.AddTab: Controller 生成タイムアウト (state=" & _
                        pane.state & ")"
        End If
        Exit Function
    End If
    Debug.Print "Wv2Browser.AddTab: Controller 生成 OK (pCtrl=" & pane.CtrlPtr & ")"

    ' --- 3. Bounds を親フレームのクライアント領域に合わせる ---
    Dim rc As RECT
    GetClientRect m_hwndHost, rc
    Dim hr As Long
    hr = pane.Ctrl_PutBounds(rc)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTab: Ctrl_PutBounds 失敗 hr=&H" & Hex(hr)
        ' Bounds 失敗は致命ではないので続行 (タブ自体は生成できている)
    End If

    ' --- 4. ICoreWebView2 (View) を取得 ---
    '   ★重要★ Controller 生成だけでは m_pView = 0 のまま。View_Navigate /
    '   View_NavigateToString は m_pView が必要なので、ここで必ず取得する。
    '   (既存 StartWebView2 のステップ 8 を踏襲。これを抜くと Navigate 系が
    '    E_POINTER=&H80004003 で失敗する。)
    hr = pane.Ctrl_GetCoreWebView2()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTab: Ctrl_GetCoreWebView2 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    If pane.ViewPtr = 0 Then
        Debug.Print "Wv2Browser.AddTab: ViewPtr が 0 (取得成功だが値が空)"
        Exit Function
    End If
    Debug.Print "Wv2Browser.AddTab: View 取得 OK (pView=" & pane.ViewPtr & ")"

    ' --- 5. 標準ナビゲーションハンドラを自動登録 (案⑩-A、第9.10c-2) ---
    '   基盤的なイベント (NavigationStarting/Completed/DocumentTitleChanged/
    '   NewWindowRequested) は全タブで自動登録する。フォーム側でこれを
    '   受けるかどうかは利用者の自由 (Wv2Pane の OnXxx が Debug.Print するだけ)。
    '   ★方針変更トリガー★ イベント受信のオーバーヘッドで動作が重くなる場合は、
    '     この自動登録を案⑩-C (ヘルパーを利用者が明示的に呼ぶ) に変更する。
    RegisterStandardHandlers pane

    ' --- 6. Collection に追加 ---
    m_tabs.Add pane
    Debug.Print "Wv2Browser.AddTab: タブ追加 OK (現在のタブ数=" & m_tabs.count & ")"

    ' --- 7. 追加したタブを自動でアクティブ化 (論点⑦ a) ---
    '   可視制御は自前で行わず ActivateTab に委譲する (責務分離)。
    '   これにより「可視制御のロジックは ActivateTab の 1 箇所だけ」を保つ。
    ActivateTab m_tabs.count

    Set AddTab = pane
End Function


' ============================================================
' RegisterStandardHandlers ? タブ 1 個に標準ナビゲーションハンドラを登録する
'   第9.10c-2 段階で追加 (案⑩-A)。既存 StartWebView2 のステップ 10?13 を
'   逐語的に踏襲。ハンドラ登録の責務をこの 1 箇所に集約する (設計原則 57 の精神)。
'
'   ★設計原則 45/52 の対象★ これらの View_AddXxx は内部で COM オブジェクトを
'   WebView2 に渡すハンドラ系メソッド。Wv2Pane 側の実装が VTABLE_OBJ_OFFSET を
'   通す前提で正しく書かれているので、ここでは Wv2Pane の公開メソッドを
'   呼ぶだけでよい (自前でスロットを触らない)。
' ============================================================
Private Sub RegisterStandardHandlers(ByVal pane As Wv2Pane)
    Dim hr As Long

    hr = pane.View_AddNavigationStarting()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationStarting 失敗 hr=&H" & Hex(hr)

    hr = pane.View_AddNavigationCompleted()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NavigationCompleted 失敗 hr=&H" & Hex(hr)

    hr = pane.View_AddDocumentTitleChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: DocumentTitleChanged 失敗 hr=&H" & Hex(hr)

    hr = pane.View_AddNewWindowRequested()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: NewWindowRequested 失敗 hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.RegisterStandardHandlers: 標準ハンドラ 4 種 登録完了"
End Sub


' ============================================================
' AddTabWithUrl ? タブを追加し、URL へ Navigate する
' ============================================================
Public Function AddTabWithUrl(ByVal url As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrl: View_Navigate 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrl: Navigate OK (" & url & ")"
    End If

    Set AddTabWithUrl = pane
End Function


' ============================================================
' AddTabWithHtml ? タブを追加し、HTML 文字列をその場で表示する
'
'   ★仕様事実 6★ NavigateToString で表示したページの origin は about:blank
'   (= "null" origin) になる。この origin では history.pushState /
'   replaceState が SecurityError で弾かれる (Web 標準仕様)。
'   pushState/replaceState を伴う本格的な SPA テストは、9.10d で http(s)
'   origin のタブ (AddTabWithUrl) を使って行う。
' ============================================================
Public Function AddTabWithHtml(ByVal htmlContent As String) As Wv2Pane
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_NavigateToString(htmlContent)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithHtml: View_NavigateToString 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithHtml: NavigateToString OK (" & _
                    Len(htmlContent) & " 文字)"
    End If

    Set AddTabWithHtml = pane
End Function


' ============================================================
' TabCount ? 現在のタブ数
' ============================================================
Public Property Get TabCount() As Long
    TabCount = m_tabs.count
End Property


' ============================================================
' ActivePane ? アクティブタブの Wv2Pane を返す
'   第9.10c-1 で activeIndex ベースに正式化 (論点④の正式化)。
'   m_activeIndex (1-based、0 = なし) が指すタブを返す。
' ============================================================
Public Property Get ActivePane() As Wv2Pane
    If m_activeIndex < 1 Or m_activeIndex > m_tabs.count Then
        Debug.Print "Wv2Browser.ActivePane: アクティブタブなし (activeIndex=" & _
                    m_activeIndex & ", tabCount=" & m_tabs.count & ")"
        Set ActivePane = Nothing
        Exit Property
    End If
    Set ActivePane = m_tabs.Item(m_activeIndex)
End Property


' ============================================================
' ActiveIndex ? アクティブタブの index (1-based、0 = なし)
' ============================================================
Public Property Get ActiveIndex() As Long
    ActiveIndex = m_activeIndex
End Property


' ============================================================
' ActivateTab ? 指定 index のタブだけを可視にし、他を不可視にする
'   第9.10c-1 段階で追加。タブ切替の唯一の入口 (single source of truth)。
'
'   引数: index (1-based)
'   戻り値: 成功で True
'
'   実装方針 (論点⑥ A 案): 全タブをループし、index のタブだけ
'   Ctrl_PutIsVisible(True)、他を False にする。愚直だが確実で、
'   「全部不可視」等の穴が出ない。タブ数は現実的に少ないので性能問題なし。
'
'   ★責務★ このメソッドは「可視制御 + activeIndex 更新」だけを担う。
'   タブの追加・削除はしない。AddTab 等はこのメソッドを呼ぶ (委譲) ことで、
'   可視制御ロジックを 1 箇所に集約する。
' ============================================================
Public Function ActivateTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.count Then
        Debug.Print "Wv2Browser.ActivateTab: index 範囲外 (index=" & index & _
                    ", tabCount=" & m_tabs.count & ")"
        Exit Function
    End If

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.count
        Set pane = m_tabs.Item(i)
        hr = pane.Ctrl_PutIsVisible(i = index)
        If hr <> 0 Then
            Debug.Print "Wv2Browser.ActivateTab: タブ" & i & " の PutIsVisible 失敗 hr=&H" & Hex(hr)
        End If
    Next i

    m_activeIndex = index
    Debug.Print "Wv2Browser.ActivateTab: タブ" & index & " をアクティブ化 (可視)"
    ActivateTab = True
End Function


' ============================================================
' ResizeAll ? 全タブの Bounds を指定矩形に再発行する
'   第9.10c-1 段階で追加。UserForm_Resize から 1 本で呼ぶ (論点⑧ a)。
'
'   不可視タブにも Bounds を再発行しておく理由: 後でアクティブ化された時に
'   サイズがずれないようにするため。可視/不可視に関わらず全タブ更新する。
' ============================================================
Public Sub ResizeAll(ByRef rc As RECT)
    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.count
        Set pane = m_tabs.Item(i)
        If pane.IsReady Then
            hr = pane.Ctrl_PutBounds(rc)
            If hr <> 0 Then
                Debug.Print "Wv2Browser.ResizeAll: タブ" & i & " の PutBounds 失敗 hr=&H" & Hex(hr)
            End If
        End If
    Next i
End Sub


' ============================================================
' PaneAt ? index (1-based) でタブを取得する (検証用)
' ============================================================
Public Property Get PaneAt(ByVal index As Long) As Wv2Pane
    If index < 1 Or index > m_tabs.count Then
        Set PaneAt = Nothing
        Exit Property
    End If
    Set PaneAt = m_tabs.Item(index)
End Property


' ============================================================
' state ? 現在の Browser 状態
' ============================================================
Public Property Get state() As BrowserState
    state = m_state
End Property


' ============================================================
' 非同期完了待ちヘルパー (private、待機ロジックを本クラスに閉じ込める)
' ============================================================
Private Function WaitEnvReady() As Boolean
    Dim t As Single
    t = Timer
    Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And ((Timer - t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitEnvReady = m_env.IsReady
End Function

Private Function WaitPaneReady(ByVal pane As Wv2Pane) As Boolean
    Dim t As Single
    t = Timer
    Do While (Not pane.IsReady) And (Not pane.IsFailed) And ((Timer - t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitPaneReady = pane.IsReady
End Function


' ============================================================
' Class_Terminate ? 全タブ解放 → Environment 解放 (設計原則 43: 子から親へ)
' ============================================================
Private Sub Class_Terminate()
    Debug.Print "Wv2Browser.Terminate 開始 (タブ数=" & m_tabs.count & ")"

    ' --- 1. 全 Pane を解放 (各 Pane の Class_Terminate が View 系を順に解放) ---
    '   Collection から取り除くと参照が切れ、参照カウント 0 で
    '   Wv2Pane.Class_Terminate が走る。末尾から順に取り除く。
    Dim i As Long
    For i = m_tabs.count To 1 Step -1
        m_tabs.Remove i
    Next i
    Set m_tabs = Nothing

    ' --- 2. Environment を解放 (全 Pane より後 = 親を最後に) ---
    Set m_env = Nothing

    m_hwndHost = 0
    m_state = Bs_Created
    Debug.Print "Wv2Browser.Terminate 完了"
End Sub


