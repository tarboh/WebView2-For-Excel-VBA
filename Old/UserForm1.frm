VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   8850.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12375
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'UserForm UserForm1.frm
'
' 第9.9c 段階 (補完群一括処理)
'   ・新規 Public Sub: OnExecuteScriptResult
'     Wv2Pane.OnExecuteScriptCompleted から直接呼ばれる結果通知の受け口。
'     現状の第一実装はログ出力のみ (将来 SPA 連携時に Form 上のコントロールへ
'     反映する形に拡張可能)。
'   ・GetActivePane プロパティに Nothing チェック警告を追加
'     m_pane が Nothing の場合、Debug.Print で警告を出してから Nothing を返す
'     (Immediate 検証で「オブジェクト変数が設定されていません」エラーで
'      迷子になるのを防ぐ、案 G1-a)。
'
' 第9.9b 段階 (SPA 対応)
'   ・新規 Public Sub: StartWebView2_SpaTest
'     SPA テストページ (BuildSpaTestHtml) を表示する生成チェーン。
'     7 種類のハンドラを登録:
'       NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'       WebMessageReceived / NewWindowRequested /
'       HistoryChanged (第9.9b 基底) / DOMContentLoaded (第9.9b 派生 _2)
'   ・StartWebView2 / StartWebView2_LocalHtml は本体変更なし
'   ・実機検証は標準モジュールの Test_SpaTest_Real から呼ぶ
'
' 第9.9a 段階 (ICoreWebView2_2 取り出し基盤)
'   ・UserForm 側はノータッチ。検証は Immediate から
'     ?UserForm1.GetActivePane.View2_Ensure で実施した
'
' 第9.8b 段階 (View 基底 IF 網羅 + Settings 系 + Ctrl_Close)
'   ・新規 Public プロパティ: GetActivePane (Immediate 検証用、m_pane を露出)
'   ・StartWebView2 / StartWebView2_LocalHtml 本体は変更なし
'   ・Wv2Pane 側で追加された View_/Settings_/Ctrl_ メソッドの叩き方は
'     GetActivePane プロパティのコメント (本ファイル内) を参照
'
' 第9.8a 段階 (クラス名正式固定、改名のみ、ロジック変更なし)
'   ・モジュール変数 m_ctrl → m_pane に改名 (Wv2Pane への型変更と整合)
'   ・As Class8 → As Wv2Environment / As Class9 → As Wv2Pane に型変更
'   ・New Class8 / New Class9 → New Wv2Environment / New Wv2Pane に変更
'   ・ヘッダ・本文の Class7 / Class8 / Class9 への言及を
'     ComCallbackHandler / Wv2Environment / Wv2Pane に置換
'   ・古いヘッダコメント (9.4c → 9.6a の改名一覧、TabWebView2 移行時の
'     役割変化説明等) は本ヘッダから削除し、詳細は開発メモ.md 参照
'
' 第9.7b 段階で追加済みの機構 (継続):
'   ・モジュール変数: hWnd_Form, m_gap{Left,Right,Top,Bottom}, m_resizeReady
'   ・UserForm_Initialize: WS_THICKFRAME 付与 + gap 保存 + フラグ立て
'   ・UserForm_Resize: gap を再現する形で Frame1 をリサイズ、その後
'     Ctrl_PutBounds 再発行で WebView2 レンダリング領域を追従
'   ・Me.Width / Me.Height の代入は撤廃済み (設計原則 36)
'
' 第9.7b で確立した設計原則 (再掲、開発メモ.md 参照):
'   ・設計原則 34: UserForm リサイズ追従は「隙間 (gap) を保存して再現」方式
'   ・設計原則 35: UserForm_Resize は Initialize 中の偽イベントに注意
'                  (m_resizeReady フラグで早期 return)
'   ・設計原則 36: UserForm のサイズはデザイナで決め、Initialize では
'                  Me.Width / Me.Height に代入しない
'   ・設計原則 37: WebView2 のリサイズ追従は UserForm_Resize から
'                  Ctrl_PutBounds 発行だけで滑らかに動作 (DWM 任せ)
'
' VBE での手作業セットアップ:
'   1. プロジェクトに UserForm を 1 個追加 (名前は UserForm1)
'   2. ツールボックスから Frame コントロールを 1 個ドロップ (名前は Frame1)
'      (サイズは適当でよい、デザイナで決めた値が gap 計算の元になる)
'   3. 以下のコードを UserForm のコードモジュールに貼り付け
'
' UserForm の役割:
'   ・hWnd_Frame に Frame1 のウィンドウハンドルを保持
'   ・hWnd_Form に UserForm 自身のウィンドウハンドルを保持 (リサイズ用)
'   ・m_gap{L,R,T,B} に Frame1 と UserForm の四方の隙間を保持
'   ・StartWebView2 で Bing 表示用の生成チェーン (4 種類のハンドラ:
'     NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'     NewWindowRequested)
'   ・StartWebView2_LocalHtml でテスト HTML 表示用の生成チェーン (5 種類:
'     上記 + WebMessageReceived)
'   ・StartWebView2_SpaTest で SPA テスト用の生成チェーン (7 種類:
'     上記 5 種類 + HistoryChanged + DOMContentLoaded、第9.9b)
'   ・SendMessageToJS で外部から VBA → JS 送信を撃てる
'   ・IsPageLoaded プロパティで NavigationCompleted を観測
'   ・UserForm_Terminate で m_pane → m_env を Nothing に倒し、
'     ComRelease チェーンを起動 (Wv2Pane.Class_Terminate 内で
'     全永続ハンドラの remove_ も走る)
'
' 将来 (TabWebView2 クラス導入後、9.10 以降) の役割:
'   ・m_env / m_pane は m_tab 1 個に集約され、UserForm のフィールドが減る
'   ・IsPageLoaded は m_tab.IsPageLoaded へ移管
'   ・StartWebView2 / StartWebView2_LocalHtml / StartWebView2_SpaTest は
'     m_tab.Init / m_tab.LoadHtml 等に置き換わる


Option Explicit

' WebView2 を配置するターゲットとなるウィンドウハンドル
Private hWnd_Frame As LongPtr

' --- 第9.7b 段階で追加 (マウスリサイズ用) ---
'   UserForm 自身の HWND (GA_ROOT で Frame1.HWND から遡って取得)
'   ・WS_THICKFRAME を付与してマウスリサイズ可能化する対象
'   ・通常の UserForm_Resize の中で SetWindowLongPtrW を再発行する必要は
'     ないので、Initialize 時に 1 回取得して保持するだけ
Private hWnd_Form    As LongPtr

' --- 第9.7b 段階で追加 (UserForm と Frame1 の隙間、設計原則 34) ---
'   Initialize 時のデザイン状態で「Frame1 と UserForm の四方の隙間」を
'   ポイント単位で保存しておく。UserForm_Resize ではこれを再現する形で
'   Frame1 をリサイズすることで、Frame1 以外のコントロール (将来追加予定)
'   の位置を破壊せずに済む。
'
'   ポイント単位なのは Frame1.Left/Top/Width/Height や Me.InsideWidth/Height
'   と単位を揃えるため (= MSForms 標準座標系)。WebView2 ICoreWebView2Controller
'   の put_Bounds はピクセル単位だが、そちらは GetClientRect (ピクセル) で
'   取り直すので変換式は登場しない。
Private m_gapLeft    As Single
Private m_gapRight   As Single
Private m_gapTop     As Single
Private m_gapBottom  As Single

' --- 第9.7b 段階で追加 (Initialize 中の偽 Resize イベント対策、設計原則 35) ---
'   VBA UserForm は Me.Width/Me.Height を設定すると Resize イベントが
'   発火するが、その時点ではまだ m_gap* が未確定。フラグで早期 return する。
Private m_resizeReady As Boolean

' 第9.3b 暫定: UserForm が Wv2Environment と Wv2Pane を直接保持する。
' 将来 Wv2Browser クラスが入ったら m_browser As Wv2Browser に集約予定 (9.10c)。
Private m_env  As Wv2Environment
Private m_pane As Wv2Pane

' 第9.10b 検証専用: Wv2Browser 骨格の実機確認に使う使い捨てフィールド。
'   正式な UserForm 接続 (m_env + m_pane → m_browser への集約) は 9.10c で行う。
'   このフィールドと StartBrowserTest は 9.10c で正式接続する際に整理・削除する。
Private m_browser As Wv2Browser


' ============================================================
' UserForm_Initialize
'   UserForm のサイズを設定し、Frame1 の HWND を取得して保持する。
'   Frame1 のサイズはデザイン時に UserForm のクライアント領域いっぱいに
'   合わせておく前提 (= Width=792, Height=572 程度を目安)。
'   生成チェーンの起動は StartWebView2 を別途呼ぶ方式。
'
'   第9.7b 段階で追加:
'     ・UserForm 自身の HWND を GetAncestor(hWnd_Frame, GA_ROOT) で取得
'     ・WS_THICKFRAME を付与してマウスリサイズ可能化
'     ・Frame1 と UserForm の四方の隙間 (gap) を保存
'     ・m_resizeReady フラグを True にして、これ以降の Resize イベントを有効化
'   注意: Me.Width / Me.Height の代入で Resize イベントが発火するが、
'         m_resizeReady = False なので UserForm_Resize は早期 return する。
' ============================================================
Private Sub UserForm_Initialize()
    ' --- 第9.7b 修正: Me.Width / Me.Height の代入を撤廃 ---
    '   起動時の UserForm サイズはデザイン時に設定された値を尊重する。
    '   デザイン時に決めたサイズを Initialize で上書きすると、
    '   ・コードがデザイナを上書きする = 設計時の整合性が崩れる
    '   ・Frame1 はデザイン時のままなので gap 計算が壊れる
    '   ・「デザイン時のレイアウトを堅持」する設計原則 34 と矛盾する
    '   起動時に異なるサイズで表示したい場合はデザイナで UserForm のサイズ
    '   (および Frame1 のサイズ・位置) を併せて調整する。

    hWnd_Frame = Frame1.[_GethWnd]
    Debug.Print "UserForm1.Initialize: Frame1.[_GethWnd] = " & hWnd_Frame

    ' --- 第9.7b: UserForm 自身の HWND を取得 ---
    hWnd_Form = GetAncestor(hWnd_Frame, GA_ROOT)
    Debug.Print "UserForm1.Initialize: UserForm hWnd = " & hWnd_Form

    ' --- 第9.7b: WS_THICKFRAME を付与してマウスリサイズ可能に ---
    Dim style As LongPtr
    style = GetWindowLongPtrW(hWnd_Form, GWL_STYLE)
    SetWindowLongPtrW hWnd_Form, GWL_STYLE, style Or WS_THICKFRAME
    SetWindowPos hWnd_Form, 0, 0, 0, 0, 0, _
                 SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER Or SWP_FRAMECHANGED
    Debug.Print "UserForm1.Initialize: WS_THICKFRAME 付与済み (style was &H" & _
                Hex(style) & ")"

    ' --- 第9.7b: Frame1 と UserForm の四方の隙間を保存 (設計原則 34) ---
    m_gapLeft = Frame1.Left
    m_gapTop = Frame1.Top
    m_gapRight = Me.InsideWidth - (Frame1.Left + Frame1.width)
    m_gapBottom = Me.InsideHeight - (Frame1.Top + Frame1.Height)
    Debug.Print "UserForm1.Initialize: gaps L=" & m_gapLeft & _
                " T=" & m_gapTop & _
                " R=" & m_gapRight & _
                " B=" & m_gapBottom

    ' --- 第9.7b: ここ以降の Resize イベントを有効化 (設計原則 35) ---
    m_resizeReady = True
End Sub


' ============================================================
' UserForm_Resize (第9.7b 段階で新規)
'
'   マウスでフォームをリサイズしたときに呼ばれる。
'   設計原則 34 に従い、保存しておいた gap を再現する形で Frame1 を
'   リサイズし、その後 GetClientRect で取得したピクセル単位の rc を
'   Ctrl_PutBounds に渡して WebView2 レンダリング領域を追従させる。
'
'   設計原則 35 に従い、Initialize 中の偽 Resize イベントは
'   m_resizeReady フラグで早期 return する。
'
'   WebView2 controller (m_pane) がまだ準備中の場合 (Initialize 直後で
'   StartWebView2 が呼ばれる前、あるいは StartWebView2 の途中でリサイズ
'   イベントが入った場合) は Frame1 だけ追従させて Ctrl_PutBounds は
'   スキップする。
' ============================================================
Private Sub UserForm_Resize()
    If Not m_resizeReady Then Exit Sub

    ' --- Frame1 を「隙間を保ったまま」リサイズ ---
    Frame1.width = Me.InsideWidth - m_gapLeft - m_gapRight
    Frame1.Height = Me.InsideHeight - m_gapTop - m_gapBottom

    ' --- Wv2Browser がまだ無ければ Frame1 リサイズのみで終了 (第9.10c-2) ---
    If m_browser Is Nothing Then Exit Sub

    ' --- Frame1 のクライアント領域をピクセル単位で取得し、全タブに再発行 ---
    '   ResizeAll が全タブ (可視/不可視問わず) に Ctrl_PutBounds を再発行する
    '   (論点⑧ a)。不可視タブも更新しておくことで、後でアクティブ化された時に
    '   サイズがずれない。
    Dim rc As RECT
    GetClientRect hWnd_Frame, rc
    m_browser.ResizeAll rc
End Sub


' ============================================================
' StartWebView2
'   Test_Controller_Real から呼ばれる生成チェーンエントリ。
'
'   流れ:
'     1. Wv2Environment を New + Init
'        WebView2Loader が同期コールバックすれば即 Es_Ready になる
'     2. env.IsReady まで DoEvents で待つ (上限 10 秒)
'     3. Wv2Pane を New + Init(env, hWnd_Frame)
'        Environment->CreateCoreWebView2Controller が dcf 経由で呼ばれる
'     4. ctrl.IsReady まで DoEvents で待つ (上限 10 秒)
'     5. 偵察ログを Debug.Print
'
'   失敗時はその場で Debug.Print して return。クリーンアップは
'   UserForm_Terminate に任せる。
' ============================================================
Public Sub StartWebView2()
    ' ============================================================
    ' 第9.10c-2 段階: Wv2Browser ベースに書き換え。
    '   UserForm のフィールドが m_env + m_pane から m_browser 1 個に集約された
    '   (第9.10 段階の当初構想の実現)。
    '   Environment 生成・タブ生成・View 取得・Navigate・標準ハンドラ登録・
    '   アクティブ化はすべて Wv2Browser の内部で行われる。
    '   旧本体 (m_env + m_pane を直接操作する版) は下部にコメントアウトで保存。
    ' ============================================================
    Debug.Print "UserForm1.StartWebView2 開始 (9.10c-2 / Wv2Browser ベース)"

    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, "C:\Temp\VBA_WebView2") Then
        Debug.Print "UserForm1.StartWebView2: Wv2Browser.Init 失敗 (state=" & _
                    m_browser.state & ")"
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2: Wv2Browser.Init OK (Environment 共有)"

    ' タブ 1 を追加して Bing を開く (AddTab 内で View 取得・標準ハンドラ登録・
    ' アクティブ化まで自動実行される)。
    Dim pane As Wv2Pane
    Set pane = m_browser.AddTabWithUrl("https://www.bing.com")
    If pane Is Nothing Then
        Debug.Print "UserForm1.StartWebView2: タブ 1 の追加に失敗"
        Exit Sub
    End If

    Debug.Print "UserForm1.StartWebView2: 完了 (TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex & ")"
    Debug.Print "  Bing が表示されたら成功。リンククリックで各ハンドラが発火します。"
End Sub

' ============================================================
' 【第9.10c-2 でコメントアウト保存】旧 StartWebView2 (m_env + m_pane 直接操作版)
'   Wv2Browser 導入前の実装。リグレッション比較・参照用に残す。
'   有効化する場合はこのブロックのコメントを外し、上の新 StartWebView2 を
'   リネーム/無効化すること。
' ============================================================
' Public Sub StartWebView2_OLD_9_9c()
'
'     Debug.Print "UserForm1.StartWebView2 開始"
'
'     ' --- 1. Environment を生成 ---
'     Set m_env = New Wv2Environment
'     m_env.Init "C:\Temp\VBA_WebView2"
'
'     ' --- 2. env.IsReady を待つ ---
'     Dim t As Single
'     t = Timer
'     Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And ((Timer - t) < 10#)
'         DoEvents
'     Loop
'
'     If Not m_env.IsReady Then
'         Debug.Print "UserForm1.StartWebView2: Environment 生成失敗 (state=" & _
'                     m_env.state & ", lastError=&H" & Hex(m_env.LastError) & ")"
'         Exit Sub
'     End If
'     Debug.Print "UserForm1.StartWebView2: Environment 生成 OK (pEnv=" & _
'                 m_env.EnvironmentPtr & ")"
'
'     ' --- 3. Controller を生成 ---
'     Set m_pane = New Wv2Pane
'     m_pane.Init m_env, hWnd_Frame
'
'     ' --- 4. ctrl.IsReady を待つ ---
'     t = Timer
'     Do While (Not m_pane.IsReady) And (Not m_pane.IsFailed) And ((Timer - t) < 10#)
'         DoEvents
'     Loop
'
'     ' --- 5. 結果ログ ---
'     If Not m_pane.IsReady Then
'         If m_pane.IsFailed Then
'             Debug.Print "UserForm1.StartWebView2: Controller 生成失敗 (state=" & _
'                         m_pane.state & ", lastError=&H" & Hex(m_pane.LastError) & ")"
'         Else
'             Debug.Print "UserForm1.StartWebView2: Controller 生成タイムアウト (state=" & _
'                         m_pane.state & ")"
'         End If
'         Exit Sub
'     End If
'     Debug.Print "UserForm1.StartWebView2: Controller 生成 OK (pCtrl=" & _
'                 m_pane.CtrlPtr & ")"
'
'     ' --- 6. Frame1 のクライアント領域を取得 ---
'     Dim rc As RECT
'     GetClientRect hWnd_Frame, rc
'     Debug.Print "UserForm1.StartWebView2: Frame1 client rect = (" & _
'                 rc.Left & "," & rc.Top & "," & rc.Right & "," & rc.Bottom & ")"
'
'     ' --- 7. Controller のサイズ設定 ---
'     Dim hr As Long
'     hr = m_pane.Ctrl_PutBounds(rc)
'     If hr <> 0 Then
'         Debug.Print "UserForm1.StartWebView2: Ctrl_PutBounds 失敗 hr=&H" & Hex(hr)
'         Exit Sub
'     End If
'     Debug.Print "UserForm1.StartWebView2: Ctrl_PutBounds OK"
'
'     ' --- 8. ICoreWebView2 を取得 ---
'     hr = m_pane.Ctrl_GetCoreWebView2()
'     If hr <> 0 Then
'         Debug.Print "UserForm1.StartWebView2: Ctrl_GetCoreWebView2 失敗 hr=&H" & Hex(hr)
'         Exit Sub
'     End If
'     If m_pane.ViewPtr = 0 Then
'         Debug.Print "UserForm1.StartWebView2: ViewPtr が 0 (取得には成功したが値が空)"
'         Exit Sub
'     End If
'     Debug.Print "UserForm1.StartWebView2: Ctrl_GetCoreWebView2 OK (pView=" & _
'                 m_pane.ViewPtr & ")"
'
'     ' --- 9. Navigate ---
'     hr = m_pane.View_Navigate("https://www.bing.com")
'     If hr <> 0 Then
'         Debug.Print "UserForm1.StartWebView2: View_Navigate 失敗 hr=&H" & Hex(hr)
'         Exit Sub
'     End If
'     Debug.Print "UserForm1.StartWebView2: View_Navigate OK (Bing が表示されるはず)"
'
'     ' --- 10?13. 標準ハンドラ登録 (現在は Wv2Browser.RegisterStandardHandlers に移設) ---
'     m_pane.View_AddNavigationStarting
'     m_pane.View_AddNavigationCompleted
'     m_pane.View_AddDocumentTitleChanged
'     m_pane.View_AddNewWindowRequested
' End Sub




' ============================================================
' StartWebView2_LocalHtml (第9.4c 段階で新規、第9.6a で View_AddNewWindowRequested 追加)
'
'   Bing ではなくテスト用 HTML を NavigateToString で表示するバリエーション。
'   Test_MessageEcho_Real から呼ばれる。
'
'   流れ:
'     1. StartWebView2 と同じく Wv2Environment + Wv2Pane を生成
'     2. View_Navigate("about:blank") は呼ばない (NavigateToString で済む)
'        ※ NavigateToString はそれ自体が「最初のナビゲーション」になる。
'     3. View_AddNavigationStarting / View_AddNavigationCompleted /
'        View_AddDocumentTitleChanged / View_AddWebMessageReceived /
'        View_AddNewWindowRequested (第9.6a 新規) を登録
'     4. View_NavigateToString(BuildTestHtml) でテスト HTML を表示
'        → Wv2Pane.m_pageLoaded が False にリセットされ、NavigationCompleted を
'          受信した時点で True に戻る (Test 側はこれを待つ)
' ============================================================
Public Sub StartWebView2_LocalHtml()
    Debug.Print "UserForm1.StartWebView2_LocalHtml 開始"

    ' --- 1. Environment を生成 ---
    Set m_env = New Wv2Environment
    m_env.Init "C:\Temp\VBA_WebView2"

    Dim t As Single
    t = Timer
    Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And ((Timer - t) < 10#)
        DoEvents
    Loop

    If Not m_env.IsReady Then
        Debug.Print "UserForm1.StartWebView2_LocalHtml: Environment 生成失敗"
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_LocalHtml: Environment 生成 OK"

    ' --- 2. Controller を生成 ---
    Set m_pane = New Wv2Pane
    m_pane.Init m_env, hWnd_Frame

    t = Timer
    Do While (Not m_pane.IsReady) And (Not m_pane.IsFailed) And ((Timer - t) < 10#)
        DoEvents
    Loop

    If Not m_pane.IsReady Then
        Debug.Print "UserForm1.StartWebView2_LocalHtml: Controller 生成失敗"
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_LocalHtml: Controller 生成 OK"

    ' --- 3. Frame1 のクライアント領域を取得して Bounds 設定 ---
    Dim rc As RECT
    GetClientRect hWnd_Frame, rc
    m_pane.Ctrl_PutBounds rc
    Debug.Print "UserForm1.StartWebView2_LocalHtml: Ctrl_PutBounds OK"

    ' --- 4. ICoreWebView2 を取得 ---
    m_pane.Ctrl_GetCoreWebView2
    Debug.Print "UserForm1.StartWebView2_LocalHtml: Ctrl_GetCoreWebView2 OK"

    ' --- 5. 5 種類のハンドラを登録 (第9.6a で NewWindowRequested を追加) ---
    m_pane.View_AddNavigationStarting
    m_pane.View_AddNavigationCompleted
    m_pane.View_AddDocumentTitleChanged
    m_pane.View_AddWebMessageReceived
    m_pane.View_AddNewWindowRequested
    Debug.Print "UserForm1.StartWebView2_LocalHtml: 5 種類のハンドラを登録"

    ' --- 6. テスト HTML を NavigateToString で表示 ---
    Dim html As String
    html = BuildTestHtml()
    Dim hr As Long
    hr = m_pane.View_NavigateToString(html)
    If hr <> 0 Then
        Debug.Print "UserForm1.StartWebView2_LocalHtml: NavigateToString 失敗 hr=&H" & Hex(hr)
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_LocalHtml: NavigateToString OK"
End Sub


' ============================================================
' StartWebView2_SpaTest (第9.9b 段階で新規、SPA 対応の検証)
'
'   SPA テストページ (BuildSpaTestHtml) を表示する生成チェーン。
'   StartWebView2_LocalHtml とほぼ同形だが、以下が異なる:
'     - 登録ハンドラが 7 種類 (5 種類 + HistoryChanged + DOMContentLoaded)
'     - 表示する HTML が BuildSpaTestHtml (SPA テストページ)
'
'   登録ハンドラ:
'     1. View_AddNavigationStarting   (基底)
'     2. View_AddNavigationCompleted  (基底)
'     3. View_AddDocumentTitleChanged (基底)
'     4. View_AddWebMessageReceived   (基底、SPA 規約 postMessage を受ける)
'     5. View_AddNewWindowRequested   (基底)
'     6. View_AddHistoryChanged       (基底、第9.9b、pushState 等を検知)
'     7. View2_AddDOMContentLoaded    (派生 _2、第9.9b、初回 DOM 構築を検知)
'
'   ★ 7 番目で View2 (m_pView2) が初投入される ★
'     View2_AddDOMContentLoaded が内部で EnsureView2 を呼び、
'     ICoreWebView2_2 を取り出してから add_DOMContentLoaded を撃つ。
'
'   実機検証は標準モジュールの Test_SpaTest_Real から呼ぶ。
' ============================================================
Public Sub StartWebView2_SpaTest()
    Debug.Print "UserForm1.StartWebView2_SpaTest 開始"

    ' --- 1. Environment を生成 ---
    Set m_env = New Wv2Environment
    m_env.Init "C:\Temp\VBA_WebView2"

    Dim t As Single
    t = Timer
    Do While (Not m_env.IsReady) And (Not m_env.IsFailed) And ((Timer - t) < 10#)
        DoEvents
    Loop

    If Not m_env.IsReady Then
        Debug.Print "UserForm1.StartWebView2_SpaTest: Environment 生成失敗"
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_SpaTest: Environment 生成 OK"

    ' --- 2. Controller を生成 ---
    Set m_pane = New Wv2Pane
    m_pane.Init m_env, hWnd_Frame

    t = Timer
    Do While (Not m_pane.IsReady) And (Not m_pane.IsFailed) And ((Timer - t) < 10#)
        DoEvents
    Loop

    If Not m_pane.IsReady Then
        Debug.Print "UserForm1.StartWebView2_SpaTest: Controller 生成失敗"
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_SpaTest: Controller 生成 OK"

    ' --- 3. Frame1 のクライアント領域を取得して Bounds 設定 ---
    Dim rc As RECT
    GetClientRect hWnd_Frame, rc
    m_pane.Ctrl_PutBounds rc
    Debug.Print "UserForm1.StartWebView2_SpaTest: Ctrl_PutBounds OK"

    ' --- 4. ICoreWebView2 を取得 ---
    m_pane.Ctrl_GetCoreWebView2
    Debug.Print "UserForm1.StartWebView2_SpaTest: Ctrl_GetCoreWebView2 OK"

    ' --- 5. 7 種類のハンドラを登録 (第9.9b で History/DOMContentLoaded を追加) ---
    m_pane.View_AddNavigationStarting
    m_pane.View_AddNavigationCompleted
    m_pane.View_AddDocumentTitleChanged
    m_pane.View_AddWebMessageReceived
    m_pane.View_AddNewWindowRequested
    m_pane.View_AddHistoryChanged        ' 第9.9b (基底 IF、vtable 13)
    m_pane.View2_AddDOMContentLoaded     ' 第9.9b (派生 _2 IF、vtable 64、View2 初投入)
    Debug.Print "UserForm1.StartWebView2_SpaTest: 7 種類のハンドラを登録"

    ' --- 6. SPA テスト HTML を NavigateToString で表示 ---
    Dim html As String
    html = BuildSpaTestHtml()
    Dim hr As Long
    hr = m_pane.View_NavigateToString(html)
    If hr <> 0 Then
        Debug.Print "UserForm1.StartWebView2_SpaTest: NavigateToString 失敗 hr=&H" & Hex(hr)
        Exit Sub
    End If
    Debug.Print "UserForm1.StartWebView2_SpaTest: NavigateToString OK"
End Sub


' ============================================================
' SendMessageToJS (第9.4c 段階で新規)
'
'   外部 (Test_MessageEcho_Real 等) から VBA → JS の文字列送信を撃つための
'   公開メソッド。Wv2Pane.View_PostWebMessageAsString の薄いラッパ。
'
'   呼び出し前に IsPageLoaded で読み込み完了を確認しているのが望ましい
'   (まだリスナーが登録されていない時点で撃つとメッセージが消える)。
' ============================================================
Public Sub SendMessageToJS(ByVal text As String)
    If m_pane Is Nothing Then
        Debug.Print "UserForm1.SendMessageToJS: m_pane が Nothing"
        Exit Sub
    End If
    If Not m_pane.IsReady Then
        Debug.Print "UserForm1.SendMessageToJS: m_pane が Ready でない"
        Exit Sub
    End If

    Dim hr As Long
    hr = m_pane.View_PostWebMessageAsString(text)
    If hr <> 0 Then
        Debug.Print "UserForm1.SendMessageToJS: 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "UserForm1.SendMessageToJS OK: """ & text & """"
    End If
End Sub


' ============================================================
' IsPageLoaded (第9.4c 段階で新規)
'
'   現在表示中のページが NavigationCompleted まで進んでいるかを返す。
'   実体は m_pane.IsPageLoaded を委譲しているだけ (薄いラッパ)。
'
'   Test_MessageEcho_Real が「ページ読み込み完了を待ってから PostMessage」
'   の同期に使う。
'
'   暫定設計の注意:
'     初回ナビゲーション完了 1 回しか意味を持たない。SPA 対応 (pushState
'     による画面更新追跡など) は将来の包括クラス導入時に再設計する。
' ============================================================
Public Property Get IsPageLoaded() As Boolean
    If m_pane Is Nothing Then
        IsPageLoaded = False
    Else
        IsPageLoaded = m_pane.IsPageLoaded
    End If
End Property


' ============================================================
' GetActivePane プロパティ (第9.8b 段階で追加、Immediate 検証用、第9.9c で改善)
'
'   現在 UserForm が保持している Wv2Pane インスタンスへの参照を返す。
'   Immediate ウィンドウから以下のように Wv2Pane のメソッドを直接叩ける:
'
'     ?UserForm1.GetActivePane.View_GetSource()
'     ?UserForm1.GetActivePane.View_GetCanGoBack()
'     ?UserForm1.GetActivePane.View_GetBrowserProcessId()
'     UserForm1.GetActivePane.View_Reload
'     UserForm1.GetActivePane.View_Stop
'     UserForm1.GetActivePane.Settings_PutIsScriptEnabled False
'     UserForm1.GetActivePane.Settings_PutAreDefaultContextMenusEnabled False
'     UserForm1.GetActivePane.Settings_PutAreDevToolsEnabled False
'     UserForm1.GetActivePane.Ctrl_Close
'     UserForm1.GetActivePane.View_ExecuteScript "1+2"
'     UserForm1.GetActivePane.View_ExecuteScriptSafe "throw new Error('x')"  (9.9c)
'
'   第9.9c 改善 (案 G1-a):
'     m_pane が Nothing の場合に Debug.Print で警告を出すようにした。
'     これにより Immediate 検証で
'       ?UserForm1.GetActivePane.View_ExecuteScript "x"
'     を叩いたとき「オブジェクト変数が設定されていません」エラーで迷子になる
'     代わりに、「m_pane が Nothing です」と警告が出るようになる。
'     Nothing 自体は変わらず返すので、利用者が Is Nothing チェックする経路も
'     維持される。
'
'   将来 TabWebView2 が入ったら GetActivePane は m_tab.ActivePane 等の
'   薄いラッパに変わる予定 (= 「アクティブタブの Pane を取る」意味)。
' ============================================================
Public Property Get GetActivePane() As Wv2Pane
    ' 第9.10c-2: m_pane 直参照から m_browser.ActivePane 経由へ繋ぎ替え。
    '   「アクティブタブの Pane を取る」という本来の意味になった。
    If m_browser Is Nothing Then
        Debug.Print "[UserForm1.GetActivePane] m_browser が Nothing (UserForm が未起動か、" & _
                    "StartWebView2 系が未呼び出しか)"
        Set GetActivePane = Nothing
        Exit Property
    End If
    Set GetActivePane = m_browser.ActivePane
End Property


' ============================================================
' OnExecuteScriptResult (第9.9c 段階で新規、案 α 第二実装の受け口)
'
'   Wv2Pane.OnExecuteScriptCompleted から直接呼ばれる ExecuteScript 結果の通知口。
'   Wv2Pane が UserForm1 への参照を保持しているわけではなく、グローバル
'   UserForm 参照 (UserForm1.OnExecuteScriptResult ...) で呼ばれる
'   (循環参照なし、案 E1-a-i)。
'
'   引数:
'     callbackId : View_ExecuteScript が返した連番 ID
'     errorCode  : HRESULT (0 = S_OK)
'     resultJson : JS の戻り値を JSON エンコードした文字列
'                  View_ExecuteScriptSafe (9.9c F) 経由なら {"ok":true,...} or
'                  {"ok":false,...} の構造化 JSON が来る
'
'   現状の実装:
'     ログ出力のみ (Wv2Pane 内のログと併存、案 E2-X)。
'     将来 SPA 連携を本格化する際は、ここから Form 上のコントロール
'     (例: TextBox や ListBox) に結果を表示する形に拡張可能。
'
'   ★ エラー時の挙動 ★
'     Wv2Pane.OnExecuteScriptCompleted は本 Sub をエラー防御なしで呼ぶので、
'     本 Sub 内で例外を投げてはいけない (= COM コールバック経路を破壊する)。
'     ログだけに留め、複雑な処理は避ける。
' ============================================================
Public Sub OnExecuteScriptResult( _
    ByVal callbackId As Long, _
    ByVal errorCode As Long, _
    ByVal resultJson As String)

    Debug.Print "UserForm1.OnExecuteScriptResult: callbackId=" & callbackId & _
                ", errorCode=" & errorCode & _
                ", resultJson=" & resultJson
End Sub


' ============================================================
' FrameHwnd (第9.10b 検証専用アクセサ)
'   Frame1 のウィンドウハンドル (hWnd_Frame) を外部に公開する。
'   Wv2Browser.Init に渡す親フレーム HWND として使う。
'   UserForm_Initialize が走った後 (= Show 後) でないと 0 のままな点に注意。
' ============================================================
Public Property Get FrameHwnd() As LongPtr
    FrameHwnd = hWnd_Frame
End Property


' ============================================================
' StartBrowserTest (第9.10b 検証専用・使い捨て)
'   Wv2Browser 骨格の実機確認。Environment 1 個 + タブ 2 個
'   (URL タブ + HTML タブ) を生成し、各 Pane が Controller/View まで
'   到達することをログで確認する。
'
'   ★スコープ注意★ 9.10b では表示切替をしないので、2 タブは同じ Frame に
'   重なって生成され、視覚的には後から追加した HTML タブが手前に見える。
'   タブ切替 (IsVisible 制御) と正式な UserForm 接続は 9.10c で行う。
'
'   使い方 (イミディエイト):
'     UserForm1.Show vbModeless
'     UserForm1.StartBrowserTest
' ============================================================
Public Sub StartBrowserTest()
    Debug.Print String(60, "=")
    Debug.Print "UserForm1.StartBrowserTest 開始 (9.10b 検証)"
    Debug.Print String(60, "=")

    ' --- 1. Wv2Browser を生成し、Environment を初期化 ---
    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, "C:\Temp\VBA_WebView2") Then
        Debug.Print "StartBrowserTest: Wv2Browser.Init 失敗 (state=" & m_browser.state & ")"
        Exit Sub
    End If
    Debug.Print "StartBrowserTest: Wv2Browser.Init OK (state=" & m_browser.state & ")"

    ' --- 2. タブ 1: URL で Bing を開く ---
    Debug.Print String(40, "-")
    Debug.Print "StartBrowserTest: タブ1 (URL) を追加 ..."
    Dim tab1 As Wv2Pane
    Set tab1 = m_browser.AddTabWithUrl("https://www.bing.com/")
    If tab1 Is Nothing Then
        Debug.Print "StartBrowserTest: タブ1 の追加に失敗"
    Else
        Debug.Print "StartBrowserTest: タブ1 OK (CtrlPtr=" & tab1.CtrlPtr & _
                    ", ViewPtr=" & tab1.ViewPtr & ")"
    End If

    ' --- 3. タブ 2: HTML 直書き ---
    Debug.Print String(40, "-")
    Debug.Print "StartBrowserTest: タブ2 (HTML) を追加 ..."
    Dim html As String
    html = "<html><body style='font-family:sans-serif;background:#eef'>" & _
           "<h1>Wv2Browser タブ2</h1>" & _
           "<p>これは NavigateToString で表示した HTML です。</p>" & _
           "<p>(origin は about:blank = 仕様事実6)</p>" & _
           "</body></html>"
    Dim tab2 As Wv2Pane
    Set tab2 = m_browser.AddTabWithHtml(html)
    If tab2 Is Nothing Then
        Debug.Print "StartBrowserTest: タブ2 の追加に失敗"
    Else
        Debug.Print "StartBrowserTest: タブ2 OK (CtrlPtr=" & tab2.CtrlPtr & _
                    ", ViewPtr=" & tab2.ViewPtr & ")"
    End If

    ' --- 4. サマリ ---
    Debug.Print String(40, "-")
    Debug.Print "StartBrowserTest: 完了。TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex
    Dim ap As Wv2Pane
    Set ap = m_browser.ActivePane
    If Not ap Is Nothing Then
        Debug.Print "StartBrowserTest: ActivePane CtrlPtr=" & ap.CtrlPtr
    End If
    Debug.Print "9.10c-1: AddTab が自動で ActivateTab するので、最後に追加した"
    Debug.Print "         タブ2 (HTML) が可視、タブ1 (Bing) は不可視のはず。"
    Debug.Print "         → 手前に HTML ページが見えれば IsVisible 制御 OK。"
    Debug.Print "タブ切替の確認: イミディエイトで UserForm1.SwitchTab 1 / 2 を実行。"
    Debug.Print "フォームを閉じると全タブ + Environment がクリーンアップされます。"
    Debug.Print String(60, "=")
End Sub


' ============================================================
' SwitchTab (第9.10c-1 検証専用・使い捨て)
'   イミディエイトから UserForm1.SwitchTab 1 / 2 でタブを切り替える。
'   ActivateTab の実機確認用。
' ============================================================
Public Sub SwitchTab(ByVal index As Long)
    If m_browser Is Nothing Then
        Debug.Print "SwitchTab: m_browser が Nothing (先に StartBrowserTest を実行)"
        Exit Sub
    End If
    If m_browser.ActivateTab(index) Then
        Debug.Print "SwitchTab: タブ" & index & " へ切替 OK (ActiveIndex=" & _
                    m_browser.ActiveIndex & ")"
    Else
        Debug.Print "SwitchTab: タブ" & index & " への切替に失敗"
    End If
End Sub


' ============================================================
' UserForm_Terminate
'   フォームが閉じられたら Wv2Pane → Wv2Environment の順で解放。
'   各クラスの Class_Terminate で ComRelease チェーンが走り、
'   ICoreWebView2Controller / ICoreWebView2Environment の参照が解放される。
' ============================================================
Private Sub UserForm_Terminate()
    Debug.Print "UserForm1.Terminate 開始"

    ' 第9.10b 検証用: Wv2Browser を先に解放 (内部で全 Pane → Environment の順)
    Set m_browser = Nothing

    Set m_pane = Nothing   ' Wv2Pane.Class_Terminate → Controller::Release
    Set m_env = Nothing    ' Wv2Environment.Class_Terminate → Environment::Release

    Debug.Print "UserForm1.Terminate 完了"
End Sub
