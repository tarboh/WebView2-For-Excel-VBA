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
' 第9.18 段階 (設定タブ + AddHostObjectToScript による検索エンジン UI)
'   ── 9.17 で本体の口 (UseSearchEngine / CurrentBrowser) を揃えた続きとして、
'      タブバーに ? 設定ボタンを付け、クリックで検索エンジン選択画面を専用タブに
'      開く。JS→VBA の連携は postMessage コマンドではなく AddHostObjectToScript を
'      使い、設定画面の JS から VBA (Wv2SettingsBridge) を直接呼ばせる。
'
'   ・UserForm1 側の変更は 2 点のみ:
'       (1) 検索エンジンのプリセット定数 (SE_*) と名前解決を Wv2SettingsBridge へ
'           移設 (DRY)。設定画面と同じ解決経路を共有する。
'       (2) UseSearchEngine を一時ブリッジ生成 → SetEngine 委譲に書き換え。
'           イミディエイト (UserForm1.UseSearchEngine "bing") の挙動は 9.17 と同じ。
'
'   ・設定タブの生成・ブリッジ attach・設定画面 HTML は Wv2Browser の責務
'       (OpenSettingsTab / BuildSettingsHtml)。TabBar は ? クリックを settings
'       コマンドで Browser に伝えるだけ (疎結合維持)。CurrentBrowser は 9.17 のまま。
'
' 第9.17 段階 (検索エンジンの切替口 + 起動中 Browser 参照の公開)
'   ── 9.16 で「本体 Property (SearchEngine) だけ作り、UI と実機確認は別ステージに
'      切り出す」と保留した宿題を回収する。UI 部品 (ドロップダウン等) はまだ置かず、
'      「起動中の Browser を外から触る足場」と「プリセット名で切り替える最小の口」を
'      生やすところまで (案 A-1 + B-1)。GUI 組み込み (NavBar への選択 UI) は次段。
'
'   ・Public Property Get CurrentBrowser() As Wv2Browser を追加 (案 A-1)。
'       m_browser は Private WithEvents で外から触れないため、素の参照を返す口を
'       1 本だけ公開する。WithEvents はこちら側の宣言制約であって、Get で参照を
'       渡すのは問題ない。以後どの拡張でも「起動中 Browser を外から掴む」足場になる。
'       起動していない (m_browser Is Nothing) 間は Nothing を返す。
'
'   ・検索エンジンのプリセット定数 (SE_GOOGLE / SE_BING / SE_DUCKDUCKGO /
'       SE_YAHOO) をモジュール冒頭に追加。いずれも「クエリ前置型テンプレート」
'       (末尾にエンコード済み検索語を連結するだけ、9.16 の案1-A と同形)。
'
'   ・Public Sub UseSearchEngine(ByVal engineName As String) を追加 (案 B-1)。
'       "google"/"bing"/"duckduckgo"/"yahoo" (大小・前後空白は無視) を
'       プリセットのテンプレートに解決し、CurrentBrowser.SearchEngine へ Let する。
'       未知の名前は既定 (Google) に落として警告を Debug.Print (空ガードは
'       Browser 側 Property Let が担保)。Browser 未起動なら何もせず注意を出す。
'       ★UI ではなくロジックの口★ 実機では StartWebView2_Full 起動後に
'       イミディエイトで  UserForm1.UseSearchEngine "bing"  と打つだけで切り替わる。
'
'   ・実機確認が「参照の口が無い」で詰まっていた点 (9.16 の Help の但し書き) を
'       CurrentBrowser で解消。9.17 の検証は Wv2Tests.Test_9_17_* に集約。
'
' 第9.13 段階 (整理: JSON ヘルパー共通化 + 検証コードの撤去)
'   ・9.11a?9.12d で積み上げた使い捨ての検証 Sub を全撤去 (35 個)。
'     タブバーとナビバーが GUI で全機能を提供している以上、検証 Sub の役目は
'     終わっている。必要になれば再生成すればよく、記録は開発メモに残っている。
'   ・旧起動 Sub も撤去し、実運用の起動口を 2 つに絞った:
'       StartWebView2_Full … 3 タブ + タブバー + ナビバー (通常の起動口)
'       StartWebView2_Spa  … SPA 雛形 (9.10e で「1 つだけ残す」と決めたもの)
'     撤去: StartWebView2 / StartWebView2_WithTabBar / StartWebView2_WithTabBar_3
'   ・行数: 1456 → 700 行
'
'   ※Wv2TabBar / Wv2NavBar 側では JSON ヘルパー (JsonEscape / GetJsonStr /
'     GetJsonNum / BoolToJson) の重複を Wv2Json.bas に共通化した。
'     UserForm 自体は JSON を触らないので影響なし。
'
' 第9.12d 段階 (JS->VBA: 入力・クリックの有効化) ? 9.12 完結
'   ・コード変更は検証用のみ (本体ロジックは Wv2NavBar 側)。
'   ・Test_9_12d_Help / Test_9_12d_Inject* を追加。起動は StartWebView2_Full を流用。
'
' 第9.12c 段階 (VBA->JS: URL とボタン状態の表示)
'   ・コード変更は検証用のみ (本体ロジックは Wv2NavBar 側)。
'   ・Test_9_12c_Help を追加。起動は StartWebView2_Full を流用。
'
' 第9.12b 段階 (Wv2NavBar の骨格 + レイアウト 3 分割)
'   ・m_navBar フィールドを追加。Wv2TabBar と同じ扱い (生成・Resize・Shutdown)。
'   ・UserForm_Resize が 3 分割になる:
'       1 段目 = タブバー   (高さ TABBAR_HEIGHT_PX)
'       2 段目 = ナビバー   (高さ NAVBAR_HEIGHT_PX) ★9.12b で追加
'       3 段目 = コンテンツ (残り全部)
'     ★分割の計算はここ (UserForm) の責務★ Browser も TabBar も NavBar も
'     互いを知らない。UserForm が矩形を配って回る (疎結合)。
'   ・StartWebView2_Full を追加 (3 タブ + タブバー + ナビバー)。
'   ・9.12b ではナビバーは静的表示のみ (ボタン押せず、URL 欄も更新されない)。
'
' 第9.12a 段階 (アドレスバー/ナビボタンの土台 ? UI なし、イミディエイト検証)
'   ・Wv2Browser の新イベント TabUrlChanged を UserForm でも購読してログ出力。
'   ・アクティブタブ委譲メソッド (NavigateActive / GoBackActive / GoForwardActive /
'     ReloadActive / StopActive / GetActiveUrl 等) をイミディエイトから検証する。
'   ・9.12b で Wv2NavBar (WebView2 製ナビバー) を載せる。ここはその土台。
'
' 第9.11d 段階 (JS->VBA クリック有効化の検証)
'   ・コード変更は検証用のみ (本体ロジックは Wv2TabBar 側)。
'   ・Test_9_11d_Help: 手動クリック検証の手順を表示。
'   ・Test_9_11d_Inject*: JS を介さず OnPaneWebMessage に直接コマンド文字列を
'     流し込む。パーサ (GetJsonStr/GetJsonNum) とコマンド処理だけを切り分けて
'     検証できる (クリック操作なしでイミディエイトから叩ける)。
'   ・タブバーの再利用のため StartWebView2_WithTabBar_3 を 9.11d でも使う。
'
' 第9.11c 段階 (VBA->JS タブ一覧描画の検証)
'   ・コード変更は検証用のみ (本体ロジックは Wv2TabBar / Wv2Pane 側)。
'   ・StartWebView2_WithTabBar_3 を追加: bing/example/msn の 3 タブ + タブバー。
'     タブバー JS が 3 タブ分を描画し、各タブのタイトルが反映されるのを見る。
'   ・Test_9_11c_Help / Test_9_11c_AddTab / Test_9_11c_CloseActive を追加。
'     AddTab/CloseTab の後にタブバー描画が全体同期で追従することを確認する。
'
' 第9.11b 段階 (WebView2 製タブバー Wv2TabBar の統合)
'   ・新規フィールド: m_tabBar As Wv2TabBar (タブバー UI。使わなければ生成
'     しないだけ = 疎結合)。
'   ・UserForm_Resize を 2 分割対応に: m_tabBar があれば上部にタブバー帯
'     (高さ TABBAR_HEIGHT_PX) を確保し、コンテンツ領域をその下にずらす。
'     タブバー帯の矩形は m_tabBar.Resize、コンテンツ領域は m_browser.ResizeAll。
'     m_tabBar が無ければ従来通り全域をコンテンツに (★Browser はタブバーを
'     知らない★ 分割計算は UserForm 側の責務)。
'   ・UserForm_Terminate に m_tabBar.Shutdown を追加 (m_browser.Shutdown より前)。
'   ・検証用: StartWebView2_WithTabBar / Test_9_11b_* を追加。
'
' 第9.11a 段階 (疎結合タブバー用イベント受信)
'   ・m_browser を WithEvents 化し、Wv2Browser の 4 イベント (TabAdded /
'     TabClosed / ActiveChanged / TabTitleChanged) を受信するハンドラを追加。
'   ・検証用: Test_9_11a_* 群 (Setup / CloseActiveMiddle / CloseActiveLast /
'     CloseBeforeActive / CloseAll / Help) と内部ヘルパ DumpState を追加。
'
' 第9.10e 段階 (NewWindowRequested 新タブ化)
'   ・UserForm_Terminate に m_browser.Shutdown を追加 (循環参照 Browser?Pane を
'     断ってから解放。案Y。UserForm 側の実装コストはこの 1 行のみ、設計原則 58)
'   ・第9.10 整理: 検証 Sub 群 (SpaSmoke / SpaMapping / SpaFull 等) を撤去。
'     SPA 起動は StartWebView2_Spa (雛形) に集約、m_env/m_pane フィールドも撤去した
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
Private Const TABBAR_HEIGHT_PX As Long = 36   ' 第9.11b: タブバー帯の高さ (px)
Private Const NAVBAR_HEIGHT_PX As Long = 34   ' 第9.12b: ナビバー帯の高さ (px)

' UserForm が保持する Wv2Browser (全タブを束ねる包括クラス、9.10c で正式接続)。
'   Environment 1 個 + 複数タブ (Wv2Pane) を管理する本体フィールド。
'   (第9.10 整理: 旧 m_env / m_pane フィールドは撤去し m_browser に一本化した)
Private WithEvents m_browser As Wv2Browser   ' 第9.11a: イベント受信のため WithEvents 化
Attribute m_browser.VB_VarHelpID = -1
Private m_tabBar As Wv2TabBar                ' 第9.11b: WebView2 製タブバー (任意)
Private m_navBar As Wv2NavBar   ' 第9.12b: WebView2 製ナビゲーションバー

' --- 第9.18: 検索エンジンのプリセット定数は Wv2SettingsBridge へ移設した ---
'   9.17 まで UserForm1 が持っていた SE_GOOGLE / SE_BING / SE_DUCKDUCKGO /
'   SE_YAHOO と名前解決ロジックは、設定画面と共有するため Wv2SettingsBridge に
'   集約した (DRY)。UseSearchEngine は一時ブリッジを生成して委譲する。


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

    ' --- 第9.11b: タブバーがあれば上部帯を確保し、コンテンツ領域を下にずらす ---
    '   ★Browser はタブバーを知らない★ 分割はここ (UserForm) で計算し、
    '   タブバー帯の矩形は m_tabBar.Resize に、コンテンツ領域の矩形は
    '   m_browser.ResizeAll に渡す。タブバーが無ければ従来通り全域をコンテンツに。
    ' --- 第9.12b: 上から順に帯を積み、残りをコンテンツにする (3 分割) ---
    '   ★分割の計算はここ (UserForm) の責務★ Browser / TabBar / NavBar は
    '   互いを知らない。UserForm が矩形を配って回る (疎結合)。
    '   バーが無ければその段を飛ばすので、タブバーだけ / ナビバーだけ / 両方なし
    '   のいずれの構成でも正しく動く。
    Dim y As Long
    y = rc.Top          ' 次に帯を置く y 座標 (上から積み上げる)

    ' --- 1 段目: タブバー ---
    If Not m_tabBar Is Nothing Then
        Dim barRect As RECT
        barRect.Left = rc.Left
        barRect.Top = y
        barRect.Right = rc.Right
        barRect.Bottom = y + TABBAR_HEIGHT_PX
        If barRect.Bottom > rc.Bottom Then barRect.Bottom = rc.Bottom
        m_tabBar.Resize barRect
        y = barRect.Bottom
    End If

    ' --- 2 段目: ナビバー (第9.12b) ---
    If Not m_navBar Is Nothing Then
        Dim navRect As RECT
        navRect.Left = rc.Left
        navRect.Top = y
        navRect.Right = rc.Right
        navRect.Bottom = y + NAVBAR_HEIGHT_PX
        If navRect.Bottom > rc.Bottom Then navRect.Bottom = rc.Bottom
        m_navBar.Resize navRect
        y = navRect.Bottom
    End If

    ' --- 3 段目: コンテンツ領域 (残り全部) ---
    Dim contentRect As RECT
    contentRect.Left = rc.Left
    contentRect.Top = y
    contentRect.Right = rc.Right
    contentRect.Bottom = rc.Bottom
    If contentRect.Top > contentRect.Bottom Then contentRect.Top = contentRect.Bottom
    m_browser.ResizeAll contentRect
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
    ' 第9.10 整理: m_pane 直参照から m_browser.ActivePane (GetActivePane) 経由へ。
    Dim pane As Wv2Pane
    Set pane = GetActivePane
    If pane Is Nothing Then
        Debug.Print "UserForm1.SendMessageToJS: アクティブな Pane が無い"
        Exit Sub
    End If
    If Not pane.IsReady Then
        Debug.Print "UserForm1.SendMessageToJS: Pane が Ready でない"
        Exit Sub
    End If

    Dim hr As Long
    hr = pane.View_PostWebMessageAsString(text)
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
    ' 第9.10 整理: m_pane 直参照から GetActivePane 経由へ。
    Dim pane As Wv2Pane
    Set pane = GetActivePane
    If pane Is Nothing Then
        IsPageLoaded = False
    Else
        IsPageLoaded = pane.IsPageLoaded
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
' BuildSpaAppHtml (第9.10d-2a、最小 SPA HTML を返す)
'   ASCII のみ (日本語を含めない) にして、VBA の Print # (システムロケール)
'   で書き出しても文字化けしないようにする。DOMContentLoaded が確実に
'   発火する実体のあるページ。pushState 等は d-1 同様 ExecuteScript 撃ちで
'   制御するので、HTML 側にはボタンを置かない (論点C: 最小構成)。
' ============================================================
Public Function BuildSpaAppHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbCrLf
    s = s & "<html><head><meta charset=""utf-8""><title>WV2 SPA d-2</title></head>" & vbCrLf
    s = s & "<body style=""font-family:sans-serif"">" & vbCrLf
    s = s & "<h1>WV2 SPA (appassets.example)</h1>" & vbCrLf
    s = s & "<p id=""msg"">initial</p>" & vbCrLf
    s = s & "<script>" & vbCrLf
    s = s & "document.addEventListener('DOMContentLoaded',function(){" & vbCrLf
    s = s & "  document.getElementById('msg').textContent='DOMContentLoaded fired';" & vbCrLf
    s = s & "});" & vbCrLf
    s = s & "</script>" & vbCrLf
    s = s & "</body></html>" & vbCrLf
    BuildSpaAppHtml = s
End Function

' ============================================================
' WriteSpaAppFolder (第9.10d-2a、フォルダ作成 + テキストファイル書き出し)
'   folderPath がなければ作成 (親 C:\Temp は既存前提)。
'   folderPath\fileName に content を書き出す (ASCII 前提、Print # で十分)。
'   戻り値: 成功で True。
' ============================================================
Public Function WriteSpaAppFolder( _
    ByVal folderPath As String, _
    ByVal fileName As String, _
    ByVal content As String) As Boolean

    On Error GoTo eh

    ' フォルダがなければ作成
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        Debug.Print "WriteSpaAppFolder: フォルダ作成 " & folderPath
    End If

    Dim fullPath As String
    fullPath = folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & fileName

    Dim fnum As Integer
    fnum = FreeFile
    Open fullPath For Output As #fnum
    Print #fnum, content
    Close #fnum

    Debug.Print "WriteSpaAppFolder: 書き出し OK " & fullPath & " (" & Len(content) & " 文字)"
    WriteSpaAppFolder = True
    Exit Function
eh:
    Debug.Print "WriteSpaAppFolder: エラー " & Err.Number & " " & Err.Description
    WriteSpaAppFolder = False
End Function

' ============================================================
' StartWebView2_Spa (SPA 起動の雛形)
'
'   ローカルフォルダを https://appassets.example/ にマッピングし、そこの
'   自作 SPA (index.html) を開く。AddTabWithUrlForSpa 1 本で
'   SetMapping + Navigate + EnableSpaHandlers (HistoryChanged/DOMContentLoaded)
'   まで一括で行う (仕様事実 13: Navigate 先はファイル名まで明示)。
'
'   SPA (history.pushState 等) を使うアプリを WebView2 で動かすときの起動口。
'   配信する HTML は BuildSpaAppHtml/WriteSpaAppFolder で用意する
'   (実運用では任意の HTML をフォルダに置き換える)。
'
'   使い方:
'     UserForm1.Show vbModeless
'     UserForm1.StartWebView2_Spa
'   アクティブタブの Pane は ?UserForm1.GetActivePane で取得できる。
' ============================================================
Public Sub StartWebView2_Spa()
    Debug.Print "UserForm1.StartWebView2_Spa 開始 (SPA 起動)"

    ' --- 1. SPA フォルダを準備 (HTML を書き出す) ---
    Dim folderPath As String
    folderPath = "C:\Temp\VBA_WebView2_SPA"
    If Not WriteSpaAppFolder(folderPath, "index.html", BuildSpaAppHtml()) Then
        Debug.Print "StartWebView2_Spa: SPA フォルダ準備に失敗"
        Exit Sub
    End If

    ' --- 2. Wv2Browser を生成し Environment を初期化 ---
    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, "C:\Temp\VBA_WebView2") Then
        Debug.Print "StartWebView2_Spa: Wv2Browser.Init 失敗 (state=" & _
                    m_browser.state & ")"
        Exit Sub
    End If
    Debug.Print "StartWebView2_Spa: Wv2Browser.Init OK"

    ' --- 3. SPA タブを一括生成 (AddTab→SetMapping→Navigate→EnableSpaHandlers) ---
    Dim pane As Wv2Pane
    Set pane = m_browser.AddTabWithUrlForSpa("appassets.example", folderPath, "index.html")
    If pane Is Nothing Then
        Debug.Print "StartWebView2_Spa: AddTabWithUrlForSpa に失敗"
        Exit Sub
    End If
    Debug.Print "StartWebView2_Spa: SPA タブ生成 OK (ViewPtr=" & pane.ViewPtr & ")"
    Debug.Print "  https://appassets.example/index.html が表示されたら成功。"
End Sub

' ============================================================
' UserForm_Terminate
'   フォームが閉じられたら Wv2Pane → Wv2Environment の順で解放。
'   各クラスの Class_Terminate で ComRelease チェーンが走り、
'   ICoreWebView2Controller / ICoreWebView2Environment の参照が解放される。
' ============================================================
' ============================================================
' ★第9.11a イベント受信ハンドラ (疎結合タブバー用イベント口の受け手)
'   Browser の RaiseEvent をここで受けて Debug.Print するだけ。
'   これが実際のタブバー実装 (WebView2 製/MsForms 製) に差し替わる。
'   ★検証の見方★ イミディエイトウィンドウ (Ctrl+G) に以下のログが出る:
'     [EVT] TabAdded / [EVT] ActiveChanged / [EVT] TabClosed / [EVT] TabTitleChanged
' ============================================================
Private Sub m_browser_TabAdded(ByVal index As Long)
    Debug.Print "[EVT] TabAdded(index=" & index & ")"
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "[EVT] ActiveChanged(index=" & index & ")"
End Sub

Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "[EVT] TabClosed(index=" & index & ")"
End Sub

Private Sub m_browser_TabTitleChanged(ByVal index As Long, ByVal title As String)
    Debug.Print "[EVT] TabTitleChanged(index=" & index & ", title=" & title & ")"
End Sub

' --- 第9.12a: URL 状態の変化 (アドレスバー更新用イベント) ---
'   NavigationCompleted と HistoryChanged の両方から撃たれるので、1 回の遷移で
'   複数回出るのが正常 (全体同期方式なので無害、設計原則 60)。
Private Sub m_browser_TabUrlChanged(ByVal index As Long, ByVal url As String, _
                                    ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    Debug.Print "[EVT] TabUrlChanged(index=" & index & ", url=" & url & _
                ", back=" & canGoBack & ", fwd=" & canGoForward & ")"
End Sub
' --- 現在状態をダンプする内部ヘルパ ---
Private Sub DumpState(ByVal tag As String)
    If m_browser Is Nothing Then
        Debug.Print "  [状態:" & tag & "] m_browser は Nothing"
        Exit Sub
    End If
    Debug.Print "  [状態:" & tag & "] TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex
End Sub


Private Sub UserForm_Terminate()
    Debug.Print "UserForm1.Terminate 開始"

    ' --- 第9.12b: ナビバーを先に落とす (生成と逆順) ---
    If Not m_navBar Is Nothing Then
        m_navBar.Shutdown
        Set m_navBar = Nothing
    End If

    ' 第9.11b: タブバーを先に解放 (WithEvents 購読を解除)。
    If Not m_tabBar Is Nothing Then
        m_tabBar.Shutdown
        Set m_tabBar = Nothing
    End If

    ' 第9.10e: 循環参照 (Browser ? Pane) を断ってから解放する。
    '   Shutdown が各 Pane の m_browser 後方参照を切る。これを呼ばずに
    '   Set m_browser = Nothing すると循環参照で解放されずリークする (設計原則 58)。
    If Not m_browser Is Nothing Then m_browser.Shutdown

    ' Wv2Browser を解放 (内部で全 Pane → Environment の順、設計原則 43)
    Set m_browser = Nothing

    Debug.Print "UserForm1.Terminate 完了"
End Sub
' --- フル構成で起動 (3 タブ + タブバー + ナビバー) ---
'   ★HTML の配信フォルダは TabBar と NavBar で別にする★
'   仮想ホスト名も別 (appassets.tabbar / appassets.navbar)。
Public Sub StartWebView2_Full()
    Debug.Print "==== StartWebView2_Full 開始 (3 タブ + タブバー + ナビバー) ===="

    ' --- 1. Browser を起動して 3 タブ追加 ---
    Set m_browser = New Wv2Browser
    If Not m_browser.Init(hWnd_Frame, "C:\Temp\VBA_WebView2") Then
        Debug.Print "StartWebView2_Full: Browser.Init 失敗 (state=" & m_browser.state & ")"
        Exit Sub
    End If
    Dim p As Wv2Pane
    Set p = m_browser.AddTabWithUrl("https://www.bing.com")
    If p Is Nothing Then Debug.Print "  タブ追加失敗 (bing)": Exit Sub
    Set p = m_browser.AddTabWithUrl("https://example.com")
    If p Is Nothing Then Debug.Print "  タブ追加失敗 (example)": Exit Sub
    Set p = m_browser.AddTabWithUrl("https://www.msn.com")
    If p Is Nothing Then Debug.Print "  タブ追加失敗 (msn)": Exit Sub

    ' --- 2. タブバーを起動 ---
    Set m_tabBar = New Wv2TabBar
    If Not m_tabBar.Init(m_browser, hWnd_Frame, "C:\Temp\VBA_WebView2_TabBar") Then
        Debug.Print "StartWebView2_Full: TabBar.Init 失敗"
        Set m_tabBar = Nothing
        Exit Sub
    End If

    ' --- 3. ナビバーを起動 (第9.12b、配信フォルダは TabBar と別) ---
    Set m_navBar = New Wv2NavBar
    If Not m_navBar.Init(m_browser, hWnd_Frame, "C:\Temp\VBA_WebView2_NavBar") Then
        Debug.Print "StartWebView2_Full: NavBar.Init 失敗"
        Set m_navBar = Nothing
        Exit Sub
    End If

    ' --- 4. レイアウトを再計算 (3 分割) ---
    UserForm_Resize

    Debug.Print "==== StartWebView2_Full 完了 (TabCount=" & m_browser.TabCount & _
                ", ActiveIndex=" & m_browser.ActiveIndex & _
                ", TabBar.IsReady=" & m_tabBar.IsReady & _
                ", NavBar.IsReady=" & m_navBar.IsReady & ") ===="
    Debug.Print "  3 段 (タブバー / ナビバー / コンテンツ) が重ならず並べば成功。"
End Sub

' ============================================================
' CurrentBrowser (第9.17、案 A-1)
'
'   起動中の Wv2Browser への読み取り専用参照を返す。m_browser は
'   Private WithEvents で外 (イミディエイト / 検索エンジン切替 / 将来の UI) から
'   触れないため、素の参照を渡す口を 1 本だけ公開する。WithEvents は宣言側の
'   制約であって、Get で参照を返すのは問題ない。
'
'   起動前 (StartWebView2_Full 未実行) は m_browser が Nothing なので
'   Nothing を返す。呼び出し側は Is Nothing を確認すること。
'
'   ★用途★
'     ・イミディエイトで  Set b = UserForm1.CurrentBrowser  として掴み、
'       b.SearchEngine = ... で検索エンジンを実機切替する (9.16 の宿題)。
'     ・今後の拡張全般で「起動中 Browser を外から触る」共通の足場。
' ============================================================
Public Property Get CurrentBrowser() As Wv2Browser
    Set CurrentBrowser = m_browser
End Property


' ============================================================
' UseSearchEngine (第9.17、案 B-1)
'
'   プリセット名で検索エンジンを切り替える最小の口。GUI 部品ではなく
'   ロジックの口 (実機ではイミディエイトから叩く)。
'
'   引数 engineName: "google" / "bing" / "duckduckgo" / "yahoo"
'     大文字小文字・前後空白は無視する。未知の名前は既定 (Google) に落として
'     警告を出す。空文字/空白のみに対する最終ガードは Browser 側 Property Let
'     (空なら Google へフォールバック) が担保するので、ここでは名前解決に専念する。
'
'   Browser 未起動 (CurrentBrowser Is Nothing) の場合は何もせず注意を出す。
'
'   使用例 (実機、StartWebView2_Full 起動後):
'     UserForm1.UseSearchEngine "bing"
'     → その後アドレスバーに  日本 天気  と打つと Bing 検索結果が出る。
' ============================================================
Public Sub UseSearchEngine(ByVal engineName As String)
    If m_browser Is Nothing Then
        Debug.Print "UseSearchEngine: Browser が未起動です (先に StartWebView2_Full を実行)。"
        Exit Sub
    End If

    ' 第9.18: プリセット名の解決は Wv2SettingsBridge に集約済み。ここでは一時
    '   ブリッジを生成して Browser を結び付け、SetEngine に委譲する。設定画面が
    '   使うのと全く同じ名前解決経路を通るので、イミディエイトと GUI で挙動が一致する。
    Dim bridge As Wv2SettingsBridge
    Set bridge = New Wv2SettingsBridge
    bridge.BindBrowser m_browser

    Dim applied As String
    applied = bridge.SetEngine(engineName)

    Debug.Print "UseSearchEngine: 検索エンジンを '" & applied & "' に設定 (" & _
                m_browser.Debug_SearchEngine() & ")"
End Sub


