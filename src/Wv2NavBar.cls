VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2NavBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2NavBar.cls  第9.25b 段階 (検証プローブの撤去 / 引数ゼロを恒久確定) ---
'
'   第9.25 の検証が「パターンA = 安全」で決着したので、一時コードを撤去した回。
'
'   ★検証結果 (第9.25 実機)★ プローブのログは次の順で出た:
'       (1) 取得の直前
'       (2) 取得の直後            ← ★この間に HostProbe の実行ログが無い★
'       ★HostProbe が実行された★ (通算 1 回目)
'       (3) f() 呼び出し 成功
'       ★HostProbe が実行された★ (通算 2 回目)
'       (5) nb.HostProbe() 直呼び 成功
'     ・(1)-(2) 間に実行ログが無い = プロパティ取得の時点では VBA 側を呼んでいない。
'     ・(3) が出た = 取得できた値は関数オブジェクトであり、後から呼び出せる。
'       つまり sync プロキシは『メソッドかプロパティか』の判断を実際の () まで遅延している。
'     ・(5) が出た = 実際に使う nb.HostProbe() の形が正常に通る。
'     ・通算 2 回ちょうど = 余計な実行はゼロ。
'     実物のボタンでも HostBack / HostForward / HostReload がそれぞれ 1 行だけで、
'     postMessage フォールバックへの落下も無く、← は 1 ページだけ戻った。
'
'   → **仕様事実 22** として記録 (詳細は開発メモ)。9.24 のダミー引数 reserved は
'     不要な保険だったことが確定したので、引数ゼロを恒久確定とする。
'     ※ただし「安全側から入って後で削る」という順序自体は正しかった。二重実行は
'       起きると『1 クリックで 2 ページ戻る』という分かりにくい形で現れ、しかも
'       ハイブリッドの保険が裏目に出るため後から気付きにくい種類の不具合になる。
'
'   ★撤去したもの (第9.25 で入れた一時コード)★
'     ・Public Function HostProbe() / HostProbeMark(phase) の 2 本
'     ・BuildNavBarHtml 内の probeParamless() と setTimeout(probeParamless, 500)
'     これで NavBar は 9.24 の構造そのままで、Host メソッド 3 本が引数ゼロになった形に
'     落ち着く。JS からの呼び出しは nb.HostBack() / nb.HostForward() / nb.HostReload() /
'     nb.HostNavigate(url)。
'
'   ★触らないもの★ Host メソッド 4 本の中身 / OnPaneWebMessage の委譲構造 /
'     navHost() の三段構え / attach / ready / Escape / editing / PushNavSyncToJs。
'     設計原則は 71 個のまま。仕様事実は 21 → 22 個。
'
' --- Wv2NavBar.cls  第9.25 段階 (検証: 引数ゼロ Host メソッドの可否 + reserved 削除) ---
'
'   第9.25 のスコープ (9.24 の宿題を潰す小さい検証ステージ。案C を採用):
'     ★問い★ 引数ゼロの Public Function を hostObjects の sync プロキシから触ったとき、
'       『プロパティ取得の時点で実行されてしまう』のか。9.24 ではこれを恐れてダミー引数
'       reserved を持たせたが、安全側に倒したので現象が出なかっただけで、要否は未検証
'       だった (「出なかった」は「不要」の証明ではない)。
'
'     ★案C = プローブ追加 + reserved 削除を同時に行う★
'       ・案A (reserved を外して実物のボタンで試すだけ) は最小変更だが、二重実行が
'         起きたとき原因が『取得で実行』なのかを推測でしか言えない。
'       ・案B (プローブだけ足して reserved は据え置き) は因果が完全に分かるが 2 段になる。
'       ・→ 案C なら 1 回の起動で「機構の事実 (プローブ)」と「実物の挙動 (ボタン)」が
'         同時に取れる。外れても実害は『← 1 回で 2 ページ戻る』だけで、しかも一度
'         観測すれば十分なので、安価に両方の答えが出る。
'
'     ★変更 1: reserved を削除★
'       HostBack / HostForward / HostReload から Optional ByVal reserved As Long を外し、
'       素の引数ゼロ関数に戻した。JS 側も nb.HostBack(0) → nb.HostBack() へ。
'       ★これが今回いちばん壊れやすい変更★ もし『取得で実行』が起きるなら、
'         (1) 取得で 1 回実行 → (2) 返り値が Boolean なので () が TypeError →
'         (3) catch して postMessage フォールバック → もう 1 回実行
'       となり、← 1 回で 2 ページ戻る。ログには HostBack が 2 行 (間に
'       OnPaneWebMessage の back) 出るので一目で分かる。
'
'     ★変更 2: 検証専用プローブを追加 (次段で撤去する一時コード)★
'       ・Public Function HostProbeMark(ByVal phase As Long) As Boolean
'           引数ありなので安全。位相を刻んでログに出すだけ。
'       ・Public Function HostProbe() As Boolean
'           ★引数ゼロ★。呼ばれたらログを出すだけで副作用ゼロ。実ブラウザ操作を
'           一切伴わないので、二重実行が起きても害がない。
'       JS 側は ready 送信の直後に setTimeout(probeParamless, 500) を仕掛け、
'       起動のたびに一度だけ次を実行する:
'           Mark(1) → var f = nb.HostProbe (取得だけ) → Mark(2)
'           → f() 成功 Mark(3) / 失敗 Mark(4)
'           → nb.HostProbe() 直呼び 成功 Mark(5) / 失敗 Mark(6)
'           (取得そのもので例外なら Mark(9))
'       ★読み方★ HostProbe の実行ログが Mark(1) と Mark(2) の『間』に出れば
'         『取得で実行されている』= reserved は必要だった。Mark(2) より後にしか
'         出なければ安全 = reserved は不要。Mark(5) が出れば実際に使う
'         nb.HostProbe() の形が通ることの直接証明になる。
'       ★500ms 遅延の理由★ ready ハンドシェイクと初期 navsync の往復を邪魔せず、
'         WebView2 初期化バーストからも少し離すため。ログのブロックも読みやすくなる。
'
'     ★トリガーを JS 内の setTimeout にした理由★
'       Wv2Pane.View_ExecuteScript (9.8c) は既にあるが、NavBar インスタンスは
'       UserForm1 が Private フィールドで握っているため、外から叩くには
'       UserForm1.frm に公開口を足すことになる。検証専用の一時コードのために
'       本番ファイルを増やしたくないので、JS 内で自走させる方を採った。
'       たーぼーさんは普通に起動してログを読むだけでよい。
'
'   ★触らないもの★ Host メソッド 4 本の中身 / OnPaneWebMessage の委譲構造 /
'     navHost() の三段構え / ready / Escape / editing / PushNavSyncToJs。
'     設計原則は 71 個のまま。仕様事実は検証結果を見てから立てる。
'
' --- Wv2NavBar.cls  第9.24 段階 (NavBar の hostObjects 化 + 処理の Host 一元化) ---
'
'   第9.24 のスコープ (9.23 で TabBar に確立した型を NavBar へ同型展開する):
'     ★狙い★ 9.21～9.23 で TabBar の activate/close/reorder は
'       「JS はハイブリッド (hostObjects 優先・postMessage フォールバック)、
'       VBA は処理を Host メソッド 1 箇所に集約、Select Case は薄いアダプタ」
'       という形に落ち着いた。NavBar の back/forward/reload/navigate も
'       まったく同じ型に揃える。TabBar で型が確立済みのため、9.21/9.22 のような
'       「まず hostObjects 化だけ」の段は踏まず、最初から案C-1 で一気に組む (案2)。
'
'     ★追加した口 (Host メソッド 4 本)★
'       ・Public Function HostBack(Optional reserved) As Boolean    → GoBackActive
'       ・Public Function HostForward(Optional reserved) As Boolean → GoForwardActive
'       ・Public Function HostReload(Optional reserved) As Boolean  → ReloadActive
'       ・Public Function HostNavigate(url) As Boolean              → NavigateActive
'       いずれも Browser 側が撃ちっぱなし (完了待ちなし = 案4-A) なので、JS からの
'       sync 直呼びでも 仕様事実 14 のデッドロックに嵌まらない。
'
'     ★★ダミー引数 reserved の理由 (論点4b)★★
'       TabBar の 3 本はすべて引数ありだったが、back/forward/reload は本来
'       引数ゼロになる。hostObjects の sync プロキシは JS 側から「メソッドか
'       プロパティか」を区別できないため、ho.NavBar.HostBack を『取得』した
'       時点で IDispatch::Invoke(DISPATCH_PROPERTYGET) が飛ぶ実装であり得る。
'       VBA は引数なし Public Function に対してこれを寛容に受け付けて実行して
'       しまうため、最悪 (1) 取得で 1 回実行 → (2) 返り値が Boolean なので ()
'       が TypeError → (3) catch して postMessage フォールバック → もう 1 回実行、
'       という「← を 1 回押すと 2 回戻る」二重実行が起こり得る。ハイブリッドの
'       保険が裏目に出る形なのでフォールバックでは防げない。
'       → 引数を 1 個持たせて曖昧さを原理的に消す (JS は nb.HostBack(0) と呼ぶ)。
'       VBA 側の Case からは Optional なので HostBack と素で呼べる。実機で
'       「引数ゼロでも安全」と判明したら次段で reserved を落としてよい。
'
'     ★空 URL ガードの置き場所 (論点3)★ Case 内にあった If Len(Trim$(url)) = 0
'       の判定を HostNavigate の中へ移した。これで hostObjects 経路・postMessage
'       経路のどちらから来ても同じガードを通る (9.23 と同じ DRY の考え方)。
'       JS 側の if(v && v.length) は保険として残す。
'
'     ★JS 側は navHost() ヘルパでハイブリッド化★ TabBar は 3 箇所を素で書き
'       下したが、NavBar は 4 箇所と多く形も揃っているのでヘルパに括る。
'         navHost(msgObj, callHost)
'           (1) ho.NavBar があれば callHost(ho.NavBar) を同期直呼び → [host]
'           (2) 例外なら catch して postCmd(msgObj)      → [host-fail→postmsg]
'           (3) そもそも取れなければ postCmd(msgObj)     → [postmsg]
'       フォールバックで送る JSON は 9.12d と完全に同一なので、VBA 側の
'       Select Case (薄いアダプタ) がそのまま受け止められる。
'       ★ho.NavBar へのアクセスはドット表記のみ★ (TabBar で実績のある形。
'       ho.NavBar[name] のような計算プロパティ表記は proxy の挙動が読めない。)
'
'     ★attach★ Init の Navigate の前に
'       m_barPane.View_AddHostObjectToScript("NavBar", Me) を置く (仕様事実 18/19
'       の順序A。TabBar の 5b と同じ位置)。失敗しても致命ではないので hr をログ
'       して Init は続行する (JS が postMessage に落ちるだけ)。
'
'   ★触らないもの★ ready ケース (通知なので postMessage が正着、設計原則 61) /
'     Case "" / Case Else / PushNavSyncToJs / SendNavSync / navsync 受信側の JS /
'     編集中フラグ (editing) と lastUrl の挙動 / Escape の編集破棄 / Browser の
'     5 イベント購読 / 3 段レイアウト。設計原則は 71 個のまま (案4-A と DRY の
'     実践であり新原則ではない)。
'
'   ★実機で確かめる新規事実の候補 (仕様事実 21)★
'     hostObjects.sync 経由で JS→VBA へ「文字列引数」を渡すのは今回が初めて
'     (TabBar は Long だけだった)。日本語や記号を含む URL / 検索語が BSTR の
'     まま欠落なく届けば 仕様事実 21 として記録する。postMessage 経路は JSON
'     エスケープを通っていた (仕様事実 16 の当事者) ので、hostObjects のほうが
'     むしろ安全になるはず、というのが事前の見立て。
'
' --- Wv2NavBar.cls  第9.13 段階 ---
'
'   第9.13 の変更 (整理: JSON ヘルパーの共通化):
'     Private に持っていた JsonEscape / GetJsonStr / BoolToJson を撤去し、
'     新設した標準モジュール Wv2Json.bas の Public 関数を使うようにした。
'     Wv2TabBar にも同じ関数が重複していたのを 1 箇所に集約する。
'       GetJsonStr → Wv2Json.JsonGetStr
'       BoolToJson → Wv2Json.JsonBool
'       JsonEscape → Wv2Json.JsonEscape (名前は同じ)
'     ★JsonEscape の \u 経路も修正された★ 高位文字 (日本語等) が \uXXXX に
'     化けなくなった (仕様事実 16。挙動は変わらないがログが読みやすくなる)。
'
' --- Wv2NavBar.cls  第9.12d 段階 ---
'
'   第9.12d のスコープ (JS->VBA: 入力・クリックの有効化)。これで 9.12 完結:
'     ナビバーのボタンと URL 入力を VBA のナビゲーション操作に繋ぐ。
'       ・←  クリック    → {"cmd":"back"}                 → Browser.GoBackActive
'       ・→  クリック    → {"cmd":"forward"}              → Browser.GoForwardActive
'       ・[更新]  クリック    → {"cmd":"reload"}               → Browser.ReloadActive
'       ・URL 欄で Enter → {"cmd":"navigate","url":"..."} → Browser.NavigateActive
'
'   ★再入の安全性 = 設計原則 63★ ナビゲーション系はすべて「完了待ちを含まない」
'     操作 (Navigate / GoBack / GoForward / Reload はいずれも撃ちっぱなし) なので、
'     WebMessage のイベントハンドラ内から同期で呼んで安全。9.11d の newtab
'     (Controller 生成待ち → 仕様事実 14 のデッドロック) のような遅延実行は不要。
'
'   ★スキーム補完は Browser 側 (NavigateActive) が行う (案3-A)★
'     JS は入力文字列をそのまま送るだけ。"example.com" → "https://example.com" の
'     正規化は Wv2Browser.NormalizeUrl の責務 (9.12a で実装済み)。
'
'   ★Escape で編集破棄 (案A: JS 側の lastUrl から即復元)★
'     当初 Escape は editing = false にして blur するだけだったが、その後 URL 欄の
'     値を戻す処理がなく、打った文字列が残ったままになった (実機で発覚)。
'     u.value が書き換わるのは次に navsync が飛んできたときだけで、Escape では
'     VBA 側で何も起きない (= sync が飛ばない) ため。
'     → JS 側に lastUrl (直近の navsync が伝えた「今表示すべき URL」) を持ち、
'        Escape で即座に復元する。VBA 往復が不要で体感も良い。lastUrl は
'        ★編集中でも必ず更新する★ (編集中に裏で遷移が起きても、Escape で
'        「今の本当の URL」に戻れるように)。
'
'   ★Enter で editing = false にする★
'     遷移が始まると TabUrlChanged → navsync で URL 欄が正規化後の URL に
'     置き換わる ("example.com" と打つと "https://example.com/" に化ける)。
'     これは Chrome/Edge と同じ挙動。editing を落とさないと URL 欄が打った文字列の
'     まま固まってしまう。
'
' --- Wv2NavBar.cls  第9.12c 段階 ---
'
'   第9.12c のスコープ (VBA->JS: URL とボタン状態の表示。入力・クリックは 9.12d):
'     Browser のイベントを受けて、ナビバー JS に「アクティブタブの URL /
'     戻れるか / 進めるか」を送って表示させる。
'
'   ★全体同期方式 (設計原則 60)★ 9.11c と同じ。状態が変わるたびに丸ごと送る:
'       {"type":"navsync","url":"...","canBack":true,"canForward":false}
'     JS 側は受け取った値で URL 欄とボタンの disabled を上書きするだけ。
'
'   ★ready ハンドシェイク (設計原則 61)★ 9.11c と同じ。JS がロード完了時に
'     {"cmd":"ready"} を postMessage → VBA が受けて初めて初期同期を返す。
'     PostWebMessageAsString は JS の message リスナー登録前に送ると落ちる
'     (仕様事実 15)。
'
'   ★★index フィルタ (9.12c 固有の論点)★★
'     TabUrlChanged は「どのタブの」URL が変わったかを index で伝える。ナビバーは
'     アクティブタブの状態だけを表示するので、index <> ActiveIndex のイベントは
'     ★無視しなければならない★。無視しないと、裏のタブが読み込み完了した瞬間に
'     URL 欄が裏タブの URL に化ける (実際 9.12b のログで、index=2 の TabUrlChanged が
'     index=3 アクティブ時に飛んでいる)。
'
'   ★ActiveChanged では Browser に現在状態を問い合わせる★
'     タブを切り替えたとき、新しいアクティブタブの URL は TabUrlChanged では
'     飛んでこない (そのタブでは何も起きていないため)。よって ActiveChanged を
'     受けたら Browser.GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward で
'     現在状態を読み直して送る。9.12a でこの読み取り口を用意してある。
'
'   ★実機で発覚したバグと修正 その3 (9.12c: 入力中の URL が消える)★
'     URL 欄に入力している最中にリンクをクリックして新タブが開くと、入力中の
'     文字列が消えた。原因は「その2」で入れた u.disabled の同期:
'       disabled を代入すると input が blur する
'         → blur ハンドラが editing = false にする
'         → 次の同期で u.value が上書きされ、入力が消える
'     ★タブ 0 対応 (その2) が編集中保護 (その1) を壊した★
'     → 対策は 2 段構え:
'        (1) 編集中は URL 欄に一切触らない (value も disabled も)
'        (2) disabled は値が変わるときだけ代入する (無駄な blur を防ぐ)
'        ボタン類はフォーカスを奪わないので常に同期してよい。
'     教訓: disabled の付け外しはフォーカスを飛ばす。フォーカス依存の状態
'           (編集中フラグ等) を持つ要素には、無条件に属性を代入しない。
'
'   ★実機で発覚したバグと修正 その2 (9.12c: タブ 0 で URL 欄が固まる)★
'     最後の 1 タブを閉じると、ナビバーの URL 欄が閉じたタブの URL のまま残った。
'     原因: Wv2Browser.CloseTab は「アクティブにするタブがない」場合 ActivateTab を
'     呼ばないため ★ActiveChanged が撃たれない★ (9.11a の 4 分岐設計。タブ 0 は
'     許容するが ActiveChanged(0) は撃たない)。NavBar は ActiveChanged しか
'     購読していなかったので、タブ 0 になったことを知らされなかった。
'     → 対処 (案A): NavBar が TabClosed も購読して PushNavSyncToJs する。
'        PushNavSyncToJs は既にタブ 0 を正しく扱う (URL 空・全ボタン無効)。
'        ★Browser のイベント契約は変えない★ 「表示に必要なイベントは表示側が
'        購読する」という疎結合の原則に従う。TabClosed と ActiveChanged が両方
'        飛ぶケースでは 2 回撃つが、全体同期方式なので無害 (設計原則 60)。
'     あわせて navsync に hasTab を追加し、リロードボタンもタブ 0 で無効にする。
'
'   ★実機で発覚したバグと修正 その1 (9.12c)★
'     当初 URL 欄の編集中保護を document.activeElement !== u で判定していたが、
'     WebView2 の単独ドキュメントでは #url が activeElement になったまま戻らず、
'     ★初期同期の 1 回だけ効いて以降 URL 欄が二度と更新されない★ という症状が出た
'     (ボタンの活性は正しく追従していたので、applyNavSync 自体は毎回走っていた)。
'     → 「ユーザーが実際に編集を始めたか」を明示フラグ (editing) で持つ方式に変更。
'        focus しただけでは編集中にせず、キー入力 (input イベント) で編集中にし、
'        blur / Enter / Escape で解除する。
'     教訓: フォーカス状態を「編集中かどうか」の代理指標に使わない (設計原則 65)。
'
'   ★JSON ヘルパーの重複について★
'     JsonEscape / GetJsonStr は Wv2TabBar にも同じものがある (Private のため
'     共有できない)。標準モジュールへの切り出しは将来の整理段階でまとめて行う
'     方針とし、現時点では 2 箇所の重複を許容する (合意済み)。
'
' --- Wv2NavBar.cls  第9.12b 段階 ---
'
'   ★WebView2 製ナビゲーションバー (アドレスバー + 戻る/進む/リロード)★
'
'   第9.11 で確立した「WebView2 を UI 部品として使う」パターン (Wv2TabBar) を
'   そのまま踏襲する。画面上部の 2 段目に細長い WebView2 をもう 1 枚置き、
'   HTML/CSS/JS でナビバーを描画する。
'
'   レイアウト (UserForm_Resize が 3 分割する):
'     +--------------------------+
'     | [Bing][MSN x] [+]        | <- Wv2TabBar   (高さ TABBAR_HEIGHT_PX)
'     +--------------------------+
'     | [<][>][R] [https://...  ]| <- Wv2NavBar   (高さ NAVBAR_HEIGHT_PX) ★本クラス
'     +--------------------------+
'     |      コンテンツ          | <- Wv2Browser のタブ群
'     +--------------------------+
'
'   ★所有構造 (Wv2TabBar と同じ、案b)★
'     Wv2NavBar が自前の Wv2Pane (m_barPane) を 1 枚持つ。この m_barPane は
'     Wv2Browser.m_tabs には入れない (タブ管理下に置かない)。よって AddTab は
'     使わず、Wv2Pane を直接生成する。
'
'   ★疎結合 (第9.12a で用意した口を使う)★
'     Wv2Browser は Wv2NavBar の存在を知らない。NavBar が Browser のイベントを
'     WithEvents で購読し、Browser の公開 API (NavigateActive / GoBackActive 等)
'     を呼ぶだけ。Browser 側は 9.12a で以下を用意済み:
'       - Public Event TabUrlChanged(index, url, canGoBack, canGoForward)
'       - Public Event ActiveChanged(index)  (タブ切替時に URL 表示を切り替える)
'       - NavigateActive / GoBackActive / GoForwardActive / ReloadActive / StopActive
'       - GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'
'   ★Environment 共有★
'     m_barPane は Wv2Browser と同じ Environment を共有する。Init で
'     browser.Environment を受け取って m_barPane.Init に渡す。
'
'   第9.12b のスコープ (骨格のみ):
'     作る  : Init (ペイン生成 → View 取得 → 仮想ホストマッピング → 静的 HTML 表示)
'             / Resize / Shutdown / IsReady / WithEvents ハンドラの空実装
'     作らない: URL とボタン状態の表示 (9.12c) / 入力・クリックの有効化 (9.12d)
'     → 9.12b では 2 段目に「見た目だけのナビバー」が表示される。ボタンは押せず、
'       URL 欄も更新されない。3 分割レイアウトが正しく共存することを確認するのが目的。
'
'   ★HTML 配信★
'     Wv2TabBar と同じく SetVirtualHostNameToFolderMapping を使う。
'     ★フォルダは TabBar と別にすること★ ホスト名は appassets.navbar
'     (TabBar は appassets.tabbar)。配信フォルダも分ける (呼び出し側が指定)。
'
'   ★仕様事実 13★ Navigate 先はファイル名まで明示する
'     (https://appassets.navbar/navbar.html)。
'   ★設計原則 56★ Ctrl_GetCoreWebView2 を省略すると Navigate が E_POINTER。
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' フィールド
' ============================================================
Private WithEvents m_browser As Wv2Browser   ' Browser のイベントを購読
Attribute m_browser.VB_VarHelpID = -1
Private m_barPane    As Wv2Pane              ' ナビバー描画用の単一ペイン
Private m_hwndHost   As LongPtr              ' ホスト (親フレーム) の hwnd
Private m_folderPath As String               ' ナビバー HTML の配信フォルダ
Private m_ready      As Boolean              ' Init 完了フラグ

' ============================================================
' 定数
' ============================================================
Private Const NAVBAR_HOST As String = "appassets.navbar"
Private Const NAVBAR_FILE As String = "navbar.html"

' Controller Ready 待ちのタイムアウト (秒)
Private Const READY_TIMEOUT As Single = 10#


Private Sub Class_Initialize()
    Set m_browser = Nothing
    Set m_barPane = Nothing
    m_hwndHost = 0
    m_folderPath = ""
    m_ready = False
End Sub


' ============================================================
' Init ? ナビバーを起動する
'
'   引数:
'     browser    : Wv2Browser (Init 済みであること。Environment を借りる)
'     hwnd       : 親フレームの hwnd (Browser と同じホストに載せる)
'     folderPath : navbar.html を書き出すフォルダ
'                  ★TabBar とは別フォルダにすること★
'
'   手順 (Wv2TabBar.Init と同じ):
'     1. ナビバー HTML を書き出す
'     2. ペイン生成 (AddTab は使わない = タブ管理下に置かない)
'     3. Controller Ready 待ち
'     4. View 取得 (設計原則 56)
'     5. 仮想ホストマッピング (Navigate の前に必須)
'     6. Navigate (仕様事実 13: ファイル名まで明示)
'
'   ★Set m_browser = browser で WithEvents 購読が始まる★
'     以降、Browser の TabUrlChanged / ActiveChanged が本クラスに届く。
'     9.12b ではハンドラは Debug.Print のみ (中身は 9.12c)。
' ============================================================
Public Function Init( _
    ByVal browser As Wv2Browser, _
    ByVal hwnd As LongPtr, _
    ByVal folderPath As String) As Boolean

    If browser Is Nothing Then
        Debug.Print "Wv2NavBar.Init: browser が Nothing"
        Exit Function
    End If

    Dim env As Wv2Environment
    Set env = browser.Environment
    If env Is Nothing Then
        Debug.Print "Wv2NavBar.Init: browser.Environment が Nothing (Browser.Init 未完了?)"
        Exit Function
    End If

    Set m_browser = browser        ' ★これで WithEvents 購読が始まる
    m_hwndHost = hwnd
    m_folderPath = folderPath

    ' --- 1. ナビバー HTML を書き出す ---
    If Not WriteNavBarHtml() Then
        Debug.Print "Wv2NavBar.Init: HTML 書き出し失敗"
        Exit Function
    End If

    ' --- 2. ナビバー用ペインを生成 (AddTab は使わない) ---
    Set m_barPane = New Wv2Pane
    m_barPane.Init env, m_hwndHost

    ' --- 3. Controller Ready 待ち ---
    If Not WaitPaneReady(m_barPane) Then
        If m_barPane.IsFailed Then
            Debug.Print "Wv2NavBar.Init: Controller 生成失敗 (lastError=&H" & _
                        Hex(m_barPane.LastError) & ")"
        Else
            Debug.Print "Wv2NavBar.Init: Controller 生成タイムアウト"
        End If
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: Controller 生成 OK (pCtrl=" & m_barPane.CtrlPtr & ")"

    ' --- 4. View 取得 (省略すると Navigate が E_POINTER、設計原則 56) ---
    Dim hr As Long
    hr = m_barPane.Ctrl_GetCoreWebView2()
    If hr <> 0 Or m_barPane.ViewPtr = 0 Then
        Debug.Print "Wv2NavBar.Init: View 取得失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: View 取得 OK (pView=" & m_barPane.ViewPtr & ")"

    ' --- 5. 仮想ホストマッピング (Navigate の前に必須) ---
    hr = m_barPane.View3_SetVirtualHostNameToFolderMapping(NAVBAR_HOST, m_folderPath, 1)  ' 1 = ALLOW
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: マッピング失敗 hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- 5b. hostObjects 公開 (第9.24、Navigate の前に必須: 仕様事実 18/19 の順序A) ---
    '   TabBar の 5b (9.21) と同じ位置・同じ形。attach するのは NavBar 自身 (Me)。
    '   JS からは window.chrome.webview.hostObjects.sync.NavBar.HostBack() /
    '   HostForward() / HostReload() / HostNavigate(url) で呼ばれる (第9.25 で引数ゼロ化)。
    '   失敗しても致命ではない (JS 側が postMessage にフォールバックする) ため、
    '   ここでは hr をログするだけで Init は続行する。
    Dim hrHost As Long
    hrHost = m_barPane.View_AddHostObjectToScript("NavBar", Me)
    If hrHost <> 0 Then
        Debug.Print "Wv2NavBar.Init: AddHostObjectToScript(NavBar) 失敗 hr=&H" & Hex(hrHost) & " (JS は postMessage にフォールバック)"
    Else
        Debug.Print "Wv2NavBar.Init: NavBar ブリッジを attach OK (back/forward/reload/navigate は hostObjects 優先)"
    End If

    ' --- 6. Navigate (仕様事実 13: ファイル名まで明示) ---
    Dim url As String
    url = "https://" & NAVBAR_HOST & "/" & NAVBAR_FILE
    hr = m_barPane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: Navigate 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: Navigate OK (" & url & ")"

    ' --- 7. WebMessage シンク登録 + 受信有効化 (第9.12c、案P) ---
    '   ナビバー JS からの postMessage (ready / 将来のクリック・URL 入力) を
    '   OnPaneWebMessage で受け取れるようにする。
    '   ★初期同期は「ここでは撃たない」★ JS の message リスナー登録前に
    '   PostMessage すると落ちる (仕様事実 15)。JS が {"cmd":"ready"} を
    '   送ってきてから OnPaneWebMessage 内で PushNavSyncToJs する (設計原則 61)。
    m_barPane.SetWebMsgSink Me
    hr = m_barPane.View_AddWebMessageReceived()
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Init: View_AddWebMessageReceived 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2NavBar.Init: WebMessage 受信を有効化 (ready 待ち)"

    m_ready = True
    Init = True
    Debug.Print "Wv2NavBar.Init: 完了 (ナビバー起動、初期同期は ready 受信後)"
End Function


' ============================================================
' Resize ? ナビバー帯の矩形を適用する
'   UserForm_Resize が計算した「2 段目の帯」の矩形をそのまま受け取る。
'   ★NavBar は自分がどこに置かれるかを計算しない★ レイアウトの真実源は
'   上位 (UserForm)。ここは言われた矩形に従うだけ (設計原則 59 の考え方)。
' ============================================================
Public Sub Resize(ByRef barRect As RECT)
    If m_barPane Is Nothing Then Exit Sub
    Dim hr As Long
    hr = m_barPane.Ctrl_PutBounds(barRect)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.Resize: Ctrl_PutBounds 失敗 hr=&H" & Hex(hr)
    End If
    m_barPane.Ctrl_PutIsVisible True   ' ナビバーは常時可視
End Sub


' ============================================================
' Shutdown ? 明示的に解放する (設計原則 58)
'   WithEvents 参照と自前ペインを断つ。UserForm_Terminate から呼ぶ。
' ============================================================
Public Sub Shutdown()
    Debug.Print "Wv2NavBar.Shutdown 開始"
    If Not m_barPane Is Nothing Then
        m_barPane.ClearWebMsgSink   ' 9.12d でシンク登録するので先に切る作法を入れておく
        m_barPane.Ctrl_Close
        Set m_barPane = Nothing
    End If
    Set m_browser = Nothing   ' WithEvents 購読を解除
    m_ready = False
    Debug.Print "Wv2NavBar.Shutdown 完了"
End Sub


Public Property Get IsReady() As Boolean
    IsReady = m_ready
End Property


' ============================================================
' ★Wv2Browser イベントハンドラ (第9.12c: 全体同期でナビバーへ反映)★
'
'   TabUrlChanged : アクティブタブの URL / 戻る・進むの可否が変わった。
'                   ★index フィルタ必須★ 裏のタブのイベントは無視する
'                   (無視しないと URL 欄が裏タブの URL に化ける)。
'                   NavigationCompleted と HistoryChanged の両方から撃たれるので
'                   1 回の遷移で複数回来るのが正常 (全体同期なので無害)。
'
'   ActiveChanged : タブが切り替わった。★新しいアクティブタブの URL は
'                   TabUrlChanged では飛んでこない★ (そのタブでは何も起きて
'                   いないため)。よって Browser に現在状態を問い合わせて送る。
' ============================================================
Private Sub m_browser_TabUrlChanged(ByVal index As Long, ByVal url As String, _
                                    ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    ' --- ★index フィルタ★ 裏タブのイベントは無視する ---
    If m_browser Is Nothing Then Exit Sub
    If index <> m_browser.ActiveIndex Then
        Debug.Print "Wv2NavBar[EVT] TabUrlChanged(index=" & index & ") → 裏タブなので無視 " & _
                    "(ActiveIndex=" & m_browser.ActiveIndex & ")"
        Exit Sub
    End If

    Debug.Print "Wv2NavBar[EVT] TabUrlChanged(index=" & index & ", url=" & url & _
                ", back=" & canGoBack & ", fwd=" & canGoForward & ") → PushNavSyncToJs"

    ' イベントが運んできた値をそのまま使う (Browser に問い合わせ直さない)
    SendNavSync url, canGoBack, canGoForward
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "Wv2NavBar[EVT] ActiveChanged(index=" & index & ") → PushNavSyncToJs"
    PushNavSyncToJs
End Sub

' --- TabClosed (第9.12c で追加、案A) ---
'   ★なぜ必要か★ CloseTab は「アクティブにするタブがない」場合 (= 最後の 1 タブを
'   閉じたとき) ActivateTab を呼ばないので ActiveChanged が撃たれない。ActiveChanged
'   だけを購読していると、タブ 0 になったことを知らされず URL 欄が固まる (実機で発覚)。
'   TabClosed も購読して PushNavSyncToJs すれば、タブ 0 なら URL 空・全ボタン無効に
'   なる (PushNavSyncToJs が既にタブ 0 を扱っている)。
'   タブが残っている場合は ActiveChanged と二重に撃たれるが、全体同期方式なので無害。
Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "Wv2NavBar[EVT] TabClosed(index=" & index & ") → PushNavSyncToJs"
    PushNavSyncToJs
End Sub


' ============================================================
' OnPaneWebMessage (第9.12c、案P: m_barPane の WebMessage シンク受け口)
'
'   ナビバー JS の postMessage が Wv2Pane 経由でここに届く。
'   第9.12c では "ready" だけを解釈する。クリック系 (back/forward/reload/navigate)
'   のコマンド解釈は 9.12d で追加する。
' ============================================================
Public Sub OnPaneWebMessage(ByVal pane As Wv2Pane, ByVal source As String, ByVal msg As String)
    Debug.Print "Wv2NavBar.OnPaneWebMessage: msg=" & msg

    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        Exit Sub
    End If

    Dim cmd As String
    cmd = JsonGetStr(msg, "cmd")

    Select Case cmd

        Case "ready"
            Debug.Print "  → ready 受信。初期同期 navsync を送信する。"
            PushNavSyncToJs

        Case "back"
            ' --- 第9.12d/9.24: ← クリック (postMessage フォールバック経路) ---
            '   ★第9.24: 処理を HostBack へ一元化 (案C-1)★ JS がハイブリッドで
            '   hostObjects を掴めなかった時だけここへ落ちる。GoBackActive の
            '   呼び出しと失敗ログは HostBack 内に集約済みなので、ここは呼び捨て
            '   委譲のみ (戻り値は受けない)。遷移が始まれば HistoryChanged /
            '   NavigationCompleted → TabUrlChanged → PushNavSyncToJs でナビバーが
            '   自動追従する (全体同期) のは従来どおり。
            Debug.Print "  → back [postmsg fallback → HostBack]"
            HostBack

        Case "forward"
            ' --- 第9.12d/9.24: → クリック (postMessage フォールバック経路) ---
            Debug.Print "  → forward [postmsg fallback → HostForward]"
            HostForward

        Case "reload"
            ' --- 第9.12d/9.24: [更新] クリック (postMessage フォールバック経路) ---
            Debug.Print "  → reload [postmsg fallback → HostReload]"
            HostReload

        Case "navigate"
            ' --- 第9.12d/9.24: URL 欄で Enter (postMessage フォールバック経路) ---
            '   ★第9.24: 処理を HostNavigate へ一元化 (案C-1)★ 空 URL ガード /
            '   NavigateActive の呼び出し / 失敗ログはすべて HostNavigate 内へ
            '   移した (論点3)。ここでは JSON から url を復元して呼び捨て委譲する
            '   だけ。JSON パース (JsonGetStr) はフォールバック経路の引数復元に
            '   必要なので Case 側に残る (hostObjects 経路では JS が文字列を直接
            '   渡すためパース不要)。
            '   ★スキーム補完・検索エンジンへの振り分けは Browser 側 (NormalizeUrl)
            '   の責務★ ここも Host メソッドも入力文字列をそのまま流す (案3-A)。
            Dim url As String
            url = JsonGetStr(msg, "url")
            Debug.Print "  → navigate(url=" & url & ") [postmsg fallback → HostNavigate]"
            HostNavigate url

        Case ""
            Debug.Print "  → cmd を取り出せなかった (msg=" & msg & ")"

        Case Else
            Debug.Print "  → 未知のコマンド: " & cmd

    End Select
End Sub


' ============================================================
' HostBack (第9.24、hostObjects 経路の back 受け口)
'
'   JS から window.chrome.webview.hostObjects.sync.NavBar.HostBack() で直接
'   呼ばれる (postMessage を介さない同期 RPC)。中身は OnPaneWebMessage の back
'   ケースと同一で、m_browser.GoBackActive を同期で呼ぶだけ。
'
'   ★同期呼び出しの安全性 (案4-A)★ GoBackActive は View_GoBack を撃つだけで
'     完了待ちを含まないため、JS->VBA の sync 直接呼び出しでも 仕様事実 14 の
'     デッドロックに嵌まらない。遷移が始まると HistoryChanged /
'     NavigationCompleted → TabUrlChanged → PushNavSyncToJs でナビバーが全体
'     同期で追従する。
'
'   ★引数ゼロで確定 (第9.25 → 9.25b)★ 9.24 では『引数ゼロの Public Function は
'     hostObjects の sync プロキシからプロパティ取得と誤認され、取得の時点で 1 回実行 →
'     フォールバックでもう 1 回実行、という二重実行を招くのではないか』という懸念から
'     Optional ByVal reserved As Long を持たせていた。9.25 で専用プローブによる実測を
'     行った結果、sync プロキシは取得の時点では VBA を呼ばず、判断を実際の () まで
'     遅延していることが確認できた (仕様事実 22)。よって reserved は不要と確定し、
'     素の引数ゼロに戻した。JS からは nb.HostBack() の形で呼ぶ。
' ============================================================
Public Function HostBack() As Boolean
    Debug.Print "Wv2NavBar.HostBack: (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostBack = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.GoBackActive()
    If Not ok Then
        Debug.Print "    [警告] GoBackActive 失敗/戻れない"
    End If
    HostBack = ok
End Function


' ============================================================
' HostForward (第9.24、hostObjects 経路の forward 受け口)
'
'   JS から hostObjects.sync.NavBar.HostForward() で直接呼ばれる。中身は
'   forward ケースと同一の m_browser.GoForwardActive。View_GoForward を撃つ
'   だけで完了待ちがない (案4-A)。第9.25 で reserved を削除し引数ゼロに戻した。
' ============================================================
Public Function HostForward() As Boolean
    Debug.Print "Wv2NavBar.HostForward: (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostForward = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.GoForwardActive()
    If Not ok Then
        Debug.Print "    [警告] GoForwardActive 失敗/進めない"
    End If
    HostForward = ok
End Function


' ============================================================
' HostReload (第9.24、hostObjects 経路の reload 受け口)
'
'   JS から hostObjects.sync.NavBar.HostReload() で直接呼ばれる。中身は
'   reload ケースと同一の m_browser.ReloadActive。View_Reload を撃つだけで
'   完了待ちがない (案4-A)。第9.25 で reserved を削除し引数ゼロに戻した。
' ============================================================
Public Function HostReload() As Boolean
    Debug.Print "Wv2NavBar.HostReload: (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostReload = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.ReloadActive()
    If Not ok Then
        Debug.Print "    [警告] ReloadActive 失敗"
    End If
    HostReload = ok
End Function


' ============================================================
' HostNavigate (第9.24、hostObjects 経路の navigate 受け口)
'
'   JS から hostObjects.sync.NavBar.HostNavigate(url) で直接呼ばれる。中身は
'   navigate ケースと同一の m_browser.NavigateActive。
'
'   ★★今回の検証の目玉 = 文字列引数★★ TabBar の 3 メソッドは Long しか
'     受け取らなかったので、hostObjects.sync 経由で JS→VBA へ String を渡すのは
'     ここが初めて。JS の string は VARIANT(BSTR) として届くはずで、postMessage
'     経路のような JSON エスケープを一切通らない。日本語の検索語や記号入り URL が
'     欠落なく届けば 仕様事実 21 として記録する。→ 判定用に受け取った url を
'     そのまま Debug.Print する (イミディエイトで文字化けの有無が見える)。
'
'   ★空 URL ガードはここに一元化 (論点3)★ hostObjects 経路・postMessage 経路の
'     どちらから来ても同じガードを通す。
'
'   ★同期呼び出しの安全性 (案4-A)★ NavigateActive は NormalizeUrl (純粋変換) と
'     View_Navigate (撃ちっぱなし) のみで完了待ちを含まない。
' ============================================================
Public Function HostNavigate(ByVal url As String) As Boolean
    Debug.Print "Wv2NavBar.HostNavigate: url=" & url & " (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostNavigate = False
        Exit Function
    End If
    If Len(Trim$(url)) = 0 Then
        Debug.Print "    [警告] URL が空 (無視)"
        HostNavigate = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.NavigateActive(url)
    If Not ok Then
        Debug.Print "    [警告] NavigateActive 失敗/無視 (アクティブタブなし?)"
    End If
    HostNavigate = ok
End Function


' ============================================================
' PushNavSyncToJs (第9.12c、Browser に現在状態を問い合わせて JS へ送る)
'
'   アクティブタブの URL / 戻れるか / 進めるか を Browser の読み取り口
'   (9.12a で用意) から取得して送る。タブ 0 のときは URL 空・両ボタン無効。
'
'   呼び出し元:
'     - ready 受信時 (初期同期)
'     - ActiveChanged (タブ切替。新タブの状態はイベントが運んでこないので問い合わせ)
' ============================================================
Private Sub PushNavSyncToJs()
    If m_browser Is Nothing Then Exit Sub

    Dim url As String
    Dim canBack As Boolean
    Dim canFwd As Boolean

    If m_browser.TabCount = 0 Or m_browser.ActiveIndex < 1 Then
        ' タブ 0 (アクティブなし): URL 空・両ボタン無効
        url = ""
        canBack = False
        canFwd = False
    Else
        url = m_browser.GetActiveUrl()
        canBack = m_browser.GetActiveCanGoBack()
        canFwd = m_browser.GetActiveCanGoForward()
    End If

    SendNavSync url, canBack, canFwd
End Sub


' ============================================================
' SendNavSync (第9.12c、navsync JSON を組み立てて JS へ送る)
'
'   送る JSON:
'     {"type":"navsync","url":"...","canBack":true,"canForward":false}
'
'   PushNavSyncToJs (Browser に問い合わせる版) と、TabUrlChanged ハンドラ
'   (イベントが運んできた値をそのまま使う版) の両方から呼ばれる共通の送信部。
' ============================================================
Private Sub SendNavSync(ByVal url As String, ByVal canBack As Boolean, ByVal canFwd As Boolean)
    If m_barPane Is Nothing Then Exit Sub
    If Not m_ready Then Exit Sub

    ' --- hasTab: アクティブタブがあるか (第9.12c) ---
    '   タブ 0 のときはリロードボタンも無効にする。戻る/進むは canBack/canForward が
    '   もともと False になるので自動的に無効になるが、リロードはそれらと独立なため
    '   明示的に伝える必要がある。
    Dim hasTab As Boolean
    If Not m_browser Is Nothing Then
        hasTab = (m_browser.TabCount > 0 And m_browser.ActiveIndex >= 1)
    End If

    ' about:blank は「まだ何も読み込んでいない」状態なので URL 欄には出さない
    ' (Controller 生成直後に一瞬だけ来る。仕様事実 6)
    Dim shown As String
    If url = "about:blank" Then
        shown = ""
    Else
        shown = url
    End If

    Dim json As String
    json = "{""type"":""navsync"",""url"":""" & JsonEscape(shown) & """" & _
           ",""canBack"":" & JsonBool(canBack) & _
           ",""canForward"":" & JsonBool(canFwd) & _
           ",""hasTab"":" & JsonBool(hasTab) & "}"

    Dim hr As Long
    hr = m_barPane.View_PostWebMessageAsString(json)
    If hr <> 0 Then
        Debug.Print "Wv2NavBar.SendNavSync: PostWebMessageAsString 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2NavBar.SendNavSync: 送信 OK"
        Debug.Print "  json=" & json
    End If
End Sub


' ============================================================
' WaitPaneReady ? Controller の生成完了を待つ
'   ★ここは Init の中 (= イベントハンドラの外) なので同期待ちしてよい★
'   イベントハンドラ内で同じことをするとデッドロックする (仕様事実 14 / 設計原則 63)。
' ============================================================
Private Function WaitPaneReady(ByVal pane As Wv2Pane) As Boolean
    Dim t As Single
    t = Timer
    Do While (Not pane.IsReady) And (Not pane.IsFailed) And ((Timer - t) < READY_TIMEOUT)
        DoEvents
    Loop
    WaitPaneReady = pane.IsReady
End Function


' ============================================================
' WriteNavBarHtml ? navbar.html を UTF-8 (BOM 無し) で書き出す
'   日本語や将来の動的差し込みに備え、Wv2TabBar と同じ ADODB.Stream 方式。
' ============================================================
Private Function WriteNavBarHtml() As Boolean
    On Error GoTo eh

    ' フォルダがなければ作成
    If Dir(m_folderPath, vbDirectory) = "" Then
        MkDir m_folderPath
        Debug.Print "Wv2NavBar.WriteNavBarHtml: フォルダ作成 " & m_folderPath
    End If

    Dim fullPath As String
    fullPath = m_folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & NAVBAR_FILE

    Dim html As String
    html = BuildNavBarHtml()

    ' ADODB.Stream で UTF-8 (BOM 無し) 書き出し
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2                ' adTypeText
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText html
    ' テキストで書き、バイナリで読み直して先頭 3 バイト (BOM) を捨てる
    bin.Position = 0
    bin.Type = 1                ' adTypeBinary
    bin.Position = 3            ' EF BB BF をスキップ
    Dim bytes() As Byte
    bytes = bin.Read
    bin.Close

    Dim outStream As Object
    Set outStream = CreateObject("ADODB.Stream")
    outStream.Type = 1          ' adTypeBinary
    outStream.Open
    outStream.Write bytes
    outStream.SaveToFile fullPath, 2   ' 2 = adSaveCreateOverWrite
    outStream.Close

    Debug.Print "Wv2NavBar.WriteNavBarHtml: 書き出し OK " & fullPath & " (" & Len(html) & " 文字, UTF-8)"
    WriteNavBarHtml = True
    Exit Function
eh:
    Debug.Print "Wv2NavBar.WriteNavBarHtml: エラー " & Err.Number & " " & Err.Description
    WriteNavBarHtml = False
End Function


' ============================================================
' BuildNavBarHtml ? ナビバー HTML を組み立てる (9.12b は静的)
'
'   9.12b の時点では「見た目だけ」。ボタンは押せず、URL 欄も更新されない。
'     - 戻る / 進む / リロード ボタン (戻る・進むは disabled 表示)
'     - URL 入力欄 (プレースホルダのみ)
'   第9.12c で navsync を受けて URL とボタン活性を反映するようにした。
'   クリック/入力の送信は 9.12d (今はコメントアウトの器のまま)。
'
'   ★URL 欄の編集中保護★ ユーザーが URL を入力している最中に裏で同期が走ると
'     入力途中の文字列が消えてしまう。document.activeElement が URL 欄のときは
'     value を書き換えない。
'
'   ★ボタンの記号★ 環境依存を避けるため HTML 数値文字参照で書く
'     (&#8592; = ← / &#8594; = → / &#10227; = 循環矢印)。
' ============================================================
Private Function BuildNavBarHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & _
            "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    s = s & "  *{box-sizing:border-box;}" & vbLf
    s = s & "  body{display:flex;align-items:center;gap:4px;padding:0 6px;" & _
            "overflow:hidden;background:#fafafa;border-bottom:1px solid #ccc;}" & vbLf
    s = s & "  button{width:30px;height:26px;border:1px solid transparent;" & _
            "border-radius:4px;background:transparent;cursor:pointer;" & _
            "font-size:15px;color:#333;line-height:1;padding:0;flex:none;}" & vbLf
    s = s & "  button:hover:not(:disabled){background:#e6e6e6;}" & vbLf
    s = s & "  button:disabled{color:#bbb;cursor:default;}" & vbLf
    s = s & "  #url{flex:1;height:26px;padding:0 10px;border:1px solid #ccc;" & _
            "border-radius:13px;background:#fff;font-size:13px;color:#222;" & _
            "outline:none;min-width:0;}" & vbLf
    s = s & "  #url:focus{border-color:#0078d4;box-shadow:0 0 0 1px #0078d4;}" & vbLf
    s = s & "  #url:disabled{background:#f0f0f0;color:#999;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "  <button id=""back""    title=""戻る""       disabled>&#8592;</button>" & vbLf
    s = s & "  <button id=""forward"" title=""進む""       disabled>&#8594;</button>" & vbLf
    s = s & "  <button id=""reload""  title=""再読み込み"" disabled>&#10227;</button>" & vbLf
    s = s & "  <input  id=""url"" type=""text"" placeholder=""URL を入力"" spellcheck=""false"" disabled>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "  var W = window.chrome && window.chrome.webview;" & vbLf
    s = s & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    ' --- 第9.24: hostObjects 優先 + postMessage フォールバック (TabBar 9.21 と同型) ---
    '   HO() は hostObjects.sync が取れれば返し、なければ null (設定タブの null
    '   ガードと同型)。navHost(msgObj, callHost) は 4 コマンド共通の三段構え。
    '   フォールバックで送る JSON は 9.12d と完全に同一 (VBA 側の薄い Case が
    '   そのまま受け止める)。
    s = s & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    s = s & "  function navHost(msgObj, callHost){" & vbLf
    s = s & "    var ho = HO();" & vbLf
    s = s & "    if(ho && ho.NavBar){" & vbLf
    s = s & "      try{ callHost(ho.NavBar); console.log('[host] '+msgObj.cmd); return; }" & vbLf
    s = s & "      catch(e){ console.log('[host-fail→postmsg] '+e); postCmd(msgObj); return; }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    console.log('[postmsg] '+msgObj.cmd); postCmd(msgObj);" & vbLf
    s = s & "  }" & vbLf
    ' ★編集中フラグは全ハンドラより先に宣言する (巻き上げに頼らない)★
    s = s & "  var editing = false;" & vbLf
    ' ★lastUrl (第9.12d)★ 直近の navsync が伝えてきた「今表示すべき URL」。
    '   Escape で編集を破棄するとき、VBA に問い合わせずここから即座に復元する。
    '   navsync のたびに更新されるので常に正しい値を持つ (表示用キャッシュ)。
    s = s & "  var lastUrl = '';" & vbLf
    ' --- ボタン 3 種 (第9.12d で有効化 → 第9.24 で hostObjects 優先へ) ---
    s = s & "  document.getElementById('back').onclick = function(){ navHost({cmd:'back'}, function(nb){ nb.HostBack(); }); };" & vbLf
    s = s & "  document.getElementById('forward').onclick = function(){ navHost({cmd:'forward'}, function(nb){ nb.HostForward(); }); };" & vbLf
    s = s & "  document.getElementById('reload').onclick = function(){ navHost({cmd:'reload'}, function(nb){ nb.HostReload(); }); };" & vbLf
    ' --- URL 欄: Enter で遷移、Escape で編集破棄 (第9.12d) ---
    '   ★Enter で editing = false★ 遷移が始まると navsync で URL 欄が正規化後の
    '   URL に置き換わる。editing を落とさないと打った文字列のまま固まる。
    '   スキーム補完は VBA 側 (NormalizeUrl) がやるので、ここは値をそのまま送る。
    s = s & "  document.getElementById('url').onkeydown = function(ev){" & vbLf
    s = s & "    if(ev.key === 'Enter'){" & vbLf
    s = s & "      var v = this.value;" & vbLf
    s = s & "      editing = false;" & vbLf
    s = s & "      if(v && v.length){ navHost({cmd:'navigate', url:v}, function(nb){ nb.HostNavigate(v); }); }" & vbLf
    s = s & "    } else if(ev.key === 'Escape'){" & vbLf
    s = s & "      editing = false;" & vbLf
    s = s & "      this.value = lastUrl;   /* 編集破棄 → 今表示すべき URL に即復元 */" & vbLf
    s = s & "      this.blur();" & vbLf
    s = s & "    }" & vbLf
    s = s & "  };" & vbLf
    ' --- 第9.12c: VBA からの navsync を受けて URL 欄とボタン活性を更新 ---
    ' ★編集中フラグ (第9.12c 修正)★
    '   当初 document.activeElement !== u で判定していたが、WebView2 の
    '   単独ドキュメントでは #url が activeElement になったまま戻らないことがあり、
    '   URL 欄が二度と更新されなくなった (実機で発覚)。
    '   「ユーザーが実際に編集を始めたか」を明示フラグで持つ方が確実。
    '     - focus しただけでは編集中にしない (クリックしただけで固まらないように)
    '     - キー入力 (input イベント) が来たら編集中にする
    '     - blur / Enter / Escape で編集中を解除する
    s = s & "  function applyNavSync(st){" & vbLf
    s = s & "    var u = document.getElementById('url');" & vbLf
    ' ★lastUrl は編集中でも必ず更新する★ 編集中に裏で遷移が起きても、
    ' Escape を押したら「今の本当の URL」に戻れるようにするため。
    s = s & "    lastUrl = st.url || '';" & vbLf
    s = s & "    var b = document.getElementById('back');" & vbLf
    s = s & "    var f = document.getElementById('forward');" & vbLf
    s = s & "    var r = document.getElementById('reload');" & vbLf
    ' --- ボタン類は常に同期してよい (フォーカスを奪わない) ---
    s = s & "    b.disabled = !st.canBack;" & vbLf
    s = s & "    f.disabled = !st.canForward;" & vbLf
    s = s & "    r.disabled = !st.hasTab;" & vbLf
    ' --- ★URL 欄は編集中なら一切触らない (value も disabled も)★ ---
    '   ★★実機で発覚したバグ★★ disabled を代入すると input が blur し、
    '   blur ハンドラが editing = false にしてしまうため、次の同期で入力中の
    '   文字列が消える (リンククリックで新タブが開いた瞬間に発生)。
    '   対策は 2 段構え:
    '     (1) 編集中は URL 欄に触らない (value も disabled も)
    '     (2) disabled は値が変わるときだけ代入する (無駄な blur を防ぐ)
    s = s & "    if(!editing){" & vbLf
    s = s & "      var want = !st.hasTab;" & vbLf
    s = s & "      if(u.disabled !== want){ u.disabled = want; }" & vbLf
    s = s & "      u.value = st.url || '';" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  (function(){" & vbLf
    s = s & "    var u = document.getElementById('url');" & vbLf
    s = s & "    u.addEventListener('input', function(){ editing = true; });" & vbLf
    s = s & "    u.addEventListener('blur',  function(){ editing = false; });" & vbLf
    s = s & "  })();" & vbLf
    s = s & "  if(W){" & vbLf
    s = s & "    W.addEventListener('message', function(ev){" & vbLf
    s = s & "      var data = ev.data;" & vbLf
    s = s & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    s = s & "      if(data && data.type === 'navsync'){ applyNavSync(data); }" & vbLf
    s = s & "    });" & vbLf
    s = s & "  }" & vbLf
    ' --- ready ハンドシェイク: リスナー登録後に ready を送る (設計原則 61) ---
    s = s & "  postCmd({cmd:'ready'});" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildNavBarHtml = s
End Function

