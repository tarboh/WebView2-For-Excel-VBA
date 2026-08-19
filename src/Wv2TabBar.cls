VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2TabBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  第9.32b 段階 (ホイール横スクロール + スクロール位置の維持) ---
'
'   第9.32b のスコープ (★JS のみ。CSS も VBA も 1 行も触らない★):
'     (1) #tabs に wheel ハンドラを付けて縦ホイールを横スクロールに変換する
'     (2) render() の前後で scrollLeft を退避・復元する
'     (3) アクティブタブが画面外なら可視範囲へ寄せる (nearest)
'
'   ★背景 1: 自動変換は起きなかった (仕様事実 28)★
'     第9.32 で overflow-x:auto のまま ::-webkit-scrollbar を隠し、
'     「Chromium は縦スクロール不可の横コンテナでホイールを横へ変換する」と
'     見込んだが、★実機では反応しなかった★。祖先の overflow 構成やヒット対象に
'     左右されるため当てにできない。明示的な wheel ハンドラを置くのが確実。
'
'   ★背景 2: render() が scrollLeft を 0 に戻す★
'     render() は先頭で tabsEl.innerHTML = '' を実行するので、★全体同期のたびに
'     横スクロール位置が左端へ戻る★。タブが少ないうちはスクロール自体が発生
'     しないので誰も気づかなかったが、スクロール可能な状態では顕在化する。
'     しかも PushSyncToJs は★タイトルが変わるたび★にも飛ぶので、右のほうを
'     見ている最中にどこかのタブの読み込みが終わっただけで左端へ戻ってしまう。
'     さらに「+ で追加した新しいタブが画面外にいて見えない」も同時に起きる。
'
'   合意した論点 (3 件すべて Claude 推奨どおり):
'     1 ホイール = ★#tabs に wheel ハンドラ★ (第9.32 の論点 4 は撤回しない)
'                  3px の細いバーに戻すとタブが潰れる問題が再発するため。
'                  ハンドラは #tabs に 1 回だけ付ける。★innerHTML を書き換えても
'                  #tabs 自体は入れ替わらないのでリスナーは生き残る★
'                  {passive:false} を指定して preventDefault する。
'     2 スクロール量 = ★deltaY をそのまま加算★ 1 ノッチ約 100px = タブ 1 枚強。
'                  足りなければ後から数値 1 つで調整できる。
'     3 位置の維持 = ★退避・復元 + アクティブを nearest で寄せる★
'                  復元だけだと新規タブが画面外のままで見えない。
'                  scrollIntoView({block:'nearest',inline:'nearest'}) は
'                  ★既に見えているときは動かない★ので副作用が最小。
'                  順序は「復元 → nearest 寄せ」。
'
'   ★事後報告 1: nearest 寄せに「アクティブが変わったときだけ」の条件を足した★
'     論点 3 のままだと、右のほうを見ている最中にタイトル更新の同期が来ただけで
'     アクティブタブの位置まで引き戻されてしまう (論点 3 の狙い「勝手に動く副作用を
'     最小にする」と逆行する)。そこでトップレベルに lastActive (数値 1 個) を持ち、
'     ★active が前回と違うときだけ★ scrollIntoView する。数値 1 個の記憶なので
'     全体同期方式 (設計原則 60) の趣旨は損なわない。
'
'   ★事後報告 2: 第9.32 のヘッダーの記述を訂正する★
'     第9.32 のブロックに「設計原則 82 個・仕様事実 26 個ともに増減なし」と
'     書いたが、実機検証の結果★設計原則 83・84 と仕様事実 27・28 が増えた★。
'     正しくは第9.32b 完了時点で設計原則 84 個・仕様事実 28 個。開発メモが正。
'       設計原則 83: 体感の遅さは「速くする」より先に「順序」を疑う
'       設計原則 84: 次の全体同期が必ず上書きするなら、先行表示は状態保持ではない
'       仕様事実 27: hostObjects.sync 呼び出し中、呼び出し元ペインの JS は完全に
'                    ブロックされ、PostWebMessageAsString で送ったメッセージも
'                    受け取れない (第9.32 のレイテンシの原因)
'       仕様事実 28: WebView2 (Chromium 151) では overflow-x:auto の横コンテナで
'                    縦ホイールが横スクロールへ自動変換されない
'
'   ★触らないもの★ CSS 全部 / VBA のロジック全部 / Browser の 5 イベント /
'     ready ハンドシェイク / 全体同期 / SafeTimer 経路 / postMessage
'     フォールバック / ドラッグ処理 / 楽観的更新 (第9.32) / Wv2Browser /
'     Wv2Pane / Wv2NavBar / UserForm1。
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  第9.32 段階 (体感: タブ切替の楽観的更新 + スクロールバーの整理) ---
'
'   第9.32 のスコープ (★JS と CSS のみ。VBA のロジックは 1 行も触らない★):
'     (1) タブ切替の★体感レイテンシ★を解消する (楽観的更新)
'     (2) タブが増えたときにタブバー内へ出る横スクロールバーを隠す
'     (3) D&D 開始時のアクティブ化は★入れない★(現状維持を明示的に決定)
'
'   ★背景: レイテンシの原因は構造的だった (第9.31 の実機検証で判明)★
'     従来の呼び出し連鎖:
'       JS: tab.onclick
'         └→ ho.TabBar.HostActivate(n)   ← ここで JS が完全にブロックする
'               └→ VBA: HostActivate → m_browser.ActivateTab   ← 重い (表示切替)
'                     └→ RaiseEvent ActiveChanged
'                           └→ TabBar.m_browser_ActiveChanged → PushSyncToJs
'                                 └→ 全タブの DocumentTitle を COM で取得
'                                 └→ View_PostWebMessageAsString(json) ← 送るだけ
'               ←── ようやく HostActivate が返る
'         └→ JS のブロック解除 → はじめて message イベントが処理される
'               └→ render() が DOM 全消し再生成 → ★ここで初めて色が変わる★
'     hostObjects.sync は同期 RPC なので、VBA の処理が全部終わるまでタブバーの
'     JS は 1 命令も進めない。PushSyncToJs が投げたメッセージも JS がブロック中
'     なので受け取れない。★「本体切替 → ログ → タブバー」の順は設計の必然★
'     だった。体感の遅れは偶然ではない。
'
'   ★VBA 側で順序を入れ替えてはいけない★
'     HostActivate は JS からの同期呼び出しの最中なので、その中で PushSyncToJs を
'     先に呼んでも JS がブロック中で受け取れない。無理に押し込むと仕様事実 14 の
'     デッドロック領域に入る。★順序の入れ替えは JS 側でやるのが正着★
'
'   合意した論点 (7 件すべて Claude 推奨どおり):
'     1 楽観的更新 = ★入れる★ クリック直後に JS が自分で active クラスを
'                    付け替えてから Host を呼ぶ。VBA から後で来る全体同期が
'                    必ず上書きするので★JS に状態を持たせることにはならず★、
'                    設計原則 60 (全体同期) と衝突しない。失敗時は次の同期で戻る。
'     2 ★描画待ち = 二重 requestAnimationFrame★ (この回で最も重要な判断)
'                    DOM を変えた直後に同期 COM 呼び出しでブロックすると、
'                    ★描画される前に固まる★ 恐れがあり論点 1 が無意味になる。
'                    setTimeout(...,0) は「次のタスク」を保証するだけで描画は
'                    保証しない。二重 rAF は「1 フレーム描かれた後」が保証される
'                    定石。代償は切替が約 16ms 遅れるだけ。
'     3 close     = ★入れない★ 閉じるは DOM から要素を消す操作で失敗時の
'                    巻き戻しが複雑。まず activate だけ入れて効き目を見る
'                    (設計原則 82 の趣旨: 何が効いたか分かる形にする)。
'     4 スクロールバー = ★完全に隠す★ #tabs::-webkit-scrollbar{display:none}
'                    overflow-x:auto は残すのでマウスホイールで横スクロールできる
'                    (Chromium は縦スクロール不可の横コンテナでホイールを横へ変換)。
'                    3px の細いバーを残す案は、タブバー直下に線が常駐して別の
'                    意味で目障りなので採らない。
'     5 タブ幅    = ★min-width:80px 維持★ 60px まで下げても「読めない幅のタブ」が
'                    増えるだけで、溢れる状況自体は変わらない。
'     6 D&D 時の  = ★アクティブにしない (現状維持)★
'       アクティブ化   ・並べ替えは「見た目の整理」であって「見る対象の変更」ではない
'                      ・ActivateTab は重いので、掴むたびに走ると体感が悪化する
'                      ・本家がアクティブ化するのはタブを引き出して別ウィンドウ化
'                        する機能があるため。その機能を持たない本実装では追随理由なし
'     7 検証      = Test_9_32_Help (目視)。
'
'   ★事後報告: HO() をトップレベルへ移した★
'     従来 HO() は render() 内の for ループの中で宣言されており、巻き上げで
'     たまたま動いていた (render のスコープにしか存在しない)。第9.32 で
'     sendActivate() を render の外に出す必要が生じたため、HO() も
'     ★postCmd の隣 (トップレベル) へ移設★した。挙動は同じだがスコープが
'     素直になる。render 内の宣言は削除済み。
'
'   ★触らないもの★ VBA のロジック全部 (Host メソッド 3 つ / OnPaneWebMessage /
'     PushSyncToJs / Init / Resize / Shutdown) / Browser の 5 イベント /
'     ready ハンドシェイク / 全体同期 / SafeTimer 経路 / postMessage
'     フォールバック / ドラッグ処理 / Wv2Browser / Wv2Pane / Wv2NavBar /
'     UserForm1。設計原則 82 個・仕様事実 26 個ともに増減なし。
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  第9.31 段階 (見た目: タブのホバー + 閉じるボタンの表示制御) ---
'
'   第9.31 のスコープ (★CSS のみ。JS も VBA ロジックも 1 行も触らない★):
'     (A-1) タブにマウスを乗せたとき背景を明るくする
'     (A-2) 閉じるボタンを「アクティブ or ホバー時のみ」表示する (Chrome 相当)
'
'   ★変更箇所は BuildTabBarHtml の <style> 内だけ★ render() / ドラッグ処理 /
'     Host メソッド / OnPaneWebMessage / Init / 全体同期 はすべて無変更。
'
'   合意した論点 (8 件すべて Claude 推奨どおり):
'     1 実現方法   = ★CSS の :hover のみ。JS は触らない★
'                    render() が innerHTML='' で DOM を毎回捨てるため、JS 側に
'                    ホバー状態を持たせると同期のたびに必ず失われる。CSS なら
'                    状態ゼロで済み、全体同期方式 (設計原則 60) と衝突しない。
'     2 ホバー背景 = ★非アクティブのみ★ #e7e7e7 → #efefef。
'                    アクティブは既に #fff で最も明るいので変えない (Chrome も同じ)。
'                    .tab:not(.active):hover と書くことで詳細度 0,3,0 となり、
'                    .tab.active (0,2,0) より強い。★記述順に依存しない★
'     3 表示条件   = ★アクティブ or ホバー★ (Chrome 相当)。
'                    アクティブタブは「今すぐ閉じたい」対象なので、マウスを
'                    乗せる前から押せるべき。
'     4 ★非表示の方法 = visibility:hidden★ (この回で最も重要な判断)
'                    ・display:none だと 18px + margin 6px が消えてタイトル幅が
'                      伸縮し、ホバーのたびに文字がガタつく。
'                    ・opacity:0 だと見えないボタンがクリックできてしまう (誤爆)。
'                    ・visibility:hidden は場所を保持したまま消え、ポインタ
'                      イベントも無効になる。
'                    代償は「タイトルに使える幅が常に 24px 減る」ことだが、
'                    min-width:80px なので許容範囲と判断した。
'     5 遷移       = ★背景色のみ .12s★
'                    visibility は離散値なので transition が効かず、閉じる
'                    ボタンは即時切替になる (これは意図どおりの挙動)。
'     6 ×の色     = ★現状維持★ (#888 / hover で #ccc 背景)。
'                    表示制御を入れた後の見え方を実機で見てから判断する方が確実。
'     7 ドラッグ中 = ★何も足さない★ .dragging 中のホバー抑制は先回りしない
'                    (第9.27 論点 7 の教訓: 先回りの補正は必要だったか永久に
'                    分からなくなる)。実機で不快なら次段で足す。
'     8 検証       = Test_9_31_Help (目視手順)。併せて★第9.30 の宿題 1 =
'                    v0_5_2 の通し動作確認★ (POST リンクの新タブ展開 /
'                    プリウォーム / ドラッグ並べ替え) を同じ手順で消化する。
'
'   ★触らないもの★ Browser の 5 イベント / ready ハンドシェイク / 全体同期 /
'     hostObjects 3 メソッド (HostActivate/HostClose/HostReorder) / postMessage
'     フォールバック / SafeTimer 経路 / 3 段レイアウト / Wv2Browser / Wv2Pane /
'     Wv2NavBar / UserForm1。設計原則 82 個・仕様事実 26 個ともに増減なし。
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  第9.23 段階 (整理: activate/close/reorder の処理を Host メソッドへ一元化) ---
'
'   第9.23 のスコープ (9.21/9.22 で出揃った 3 コマンドの処理重複を解消。案C-1):
'     ★背景★ 9.21/9.22 で activate/close/reorder を hostObjects 経路へ移した際、
'       同じ処理 (ActivateTab/CloseTab/MoveTab の呼び出しと失敗ログ) が
'       「OnPaneWebMessage の Case 内」と「HostActivate/HostClose/HostReorder 内」の
'       2 箇所に重複していた。片方だけ直して片方を直し忘れる事故の温床。
'
'     ★決定 = 案C (保険を残して処理一元化)★ 3 コマンドの JS 側ハイブリッド
'       (hostObjects 優先・postMessage フォールバック) はそのまま残す。フォール
'       バックはブラウザ根幹操作 (切替/閉じる/並べ替え) の安全網として捨てない。
'       その上で、フォールバック時に通る Case activate/close/reorder の中身を、
'       対応する Host メソッド (HostActivate/HostClose/HostReorder) の呼び出しへ
'       委譲する。これで ActivateTab/CloseTab/MoveTab を呼ぶロジックと失敗ログが
'       Host メソッド 1 箇所に集約される (DRY)。
'
'     ★案C-1 (戻り値を受けない呼び捨て)★ Case 側は HostXxx を呼び捨てる。失敗
'       ログは Host メソッド内部で既に出るため Case 側の二重ログは消す。ただし
'       postMessage 経路を通った証跡として Case 入口の Debug.Print (→ activate 等)
'       は残す。JSON パース (JsonGetNum) はフォールバック経路の引数復元に必要な
'       ので Case 側に残る (これは正しい姿。hostObjects 経路では JS が数値を直接
'       渡すためパース不要、フォールバック時のみ JSON から復元する)。
'
'   ★触らないもの★ ready/newtab/settings の 3 ケース (postMessage が正着) /
'     JS 側の BuildTabBarHtml (ハイブリッド構造は無変更) / Host メソッド 3 つの
'     処理本体 / Init の attach / Browser の 5 イベント / 全体同期 / ready ハンド
'     シェイク / SafeTimer 経路 / 3 段レイアウト。設計原則 71 個・仕様事実 20 個
'     ともに増減なし (既存原則 = 案4-A / DRY の実践。新しい仕様事実なし)。
'

''''''''''''''''''''''''''''''''''
' --- Wv2TabBar.cls  第9.22 段階 (試験: close / reorder の hostObjects 化) ---
'
'   第9.22 のスコープ (9.21 の自然な発展。activate に続き close/reorder も
'     hostObjects 直接呼び出しへ試験移行し、postMessage 版と並走比較する):
'     ★狙い★ 9.21 で activate 単独の実 URL hostObjects 成立 (仕様事実 19) と
'       分離可能性が実証できた。次の一歩として close / reorder (どちらも
'       完了待ちなしの同期操作 = 案4-A) を hostObjects 経路へ横展開する。
'       これで「hostObjects に素直に向く 3 コマンド (activate/close/reorder)」が
'       出揃い、次段 (9.23 予定) で Wv2Json パーサ + OnPaneWebMessage の
'       Select Case を実際に縮約する価値があるかを、データを持って判断できる。
'
'     ★触らないもの (安定性担保)★ ready / newtab / settings は postMessage の
'       ままにする。ready は通知 (RPC でない)。newtab/settings は内部で
'       WaitPaneReady (Controller 生成の完了待ち) を通るため sync 直呼びだと
'       仕様事実 14 のデッドロック危険が上がる。→ この 3 つは 9.21 同様に不変。
'
'     ★追加した口★ (いずれも 9.21 の HostActivate と同型の写経)
'       ・Public Function HostClose(index) As Boolean
'         JS から hostObjects.sync.TabBar.HostClose(n) で直接呼ばれる。
'         中身は close ケースと同じ m_browser.CloseTab(index)。CloseTab は
'         Ctrl_Close のみで完了待ちがない (案4-A) ため sync 直呼びで安全。
'       ・Public Function HostReorder(fromIdx, toIdx) As Boolean
'         JS から hostObjects.sync.TabBar.HostReorder(from, to) で直接呼ばれる。
'         中身は reorder ケースと同じ m_browser.MoveTab(fromIdx, toIdx)。
'         MoveTab は配列操作 + PutIsVisible のみで完了待ちがない (案4-A)。
'         ★VBA 引数名は fromIdx/toIdx とする★ (From は方言で衝突報告あり。
'         JS 側は位置引数呼びなので引数名は JS から見えず自由に付けられる)。
'
'     ★JS 側はハイブリッド (9.21 の activate と同一の三段構え)★
'       close の × クリック / reorder の drop で、hostObjects.sync.TabBar が
'       取れれば HostClose/HostReorder を同期直呼びし console に [host]、
'       host 呼びが例外を投げれば catch して postCmd に落ち [host-fail→postmsg]、
'       hostObjects が無ければ従来の postCmd に落ち [postmsg]。既存の HO()
'       ヘルパ (9.21) をそのまま流用する。newtab/settings/ready の送信は無変更。
'
'   ★触らないもの★ Browser の 5 イベント / 全体同期 (VBA->JS push は postMessage
'     のまま) / ready ハンドシェイク / SafeTimer 経路 (newtab/settings) / 3 段
'     レイアウト。設計原則 71 個・仕様事実 19 個ともに増減なし (仕様事実 19 の
'     射程内。close/reorder も activate と同じ実 URL + Navigate 前 attach で成立)。
'
' --- Wv2TabBar.cls  第9.21 段階 (試験: activate の hostObjects 化) ---
'
'   第9.21 のスコープ (JS->VBA の activate だけを hostObjects 直接呼び出しへ):
'     ★狙い★ TabBar の 6 コマンド (ready/activate/close/newtab/settings/
'       reorder) のうち activate 1 個だけを試験的に hostObjects 経路へ移し、
'       既存の postMessage 版と並走比較する。close/reorder/newtab/settings/
'       ready は一切触らず postMessage のまま (安定性担保)。
'
'     ★検証の核 = 実 URL Navigate でも仕様事実 18 が成立するか★
'       仕様事実 18 は 9.18 設定タブ (NavigateToString / data:URL) で確立した。
'       TabBar は appassets.tabbar への実 URL を View_Navigate で読み込む。
'       『Navigate の前に AddHostObjectToScript しておけば、その後読み込まれた
'       JS が初回実行から hostObjects を掴める』が実 URL でも効くかを実機確認する。
'       → 効けば仕様事実 18 の射程が実 URL にも及ぶことが裏づけられ、以降の
'         close/reorder 横展開や NavBar への波及の土台になる。
'
'     ★追加した口★
'       ・Public Function HostActivate(index) As Boolean
'         JS から hostObjects.sync.TabBar.HostActivate(n) で直接呼ばれる。
'         中身は既存 activate ケースと同じ m_browser.ActivateTab(index)。
'         ActivateTab は Ctrl_PutIsVisible のみで完了待ちがない (案4-A) ため、
'         JS->VBA の sync 直接呼び出しでもデッドロックしない。
'         成功で ActiveChanged→PushSyncToJs が走り強調が追従する (全体同期不変)。
'       ・Init の Navigate 前に m_barPane.View_AddHostObjectToScript("TabBar", Me)。
'         マッピング直後・Navigate 直前に置く (仕様事実 18 の順序A)。
'
'     ★JS 側はハイブリッド (フォールバック付き)★
'       タブ本体クリックで、hostObjects.sync.TabBar が取れれば HostActivate を
'       呼び console に [host] を残す。取れなければ従来の postCmd(activate) に
'       落ち [postmsg] を残す (順序A 失敗時の保険。設定タブの null ガードと同型)。
'       activate 以外の送信は無変更。
'
'   ★触らないもの★ Browser の 5 イベント / 全体同期 (VBA->JS push は postMessage
'     のまま) / ready ハンドシェイク / SafeTimer 経路 / 3 段レイアウト /
'     close・reorder・newtab・settings の既存経路。設計原則・仕様事実は増減なし
'     (仕様事実 18 の射程内)。
'
' --- Wv2TabBar.cls  第9.18 段階 ---
'
'   第9.18 の追加事項 (設定ボタン → 設定タブ):
'     ★タブバー右端に ? (設定) ボタンを追加★ (部品3)
'       BuildTabBarHtml に #settings 要素 (margin-left:auto で右端寄せ) と
'       hover スタイル、onclick を追加。クリックで {"cmd":"settings"} を postMessage。
'     ★OnPaneWebMessage に "settings" ケースを追加★
'       Browser.RequestSettingsTabAsync を呼ぶ。★同期で開いてはいけない★
'       設定タブ生成も内部で WaitPaneReady を通るため、newtab と同じデッドロック
'       (仕様事実 14) を避けて SafeTimer 経由の遅延実行にする。
'     ★TabBar 自体は設定内容を一切知らない (疎結合維持)★
'       TabBar は「設定を開きたい」という意図を Browser に伝えるだけ。実際に
'       設定タブを組み立て・ブリッジを attach するのは Browser の責務。TabBar の
'       5 イベント購読構造・全体同期・ready ハンドシェイクは無変更。
'
' --- Wv2TabBar.cls  第9.15 段階 ---
'
'   第9.15 の追加事項 (タブのドラッグ並べ替え):
'     ★HTML5 Drag & Drop でタブを掴んで並べ替える (案B-1)★
'       BuildTabBarHtml の各タブ div に draggable=true を付け、dragstart /
'       dragover / drop / dragend を張る。ドラッグ中のタブに .dragging、ドロップ
'       先候補に .dragover の視覚クラスを付けるだけの最小実装 (磨きは候補②へ)。
'
'     ★ドロップで {"cmd":"reorder","from":N,"to":M} を postMessage★
'       from/to は 1-based。VBA 側 OnPaneWebMessage に reorder ケースを追加し、
'       Browser.MoveTab from, to を同期で呼ぶ (配列操作のみで完了待ちなし →
'       activate/close と同じく安全。newtab のような SafeTimer 迂回は不要)。
'       MoveTab 成功時に ActiveChanged → PushSyncToJs が走り (全体同期・
'       設計原則 60、案A-2b)、タブバーが並べ替え後の順序で再描画される。JS 側は
'       楽観的に DOM を動かさず、VBA からの sync を待って描き直す (真実源は VBA)。
'
' --- Wv2TabBar.cls  第9.13 段階 ---
'
'   第9.13 の変更 (整理: JSON ヘルパーの共通化):
'     Private に持っていた JsonEscape / GetJsonStr / GetJsonNum を撤去し、
'     新設した標準モジュール Wv2Json.bas の Public 関数を使うようにした。
'     Wv2NavBar にも同じ関数が重複していたのを 1 箇所に集約する。
'       GetJsonStr → Wv2Json.JsonGetStr
'       GetJsonNum → Wv2Json.JsonGetNum
'       JsonEscape → Wv2Json.JsonEscape (名前は同じ)
'     ★JsonEscape の \u 経路も修正された★ 高位文字 (日本語等) が \uXXXX に
'     化けなくなった (仕様事実 16。挙動は変わらないがログが読みやすくなる)。
'
' --- Wv2TabBar.cls  第9.11d 段階 ---
'
'   第9.11d のスコープ (JS->VBA 方向、クリック有効化):
'     タブバー上のクリックを VBA のタブ操作に繋ぐ。これで Layer 2
'     (WebView2 製タブバー) が双方向に完成する。
'       ・タブ本体クリック → {"cmd":"activate","index":N} → Browser.ActivateTab N
'       ・× ボタンクリック → {"cmd":"close","index":N}    → Browser.CloseTab N
'       ・＋ ボタンクリック → {"cmd":"newtab"}            → Browser.RequestNewTabAsync
'
'   ★コマンド解釈 = 案1-A (最小自前パーサ)★ VBA に標準 JSON パーサはなく、
'     外部 COM (ScriptControl 等) も本プロジェクトでは使わない。送信側は自前
'     HTML の JS だけなので入力は完全に信頼でき、汎用パーサは過剰。cmd と index
'     を抜き出す最小関数 (GetJsonStr / GetJsonNum) を自前で用意する。VBA->JS が
'     JSON なので JS->VBA も JSON で揃え、将来コマンドが増えても自然に拡張できる。
'
'   ★再入の安全性 = 案4-A★ OnPaneWebMessage は WebView2 のイベントハンドラの
'     中で走る。ここから別の WebView2 操作を呼ぶときは 仕様事実 14 (ハンドラ内で
'     非同期 WebView2 操作の完了を待つとデッドロック) に警戒が要る。
'       ・activate (ActivateTab)  : Ctrl_PutIsVisible のみ。完了待ちなし → 同期で安全
'       ・close    (CloseTab)     : Ctrl_Close のみ。完了待ちなし     → 同期で安全
'       ・newtab   (AddTabWithUrl): ★WaitPaneReady で Controller 生成の完了を待つ★
'                                   → 仕様事実 14 に該当。デッドロックする。
'     よって newtab だけは Browser.RequestNewTabAsync (9.10e の SafeTimer 一発撮り
'     機構) に流し、ハンドラを即返す。NewWindowRequested で実績のある経路をそのまま
'     再利用する (設計原則 63)。
'
'   ★close のポリシー = 案2-A (タブ 0 許容)★ 最後の 1 タブも閉じられる。
'     「最低タブ数ポリシーは上位に委ねる」という 9.11a の CloseTab の判断と整合。
'     タブ 0 になったら JS 側が「(タブなし)」を表示する (9.11c の render で対応済み)。
'     ＋ ボタンから復帰できる。
'
'   ★newtab の初期 URL = 案3-B (定数 HOME_URL)★ 将来 UserForm から設定可能に
'     できるよう定数で持つ。
'
' --- Wv2TabBar.cls  第9.11c 段階 ---
'
'   第9.11c のスコープ (VBA->JS 方向、タブ一覧の描画。表示のみ):
'     Browser のイベント (TabAdded/TabClosed/ActiveChanged/TabTitleChanged) を
'     受けて、タブバー JS にタブ一覧を丸ごと送って再描画させる。クリックは
'     まだ効かない (器だけ用意。activate/close/newtab は 9.11d)。
'
'   ★全体同期方式 (差分ではない)★ VBA->JS は毎回タブ全体の状態を送る:
'       {"type":"sync","tabs":[{"title":"..."},...],"active":N}
'     JS 側は受け取った配列で DOM を丸ごと再描画する (宣言的レンダリング)。
'     index 補正ロジックが JS に不要で堅牢 (VBA が真実源)。初期同期も sync 一発。
'
'   ★ready ハンドシェイク (案X)★ JS はロード完了時に {"cmd":"ready"} を
'     postMessage する。VBA はそれを受けて初めて初期同期 sync を返す。
'     PostWebMessageAsString は JS 側の message リスナー登録前に送ると落ちる
'     (9.11b でタイトルイベントが未処理で流れたのがこの時間帯) ため、
'     DOMContentLoaded より確実な ready 合図で初期同期のタイミングを取る。
'     9.11d のクリック受信の土台にもなる。
'
'   ★タイトルの真実源 = Pane (選択肢3)★ タイトルは Wv2Pane.m_title に
'     キャッシュされる (9.11c で Pane 側に追加)。PushSyncToJs は
'     m_browser.PaneAt(i).DocumentTitle で読む。index ずれの影響を受けない。
'
'   ★Pane->TabBar 後方参照 = 案P (WebMessage シンク)★ タブバーペイン
'     (m_barPane) の JS から来る postMessage を受けるため、TabBar は自分を
'     m_barPane の WebMessage シンクに登録する (m_barPane.SetWebMsgSink Me)。
'     受信メソッドは OnPaneWebMessage(pane, source, msg)。解放時は Shutdown が
'     ClearWebMsgSink 経由で切る (設計原則 58)。
'
'   第9.11c の実装内訳:
'     - Init: m_barPane に自分をシンク登録 + View_AddWebMessageReceived で受信有効化。
'     - 4 つのイベントハンドラ: 共通の PushSyncToJs 呼び出しに置換。
'     - OnPaneWebMessage: {"cmd":"ready"} を受けたら PushSyncToJs で初期同期。
'     - PushSyncToJs: TabCount/PaneAt(i).DocumentTitle/ActiveIndex を JSON 化して
'       m_barPane.View_PostWebMessageAsString で JS へ送る。JsonEscape で最小
'       エスケープ。
'     - BuildTabBarHtml: ready 送信 + sync 受信描画を追加 (クリックは器だけ)。
'
'   WebView2 製タブバー UI。Wv2Browser のイベントを購読 (WithEvents) して
'   タブバーの見た目を更新し、タブバー上のクリックを Wv2Browser への操作
'   (ActivateTab / CloseTab / 新規タブ) に変換する。
'
'   ★疎結合 (第9.11 段階の中核方針)★
'     Wv2Browser は Wv2TabBar の存在を一切知らない。本クラスは Browser の
'     Public Event を「聞く側」であり、Browser から見れば聞き手ゼロと同じ。
'     タブバー不要なユーザーは本クラスをインスタンス化しなければよいだけ。
'     MsForms 製タブバーが欲しいユーザーは本クラスの代わりに独自クラスを
'     WithEvents で Browser に繋げばよい (本クラスは WebView2 製の一実装)。
'
'   ★所有構造 (案b、第9.11 段階で確定)★
'     UserForm1
'       ├ m_browser As Wv2Browser   (コンテンツタブ群を管理)
'       └ m_tabBar  As Wv2TabBar    (本クラス。タブバー UI)
'     本クラスは内部に m_barPane As Wv2Pane (タブバー描画用の単一ペイン) を
'     持つ。この m_barPane は Wv2Browser.m_tabs には入れない (通常タブと
'     混ざらないよう、タブ管理下に置かない)。よって AddTab は使わず、
'     Wv2Pane を直接生成する。
'
'   ★Environment 共有 (第9.11b)★
'     m_barPane は Wv2Browser と同じ Environment を共有する
'     (Environment は 1 個を全ペインで共有)。Init で browser.Environment を
'     受け取って m_barPane.Init に渡す。
'
'   第9.11b のスコープ (骨格のみ):
'     作る  : Init (ペイン生成 → View 取得 → 仮想ホストマッピング →
'             静的 HTML 表示) / Resize / Shutdown / WithEvents ハンドラの
'             空実装 (Debug.Print のみ)
'     作らない: タブ一覧の描画 (9.11c) / クリック → 操作の双方向通信 (9.11d)
'     → 9.11b では上部帯に "TabBar" と表示されるだけ。コンテンツ領域が
'       その下にずれて共存することを確認するのが目的。
'
'   ★ペイン生成手順 (Wv2Browser.AddTab を参考、ただしタブ管理部分は除く)★
'     1. New Wv2Pane → Init(env, hwnd)
'     2. Controller Ready 待ち (IsReady / IsFailed をポーリング、DoEvents)
'     3. Ctrl_GetCoreWebView2 (View 取得。省略すると Navigate が E_POINTER、設計原則 56)
'     4. View3_SetVirtualHostNameToFolderMapping (Navigate の前に必須)
'     5. View_Navigate(https://<host>/tabbar.html) (仕様事実 13: ファイル名明示)
'
'   ★HTML 配信 (第9.10d の SetVirtualHostNameToFolderMapping を流用)★
'     タブバー HTML をローカルフォルダに UTF-8 で書き出し、
'     https://appassets.tabbar/tabbar.html として表示する。日本語タイトルを
'     将来 (9.11c) 流し込むため、書き出しは ADODB.Stream で UTF-8 (BOM 無し)。
''''''''''''''''''''''''''''''''''

Option Explicit

' ============================================================
' フィールド
' ============================================================
Private WithEvents m_browser As Wv2Browser   ' Browser のイベントを購読
Attribute m_browser.VB_VarHelpID = -1
Private m_barPane   As Wv2Pane                ' タブバー描画用の単一ペイン
Private m_hwndHost  As LongPtr               ' ホスト (親フレーム) の hwnd
Private m_folderPath As String               ' タブバー HTML の配信フォルダ
Private m_ready     As Boolean               ' Init 完了フラグ

' タブバーの配信ホスト名 (コンテンツ側の appassets.example とは分ける)
Private Const TABBAR_HOST As String = "appassets.tabbar"
Private Const TABBAR_FILE As String = "tabbar.html"

' --- 第9.11d: ＋ ボタン (newtab) で開く初期 URL (案3-B) ---
'   定数で持つことで、将来 UserForm から設定可能にする余地を残す。
'   about:blank だとタイトルが空でタブが「(読み込み中)」のままになるため、
'   検証の見栄えを優先して実サイトを既定にする。
Private Const HOME_URL As String = "https://www.bing.com"

' Controller Ready 待ちのタイムアウト (秒)
Private Const READY_TIMEOUT As Single = 10#


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    Set m_browser = Nothing
    Set m_barPane = Nothing
    m_hwndHost = 0
    m_folderPath = ""
    m_ready = False
End Sub


' ============================================================
' Init ? タブバーを起動する
'
'   引数:
'     browser    : 購読対象の Wv2Browser (コンテンツタブ群を管理しているもの)
'     hwnd       : ホスト (親フレーム) の hwnd (m_barPane の親)
'     folderPath : タブバー HTML を書き出すローカルフォルダ
'   戻り値: 成功で True
'
'   ★Environment は browser.Environment から共有取得する (第9.11b)★
' ============================================================
Public Function Init( _
    ByVal browser As Wv2Browser, _
    ByVal hwnd As LongPtr, _
    ByVal folderPath As String) As Boolean

    If browser Is Nothing Then
        Debug.Print "Wv2TabBar.Init: browser が Nothing"
        Exit Function
    End If

    Dim env As Wv2Environment
    Set env = browser.Environment
    If env Is Nothing Then
        Debug.Print "Wv2TabBar.Init: browser.Environment が Nothing (Browser.Init 未完了?)"
        Exit Function
    End If

    Set m_browser = browser        ' ★これで WithEvents 購読が始まる
    m_hwndHost = hwnd
    m_folderPath = folderPath

    ' --- 1. タブバー HTML を書き出す ---
    If Not WriteTabBarHtml() Then
        Debug.Print "Wv2TabBar.Init: HTML 書き出し失敗"
        Exit Function
    End If

    ' --- 2. タブバー用ペインを生成 (AddTab は使わない = タブ管理下に置かない) ---
    Set m_barPane = New Wv2Pane
    m_barPane.Init env, m_hwndHost

    ' --- 3. Controller Ready 待ち ---
    If Not WaitPaneReady(m_barPane) Then
        If m_barPane.IsFailed Then
            Debug.Print "Wv2TabBar.Init: Controller 生成失敗 (lastError=&H" & _
                        Hex(m_barPane.LastError) & ")"
        Else
            Debug.Print "Wv2TabBar.Init: Controller 生成タイムアウト"
        End If
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: Controller 生成 OK (pCtrl=" & m_barPane.CtrlPtr & ")"

    ' --- 4. View 取得 (省略すると Navigate が E_POINTER、設計原則 56) ---
    Dim hr As Long
    hr = m_barPane.Ctrl_GetCoreWebView2()
    If hr <> 0 Or m_barPane.ViewPtr = 0 Then
        Debug.Print "Wv2TabBar.Init: View 取得失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: View 取得 OK (pView=" & m_barPane.ViewPtr & ")"

    ' --- 5. 仮想ホストマッピング (Navigate の前に必須) ---
    hr = m_barPane.View3_SetVirtualHostNameToFolderMapping(TABBAR_HOST, m_folderPath, 1)  ' 1 = ALLOW
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: マッピング失敗 hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- 5b. hostObjects 公開 (第9.21、Navigate の前に必須: 仕様事実 18 の順序A) ---
    '   ★実 URL Navigate でも成立するかの検証点★ 設定タブ (NavigateToString)
    '   で確立した『Navigate 前に attach すれば JS 初回実行から hostObjects を
    '   掴める』が、appassets.tabbar への実 URL でも効くかを試す。
    '   attach するのは TabBar 自身 (Me)。JS からは
    '   window.chrome.webview.hostObjects.sync.TabBar.HostActivate(n) で呼ばれる。
    '   失敗しても致命ではない (JS 側が postMessage にフォールバックする) ため、
    '   ここでは hr をログするだけで Init は続行する。
    Dim hrHost As Long
    hrHost = m_barPane.View_AddHostObjectToScript("TabBar", Me)
    If hrHost <> 0 Then
        Debug.Print "Wv2TabBar.Init: AddHostObjectToScript(TabBar) 失敗 hr=&H" & Hex(hrHost) & " (JS は postMessage にフォールバック)"
    Else
        Debug.Print "Wv2TabBar.Init: TabBar ブリッジを attach OK (activate は hostObjects 優先)"
    End If

    ' --- 6. Navigate (仕様事実 13: ファイル名まで明示) ---
    Dim url As String
    url = "https://" & TABBAR_HOST & "/" & TABBAR_FILE
    hr = m_barPane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: Navigate 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: Navigate OK (" & url & ")"

    ' --- 7. WebMessage シンク登録 + 受信有効化 (第9.11c、案P) ---
    '   タブバー JS からの postMessage (ready / 将来のクリック) を
    '   OnPaneWebMessage で受け取れるようにする。
    '   ★初期同期は「ここでは撃たない」★ JS の message リスナー登録前に
    '   PostMessage すると落ちる (9.11b の教訓)。JS が {"cmd":"ready"} を
    '   送ってきてから OnPaneWebMessage 内で PushSyncToJs する (案X)。
    m_barPane.SetWebMsgSink Me
    hr = m_barPane.View_AddWebMessageReceived()
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Init: View_AddWebMessageReceived 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2TabBar.Init: WebMessage 受信を有効化 (ready 待ち)"

    m_ready = True
    Init = True
    Debug.Print "Wv2TabBar.Init: 完了 (タブバー起動、初期同期は ready 受信後)"
End Function


' ============================================================
' Resize ? タブバーペインの Bounds を設定する
'   UserForm_Resize から、上部帯の矩形を渡してもらう。
'   (レイアウト分割は UserForm 側が計算し、Browser はタブバーを知らない)
' ============================================================
Public Sub Resize(ByRef barRect As RECT)
    If m_barPane Is Nothing Then Exit Sub
    Dim hr As Long
    hr = m_barPane.Ctrl_PutBounds(barRect)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.Resize: Ctrl_PutBounds 失敗 hr=&H" & Hex(hr)
    End If
    m_barPane.Ctrl_PutIsVisible True   ' タブバーは常時可視
End Sub


' ============================================================
' Shutdown ? タブバーペインを解放する
'   UserForm_Terminate から m_browser.Shutdown より前に呼ぶ。
'   本クラスは m_browser を WithEvents で参照しているが、これは循環では
'   ない (Browser は本クラスを知らない) ので、Set m_browser = Nothing で
'   購読を解除するだけでよい。
' ============================================================
Public Sub Shutdown()
    Debug.Print "Wv2TabBar.Shutdown 開始"
    If Not m_barPane Is Nothing Then
        m_barPane.ClearWebMsgSink   ' 第9.11c: シンク登録を先に切る (設計原則 58)
        m_barPane.Ctrl_Close
        Set m_barPane = Nothing
    End If
    Set m_browser = Nothing   ' WithEvents 購読を解除
    m_ready = False
    Debug.Print "Wv2TabBar.Shutdown 完了"
End Sub


' ============================================================
' IsReady
' ============================================================
Public Property Get IsReady() As Boolean
    IsReady = m_ready
End Property


' ============================================================
' ★Wv2Browser イベントハンドラ (第9.11c: 全体同期方式)★
'   どのイベントでも「タブ全体の状態を丸ごと JS へ送る」PushSyncToJs を
'   呼ぶだけ。差分は扱わない (VBA が真実源、JS は宣言的に再描画)。
'   ★注意★ これらは Browser のイベントで発火するが、初期同期 (起動直後の
'   1 回目) は JS の ready 合図 (OnPaneWebMessage) をトリガーにする。ready
'   前に来たイベントの sync は JS のリスナー未登録で取りこぼれる可能性が
'   あるが、ready 受信時に最新状態を丸ごと送るため最終的に整合する
'   (全体同期の利点)。
' ============================================================
Private Sub m_browser_TabAdded(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] TabAdded(index=" & index & ") → PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_TabClosed(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] TabClosed(index=" & index & ") → PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_ActiveChanged(ByVal index As Long)
    Debug.Print "Wv2TabBar[EVT] ActiveChanged(index=" & index & ") → PushSyncToJs"
    PushSyncToJs
End Sub

Private Sub m_browser_TabTitleChanged(ByVal index As Long, ByVal title As String)
    Debug.Print "Wv2TabBar[EVT] TabTitleChanged(index=" & index & ", title=" & title & ") → PushSyncToJs"
    PushSyncToJs
End Sub


' ============================================================
' OnPaneWebMessage (第9.11c、案P: m_barPane の WebMessage シンク受け口)
'   タブバー JS の window.chrome.webview.postMessage(...) が Wv2Pane 経由で
'   ここに届く。Wv2Pane.SetWebMsgSink Me で登録した結果、Wv2Pane の
'   View_OnWebMessageReceived から OnPaneWebMessage(pane, source, msg) が
'   呼ばれる。
'
'   引数:
'     pane   : 送信元ペイン (ここでは常に m_barPane)
'     source : 送信元 URL
'     msg    : JS が送った文字列 (本実装では {"cmd":"ready"} 等の JSON)
'
'   第9.11c では "ready" だけを解釈する。クリック系 (activate/close/newtab) の
'   コマンド解釈は 9.11d で追加する。
'   ★JSON パーサは使わない★ VBA に標準 JSON パーサがないので、9.11c では
'   "ready" が含まれるかの単純判定で足りる (受信コマンドは自前 HTML の JS が
'   送るものだけなので信頼できる)。9.11d でコマンドが増えたら最小パースに拡張。
' ============================================================
Public Sub OnPaneWebMessage(ByVal pane As Wv2Pane, ByVal source As String, ByVal msg As String)
    Debug.Print "Wv2TabBar.OnPaneWebMessage: msg=" & msg

    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        Exit Sub
    End If

    Dim cmd As String
    cmd = JsonGetStr(msg, "cmd")

    ' idx は activate/close のフォールバック経路で JSON から index を復元するのに使う。
    ' ok は 9.23 で不要化 (処理と失敗ログを Host メソッドへ一元化したため Case 側で成否を受けない)。
    Dim idx As Long

    Select Case cmd

        Case "ready"
            ' --- 第9.11c: JS のリスナー登録完了。初期同期 sync を返す (案X) ---
            Debug.Print "  → ready 受信。初期同期 sync を送信する。"
            PushSyncToJs

        Case "activate"
            ' --- 第9.11d/9.23: タブ本体クリック → アクティブ切替 (postMessage フォールバック経路) ---
            '   ★第9.23: 処理を HostActivate へ一元化 (案C-1)★ JS がハイブリッドで
            '   hostObjects を掴めなかった時だけここへ落ちる。ActivateTab 呼び出しと
            '   失敗ログは HostActivate 内に集約済みなので、ここでは JSON から index を
            '   復元して HostActivate に呼び捨てで委譲するだけ (戻り値は受けない)。
            idx = JsonGetNum(msg, "index")
            Debug.Print "  → activate(index=" & idx & ") [postmsg fallback → HostActivate]"
            HostActivate idx

        Case "close"
            ' --- 第9.11d/9.23: × ボタンクリック → タブを閉じる (postMessage フォールバック経路) ---
            '   ★第9.23: 処理を HostClose へ一元化 (案C-1)★ hostObjects を掴めなかった
            '   時だけここへ落ちる。CloseTab 呼び出し・タブ 0 許容 (案2-A)・失敗ログは
            '   すべて HostClose 内に集約済み。ここでは index を復元して呼び捨て委譲する。
            idx = JsonGetNum(msg, "index")
            Debug.Print "  → close(index=" & idx & ") [postmsg fallback → HostClose]"
            HostClose idx

        Case "newtab"
            ' --- 第9.11d: ＋ ボタンクリック → 新規タブ ---
            '   ★★ 同期で AddTabWithUrl を呼んではいけない ★★
            '   AddTabWithUrl は内部で WaitPaneReady (Controller 生成の完了待ち)
            '   をするため、WebMessage ハンドラ内から呼ぶと 仕様事実 14 の
            '   デッドロックに嵌まる (9.10e の NewWindowRequested と同じ構造)。
            '   RequestNewTabAsync は URL をキューに積んで SafeTimer をワンショット
            '   起動し即座に返るので、ハンドラ完了後に実際のタブ生成が走る (案4-A)。
            Debug.Print "  → newtab (RequestNewTabAsync 経由、" & HOME_URL & ")"
            m_browser.RequestNewTabAsync HOME_URL

        Case "settings"
            ' --- 第9.18: ? ボタンクリック → 設定タブを開く ---
            '   ★★ 同期で OpenSettingsTab を呼んではいけない ★★
            '   設定タブ生成も内部で AddTab → WaitPaneReady (Controller 生成の
            '   完了待ち) を通るため、WebMessage ハンドラ内から直接呼ぶと newtab と
            '   同じ 仕様事実 14 のデッドロックに嵌まる。RequestSettingsTabAsync は
            '   保留フラグを立てて SafeTimer をワンショット起動し即座に返すので、
            '   ハンドラ完了後に実際の設定タブ生成が走る (案4-A と同じ判断)。
            Debug.Print "  → settings (RequestSettingsTabAsync 経由)"
            m_browser.RequestSettingsTabAsync

        Case "reorder"
            ' --- 第9.15/9.23: ドラッグ並べ替え → タブを移動 (postMessage フォールバック経路) ---
            '   ★第9.23: 処理を HostReorder へ一元化 (案C-1)★ hostObjects を掴めなかった
            '   時だけここへ落ちる。from/to はいずれも 1-based。MoveTab 呼び出しと失敗ログは
            '   HostReorder 内に集約済み。ここでは from/to を復元して呼び捨て委譲する。
            Dim fromIdx As Long, toIdx As Long
            fromIdx = JsonGetNum(msg, "from")
            toIdx = JsonGetNum(msg, "to")
            Debug.Print "  → reorder(from=" & fromIdx & ", to=" & toIdx & ") [postmsg fallback → HostReorder]"
            HostReorder fromIdx, toIdx

        Case ""
            Debug.Print "  → cmd を取り出せなかった (msg=" & msg & ")"

        Case Else
            Debug.Print "  → 未知のコマンド: " & cmd

    End Select
End Sub


' ============================================================
' HostActivate (第9.21、hostObjects 経路の activate 受け口)
'
'   JS から window.chrome.webview.hostObjects.sync.TabBar.HostActivate(n) で
'   直接呼ばれる (postMessage を介さない同期 RPC)。中身は OnPaneWebMessage の
'   activate ケースと同一で、m_browser.ActivateTab を同期で呼ぶだけ。
'
'   ★同期呼び出しの安全性 (案4-A)★ ActivateTab は Ctrl_PutIsVisible のみで
'     完了待ちを含まないため、JS->VBA の sync 直接呼び出しでも 仕様事実 14 の
'     デッドロックに嵌まらない。成功すると RaiseEvent ActiveChanged →
'     PushSyncToJs が走り、タブバーのアクティブ強調が全体同期で追従する。
'
'   引数 index は 1-based (JS 側の data-index と同じ)。戻り値は成否 (JS 側は
'   参照しなくてよいが、比較検証で戻り値の往復も確認できるよう Boolean で返す)。
' ============================================================
Public Function HostActivate(ByVal index As Long) As Boolean
    Debug.Print "Wv2TabBar.HostActivate: index=" & index & " (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostActivate = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.ActivateTab(index)
    If Not ok Then
        Debug.Print "    [警告] ActivateTab 失敗 (index=" & index & ")"
    End If
    HostActivate = ok
End Function


' ============================================================
' HostClose (第9.22、hostObjects 経路の close 受け口)
'
'   JS から window.chrome.webview.hostObjects.sync.TabBar.HostClose(n) で
'   直接呼ばれる (postMessage を介さない同期 RPC)。中身は OnPaneWebMessage の
'   close ケースと同一で、m_browser.CloseTab を同期で呼ぶだけ。
'
'   ★同期呼び出しの安全性 (案4-A)★ CloseTab は Ctrl_Close のみで完了待ちを
'     含まないため、JS->VBA の sync 直接呼び出しでも 仕様事実 14 のデッドロックに
'     嵌まらない。★タブ 0 許容 (案2-A)★ 最後の 1 タブも閉じられ、CloseTab 内の
'     アクティブ補正 (9.11a の 4 分岐) が走り、RaiseEvent ActiveChanged /
'     TabClosed → PushSyncToJs でタブバーが全体同期で追従する。
'
'   引数 index は 1-based (JS 側の data-index と同じ)。戻り値は成否。
' ============================================================
Public Function HostClose(ByVal index As Long) As Boolean
    Debug.Print "Wv2TabBar.HostClose: index=" & index & " (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostClose = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.CloseTab(index)
    If Not ok Then
        Debug.Print "    [警告] CloseTab 失敗 (index=" & index & ")"
    End If
    HostClose = ok
End Function


' ============================================================
' HostReorder (第9.22、hostObjects 経路の reorder 受け口)
'
'   JS から window.chrome.webview.hostObjects.sync.TabBar.HostReorder(from, to)
'   で直接呼ばれる (postMessage を介さない同期 RPC)。中身は OnPaneWebMessage の
'   reorder ケースと同一で、m_browser.MoveTab を同期で呼ぶだけ。
'
'   ★同期呼び出しの安全性 (案4-A)★ MoveTab は配列操作 (Remove/Add) と
'     PutIsVisible のみで完了待ちを含まないため、sync 直接呼び出しでも
'     デッドロックしない。成功で RaiseEvent ActiveChanged → PushSyncToJs が
'     走り、並べ替え後の順序で全体同期の再描画が走る。
'
'   ★引数名 fromIdx/toIdx★ From は一部方言で衝突報告があるため回避。JS 側は
'     位置引数呼び (HostReorder(from,to)) なので VBA 引数名は JS から見えない。
'     いずれも 1-based (JS の from/to と同じ意味)。戻り値は成否。
' ============================================================
Public Function HostReorder(ByVal fromIdx As Long, ByVal toIdx As Long) As Boolean
    Debug.Print "Wv2TabBar.HostReorder: from=" & fromIdx & " to=" & toIdx & " (host経由)"
    If m_browser Is Nothing Then
        Debug.Print "  → m_browser が Nothing のため無視"
        HostReorder = False
        Exit Function
    End If
    Dim ok As Boolean
    ok = m_browser.MoveTab(fromIdx, toIdx)
    If Not ok Then
        Debug.Print "    [警告] MoveTab 失敗 (from=" & fromIdx & ", to=" & toIdx & ")"
    End If
    HostReorder = ok
End Function


' ============================================================
' PushSyncToJs (第9.11c、全体同期: タブ全体の状態を JSON 化して JS へ送る)
'
'   送る JSON の形:
'     {"type":"sync","tabs":[{"title":"..."},{"title":"..."}],"active":N}
'   active は 1-based (0 = アクティブなし)。JS 側はこの配列で DOM を丸ごと
'   再描画する。
'
'   タイトルの真実源は Wv2Pane.DocumentTitle (選択肢3)。m_browser.PaneAt(i) で
'   各ペインを取り、DocumentTitle を読む。まだタイトル未取得のペインは空文字
'   なので、JS 側で (読み込み中) 等にフォールバックしてよい (本実装では
'   空文字のまま送り、表示は JS に委ねる)。
' ============================================================
Private Sub PushSyncToJs()
    If m_barPane Is Nothing Then Exit Sub
    If m_browser Is Nothing Then Exit Sub
    If Not m_ready Then Exit Sub

    Dim n As Long
    n = m_browser.TabCount

    Dim json As String
    json = "{""type"":""sync"",""tabs"":["

    Dim i As Long
    Dim pane As Wv2Pane
    Dim title As String
    For i = 1 To n
        Set pane = m_browser.PaneAt(i)
        If pane Is Nothing Then
            title = ""
        Else
            title = pane.DocumentTitle
        End If
        If i > 1 Then json = json & ","
        json = json & "{""title"":""" & JsonEscape(title) & """}"
    Next i

    json = json & "],""active"":" & m_browser.ActiveIndex & "}"

    Dim hr As Long
    hr = m_barPane.View_PostWebMessageAsString(json)
    If hr <> 0 Then
        Debug.Print "Wv2TabBar.PushSyncToJs: PostWebMessageAsString 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2TabBar.PushSyncToJs: 送信 OK (" & n & " タブ, active=" & _
                    m_browser.ActiveIndex & ")"
        Debug.Print "  json=" & json
    End If
End Sub


' ============================================================
' WaitPaneReady ? Controller Ready 待ち (Wv2Browser.WaitPaneReady と同形)
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
' WriteTabBarHtml ? タブバー HTML をフォルダに UTF-8 (BOM 無し) で書き出す
'   9.11b は静的 (中央に "TabBar" と出すだけ)。9.11c 以降で JS を足す。
'   ★UTF-8 書き出し★ 将来日本語タイトルを流し込むため ADODB.Stream を使う
'   (VBA の Print # は既定コードページ依存で日本語が化ける)。
' ============================================================
Private Function WriteTabBarHtml() As Boolean
    On Error GoTo eh

    ' フォルダがなければ作成
    If Dir(m_folderPath, vbDirectory) = "" Then
        MkDir m_folderPath
        Debug.Print "Wv2TabBar.WriteTabBarHtml: フォルダ作成 " & m_folderPath
    End If

    Dim fullPath As String
    fullPath = m_folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & TABBAR_FILE

    Dim html As String
    html = BuildTabBarHtml()

    ' ADODB.Stream で UTF-8 (BOM 無し) 書き出し
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.Charset = "UTF-8"
    st.Open
    st.WriteText html

    ' UTF-8 BOM を除去して保存する
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText html
    ' 一旦テキストで書き、バイナリで読み直して先頭 3 バイト (BOM) を捨てる
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

    st.Close

    Debug.Print "Wv2TabBar.WriteTabBarHtml: 書き出し OK " & fullPath & " (" & Len(html) & " 文字, UTF-8)"
    WriteTabBarHtml = True
    Exit Function
eh:
    Debug.Print "Wv2TabBar.WriteTabBarHtml: エラー " & Err.Number & " " & Err.Description
    WriteTabBarHtml = False
End Function


' ============================================================
' BuildTabBarHtml ? タブバー HTML を組み立てる (第9.11c: sync 描画対応)
'
'   第9.11c の JS 動作:
'     1. ロード完了時に window.chrome.webview.postMessage('{"cmd":"ready"}') を
'        送る (ready ハンドシェイク、案X)。これで VBA が初期同期 sync を返す。
'     2. window.chrome.webview.addEventListener('message', ...) で sync を受信し、
'        {type:"sync", tabs:[{title}], active:N} からタブ DOM を丸ごと再描画する
'        (全体同期、宣言的レンダリング)。
'     3. 第9.11d: クリックを VBA へ送る。タブ本体 → activate、× → close、
'        ＋ → newtab。× のクリックはタブ本体の onclick に伝播しないよう
'        stopPropagation する (親の activate が誤発火しないように)。
'
'   ★JSON 文字列は VBA の二重引用符エスケープ ("") でリテラルを埋め込む★
' ============================================================
Private Function BuildTabBarHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  html,body{margin:0;padding:0;height:100%;box-sizing:border-box;" & _
            "font-family:'Segoe UI','Meiryo',sans-serif;}" & vbLf
    s = s & "  *{box-sizing:border-box;}" & vbLf
    s = s & "  body{display:flex;align-items:stretch;overflow:hidden;" & _
            "background:#f3f3f3;border-bottom:1px solid #ccc;}" & vbLf
    s = s & "  #tabs{display:flex;align-items:stretch;height:100%;overflow-x:auto;}" & vbLf
    ' --- 第9.32 (論点 4): 横スクロールバーを隠す ---
    '     overflow-x:auto のままなので★ホイールで横スクロールはできる★
    '     (Chromium は縦スクロール不可の横コンテナでホイールを横へ変換する)。
    '     バーが出るとタブバーの高さを内側から食ってタブが潰れるので隠す。
    s = s & "  #tabs::-webkit-scrollbar{display:none;}" & vbLf
    ' --- 第9.31 (A-1): 背景色だけ 0.12 秒でなめらかに切り替える (論点 5) ---
    s = s & "  .tab{display:flex;align-items:center;max-width:200px;min-width:80px;" & _
            "height:100%;padding:0 6px 0 12px;border-right:1px solid #ccc;" & _
            "background:#e7e7e7;cursor:pointer;font-size:13px;color:#333;" & _
            "white-space:nowrap;transition:background-color .12s;}" & vbLf
    s = s & "  .tab.active{background:#fff;color:#000;font-weight:600;}" & vbLf
    ' --- 第9.31 (A-1): ホバー背景。★非アクティブのみ★ (論点 2) ---
    '     :not(.active) で詳細度 0,3,0。.tab.active (0,2,0) より強いので
    '     記述順に関係なくアクティブタブの #fff が保たれる。
    s = s & "  .tab:not(.active):hover{background:#efefef;}" & vbLf
    s = s & "  .tab .title{overflow:hidden;text-overflow:ellipsis;flex:1;}" & vbLf
    ' --- 第9.31 (A-2): 閉じるボタンは既定で隠す (論点 3・4) ---
    '     ★visibility:hidden を使う理由★
    '       display:none だと 18px + margin 6px が消えてタイトル幅が伸縮し、
    '       ホバーのたびに文字がガタつく。opacity:0 だと見えないボタンを
    '       クリックできてしまう。visibility:hidden なら場所を保ったまま消え、
    '       ポインタイベントも無効になる。
    s = s & "  .tab .close{margin-left:6px;width:18px;height:18px;line-height:18px;" & _
            "text-align:center;border-radius:3px;color:#888;visibility:hidden;}" & vbLf
    ' --- 第9.31 (A-2): アクティブ or ホバーのときだけ見せる (Chrome 相当) ---
    '     どちらも詳細度 0,3,0 で .tab .close (0,2,0) より強い。
    s = s & "  .tab.active .close,.tab:hover .close{visibility:visible;}" & vbLf
    s = s & "  .tab .close:hover{background:#ccc;color:#000;}" & vbLf
    s = s & "  .tab.dragging{opacity:0.45;}" & vbLf
    s = s & "  .tab.dragover{box-shadow:inset 3px 0 0 #4a90e2;}" & vbLf
    s = s & "  .tab.dragover-after{box-shadow:inset -3px 0 0 #4a90e2;}" & vbLf
    s = s & "  #newtab{display:flex;align-items:center;justify-content:center;" & _
            "width:32px;height:100%;cursor:pointer;font-size:18px;color:#555;}" & vbLf
    s = s & "  #newtab:hover{background:#e0e0e0;}" & vbLf
    ' --- 第9.18: 設定 (?) ボタン。newtab の右隣。右端に寄せる ---
    s = s & "  #settings{display:flex;align-items:center;justify-content:center;" & _
            "width:34px;height:100%;cursor:pointer;font-size:16px;color:#666;margin-left:auto;}" & vbLf
    s = s & "  #settings:hover{background:#e0e0e0;color:#333;}" & vbLf
    s = s & "  #empty{padding:0 12px;display:flex;align-items:center;color:#999;font-size:13px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "  <div id=""tabs""></div>" & vbLf
    s = s & "  <div id=""newtab"" title=""新しいタブ"">+</div>" & vbLf
    ' --- 第9.18: 設定ボタン (?)。margin-left:auto で右端へ ---
    s = s & "  <div id=""settings"" title=""設定"">&#9881;</div>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "  var W = window.chrome && window.chrome.webview;" & vbLf
    s = s & "  function postCmd(obj){ if(W){ W.postMessage(JSON.stringify(obj)); } }" & vbLf
    ' --- 第9.15: ドラッグ状態と挿入マーカー掃除ヘルパー ---
    ' --- 第9.32: HO() を render 内からトップレベルへ移設 (事後報告) ---
    s = s & "  function HO(){ return (W && W.hostObjects && W.hostObjects.sync) ? W.hostObjects.sync : null; }" & vbLf
    ' --- 第9.32 (論点 1): 楽観的更新。VBA を待たずに先に見た目だけ切り替える ---
    '     次の全体同期が必ず上書きするので JS 側に状態は残らない (設計原則 60 と両立)。
    s = s & "  function markActive(el){" & vbLf
    s = s & "    var xs = document.querySelectorAll('.tab.active');" & vbLf
    s = s & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('active'); }" & vbLf
    s = s & "    if(el){ el.classList.add('active'); }" & vbLf
    s = s & "  }" & vbLf
    ' --- 第9.32 (論点 2): 二重 rAF で「1 フレーム描かれた後」に fn を走らせる ---
    '     DOM を変えた直後に同期 COM 呼び出しでブロックすると描画前に固まるため。
    s = s & "  function afterPaint(fn){" & vbLf
    s = s & "    if(window.requestAnimationFrame){" & vbLf
    s = s & "      requestAnimationFrame(function(){ requestAnimationFrame(fn); });" & vbLf
    s = s & "    } else {" & vbLf
    s = s & "      setTimeout(fn, 0);" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    ' --- 第9.32: activate の送信部を切り出し (中身は第9.31 までと同一) ---
    s = s & "  function sendActivate(n){" & vbLf
    s = s & "    var ho = HO();" & vbLf
    s = s & "    if(ho && ho.TabBar){" & vbLf
    s = s & "      try{ ho.TabBar.HostActivate(n); console.log('[host] activate '+n); }" & vbLf
    s = s & "      catch(e){ console.log('[host-fail→postmsg] '+e); postCmd({cmd:'activate', index:n}); }" & vbLf
    s = s & "    } else {" & vbLf
    s = s & "      console.log('[postmsg] activate '+n); postCmd({cmd:'activate', index:n});" & vbLf
    s = s & "    }" & vbLf
    s = s & "  }" & vbLf
    ' --- 第9.32b: 直前の active を覚える (nearest 寄せの発火条件。事後報告 1) ---
    s = s & "  var lastActive = -1;" & vbLf
    s = s & "  var dragSrc = null;" & vbLf
    s = s & "  function clearMarks(){" & vbLf
    s = s & "    var xs = document.querySelectorAll('.dragover, .dragover-after');" & vbLf
    s = s & "    for(var k=0;k<xs.length;k++){ xs[k].classList.remove('dragover'); xs[k].classList.remove('dragover-after'); }" & vbLf
    s = s & "  }" & vbLf
    ' --- タブ一覧を丸ごと描画 (全体同期) ---
    s = s & "  function render(state){" & vbLf
    s = s & "    var tabsEl = document.getElementById('tabs');" & vbLf
    ' --- 第9.32b (論点 3): innerHTML='' は scrollLeft を 0 に戻すので退避する ---
    s = s & "    var keepLeft = tabsEl.scrollLeft;" & vbLf
    s = s & "    tabsEl.innerHTML = '';" & vbLf
    s = s & "    var list = (state && state.tabs) ? state.tabs : [];" & vbLf
    s = s & "    var active = (state && state.active) ? state.active : 0;" & vbLf
    s = s & "    if(list.length === 0){" & vbLf
    s = s & "      var e = document.createElement('div');" & vbLf
    s = s & "      e.id = 'empty'; e.textContent = '(タブなし)';" & vbLf
    s = s & "      tabsEl.appendChild(e); return;" & vbLf
    s = s & "    }" & vbLf
    s = s & "    for(var i=0; i<list.length; i++){" & vbLf
    s = s & "      var idx = i + 1;" & vbLf                     ' 1-based
    s = s & "      var t = list[i] || {};" & vbLf
    s = s & "      var title = (t.title && t.title.length) ? t.title : '(読み込み中)';" & vbLf
    s = s & "      var tab = document.createElement('div');" & vbLf
    s = s & "      tab.className = 'tab' + (idx === active ? ' active' : '');" & vbLf
    s = s & "      tab.setAttribute('data-index', idx);" & vbLf
    s = s & "      tab.draggable = true;" & vbLf
    s = s & "      var ti = document.createElement('span');" & vbLf
    s = s & "      ti.className = 'title'; ti.textContent = title; ti.title = title;" & vbLf
    s = s & "      var cl = document.createElement('span');" & vbLf
    s = s & "      cl.className = 'close'; cl.textContent = '×';" & vbLf
    ' --- クリックは器だけ (9.11d で中身)。今はデバッグログのみで送信しない ---
    ' --- 第9.21: activate は hostObjects 優先 + postMessage フォールバック ---
    '   HO() は hostObjects.sync が取れれば返し、なければ null。設定タブの
    '   null ガードと同型。取れれば TabBar.HostActivate(n) を同期直呼び。
    ' --- 第9.32 (論点 1・2): 先に見た目を変え、1 フレーム描かせてから VBA を呼ぶ ---
    s = s & "      tab.onclick = function(ev){" & vbLf
    s = s & "        if(ev.target && ev.target.className === 'close') return;" & vbLf
    s = s & "        var n = parseInt(this.getAttribute('data-index'), 10);" & vbLf
    s = s & "        markActive(this);" & vbLf
    s = s & "        afterPaint(function(){ sendActivate(n); });" & vbLf
    s = s & "      };" & vbLf
    ' --- 第9.22: close も hostObjects 優先 + postMessage フォールバック ---
    '   activate と同一の三段構え (HO() は 9.21 で定義済みのものを流用)。
    s = s & "      cl.onclick = function(ev){" & vbLf
    s = s & "        ev.stopPropagation();" & vbLf
    s = s & "        var n = parseInt(this.parentNode.getAttribute('data-index'), 10);" & vbLf
    s = s & "        var ho = HO();" & vbLf
    s = s & "        if(ho && ho.TabBar){" & vbLf
    s = s & "          try{ ho.TabBar.HostClose(n); console.log('[host] close '+n); }" & vbLf
    s = s & "          catch(e){ console.log('[host-fail→postmsg] '+e); postCmd({cmd:'close', index:n}); }" & vbLf
    s = s & "        } else {" & vbLf
    s = s & "          console.log('[postmsg] close '+n); postCmd({cmd:'close', index:n});" & vbLf
    s = s & "        }" & vbLf
    s = s & "      };" & vbLf
    ' --- 第9.15: HTML5 Drag & Drop で並べ替え (案B-1) ---
    '   dragstart で掴んだ元 index を dragSrc に控える。dragover で「相手タブの
    '   左半分か右半分か」を見て挿入位置マーカー (.dragover / .dragover-after) を
    '   付ける。drop で from/to を計算し reorder を postMessage して VBA に委ねる。
    '   DOM は自前で動かさず、VBA からの sync 再描画に任せる (真実源は VBA)。
    s = s & "      tab.ondragstart = function(ev){" & vbLf
    s = s & "        dragSrc = parseInt(this.getAttribute('data-index'),10);" & vbLf
    s = s & "        this.classList.add('dragging');" & vbLf
    s = s & "        if(ev.dataTransfer){ ev.dataTransfer.effectAllowed='move'; try{ev.dataTransfer.setData('text/plain', String(dragSrc));}catch(e){} }" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragover = function(ev){" & vbLf
    s = s & "        ev.preventDefault();" & vbLf
    s = s & "        if(ev.dataTransfer){ ev.dataTransfer.dropEffect='move'; }" & vbLf
    s = s & "        clearMarks();" & vbLf
    s = s & "        var r = this.getBoundingClientRect();" & vbLf
    s = s & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    s = s & "        this.classList.add(after ? 'dragover-after' : 'dragover');" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragleave = function(ev){ this.classList.remove('dragover'); this.classList.remove('dragover-after'); };" & vbLf
    s = s & "      tab.ondrop = function(ev){" & vbLf
    s = s & "        ev.preventDefault();" & vbLf
    s = s & "        var target = parseInt(this.getAttribute('data-index'),10);" & vbLf
    s = s & "        var r = this.getBoundingClientRect();" & vbLf
    s = s & "        var after = (ev.clientX - r.left) > (r.width/2);" & vbLf
    s = s & "        clearMarks();" & vbLf
    s = s & "        if(dragSrc == null) return;" & vbLf
    ' --- from/to を 1-based で計算 ---
    '   after=false → target の位置に入れる (target 番目にする)。
    '   after=true  → target の直後に入れる (target+1 番目)。
    '   ただし元位置 dragSrc が挿入位置より前にある場合、抜けた分だけ
    '   目標 index が 1 つ手前にずれるので補正する (VBA MoveTab は
    '   「並べ替え後に to 番目」を期待するため、ここで最終 to を確定させる)。
    s = s & "        var dest = after ? target + 1 : target;" & vbLf
    s = s & "        if(dragSrc < dest){ dest = dest - 1; }" & vbLf
    s = s & "        if(dest < 1) dest = 1;" & vbLf
    ' --- 第9.22: reorder も hostObjects 優先 + postMessage フォールバック ---
    '   activate/close と同一の三段構え。HostReorder(from, to) を位置引数で呼ぶ。
    s = s & "        if(dragSrc !== dest){" & vbLf
    s = s & "          var ho = HO();" & vbLf
    s = s & "          if(ho && ho.TabBar){" & vbLf
    s = s & "            try{ ho.TabBar.HostReorder(dragSrc, dest); console.log('[host] reorder '+dragSrc+'->'+dest); }" & vbLf
    s = s & "            catch(e){ console.log('[host-fail→postmsg] '+e); postCmd({cmd:'reorder', from:dragSrc, to:dest}); }" & vbLf
    s = s & "          } else {" & vbLf
    s = s & "            console.log('[postmsg] reorder '+dragSrc+'->'+dest); postCmd({cmd:'reorder', from:dragSrc, to:dest});" & vbLf
    s = s & "          }" & vbLf
    s = s & "        }" & vbLf
    s = s & "        dragSrc = null;" & vbLf
    s = s & "      };" & vbLf
    s = s & "      tab.ondragend = function(ev){ this.classList.remove('dragging'); clearMarks(); dragSrc = null; };" & vbLf
    s = s & "      tab.appendChild(ti); tab.appendChild(cl);" & vbLf
    s = s & "      tabsEl.appendChild(tab);" & vbLf
    s = s & "    }" & vbLf
    ' --- 第9.32b (論点 3): 復元 → アクティブが変わったときだけ nearest で寄せる ---
    s = s & "    tabsEl.scrollLeft = keepLeft;" & vbLf
    s = s & "    if(active !== lastActive){" & vbLf
    s = s & "      var act = tabsEl.querySelector('.tab.active');" & vbLf
    s = s & "      if(act && act.scrollIntoView){" & vbLf
    s = s & "        try{ act.scrollIntoView({block:'nearest', inline:'nearest'}); }catch(e){}" & vbLf
    s = s & "      }" & vbLf
    s = s & "    }" & vbLf
    s = s & "    lastActive = active;" & vbLf
    s = s & "  }" & vbLf
    ' --- 新規タブボタンも器だけ ---
    ' --- 第9.32b (論点 1・2): 縦ホイールを横スクロールへ変換する ---
    '     #tabs は innerHTML を書き換えても要素自体は入れ替わらないので、
    '     リスナーは 1 回付ければ生き残る。preventDefault のため passive:false。
    s = s & "  var tabsRoot = document.getElementById('tabs');" & vbLf
    s = s & "  if(tabsRoot){" & vbLf
    s = s & "    tabsRoot.addEventListener('wheel', function(ev){" & vbLf
    s = s & "      if(tabsRoot.scrollWidth <= tabsRoot.clientWidth) return;" & vbLf
    s = s & "      var d = ev.deltaY || 0;" & vbLf
    s = s & "      if(d === 0) return;" & vbLf
    s = s & "      tabsRoot.scrollLeft += d;" & vbLf
    s = s & "      ev.preventDefault();" & vbLf
    s = s & "    }, {passive:false});" & vbLf
    s = s & "  }" & vbLf
    s = s & "  document.getElementById('newtab').onclick = function(){" & vbLf
    s = s & "    postCmd({cmd:'newtab'});" & vbLf
    s = s & "  };" & vbLf
    ' --- 第9.18: 設定ボタン → settings コマンド ---
    s = s & "  document.getElementById('settings').onclick = function(){" & vbLf
    s = s & "    postCmd({cmd:'settings'});" & vbLf
    s = s & "  };" & vbLf
    ' --- VBA からの sync 受信 ---
    s = s & "  if(W){" & vbLf
    s = s & "    W.addEventListener('message', function(ev){" & vbLf
    s = s & "      var data = ev.data;" & vbLf
    s = s & "      if(typeof data === 'string'){ try{ data = JSON.parse(data); }catch(e){ return; } }" & vbLf
    s = s & "      if(data && data.type === 'sync'){ render(data); }" & vbLf
    s = s & "    });" & vbLf
    s = s & "  }" & vbLf
    ' --- ready ハンドシェイク: リスナー登録後に ready を送る (案X) ---
    s = s & "  render({tabs:[], active:0});" & vbLf
    s = s & "  postCmd({cmd:'ready'});" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildTabBarHtml = s
End Function


