VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Browser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  K-2 段階 (設定タブが検索で潰れるバグの修正) ---
'
'   ★症状★ 設定タブがアクティブな状態でアドレスバーに検索語を入れると、
'     設定タブ自身が検索結果へ遷移して設定画面が消える。さらに歯車を押しても
'     「既に設定タブがある」と判定されて新しい設定タブが開かない。
'
'   ★真因★ IsSettings は生成時に立てたきり、遷移しても False に戻らなかった。
'     フラグが実態を反映していないので、OpenSettingsTab の重複判定が
'     「元・設定タブ」を掴んでいた。
'     ★アドレスバーだけの問題ではない★ リンク・戻る・リロード・JS 由来の遷移でも
'     同じことが起きる。だから NavigateActive を塞ぐ対症療法では直らない。
'
'   ★修正 (A-2)★ 判定をフラグから ★ページ自身に聞く★ 方式へ変える。
'     設定画面の HTML に window.__wv2Settings = true を埋め、OpenSettingsTab は
'     各タブへ EvalSync (D-1) を撃って「今そのタブが設定画面を表示しているか」を
'     直接尋ねる。アドレスの表現にも遷移経路にも依存せず、戻るで設定画面へ
'     復帰したタブも正しく認識できる。
'     ★判定の権威を、状態を実際に持っている側に置く (設計原則91)★
'
'   ★修正 (B-1)★ 判定を正確にしても、戻る操作は OpenSettingsTab を通らないので
'     設定ビューが 2 枚並ぶことはありうる。その場合に値が食い違わないよう、
'     ★検索エンジンが変わったら設定画面を表示中の全タブへ現在値を押し込む★
'     (設計原則60 の全体同期そのまま)。
'     ★仕様事実27 のデッドロックを避けるため SafeTimer で一拍置く★
'     SetEngine は hostObjects 経由 (JS がブロック中) に呼ばれるので、
'     その場で EvalSync を撃ってはいけない。
'
'   ★IsSettings フラグは残してあるが、もう重複判定には使っていない★
'     「生成時に設定タブとして作られた」という由来の記録でしかない。
''''''''''''''''''''''''''''''''''
' --- Wv2Browser.cls  第9.29 段階 (設定画面にカスタム検索 URL 入力欄 + 予備予約の掃除) ---
'
'   第9.29 のスコープ (本体は Wv2SettingsBridge 側。ここは UI と小掃除):
'
'   ★(A) BuildSettingsHtml にカスタム入力欄を追加 (論点4 案a)★
'     カードグリッドの★下に独立セクション★として置く。5 枚目のカードにしない
'     のは、EngineList() を「プリセット定義」として純粋なまま保つため
'     (第9.18 の「HTML 側にエンジン名を重複させない」設計が崩れない)。
'       ・入力欄 + [適用] ボタン。custom 使用中はセクション枠が光り「使用中」バッジ。
'       ・入力中は Settings.PreviewUrlForTemplate で★適用前★のライブプレビュー。
'         エンコードを JS で自作せず VBA を通すので本番と完全同一。
'       ・不正な入力 (http:// / https:// で始まらない) はその場で赤字表示し、
'         適用しても Browser・保存ファイルには一切触らない。
'     ★<form> は使わない★ Enter で submit されて設定画面自体が遷移してしまう。
'       <input> + <button> の onclick / onkeydown で処理する。
'     ★JS の文字列・属性値はシングルクォート★ VBA 文字列内にバックスラッシュ+二重引用符は書けない
'       (VBA が " を文字列終端とみなす)。静的な HTML 部分は VBA の "" で書く。
'
'   ★(B) EnsurePrewarmScheduled の無駄な予約を止めた (第9.28b の宿題 1)★
'     ★実装時に確定した原因 (合意時は「m_prewarmClaimed なら張らない」の予定)★
'       m_adoptTimer_TimerCall は AdoptPrewarmedNow → CreatePrewarmPaneNow の順に
'       走る。後者の CreatePaneCore は WaitPaneReady で DoEvents を回すので、
'       その最中に予備切れフォールバックの AddTab が★入れ子で★割り込む。この
'       AddTab 末尾の EnsurePrewarmScheduled から見ると m_prewarmPane はまだ
'       Nothing (外側が代入前) なので予約が張られ、1.5 秒後に外側が温め終えた
'       予備を見て「既に予備タブがある (生成をスキップ)」が出ていた。
'       → m_prewarmClaimed はこの時点で既に False なので条件にならない。
'         正しい条件は★入れ子かどうか★= m_creatingPane > 0 だった。
'     ★トレードオフ★ 外側の予備生成が失敗した場合、この AddTab では予約を
'       張らないので次の AddTab / 切替まで予備が空のままになる。自動リトライ
'       (失敗時に RequestPrewarmAsync を撃つ) は、生成が恒常的に失敗する環境で
'       1.5 秒ごとの無限リトライになるため入れていない。
'
' --- Wv2Browser.cls  第9.28 段階 (プリウォーム委譲を既定化 + 予備まわりの小修正 2 件) ---
'
'   第9.28 のスコープ (9.27 の「未解決 / 次段の宿題」から 3 件を回収する回。すべて小):
'
'   ★(A) NewWindowMode の既定値を 0 → 2 に変更★
'     9.27 で本丸 (put_NewWindow によるプリウォーム委譲) が実機で通ったのに、既定は
'     0 (現行動作) のままだった。モード切替を忘れると POST リンクが 405 に戻ることが
'     9.27 ラウンド3 で実演されたので、★POST が正しく開く状態を素の状態にする★。
'       ・起動後 PREWARM_DELAY_MS (1500ms) の間は予備が無く現行動作にフォールバックする。
'         これは仕様どおり。その間はまだ最初のページのロード中でリンクを踏めないので
'         実害はなく、踏んだ場合も壊れず★印つきのログが出るだけ。
'       ・遅延の短縮はしない。1500ms は 9.27 論点1 で「Init 直後は Controller 生成が
'         連続する最混雑区間で 仕様事実 20 の窓が開く」から意図的に置いた安全マージンで、
'         既定化のために削るのは順序が逆になる。
'       ・コスト: 常時 WebView2 が 1 枚余分に常駐する (msedgewebview2.exe が 1 個多い)。
'         許容できない環境では起動後に NewWindowMode = 0 を撃てば従来どおりになる。
'
'   ★(B) m_creatingPane を Boolean → Long (カウンタ) に変更★
'     9.27 ラウンド3 で CreatePaneCore の入れ子が実際に発生した (プリウォーム生成の
'     WaitPaneReady が DoEvents を回している最中に m_newTabTimer_TimerCall が
'     AddTab → CreatePaneCore を呼んだ)。結果は正常だったが、Boolean だと
'     ★内側が終わった時点でフラグが下りる★ ため、外側がまだ生成中でも次が走りうる。
'     加算/減算にすれば「1 つでも生成中なら > 0」が保たれる。
'
'   ★(C) RequestPrewarmAsync の早期 return ログを実態に合わせた★
'     予備が「使える」のか「Claim 済み (採用待ち)」なのかで意味がまるで違うのに、
'     9.27 ではどちらも「既に予備タブがあるので何もしない」と出していた。分岐して
'     別々の文言にする。★動作は変えない★ (下記の事後報告を参照)。
'
'   ★触らないもの★ put_NewWindow の経路そのもの / プリウォームの生成・採用・補充の
'     ロジック / Deferral (使わない方針のまま) / 他のクラス。
'     Wv2Pane.cls は 9.27 のまま無変更 (モードを毎回 Browser から読む作りなので、
'     既定値の変更だけで新しい既定が効く)。
'
' --- Wv2Browser.cls  第9.27 段階 (本丸: put_NewWindow による自前タブへの委譲 / プリウォーム方式) ---
'
'   第9.27 のスコープ:
'     仕様事実 23 (9.26b) で「ランタイムに遷移そのものを渡せば POST は完全に保たれる」ことが
'     確定した。残る課題は UX ―― ランタイム任せのポップアップではなく、★自前のタブ★ に
'     その遷移を載せること。args.put_NewWindow (vtable index 4) に「こちらが用意した
'     CoreWebView2」を渡せば、ランタイムはその WebView に元のリクエストを載せてくれる。
'
'   ★採用した方式 = プリウォーム (案B / 論点 1～7 合意済み)★
'     空の Pane を常時 1 枚温めておき、NewWindowRequested が来た瞬間に
'     put_NewWindow(予備の ViewPtr) を撃って即 Exit する。
'     → ★Deferral (GetDeferral / Complete) を一切使わない★ のが最大の利点。
'       Deferral 方式は「イベントの外へ args / deferral の COM 参照を持ち越す」ため
'       自前の AddRef/Release 管理が要り、仕様事実 10 (Sentinel MEM_FREE) の隣を歩く。
'       プリウォームならイベント内でやるのは vtable 呼び 1～2 回だけで、待ちがゼロ。
'       仕様事実 14 (SafeTimer デッドロック) にも 20 (EBMode の窓) にも触れない。
'
'   ★追加したもの★
'     [宣言部]
'       ・m_newWindowMode           : 0=現行(GET横取り) / 1=ポップアップ委譲 / 2=プリウォーム委譲
'       ・m_setHandledWithNewWindow : put_NewWindow と put_Handled を併用するか (既定 True)
'       ・m_prewarmPane / m_prewarmClaimed : 予備タブ本体と「予約済み」フラグ
'       ・m_prewarmTimer / m_prewarmPending : 予備タブ生成のワンショット SafeTimer
'       ・m_adoptTimer / m_adoptPending / m_adoptUrl : 採用処理のワンショット SafeTimer
'       ・m_creatingPane            : Pane 生成中フラグ (プリウォーム TimerCall の再入ガード)
'     [切り出し = 論点2 案a]
'       ・CreatePaneCore / CreatePaneCoreBody : 旧 AddTab の前半 (Pane 生成～ハンドラ登録)。
'         AddTab とプリウォーム生成の両方から呼ぶ。AddTab は「CreatePaneCore + 後半 3 つ
'         (m_tabs.Add / TabAdded / ActivateTab)」だけの薄い関数になった。
'     [新規メソッド]
'       ・NewWindowMode (Get/Let)         : モード切替 (3 値統合 = 論点4 案b)
'       ・PopupDelegateMode (Get/Let)     : ★互換ラッパ★ 内部で NewWindowMode に委譲。
'                                           Test_9_26_Popup_On/_Off がそのまま動く。
'       ・SetHandledWithNewWindow (Get/Let): 論点5 の保険スイッチ
'       ・ClaimPrewarmViewPtr             : 予備の ViewPtr を「予約付き」で取り出す
'       ・ReleasePrewarmClaim             : put_NewWindow 失敗時に予約を戻す
'       ・RequestPrewarmAsync             : 予備タブ生成の予約 (ワンショット Timer)
'       ・RequestAdoptPrewarmedAsync      : 採用処理の予約 (ワンショット Timer)
'       ・Debug_PrewarmStatus             : 検証用の状態ダンプ
'
'   ★処理の流れ★
'     [起動] Init 成功 → RequestPrewarmAsync (1500ms 後) → CreatePrewarmPaneNow
'            → 予備 Pane を生成し Ctrl_PutIsVisible False で裏に伏せる (論点3)
'     [発火] Wv2Pane.View_OnNewWindowRequested (モード 2)
'            → ClaimPrewarmViewPtr で予備を予約 → put_NewWindow → put_Handled
'            → RequestAdoptPrewarmedAsync → 即 Exit Sub (★イベント内で何も待たない★)
'     [採用] m_adoptTimer_TimerCall (ハンドラ完了後)
'            → AdoptPrewarmedNow : m_tabs.Add → RaiseEvent TabAdded → ActivateTab
'            → CreatePrewarmPaneNow : 次の予備を補充 (論点6: 採用が先、補充が後)
'
'   ★予備が使えないとき★ 未生成 / 未 Ready / 予約済み / put_NewWindow 失敗 のいずれでも
'     Wv2Pane は現行動作 (put_Handled + RequestNewTabAsync = GET 横取り) にフォールバック
'     する。POST リンクなら 405 に戻るが、★壊れるよりはよい★。フォールバックした事実は
'     必ず ★印つきでログに出す★ ので、実機で見分けられる。
'
'   ★ログ文言の変更点★ 旧 AddTab の前半が CreatePaneCore に移ったため、
'     "Wv2Browser.AddTab: Controller 生成 OK" 等の行は
'     "Wv2Browser.CreatePaneCore: ..." に変わる。過去ログとの突き合わせ時は注意。
'
'   ★触らないもの★ RequestNewTabAsync / m_newTabTimer / 設定タブ系 / イベント 5 種 /
'     検索エンジン系 / ActivateTab / CloseTab / MoveTab。
'
' --- Wv2Browser.cls  第9.26 段階 (NewWindowRequested 委譲プローブのトグル) ---
'
'   第9.26 のスコープ (POST リンクの新タブ展開が 405 になる件の「切り分け」だけを行う回):
'     ★背景★ 職場の基幹システム (Apache Tomcat/7.0.76) で、ログイン後の「アポイント管理」
'       クリックが HTTP 405 (Request method 'GET' not supported) になる。原因は実装バグでは
'       なく方式の構造的な穴で、現行の Wv2Pane.View_OnNewWindowRequested が
'         (1) args.get_Uri で URL 文字列「だけ」を取り出し
'         (2) args.put_Handled(TRUE) でランタイムの遷移そのものを破棄し
'         (3) 新タブで Navigate (= 常に GET) し直す
'       という作りのため、元の POST のメソッド/ボディ/Referer が (2) の時点で失われる。
'
'     ★この段でやること = 事実の確認だけ★ put_Handled を「立てない」モードを足し、
'       WebView2 ランタイム自身にポップアップウィンドウで開かせる (第9.6a の挙動)。
'       そこで業務画面が正常に開けば「ランタイムに委譲すれば POST は保たれる」が確定し、
'       本丸 (args.put_NewWindow で別タブの CoreWebView2 を渡す方式) の投資が回収できると
'       分かる。逆に 405 のままなら原因は POST ボディ以外 (Referer/Cookie/セッション) に
'       あることになり、本丸に着手する前に方向転換できる。
'       ★ここを確かめずに Deferral の入れ子 (デッドロック難所) に突っ込むのが一番高くつく★
'
'     ★このモードは恒久策ではない★ ポップアップは自前のタブ UI の外に出る別ウィンドウで、
'       こちらは Controller を握らないので閉じる・並べる等の制御ができない。あくまで
'       検証 (と、どうしても業務で必要なときの応急) 用。既定は OFF。
'
'   ★追加したもの (Browser 側で一括管理する = たーぼーさん合意)★
'     ・Private m_popupDelegateMode As Boolean        (モジュール冒頭の宣言部)
'     ・Public Property Get/Let PopupDelegateMode     (既定 False = 現行動作のまま)
'     フラグを Pane ごとではなく Browser に置いたので、全タブの Wv2Pane が
'     m_browser.PopupDelegateMode を見るだけで足並みが揃う。切り替えもタブ数に依らず 1 回。
'
'   ★イミディエイトからの叩き方★ 既存の UserForm1.CurrentBrowser (第9.17 の足場) を
'     使うので UserForm1.frm は無変更で済む。実際の操作は Wv2Tests.bas の
'     Test_9_26_Popup_On / _Off / _Help に包んである。
'
'   ★触らないもの★ AddTab / RequestNewTabAsync / SafeTimer 遅延実行 (設計原則 63) /
'     イベント 5 種 / 検索エンジン系。設計原則・仕様事実の追加も無し (この段は観測のみ)。
'
' --- Wv2Browser.cls  第9.20 段階 (検索エンジン選択の永続化・復元側) ---
'
'   第9.20 の追加事項 (前回選んだ検索エンジンで起動する):
'     ★Class_Initialize で保存済みエンジンを復元する (案イ)★
'       Browser が生まれた瞬間に、一時的な Wv2SettingsBridge を New して
'       LoadEngineTemplate() を呼び、前回保存した検索エンジンの
'       テンプレート文字列があれば m_searchTemplate をそれで上書きする。
'       保存が無ければ既定 (Google) のまま。UserForm 経由でも直起動でも
'       Browser 誕生時に確定するので、どの経路でも一貫して前回値で始まる。
'
'     ★名前⇔テンプレート解決は Bridge に閉じたまま★
'       Browser はエンジン名 ("bing" 等) の語彙を知らない。Bridge の
'       LoadEngineTemplate() が「テンプレート文字列 or 空」を返すので、
'       Browser はそれを m_searchTemplate に入れるだけ。保存側 (SetEngine の
'       直後に SaveEngineName) も Bridge に閉じている (第9.20 はブリッジに集約)。
'
'     ★依存の向き★ Class_Initialize で New Wv2SettingsBridge するが、これは
'       OpenSettingsTab 内で既にやっているのと同じ作法 (型参照だけで
'       循環参照にはならない)。LoadEngineTemplate は副作用のない読取りなので
'       BindBrowser も不要 (Browser 参照なしでファイルだけ読む)。
'
' --- Wv2Browser.cls  第9.19 段階 ---
'
'   第9.19 の追加事項 (設定タブの重複防止 + 設定画面プレビューの既定表示):
'     ★設定画面 hover プレビューの既定を『現在選択中エンジン』にする★
'       BuildSettingsHtml の JS に showCurrentPreview() を追加し、(a) 設定画面を
'       開いた直後 (未 hover) と (b) いずれのカードからもカーソルが外れたとき
'       に、現在選択中のエンジンでのプレビューを表示する。以前は最後に hover
'       したカードのプレビューが残り続けていた。
'       ★戻し判定はカード個別の mouseenter/mouseleave + hoverId フラグで行う★
'       グリッドコンテナ全体の mouseleave だと、カード間の隙間や別カードへ移る
'       間はまだコンテナ内にいて発火せず、直感に反した (グリッド領域を完全に
'       抜けるまで戻らない)。各カードで hoverId を立て、mouseleave 時に
'       setTimeout(0) 越しに『まだどこかのカードに hover していなければ選択中へ
'       戻す』ことで、カード間移動のチラつきを避けつつカードから離れた瞬間に
'       戻す (次カードの mouseenter が同ティックで先に hoverId を立て直すため
'       戻し処理はスキップされる)。
'
'     ★OpenSettingsTab の冒頭で既存の設定タブを探す★
'       9.18 では [歯車] を押すたびに設定タブが新規生成され、連打すると設定タブが
'       いくつも増えてしまった。9.19 では OpenSettingsTab の冒頭で m_tabs を
'       頭から走査し、Pane.IsSettings = True のタブが既にあれば、それを
'       ActivateTab するだけで Exit Sub する (新規生成しない)。無ければ従来
'       どおり生成し、生成した Pane に IsSettings = True を立てる。
'
'     ★識別は Wv2Pane.IsSettings フラグ (案A、設計は 9.19 開発メモ参照)★
'       走査で見つからない = 設定タブが無い / 閉じられた、なので生存確認が
'       走査と一体化する。Browser 側に設定 Pane 参照を別途持ってダングリングを
'       見張る必要がない。m_settingsBridge は従来どおり最後に開いた設定タブの
'       ブリッジを 1 本保持する (既存タブは attach 時点の参照を各自保持)。
'
' --- Wv2Browser.cls  第9.18 段階 ---
'
'   第9.18 の追加事項 (設定タブ + AddHostObjectToScript による検索エンジン UI):
'     ★設定タブを開く経路を新設★ (部品3)
'       TabBar の [歯車] ボタン → OnPaneWebMessage "settings" → RequestSettingsTabAsync。
'       設定タブ生成も内部で AddTab → WaitPaneReady を通るため、newtab と同じく
'       WebMessage ハンドラ内から同期呼びするとデッドロックする (仕様事実 14)。
'       専用のワンショット SafeTimer (m_settingsTimer) で遅延実行する。
'       ・RequestSettingsTabAsync(): 保留フラグを立て New SafeTimer をワンショット起動
'       ・m_settingsTimer_TimerCall: ハンドラ完了後に OpenSettingsTab を呼ぶ
'       ・OpenSettingsTab(): AddTab → ブリッジ attach → NavigateToString → ActivateTab
'       ・BuildSettingsHtml(): 検索エンジン選択カード UI の HTML を組む
'
'     ★AddHostObjectToScript で Wv2SettingsBridge を JS に公開★ (部品2-b)
'       設定タブの Pane に Wv2SettingsBridge を attach し、設定画面の JS から
'       hostObjects.sync.Settings.SetEngine("bing") のように VBA を直接呼ばせる。
'       Wv2Browser を丸ごと渡さずブリッジ越しにするのは「JS に晒す面」を検索
'       エンジンだけに絞るため (設定画面から CloseTab 等を叩けないようにする)。
'       プリセット名⇔テンプレートの解決はブリッジに集約した (9.17 まで UserForm1 が
'       持っていた SE_* 定数と名前解決を移設。DRY)。
'
'     ★attach は NavigateToString の前 (順序A) を第一候補に実機検証★
'       登録が先なら読み込まれた JS が最初から hostObjects を使える、という仮説。
'       ダメなら ready ハンドシェイク (順序B) にフォールバックする。
'
'     ★ライフタイム★ attach したブリッジは m_settingsBridge に 1 本保持し、
'       設定タブが生きている間 WebView2 から参照される。Shutdown で解放する。
'
'   第9.16 の追加事項 (検索エンジン切替):
'     ★検索 URL テンプレートを設定値化する (案1-A + 案2-A)★
'       9.14 で BuildSearchUrl にハードコードしていた Google 固定の検索 URL を、
'       モジュール変数 m_searchTemplate に持たせて差し替え可能にする。既定は
'       Google のまま (無設定なら 9.14 と完全に同じ挙動)。
'
'     ★テンプレートは「クエリ前置型」の文字列★ (案1-A)
'       "https://www.google.com/search?q=" のように、末尾に検索語 (エンコード済) を
'       連結するだけの前置テンプレート。Google / Bing / DuckDuckGo / Yahoo など主要
'       エンジンはすべてこの形で表現できる。Enum + 対応表 (案1-B) は状態が増える
'       ので採らず、テンプレート文字列を直に持つ (YAGNI)。将来 "{q}" が末尾でない
'       稀なエンジンが必要になったら、そのとき初めてプレースホルダ方式に拡張する。
'
'     ★設定の口は本体 Property のみ、UI は付けない (案2-A)★
'       Public Property Get/Let SearchEngine で読み書きする。空文字を渡されたら
'       既定の Google テンプレートにフォールバックする軽いガードのみ入れる (URL の
'       妥当性検証はしない。テンプレートは開発者が入れる信頼できる値のため)。UI
'       からの設定 (コンボボックス等) は別ステージに切り出す (本ステージの粒度を
'       小さく保ち単独で実機検証できるようにするため)。
'
'     ★Debug_SearchEngine を追加 (ロジック検証口)★
'       現在の m_searchTemplate を返す薄い公開ラッパ。Wv2Thunks.Test_9_16_SearchEngine
'       が「エンジンを切り替えると BuildSearchUrl の結果が追従するか」を WebView2
'       起動なしで検証するために使う (Debug_NormalizeUrl と同じ位置づけ)。
'
'   第9.16 の整理事項 (同ステージで実施):
'     ★URL 純関数を Wv2Url.bas へ切り出した (案2-A / 案3-A)★
'       IsAlphaOnly / UrlEncodeComponent / Utf8Bytes の 3 つ (入力だけで出力が
'       決まる純粋なデータ変換) を Wv2Url.bas へ移設した。Wv2Json.bas と対をなす
'       「状態を持たない変換ユーティリティ群」。NormalizeUrl / BuildSearchUrl は
'       検索エンジン設定 (m_searchTemplate) というインスタンス状態に依存する
'       「ブラウザの振る舞い (ポリシー)」なので Wv2Browser に残し、内部で
'       Wv2Url.IsAlphaOnly / Wv2Url.UrlEncodeComponent を呼ぶ (ポリシーと
'       メカニズムの分離)。ロジックは無変更、置き場所と呼び出し修飾のみ変更。
'
'     ★フィーチャー検証 Sub を Wv2Tests.bas へ集約した (案1-A)★
'       Test_9_14/15/16 等 (Wv2Browser の Debug_* を叩くだけの純ロジック検証) は
'       Wv2Thunks.bas から Wv2Tests.bas へ移した。Wv2Thunks はサンク基盤に純化。
'       Wv2Browser 側は検証口 (Debug_*) をそのまま公開し続ける (変更なし)。
'
' --- Wv2Browser.cls  第9.15 段階 ---
'
'   第9.15 の追加事項 (タブのドラッグ並べ替え):
'     ★MoveTab(fromIndex, toIndex) を新設★ (案A-1)
'       m_tabs (1-based Collection) の要素を fromIndex から toIndex へ移動する。
'       Collection には要素移動 API がないため Remove → 目的位置に Add (before/
'       after 指定) で組み替える。タブ本体 (Pane) はそのまま、順序だけが変わる。
'
'     ★アクティブ追従はオブジェクト参照で行う (案A-1)★
'       並べ替えでアクティブタブの index が動くため、並べ替え前にアクティブ Pane
'       を掴んでおき、並べ替え後に m_tabs を線形走査して新しい index を逆引きする
'       (NotifyTitleChanged と同じ参照一致パターン)。可視状態は変わらないが
'       最後に ActivateTab を呼んで正す (CloseTab と同じ流儀)。
'
'     ★同期は既存経路を流用、新イベントは足さない (案A-2b)★
'       並べ替え完了後に ActivateTab(新active) を呼ぶと RaiseEvent ActiveChanged
'       が出て、TabBar/NavBar の PushSyncToJs (全体同期・設計原則 60) が走る。
'       全体同期は「状態が変わったら全タブ状態を丸ごと送り直す」思想なので、
'       アクティブ自体は変わらなくても並べ替え結果がそのまま JS に反映される。
'       → Public Event を 5 種のまま据え置ける (TabsReordered は作らない)。
'
'     ★Debug_TabOrderTitles を追加 (ロジック検証口)★
'       現在のタブ順のタイトル (なければ (t<index>)) を "|" 連結で返す薄い公開
'       ラッパ。Wv2Thunks.Test_9_15_Reorder が MoveTab の並べ替え結果を
'       WebView2 起動なしで検証するために使う (Debug_NormalizeUrl と同じ位置づけ)。
'
' --- Wv2Browser.cls  第9.14 段階 ---
'
'   第9.14 の追加事項 (検索語フォールバック、既定 Google):
'     ★アドレスバーに「検索したい言葉」を打つと Google 検索に流す★
'       9.12a で「スコープ外・次段階」と保留していた宿題 (NormalizeUrl のコメント)
'       を回収する。ドット無し / 空白ありの入力は URL ではなく検索語とみなし、
'       https://www.google.com/search?q=... に変換して Navigate する。
'
'     ★NormalizeUrl を三分岐に拡張 (案A)★
'       (1) スキーム有り ("://" を含む / "about:" 等)      → そのまま
'       (2) ドットを含み かつ 空白を含まない               → https:// を前置 (URL)
'       (3) それ以外 (空白を含む / ドット無し)             → 検索語 → BuildSearchUrl
'       Chrome/Edge のアドレスバーとほぼ同じ体感。"localhost:8080" は既存のコロン
'       分岐 (IsAlphaOnly) で (1) に落ちるので誤って検索されない。
'
'     ★URL エンコードは外部 COM ライブラリなしで実装 (案E-1)★
'       percent-encoding には UTF-8 バイト列化が必須 (日本語検索語 →
'       %E6%97%A5%E6%9C%AC 等)。本プロジェクトの WebView2 制御は外部 COM を一切
'       使わない方針だが、UTF-8 変換は「マシンコードサンク・vtable・SAFEARRAY」層
'       とは無関係な単なるデータ変換ユーティリティであり、既に HTML 書き出しで
'       ADODB.Stream を採用済み (思想的にグレーではない)。ADODB.Stream に文字列を
'       Charset="utf-8" で書き、バイト列を読み戻して RFC3986 の非予約文字以外を
'       %XX にエスケープする。
'
'     ★置き場所 = Wv2Browser の Private ヘルパー (案P-1)★
'       検索フォールバックは Browser の責務 (NormalizeUrl と同じ場所) なので、
'       UrlEncodeComponent / Utf8Bytes / BuildSearchUrl を Wv2Browser 内に Private で
'       持つ。将来 URL 処理が増えたら整理段階で Wv2Url.bas に抽出する (9.13 で JSON を
'       後追い抽出したのと同じリズム)。Wv2Json.bas は「JSON 専用」宣言なので混ぜない。
'
' --- Wv2Browser.cls  第9.12a 段階 ---
'
'   第9.12a の追加事項 (アドレスバー/ナビボタンの土台):
'     ★新規イベント TabUrlChanged(index, url, canGoBack, canGoForward)★
'       9.11a の TabTitleChanged と全く同じ疎結合パターン。Browser はナビバーの
'       存在を知らず、状態変化を撃つだけ。ナビバーは WithEvents で購読する。
'     - NotifyUrlChanged(senderPane, url, canBack, canFwd): Pane から呼ばれ、
'       オブジェクト参照の一致で index を逆引きして RaiseEvent TabUrlChanged。
'       (NotifyTitleChanged と同形)
'
'     ★HistoryChanged を標準ハンドラに昇格 (案2-A-1)★
'       従来 HistoryChanged は「SPA 用」として EnableSpaHandlers でのみ登録して
'       いた (9.10d-2b、9.10c-2 論点⑩)。しかし戻る/進むボタンの活性状態は SPA に
'       限らず全タブで追跡する必要があるため、RegisterStandardHandlers に移して
'       全タブで自動登録する。→ 標準ハンドラは 4 種 → 5 種に。
'       EnableSpaHandlers は DOMContentLoaded のみを担当する形に変更 (SPA 固有の
'       イベントはこちらに残す)。二重登録を避けるため EnableSpaHandlers から
'       HistoryChanged の登録を外した。
'
'     ★アクティブタブへの操作委譲 (ナビバー用の公開 API)★
'       ナビバーは「アクティブタブ」に操作を送りたいが、Pane を直接触らせると
'       疎結合が壊れる。Browser が代理する薄いメソッドを用意する。
'         NavigateActive(url) / GoBackActive / GoForwardActive /
'         ReloadActive / StopActive
'         GetActiveUrl / GetActiveCanGoBack / GetActiveCanGoForward
'       ★設計原則 63★ いずれも「完了待ちを含まない」操作なので、イベント
'         ハンドラ内から同期で呼んで安全 (9.11d の newtab のような遅延実行は不要)。
'
'     ★URL のスキーム補完 (案3-A)★ NavigateActive はスキームがなければ https:// を
'       前置する ("example.com" → "https://example.com")。検索語フォールバックは
'       スコープ外 (次段階)。
'
' --- Wv2Browser.cls  第9.11c 段階 ---
'
'   第9.11c の追加事項 (新規タブの初期 Bounds をコンテンツ矩形に合わせる、案A):
'     ★問題★ AddTab は新タブの初期 Bounds を GetClientRect(m_hwndHost)=フレーム
'       全域にしていた。タブバー帯 (上部 36px) を差し引かないため、起動後に
'       追加したタブがタブバーと重なる (UserForm_Resize が走るまで直らない)。
'     ★原因の構造★ Browser はタブバーの存在を知らない (疎結合) ので、実際の
'       コンテンツ矩形 (タブバー分を除いた領域) は上位 (UserForm) しか知らない。
'     ★対処 (案A)★ レイアウトの真実源は上位のままにしつつ、Browser は「上位が
'       最後に適用したコンテンツ矩形」をキャッシュして新タブに一貫適用する。
'       - 新規フィールド: m_lastContentRect As RECT / m_hasContentRect As Boolean
'       - ResizeAll: 渡された rc を m_lastContentRect に保存し m_hasContentRect=True。
'       - AddTab: m_hasContentRect が True ならその矩形を初期 Bounds に使う。
'         未設定時 (ResizeAll 未実行) は従来通り GetClientRect(全域) にフォールバック。
'       ★疎結合は不変★ Browser はタブバーを知らないまま「上位が指定した領域」を
'       覚えるだけ。UserForm は無変更 (起動時 UserForm_Resize が既に ResizeAll を
'       呼ぶので、以降の AddTab は自動的に正しい矩形で生まれる)。
'
' --- Wv2Browser.cls  第9.11b 段階 ---
'
'   第9.11b の追加事項 (Wv2TabBar に Environment を貸す口):
'     - Property Get Environment() As Wv2Environment を追加。Wv2TabBar が
'       タブバー描画用ペインを生成する際、Browser と同じ Environment を
'       共有するため (Environment は 1 個を全ペインで共有)。
'       ★疎結合★ Browser は「Environment を貸す」だけで、借りた側が
'       何に使うか (タブバーか別用途か) を知らない。
'
'   第9.11a の追加事項 (疎結合タブバー用イベント口 + CloseTab、案A):
'     ★疎結合★ Browser はタブバーの実装 (WebView2 製/MsForms 製/不要) を
'       一切知らない。状態変化を Public Event で外に通知するだけ (RaiseEvent は
'       聞き手ゼロでも無害に素通りする)。受け手は WithEvents m_browser で受信。
'     - 新規イベント: TabAdded(index) / TabClosed(index) / ActiveChanged(index) /
'       TabTitleChanged(index, title)
'     - AddTab: Collection 追加直後に RaiseEvent TabAdded (ActivateTab の前)。
'       直後に ActivateTab を呼ぶので、タブ追加時は TabAdded → ActiveChanged の
'       順で 2 発出る。
'     - ActivateTab: 成功時に RaiseEvent ActiveChanged。
'     - CloseTab(index): 新設。ClearBrowserRef (設計原則 58) → Ctrl_Close (vtable 24、
'       二重 Close 無害) → m_tabs.Remove → アクティブ補正 → RaiseEvent TabClosed。
'       アクティブ補正は 4 分岐 (0 個=タブ 0 許容 / 閉じたのがアクティブより前 /
'       アクティブ自身=右隣優先なければ左隣 / アクティブより後)。可視修復は
'       ActivateTab に集約 (設計原則 57)。最低タブ数ポリシーは上位に委ねる。
'     - NotifyTitleChanged(senderPane, title): 新設。Pane からタイトル変更を
'       受け、オブジェクト参照の一致で index を逆引いて RaiseEvent TabTitleChanged。
'
'   第9.10e の追加事項 (NewWindowRequested の新タブ化サポート、案Y + 案β Timer):
'     ★デッドロック回避 (案β)★ Wv2Pane から NewWindowRequested ハンドラ内で
'       AddTabWithUrl を直接呼ぶとデッドロックする (仕様事実 14) ため、
'       RequestNewTabAsync + SafeTimer (ワンショット) でハンドラ完了後に実行する。
'     - 新規フィールド: m_newTabTimer (WithEvents SafeTimer) / m_pendingUrls (保留キュー)
'     - RequestNewTabAsync(url): 保留キューに積み、New SafeTimer でワンショット起動
'       (SafeTimer.StartTimer は同間隔の再起動を無視し停止後も m_uElapse が残るため、
'        要求ごとに New して確実に起動させる)
'     - m_newTabTimer_TimerCall: ハンドラ外でキューを順に AddTabWithUrl 処理
'     - AddTab で pane.SetBrowser Me を実行し、各 Pane に自分 (Browser) への
'       後方参照を張る (Wv2Pane が NewWindowRequested 時に AddTabWithUrl を
'       呼べるようにする)。
'     - Shutdown: 循環参照 (Browser.m_tabs ⇔ Pane.m_browser) を断つ。各 Pane の
'       ClearBrowserRef を呼んでから解放する。UserForm_Terminate から明示的に
'       呼ぶ (設計原則 58)。UserForm 側の実装コストは Shutdown 呼び出し 1 行のみ。
'
'   第9.10d-2b の追加事項 (SPA 用ヘルパー、案②-c + 案F-1):
'     - EnableSpaHandlers(pane): SPA 用イベント (HistoryChanged + DOMContentLoaded)
'       を pane に登録する公開ヘルパー。内部で pane.View_AddHistoryChanged +
'       pane.View2_AddDOMContentLoaded を呼ぶだけ。標準 4 ハンドラ
'       (RegisterStandardHandlers) には含めない SPA 専用イベントをここに集約
'       (9.10c-2 論点⑩: SPA 用は明示登録)。既存タブへの後付けも可能。
'     - AddTabWithUrlForSpa(hostName, folderPath, fileName): SPA タブを一括生成。
'       AddTab → View3_SetVirtualHostNameToFolderMapping (Navigate 前・View 取得後) →
'       Navigate(https://<host>/<file>) → EnableSpaHandlers の順。
'       ★仕様事実 13★ Navigate 先はファイル名まで明示する (ルート "/" は
'       ディレクトリに解決され ERR_ACCESS_DENIED)。
'
'   第9.10c-2 の追加事項:
'     - RegisterStandardHandlers(pane): 標準 4 ハンドラ (NavigationStarting/
'       Completed/DocumentTitleChanged/NewWindowRequested) を AddTab 内で自動登録
'       (案⑩-A)。
'
'   第9.10c-1 の追加事項:
'     - ActivateTab(index): タブの可視制御 (アクティブタブだけ可視) を
'       pane.Ctrl_PutIsVisible で実現。可視制御はこの 1 箇所に集約 (設計原則 57)。
'       AddTab はこれに委譲する。
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

' --- 第9.16 段階で追加: 既定の検索 URL テンプレート (前置型、末尾に検索語を連結) ---
Private Const DEFAULT_SEARCH_TEMPLATE As String = "https://www.google.com/search?q="

' ============================================================
' フィールド
' ============================================================
Private m_env      As Wv2Environment   ' 1 個、全タブ共有
Private m_tabs     As Collection       ' Wv2Pane を格納 (1-based)
Private m_dbgTagMap As Collection            ' 第9.15 検証用: ObjPtr(key) → タグ ("A","B",..)。本番未使用時は Nothing
Private m_hwndHost As LongPtr          ' 親フレーム HWND (全タブ共通)
Private m_state    As BrowserState

' --- 第9.16 段階で追加: 検索エンジンの検索 URL テンプレート (案1-A) ---
'   BuildSearchUrl が末尾に「エンコード済み検索語」を連結する前置テンプレート。
'   既定は Google。Class_Initialize でセットし、Property SearchEngine で差し替える。
'   例: "https://www.google.com/search?q="  /  "https://www.bing.com/search?q="
Private m_searchTemplate As String

' 第9.10c-1 段階で追加: アクティブタブの index (1-based、0 = アクティブなし)。
'   ActivateTab が更新する唯一の状態。ActivePane はこれを参照する。
Private m_activeIndex As Long

' --- 第9.11c 段階で追加: コンテンツ矩形キャッシュ (案A) ---
'   上位 (UserForm) が ResizeAll で最後に適用した矩形を覚えておき、AddTab の
'   初期 Bounds に流用する。これで起動後に追加したタブもタブバー帯を避けた
'   正しい領域に置かれる。m_hasContentRect が False の間 (ResizeAll 未実行) は
'   AddTab は GetClientRect(全域) にフォールバックする。
'   ★Browser はタブバーを知らない★ ここに入るのは「上位が指定した領域」で
'   あって、タブバーの高さを Browser 自身が計算するわけではない (疎結合維持)。
Private m_lastContentRect As RECT
Private m_hasContentRect  As Boolean

' 第9.10e 段階 (Timer 遅延実行) で追加:
'   NewWindowRequested の新タブ生成を、ハンドラ完了後に遅延実行するための
'   ワンショット SafeTimer と保留 URL キュー。WithEvents で TimerCall を受ける。
Private WithEvents m_newTabTimer As SafeTimer
Attribute m_newTabTimer.VB_VarHelpID = -1
Private m_pendingUrls As Collection

' 第9.18 段階 (設定タブの遅延生成) で追加:
'   TabBar の [歯車] ボタン → OnPaneWebMessage の "settings" ケースから
'   RequestSettingsTabAsync が呼ばれる。設定タブを開く処理も内部で
'   AddTab → WaitPaneReady (Controller 生成完了待ち) を通るため、WebMessage
'   ハンドラ内から同期で呼ぶと 仕様事実 14 のデッドロックに嵌まる (newtab と
'   同じ構造)。そこで newtab と同型のワンショット SafeTimer で遅延実行する。
'   ★newtab のキュー (m_pendingUrls) には相乗りしない★ あちらは URL 文字列
'     専用。設定タブは「開く要求が来たフラグ」だけで足りるので専用に持つ
'     (種別を混ぜるより読みやすく、設計原則63 の一発撮りに素直)。
'   ★m_settingsBridge★ 設定タブに attach する Wv2SettingsBridge を保持する。
'     設定タブが生きている間 WebView2 から参照されるので Browser も 1 本持つ。
Private WithEvents m_settingsTimer As SafeTimer
Attribute m_settingsTimer.VB_VarHelpID = -1

' --- K-2: 設定ビューの全体同期 (設計原則60) ---
'   検索エンジンが変わったら、設定画面を表示中の全タブへ現在値を押し込む。
'   ★SafeTimer を挟むのは仕様事実27 のデッドロック回避★
Private WithEvents m_syncTimer As SafeTimer
Attribute m_syncTimer.VB_VarHelpID = -1
Private m_syncPending As Boolean
Private m_settingsPending As Boolean
Private m_settingsBridge As Wv2SettingsBridge

' --- 第9.26 段階で追加 → 第9.27 で撤去: m_popupDelegateMode (Boolean) ---
'   9.26 の委譲トグルは Boolean 1 個だったが、9.27 でモードが 3 通りになったため
'   m_newWindowMode (Long) に統合した (論点4 案b)。Boolean を 2 個並べると意味の
'   ない組み合わせが 4 通りできてしまうため。
'   ★PopupDelegateMode プロパティ自体は互換ラッパとして残っている★ ので、
'     Test_9_26_Popup_On / _Off は 9.27 でもそのまま動く。

' --- 第9.27 段階で追加: NewWindowRequested の 3 値モードとプリウォーム機構 ---
'   ★モード値 (Wv2Pane 側の Private Const と対応)★
'     0 = NWM_LEGACY  : 現行動作。put_Handled(TRUE) + RequestNewTabAsync で自前の新タブ。
'                       URL しか引き継がないので POST は GET に化ける (= 405 の原因)。
'     1 = NWM_POPUP   : 9.26 のプローブ。put_Handled を立てずランタイムに丸投げ。
'                       POST は保たれるがポップアップになりタブバーに乗らない。
'     2 = NWM_PREWARM : 9.27 の本命。予備タブの CoreWebView2 を put_NewWindow に渡す。
'                       POST が保たれ、かつ自前のタブに収まる。
'   ★m_setHandledWithNewWindow (論点5 の保険)★
'     put_NewWindow をセットしたとき put_Handled(TRUE) も併せて立てるかどうか。
'     公式サンプルは両方立てる形なので既定 True。万一「白いまま」「二重に開く」等が
'     起きたときの切り分け用に False にできるようにしてある。
Private m_newWindowMode As Long
Private m_setHandledWithNewWindow As Boolean

'   ★プリウォーム (予備タブ)★
'     m_prewarmPane    : 温めてある空の Pane (m_tabs には入れない = タブバーに出ない)
'     m_prewarmClaimed : put_NewWindow に渡し済み (= 採用待ち) を示す。二重に渡すと
'                        同じ WebView に 2 つの遷移が載るので、Claim で 1 回だけに絞る。
Private m_prewarmPane As Wv2Pane
Private m_prewarmClaimed As Boolean

'   ★予備タブ生成のワンショット Timer (論点1 案b)★
'     Init 成功直後に予約し、起動シーケンス (Env → 最初のタブ → TabBar/NavBar) が
'     落ち着いた頃に発火させる。起動の最も混雑した区間に 4 枚目の Controller 生成を
'     差し込まないための遅延 (仕様事実 20 の窓を避ける)。
Private WithEvents m_prewarmTimer As SafeTimer
Attribute m_prewarmTimer.VB_VarHelpID = -1
Private m_prewarmPending As Boolean

'   ★採用処理のワンショット Timer★
'     NewWindowRequested ハンドラの中で m_tabs.Add / ActivateTab までやると、
'     イベント内で重い処理を回すことになる。RequestNewTabAsync と同型に、
'     ハンドラ完了後の TimerCall で採用 → 補充の順に処理する。
Private WithEvents m_adoptTimer As SafeTimer
Attribute m_adoptTimer.VB_VarHelpID = -1
Private m_adoptPending As Boolean
Private m_adoptUrl As String

'   ★Pane 生成中カウンタ (再入ガード)★
'     CreatePaneCore は内部の WaitPaneReady で DoEvents を回すため、その最中に
'     m_prewarmTimer_TimerCall が割り込みうる。生成中に更に生成を始めないよう見張る。
'     ★第9.28 で Boolean → Long に変更★ 9.27 ラウンド3 で入れ子が実際に発生した。
'     Boolean だと内側の生成が終わった時点でフラグが下りてしまい、外側がまだ生成中でも
'     次の生成が走りうる。加算/減算なら「1 つでも生成中なら > 0」が保たれる。
Private m_creatingPane As Long

'   予備タブ生成を予約してから実際に生成するまでの待ち時間 (ms)。
'   起動シーケンスが終わるのを待つのが目的なので、短すぎると再入ガードで見送りになり、
'   長すぎると最初の POST リンクに間に合わない。1.5 秒を初期値とする。
Private Const PREWARM_DELAY_MS As Long = 1500

' --- 第9.11a 段階で追加: 疎結合タブバー用イベント口 (案A) ---
'   Browser は「誰が聞いているか」を一切知らない。タブバーが WebView2 製でも
'   MsForms 製でも不要でも、この口は同じ。RaiseEvent は聞き手ゼロでも無害に
'   素通りする。受け手は Private WithEvents m_browser As Wv2Browser で受信する。
'   ・TabAdded(index)              : タブが追加された (Collection 追加直後)
'   ・TabClosed(index)             : タブが閉じられた (Collection 除去直後)
'   ・ActiveChanged(index)         : アクティブタブが変わった (ActivateTab 成功時)
'   ・TabTitleChanged(index,title) : タブのタイトルが変わった (Pane から NotifyTitleChanged 経由)
Public Event TabAdded(ByVal index As Long)
Public Event TabClosed(ByVal index As Long)
Public Event ActiveChanged(ByVal index As Long)
Public Event TabTitleChanged(ByVal index As Long, ByVal title As String)

' --- 第9.12a: URL 状態の変化を通知 (アドレスバー / 戻る・進むボタン用) ---
'   Pane の NavigationCompleted / HistoryChanged から NotifyUrlChanged 経由で撃たれる。
'   ★疎結合★ Browser はナビバーの存在を知らない。聞き手ゼロでも RaiseEvent は
'   無害に素通りする。
Public Event TabUrlChanged(ByVal index As Long, ByVal url As String, _
                           ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)


' ============================================================
' Class_Initialize
' ============================================================
Private Sub Class_Initialize()
    Set m_tabs = New Collection
    Set m_env = Nothing
    m_hwndHost = 0
    m_state = Bs_Created
    m_activeIndex = 0
    Set m_pendingUrls = New Collection   ' 第9.10e (NewWindowRequested 保留 URL キュー)
    Set m_newTabTimer = Nothing          ' 第9.10e (要求ごとに New する)
    m_hasContentRect = False             ' 第9.11c (コンテンツ矩形 未キャッシュ)
    m_searchTemplate = DEFAULT_SEARCH_TEMPLATE  ' 第9.16 (既定 Google。Property で差し替え可)
    Set m_settingsTimer = Nothing        ' 第9.18 (要求ごとに New する)
    m_settingsPending = False            ' 第9.18 (設定タブ生成の保留フラグ)
    Set m_settingsBridge = Nothing       ' 第9.18 (attach するブリッジ)

    ' 第9.27: NewWindowRequested のモードとプリウォーム機構の初期化
    '   ★第9.28 で既定を 0 → 2 に変更★ 本丸が実機で通ったので、POST リンクが
    '     正しく自前タブで開く状態を「素の状態」にする。起動後 PREWARM_DELAY_MS の
    '     間だけ予備が未生成で現行動作にフォールバックするが、その間はまだ最初の
    '     ページのロード中でリンクを踏めないため実害はない。
    '     従来どおりにしたい場合は起動後に NewWindowMode = 0 を撃つ。
    m_newWindowMode = 2                  ' 既定はプリウォーム委譲 (NWM_PREWARM)
    m_setHandledWithNewWindow = True     ' put_NewWindow と put_Handled を併用 (論点5)
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False
    Set m_prewarmTimer = Nothing         ' 要求ごとに New する
    m_prewarmPending = False
    Set m_adoptTimer = Nothing           ' 要求ごとに New する
    m_adoptPending = False
    m_adoptUrl = ""
    m_creatingPane = 0                   ' 第9.28: Boolean → Long (入れ子カウンタ)

    ' 第9.20: 保存済みの検索エンジンを復元する (案イ)。
    '   一時ブリッジを New し、LoadEngineTemplate() で保存値のテンプレートを
    '   取る。非空なら m_searchTemplate を上書き、空なら既定 (Google) のまま。
    '   ブリッジは Browser 参照を持たずファイルだけ読むので BindBrowser 不要。
    Dim initBridge As Wv2SettingsBridge
    Set initBridge = New Wv2SettingsBridge
    Dim savedTemplate As String
    savedTemplate = initBridge.LoadEngineTemplate()
    If Len(savedTemplate) > 0 Then
        m_searchTemplate = savedTemplate
        Debug.Print "Wv2Browser.Class_Initialize: 保存済みエンジンを復元 = " & savedTemplate
    End If
    Set initBridge = Nothing
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

    ' --- 第9.27: 予備タブ (プリウォーム) の生成を予約する (論点1 案b) ---
    '   ここで同期生成しないのは、起動シーケンスの最も混雑した区間 (この直後に
    '   最初のタブ / TabBar / NavBar の Controller が次々生成される) を避けるため。
    '   ワンショット Timer で PREWARM_DELAY_MS 後に静かに温める。
    RequestPrewarmAsync

    Init = True
End Function


' ============================================================
' AddTab ? 器だけのタブを 1 個追加する (Navigate しない)
'   新 Wv2Pane を生成し、共有 Environment から Controller を生成、
'   Ready まで待って Bounds を設定し、Collection に追加する。
'   戻り値: 生成された Wv2Pane (失敗時 Nothing)
' ============================================================
Public Function AddTab() As Wv2Pane
    ' --- 前半 (Pane 生成 ～ 標準ハンドラ登録) は CreatePaneCore に切り出した (第9.27) ---
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then Exit Function

    ' --- 6. Collection に追加 ---
    m_tabs.Add pane
    Debug.Print "Wv2Browser.AddTab: タブ追加 OK (現在のタブ数=" & m_tabs.Count & ")"

    ' --- 6.5 タブ追加を通知 (第9.11a、疎結合イベント) ---
    '   ActivateTab より前に出す。これにより聞き手は「まず追加、次に
    '   アクティブ変更」の順で受け取れる (AddTab は直後に ActivateTab を呼ぶ)。
    RaiseEvent TabAdded(m_tabs.Count)

    ' --- 7. 追加したタブを自動でアクティブ化 (論点⑦ a) ---
    '   可視制御は自前で行わず ActivateTab に委譲する (責務分離)。
    '   これにより「可視制御のロジックは ActivateTab の 1 箇所だけ」を保つ。
    ActivateTab m_tabs.Count

    ' --- 8. 予備タブが無ければ生成を予約し直す (第9.27) ---
    '   起動シーケンス中は再入ガードでプリウォームが見送られることがある。
    '   AddTab が終わるたびにここで張り直せば、最終的に必ず 1 枚温まる。
    EnsurePrewarmScheduled

    Set AddTab = pane
End Function


' ============================================================
' CreatePaneCore / CreatePaneCoreBody (第9.27、旧 AddTab 前半の切り出し = 論点2 案a)
'
'   「Pane を 1 個作って使える状態にする」までを担当する。m_tabs への登録・
'   TabAdded の発火・ActivateTab は【やらない】。呼び出し側は 2 つ:
'     ・AddTab            → 作った Pane をそのままタブとして登録する
'     ・CreatePrewarmPaneNow → 作った Pane を予備として裏に伏せる
'
'   ★2 段構え (Core / CoreBody) の理由★
'     CoreBody には失敗時の Exit Function が複数ある。再入ガード (m_creatingPane) を
'     CoreBody の中で上げ下げすると、どの Exit でも必ず下ろす保証が要りコードが汚れる。
'     薄いラッパ (Core) で挟めば、Body がどこで抜けてもガードは 1 箇所で下りる。
'
'   ★このガードが要る理由★ CoreBody 内の WaitPaneReady は DoEvents を回す。その最中に
'     m_prewarmTimer_TimerCall が割り込むと Pane 生成が入れ子になる。ガードを見て
'     見送らせる (仕様事実 14 / 20 に近付かないための予防)。
' ============================================================
Private Function CreatePaneCore() As Wv2Pane
    ' ★第9.28: フラグ → カウンタ★ 入れ子で呼ばれても、外側が終わるまで 0 に戻らない。
    m_creatingPane = m_creatingPane + 1
    Set CreatePaneCore = CreatePaneCoreBody()
    m_creatingPane = m_creatingPane - 1
    If m_creatingPane < 0 Then m_creatingPane = 0   ' 念のための下限クリップ
End Function

Private Function CreatePaneCoreBody() As Wv2Pane
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.CreatePaneCore: Environment が未 Ready (state=" & m_state & ")"
        Exit Function
    End If

    ' --- 1. Pane を生成し、共有 Environment から Controller を生成 ---
    Dim pane As Wv2Pane
    Set pane = New Wv2Pane
    pane.Init m_env, m_hwndHost

    ' --- 2. Controller Ready 待ち (内部に閉じ込め) ---
    If Not WaitPaneReady(pane) Then
        If pane.IsFailed Then
            Debug.Print "Wv2Browser.CreatePaneCore: Controller 生成失敗 (state=" & _
                        pane.state & ", lastError=&H" & Hex(pane.LastError) & ")"
        Else
            Debug.Print "Wv2Browser.CreatePaneCore: Controller 生成タイムアウト (state=" & _
                        pane.state & ")"
        End If
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: Controller 生成 OK (pCtrl=" & pane.CtrlPtr & ")"

    ' --- 3. Bounds を初期設定する (第9.11c: コンテンツ矩形キャッシュ優先、案A) ---
    '   上位 (UserForm) が ResizeAll で矩形を教えてくれていれば、それを使う
    '   (タブバー帯を避けた正しい領域)。まだなら従来通りフレーム全域。
    '   いずれにせよ最終的なレイアウトは UserForm_Resize が再発行して整える。
    Dim rc As RECT
    If m_hasContentRect Then
        rc = m_lastContentRect
        Debug.Print "Wv2Browser.CreatePaneCore: 初期 Bounds にキャッシュ済みコンテンツ矩形を使用"
    Else
        GetClientRect m_hwndHost, rc
        Debug.Print "Wv2Browser.CreatePaneCore: 初期 Bounds にフレーム全域を使用 (ResizeAll 未実行)"
    End If
    Dim hr As Long
    hr = pane.Ctrl_PutBounds(rc)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_PutBounds 失敗 hr=&H" & Hex(hr)
        ' Bounds 失敗は致命ではないので続行 (タブ自体は生成できている)
    End If

    ' --- 4. ICoreWebView2 (View) を取得 ---
    '   ★重要★ Controller 生成だけでは m_pView = 0 のまま。View_Navigate /
    '   View_NavigateToString は m_pView が必要なので、ここで必ず取得する。
    '   ★第9.27 では更に重要★ put_NewWindow に渡すのはこの ViewPtr そのもの。
    hr = pane.Ctrl_GetCoreWebView2()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: Ctrl_GetCoreWebView2 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    If pane.ViewPtr = 0 Then
        Debug.Print "Wv2Browser.CreatePaneCore: ViewPtr が 0 (取得成功だが値が空)"
        Exit Function
    End If
    Debug.Print "Wv2Browser.CreatePaneCore: View 取得 OK (pView=" & pane.ViewPtr & ")"

    ' --- 5. 標準ナビゲーションハンドラを自動登録 (案⑩-A、第9.10c-2) ---
    '   基盤的なイベント (NavigationStarting/Completed/DocumentTitleChanged/
    '   NewWindowRequested/HistoryChanged) は全タブで自動登録する。
    '   ★予備タブにも登録する★ 採用された瞬間からタイトル・URL が追跡できるように。
    RegisterStandardHandlers pane

    ' --- 5.5 Wv2Pane に後方参照をセット (第9.10e、案Y) ---
    '   NewWindowRequested で Pane が m_browser の機能を呼べるようにする。
    '   ★循環参照になるので、解放時は Shutdown → ClearBrowserRef で断つ (設計原則 58)。
    pane.SetBrowser Me

    Set CreatePaneCoreBody = pane
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

    ' --- 第9.12a: HistoryChanged を標準ハンドラに昇格 (案2-A-1) ---
    '   戻る/進むボタンの活性状態は SPA に限らず全タブで追跡する必要があるため、
    '   SPA 専用扱い (EnableSpaHandlers) をやめて全タブで自動登録する。
    '   pushState/replaceState/popstate に加え、通常の戻る/進むでも発火する。
    hr = pane.View_AddHistoryChanged()
    If hr <> 0 Then Debug.Print "Wv2Browser.RegisterStandardHandlers: HistoryChanged 失敗 hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.RegisterStandardHandlers: 標準ハンドラ 5 種 登録完了"
End Sub


' ============================================================
' EnableSpaHandlers ? SPA 用イベントハンドラを pane に登録する (第9.10d-2b、案②-c)
'
'   SPA (history 操作 + DOM 構築) を観察するためのイベント 2 種を登録する:
'     - HistoryChanged   (基底 ICoreWebView2、vtable 13) : pushState/replaceState/
'                          popstate すべてで発火 (仕様事実 12)
'     - DOMContentLoaded (ICoreWebView2_2、vtable 64)     : 初回ロード時のみ発火
'                          (pushState では発火しない = HistoryChanged との役割分担)
'
'   標準 4 ハンドラ (RegisterStandardHandlers) には含めず、SPA タブにだけ
'   明示登録する (9.10c-2 論点⑩)。AddTabWithUrlForSpa が内部で呼ぶが、
'   Public にして既存タブへの後付けもできるようにする (案②-c)。
'
'   ★設計原則 45★ View_AddHistoryChanged / View2_AddDOMContentLoaded は
'   Wv2Pane 側で VTABLE_OBJ_OFFSET を通す前提で正しく実装済み。ここでは
'   その公開メソッドを呼ぶだけ (自前でスロットを触らない)。
' ============================================================
Public Sub EnableSpaHandlers(ByVal pane As Wv2Pane)
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.EnableSpaHandlers: pane が Nothing"
        Exit Sub
    End If

    ' --- 第9.12a: HistoryChanged はここから外した ---
    '   RegisterStandardHandlers (全タブ自動登録) に昇格したため、ここで再登録すると
    '   二重登録になる。SPA 固有のイベント (DOMContentLoaded) だけを残す。
    Dim hr As Long
    hr = pane.View2_AddDOMContentLoaded()
    If hr <> 0 Then Debug.Print "Wv2Browser.EnableSpaHandlers: DOMContentLoaded 失敗 hr=&H" & Hex(hr)

    Debug.Print "Wv2Browser.EnableSpaHandlers: SPA ハンドラ 1 種 " & _
                "(DOMContentLoaded) 登録完了 ※HistoryChanged は標準ハンドラに昇格 (9.12a)"
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
' AddTabWithUrlForSpa ? SPA タブを一括生成する (第9.10d-2b、案F-1)
'
'   ローカルフォルダを仮想ホストにマッピングし、そこの HTML を https origin で
'   開いて SPA 用ハンドラまで登録する、SPA 検証・運用の入口メソッド。
'
'   手順 (順序が重要):
'     1. AddTab            : Controller/View 取得・標準 4 ハンドラ登録・アクティブ化
'     2. SetVirtualHostNameToFolderMapping : ★Navigate の前・View 取得後に必須★
'                            (View 取得は AddTab 内で完了しているので順序 OK)
'     3. View_Navigate     : https://<hostName>/<fileName>
'                            ★仕様事実 13★ ルート "/" は ERR_ACCESS_DENIED になるので
'                            ファイル名まで明示する
'     4. EnableSpaHandlers : HistoryChanged + DOMContentLoaded を登録
'
'   引数:
'     hostName   : 仮想ホスト名 (例 "appassets.example")
'     folderPath : マッピング先ローカルフォルダの絶対パス
'     fileName   : フォルダ内の開始ファイル名 (例 "index.html")
'
'   accessKind は ALLOW (=1) 固定。必要になれば引数化する。
'
'   戻り値: 生成した Wv2Pane (失敗時は Nothing、ただしタブ生成後の
'           マッピング/Navigate 失敗はタブ自体は返す = 部分成功を握れる)
' ============================================================
Public Function AddTabWithUrlForSpa( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal fileName As String) As Wv2Pane

    ' --- 1. 器のタブを追加 (View 取得・標準ハンドラ登録・アクティブ化まで) ---
    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then Exit Function

    ' --- 2. 仮想ホストマッピングを設定 (★Navigate の前に必須) ---
    Dim hr As Long
    hr = pane.View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, 1)   ' 1 = ALLOW
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: マッピング失敗 hr=&H" & Hex(hr)
        Set AddTabWithUrlForSpa = pane   ' タブ自体は生成済みなので返す
        Exit Function
    End If

    ' --- 3. https://<host>/<file> へ Navigate (仕様事実 13: ファイル名まで明示) ---
    Dim url As String
    url = "https://" & hostName & "/" & fileName
    hr = pane.View_Navigate(url)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.AddTabWithUrlForSpa: Navigate OK (" & url & ")"
    End If

    ' --- 4. SPA 用ハンドラ (HistoryChanged + DOMContentLoaded) を登録 ---
    EnableSpaHandlers pane

    Set AddTabWithUrlForSpa = pane
End Function


' ============================================================
' TabCount ? 現在のタブ数
' ============================================================
Public Property Get TabCount() As Long
    TabCount = m_tabs.Count
End Property


' ============================================================
' ActivePane ? アクティブタブの Wv2Pane を返す
'   第9.10c-1 で activeIndex ベースに正式化 (論点④の正式化)。
'   m_activeIndex (1-based、0 = なし) が指すタブを返す。
' ============================================================
Public Property Get ActivePane() As Wv2Pane
    If m_activeIndex < 1 Or m_activeIndex > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivePane: アクティブタブなし (activeIndex=" & _
                    m_activeIndex & ", tabCount=" & m_tabs.Count & ")"
        Set ActivePane = Nothing
        Exit Property
    End If
    Set ActivePane = m_tabs.Item(m_activeIndex)
End Property

' ============================================================
' Environment ? 共有 Environment を貸し出す (第9.11b、Wv2TabBar 用)
'   Wv2TabBar がタブバー描画用の単一ペインを生成するのに、この
'   Browser が保持する Environment (全ペイン共有) を使う。Init 完了
'   (Bs_EnvReady) 後に使うこと。未 Ready の間は Nothing を返す。
' ============================================================
Public Property Get Environment() As Wv2Environment
    Set Environment = m_env
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
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.ActivateTab: index 範囲外 (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
        Set pane = m_tabs.Item(i)
        hr = pane.Ctrl_PutIsVisible(i = index)
        If hr <> 0 Then
            Debug.Print "Wv2Browser.ActivateTab: タブ" & i & " の PutIsVisible 失敗 hr=&H" & Hex(hr)
        End If
    Next i

    m_activeIndex = index
    Debug.Print "Wv2Browser.ActivateTab: タブ" & index & " をアクティブ化 (可視)"

    ' --- アクティブ変更を通知 (第9.11a、疎結合イベント) ---
    RaiseEvent ActiveChanged(index)
    ActivateTab = True
End Function


' ============================================================
' CloseTab ? 指定 index のタブを閉じる (第9.11a 段階で追加)
'
'   引数: index (1-based)
'   戻り値: 成功で True
'
'   手順:
'     1. index 範囲チェック
'     2. 閉じる Pane の後方参照を断ち (ClearBrowserRef、設計原則 58) →
'        Ctrl_Close で Controller を明示クローズ (vtable 24、二重 Close は無害)
'     3. m_tabs.Remove index (参照カウント 0 で Pane.Class_Terminate)
'     4. m_activeIndex の補正 + 可視状態の修復 (ActivateTab に集約、設計原則 57):
'        ・閉じたのがアクティブ : 右隣優先、なければ左隣 (末尾) をアクティブ化
'        ・閉じたのがアクティブより前 : m_activeIndex を 1 減らして同じタブを維持
'        ・タブが 0 個になった : m_activeIndex = 0 (アクティブなし、タブ 0 許容)
'     5. RaiseEvent TabClosed(index) (アクティブが変われば ActivateTab 経由で ActiveChanged も出る)
'
'   ★タブ 0 許容★ 最低タブ数のポリシーは Browser は持たない。上位 (タブバー
'   実装側) に委ねることで、タブバー不要な使い方でも余計な制約を持ち込まない。
' ============================================================
Public Function CloseTab(ByVal index As Long) As Boolean
    If index < 1 Or index > m_tabs.Count Then
        Debug.Print "Wv2Browser.CloseTab: index 範囲外 (index=" & index & _
                    ", tabCount=" & m_tabs.Count & ")"
        Exit Function
    End If

    Dim pane As Wv2Pane
    Set pane = m_tabs.Item(index)

    ' --- 1. 後方参照を先に断つ (設計原則 58) ---
    '   循環参照 Browser⇔Pane を解放前に切る。Shutdown と同じ手順。
    pane.ClearBrowserRef

    ' --- 2. Controller を明示クローズ ---
    '   Ctrl_Close は View3/View2/View/Ctrl を順に Release し、登録ハンドラも
    '   清算する (二重 Close は S_OK で無害)。
    Dim hr As Long
    hr = pane.Ctrl_Close()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.CloseTab: Ctrl_Close 失敗 hr=&H" & Hex(hr) & " (続行)"
    End If

    ' --- 3. Collection から除去 (参照カウント 0 で Pane.Class_Terminate) ---
    Set pane = Nothing
    m_tabs.Remove index
    Debug.Print "Wv2Browser.CloseTab: タブ" & index & " を閉じた (残りタブ数=" & m_tabs.Count & ")"

    ' --- 4. アクティブ補正 + 可視状態の修復 ---
    Dim newActive As Long
    If m_tabs.Count = 0 Then
        ' タブ 0 個: アクティブなし (タブ 0 許容)
        m_activeIndex = 0
        Debug.Print "Wv2Browser.CloseTab: 全タブが閉じられた (activeIndex=0)"
    ElseIf index < m_activeIndex Then
        ' 閉じたのがアクティブより前: index が 1 つ前にずれるので補正して
        ' 「同じタブ」を維持する。可視状態は変わらないが ActivateTab で正す。
        newActive = m_activeIndex - 1
        ActivateTab newActive
    ElseIf index = m_activeIndex Then
        ' 閉じたのがアクティブ: 右隣優先、なければ左隣 (末尾)。
        '   Remove 後、元の右隣は今や同じ index にいる。index が残りタブ数を
        '   超えたら (末尾を閉じた場合) は末尾をアクティブ化する。
        If index <= m_tabs.Count Then
            newActive = index
        Else
            newActive = m_tabs.Count
        End If
        ActivateTab newActive
    Else
        ' 閉じたのがアクティブより後: activeIndex は変わらない。
        '   可視状態も変わらないので ActivateTab は不要だが、念のため正しておく。
        ActivateTab m_activeIndex
    End If

    ' --- 5. タブ閉じを通知 (第9.11a、疎結合イベント) ---
    RaiseEvent TabClosed(index)
    CloseTab = True
End Function


' ============================================================
' MoveTab ? タブを fromIndex から toIndex へ並べ替える (第9.15 段階で追加)
'
'   呼び出し元: Wv2TabBar.OnPaneWebMessage の {"cmd":"reorder","from":N,"to":M}
'   引数はいずれも 1-based。from == to や範囲外は何もせず False を返す。
'
'   ★手順★
'     1. アクティブ Pane を参照で掴む (index は並べ替えで動くため)。
'     2. 移動対象 Pane を Remove して、目的位置に Add (before/after) で挿し直す。
'        ・to == 1            : 先頭へ (Add pane, , 1)
'        ・それ以外            : (to-1) 番目の後ろへ (Add pane, , , to-1)
'          ※ Remove 後の Collection 上での基準 index に注意して計算する。
'     3. 掴んでおいたアクティブ Pane を線形走査で逆引きし m_activeIndex を更新。
'     4. ActivateTab で可視状態を正す → RaiseEvent ActiveChanged → 全体同期
'        (設計原則 60、案A-2b)。新イベントは撃たない。
'
'   ★完了待ちなし★ 配列操作 (Remove/Add) と PutIsVisible のみで Controller
'     生成のような待ちを含まないため、WebMessage ハンドラ内から同期で呼んで安全
'     (activate/close と同じ。newtab のような SafeTimer 迂回は不要)。
' ============================================================
Public Function MoveTab(ByVal fromIndex As Long, ByVal toIndex As Long) As Boolean
    Dim n As Long
    n = m_tabs.Count

    ' --- 範囲チェック / 無変化チェック ---
    If fromIndex < 1 Or fromIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: from 範囲外 (from=" & fromIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If toIndex < 1 Or toIndex > n Then
        Debug.Print "Wv2Browser.MoveTab: to 範囲外 (to=" & toIndex & ", tabCount=" & n & ")"
        Exit Function
    End If
    If fromIndex = toIndex Then
        ' 位置が変わらないので何もしない (成功扱い)。
        MoveTab = True
        Exit Function
    End If

    ' --- 1. アクティブ Pane を参照で掴む (index は並べ替えで動くため) ---
    Dim activePaneRef As Wv2Pane
    If m_activeIndex >= 1 And m_activeIndex <= n Then
        Set activePaneRef = m_tabs.Item(m_activeIndex)
    End If

    ' --- 2. 対象 Pane を取り出して Remove ---
    Dim moving As Wv2Pane
    Set moving = m_tabs.Item(fromIndex)
    m_tabs.Remove fromIndex

    ' Remove 後の Collection は要素数 n-1。ここに moving を toIndex の位置へ挿す。
    '   Collection.Add の before/after は「Remove 後の Collection 上での index」を指す。
    '   目標は「並べ替え後に moving が toIndex 番目にいる」こと。
    If toIndex = 1 Then
        ' 先頭へ: 現在の先頭 (before:=1) の前に入れる。
        m_tabs.Add moving, , 1
    Else
        ' toIndex 番目にしたい = (toIndex-1) 番目の要素の後ろに入れる。
        '   Remove 後の Collection でも 1..(toIndex-1) の要素は元の並びのまま残って
        '   いる (fromIndex を抜いた影響は toIndex 以降にしか及ばないか、または
        '   fromIndex < toIndex のとき (toIndex-1) 番目が元の toIndex 番目に一致する)。
        '   いずれの場合も after:=(toIndex-1) で「並べ替え後に toIndex 番目」になる。
        m_tabs.Add moving, , , toIndex - 1
    End If

    Debug.Print "Wv2Browser.MoveTab: タブを " & fromIndex & " → " & toIndex & " へ移動 (tabCount=" & m_tabs.Count & ")"

    ' --- 3. アクティブ Pane の新しい index を逆引き ---
    If Not activePaneRef Is Nothing Then
        Dim i As Long
        For i = 1 To m_tabs.Count
            If m_tabs.Item(i) Is activePaneRef Then
                m_activeIndex = i
                Exit For
            End If
        Next i
    End If

    ' --- 4. 可視状態を正す → ActiveChanged → 全体同期 (案A-2b) ---
    '   可視状態自体は変わらないが、ActivateTab を通すことで PushSyncToJs を
    '   促し、並べ替え結果を TabBar/NavBar に反映させる。
    If m_activeIndex >= 1 And m_activeIndex <= m_tabs.Count Then
        ActivateTab m_activeIndex
    End If

    MoveTab = True
End Function


' ============================================================
' Debug_TabOrderTitles ? 現在のタブ順を "|" 連結の識別タグで返す (第9.15、検証口)
'
'   MoveTab の並べ替え結果を WebView2 起動なしで検証するための薄い公開ラッパ。
'   通常タブはタイトル (未設定なら "(t<index>)") を返すが、検証時は
'   Debug_SeedDummyTabs が各ダミー Pane の ObjPtr → 連番タグ ("A","B",...) の
'   対応を m_dbgTagMap に記録しておくので、並べ替えで index が動いても
'   「そのタブが元々何番だったか」を安定して追跡できる。
'   Debug_NormalizeUrl と同じ「生きた検証資産」の位置づけ。
' ============================================================
Public Function Debug_TabOrderTitles() As String
    Dim i As Long
    Dim parts As String
    Dim tag As String
    For i = 1 To m_tabs.Count
        tag = ""
        ' 検証用タグマップがあれば ObjPtr で引く (並べ替えに強い安定識別)
        If Not m_dbgTagMap Is Nothing Then
            Dim key As String
            key = CStr(ObjPtr(m_tabs.Item(i)))
            On Error Resume Next
            tag = m_dbgTagMap.Item(key)
            On Error GoTo 0
        End If
        ' タグが無ければ本番タブとみなしタイトル (未設定はプレースホルダ)
        If Len(tag) = 0 Then
            On Error Resume Next
            tag = m_tabs.Item(i).DocumentTitle
            On Error GoTo 0
            If Len(tag) = 0 Then tag = "(t" & i & ")"
        End If
        If i > 1 Then parts = parts & "|"
        parts = parts & tag
    Next i
    Debug_TabOrderTitles = parts
End Function


' ============================================================
' Debug_SeedDummyTabs ? 検証用にタグ付きダミータブを n 個仕込む (第9.15、検証口)
'
'   Test_9_15_Reorder が WebView2 を起動せずに MoveTab の並べ替えロジックだけを
'   検証するための代替口。空の Wv2Pane を n 個 m_tabs に積み、各 Pane の ObjPtr に
'   連番タグ ("A","B","C",...) を対応させた m_dbgTagMap を作る。activeIndex は 1。
'   本番 MoveTab をそのまま通して検証するので二重管理にならない。
'   可視制御 (ActivateTab の PutIsVisible) はダミー Pane の View が Nothing のため
'   内部で失敗ログを出すが無害 (順序検証には影響しない)。
' ============================================================
Public Sub Debug_SeedDummyTabs(ByVal n As Long)
    Set m_tabs = New Collection
    Set m_dbgTagMap = New Collection
    Dim i As Long
    Dim p As Wv2Pane
    Dim tag As String
    For i = 1 To n
        Set p = New Wv2Pane
        m_tabs.Add p
        tag = Chr$(64 + i)   ' 1→"A", 2→"B", ...
        m_dbgTagMap.Add tag, CStr(ObjPtr(p))
    Next i
    If n >= 1 Then
        m_activeIndex = 1
    Else
        m_activeIndex = 0
    End If
End Sub


' ============================================================
' Debug_ActiveIndex ? 現在の m_activeIndex を返す (第9.15、検証口)
'   MoveTab 後のアクティブ追従を検証するために使う。
' ============================================================
Public Function Debug_ActiveIndex() As Long
    Debug_ActiveIndex = m_activeIndex
End Function


' ============================================================
' NotifyTitleChanged ? Pane からタイトル変更を受け取り、index を逆引いて
'   RaiseEvent TabTitleChanged を発火する (第9.11a 段階で追加)
'
'   呼び出し元: Wv2Pane.View_OnDocumentTitleChanged (後方参照 m_browser 経由)
'   引数: senderPane = タイトルが変わった Pane (自分自身 Me)、title = 新タイトル
'
'   ★逆引き★ Collection はオブジェクト参照の一致で index を探す。タブ数は
'   現実的に少ないので線形探索で十分。見つからなければ (既に閉じられた
'   タブの遅延イベント等) 何もしない。
' ============================================================
Public Sub NotifyTitleChanged(ByVal senderPane As Wv2Pane, ByVal title As String)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabTitleChanged(i, title)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyTitleChanged: senderPane が m_tabs に見当たらない (無視)"
End Sub


' ============================================================
' NotifyUrlChanged (第9.12a、Pane からの URL 状態通知を受けて index 逆引き)
'
'   Pane の NavigationCompleted / HistoryChanged から呼ばれる。オブジェクト参照の
'   一致で index を逆引きし、RaiseEvent TabUrlChanged する。
'   (9.11a の NotifyTitleChanged と全く同じ構造)
'
'   ★タブバー用ペインからは呼ばれない★ タブバー/ナビバーのペインは m_tabs に
'   入っていないうえ、そもそも m_browser 後方参照を持たない (SetBrowser されない)
'   ので、ここには到達しない。
' ============================================================
Public Sub NotifyUrlChanged(ByVal senderPane As Wv2Pane, ByVal url As String, _
                            ByVal canGoBack As Boolean, ByVal canGoForward As Boolean)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is senderPane Then
            RaiseEvent TabUrlChanged(i, url, canGoBack, canGoForward)
            Exit Sub
        End If
    Next i
    Debug.Print "Wv2Browser.NotifyUrlChanged: senderPane が m_tabs に見当たらない (無視)"
End Sub


' ============================================================
' アクティブタブへの操作委譲 (第9.12a、ナビバー用の公開 API)
'
'   ナビバーは「アクティブタブ」に操作を送りたいが、Pane を直接触らせると疎結合が
'   壊れる (ナビバーが Browser の内部構造を知ることになる)。Browser が代理する
'   薄いメソッドを用意し、ナビバーは Browser とだけ話すようにする。
'
'   ★設計原則 63★ いずれも「完了待ちを含まない」操作 (Navigate/GoBack/GoForward/
'     Reload/Stop はすべて撃ちっぱなし) なので、WebMessage のイベントハンドラ内から
'     同期で呼んで安全。9.11d の newtab (Controller 生成待ち) のような遅延実行は不要。
'
'   タブ 0 (アクティブなし) のときは何もせず False / 空文字を返す。
' ============================================================
Public Function NavigateActive(ByVal url As String) As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.NavigateActive: アクティブタブなし (無視)"
        Exit Function
    End If

    Dim target As String
    target = NormalizeUrl(url)
    If Len(target) = 0 Then
        Debug.Print "Wv2Browser.NavigateActive: URL が空 (無視)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_Navigate(target)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.NavigateActive: Navigate 失敗 hr=&H" & Hex(hr) & " (" & target & ")"
        Exit Function
    End If

    Debug.Print "Wv2Browser.NavigateActive: Navigate OK (" & target & ")"
    NavigateActive = True
End Function

Public Function GoBackActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoBack() Then
        Debug.Print "Wv2Browser.GoBackActive: 戻れない (履歴なし)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoBack()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoBackActive: GoBack 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoBackActive: GoBack OK"
    GoBackActive = True
End Function

Public Function GoForwardActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    If Not pane.View_GetCanGoForward() Then
        Debug.Print "Wv2Browser.GoForwardActive: 進めない (履歴なし)"
        Exit Function
    End If

    Dim hr As Long
    hr = pane.View_GoForward()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.GoForwardActive: GoForward 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.GoForwardActive: GoForward OK"
    GoForwardActive = True
End Function

Public Function ReloadActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Reload()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.ReloadActive: Reload 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.ReloadActive: Reload OK"
    ReloadActive = True
End Function

Public Function StopActive() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function

    Dim hr As Long
    hr = pane.View_Stop()
    If hr <> 0 Then
        Debug.Print "Wv2Browser.StopActive: Stop 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    Debug.Print "Wv2Browser.StopActive: Stop OK"
    StopActive = True
End Function


' --- アクティブタブの現在状態 (ナビバーの初期表示用) ---
'   イベント (TabUrlChanged) が来る前や、タブ切替直後にナビバーが
'   「今の状態」を問い合わせるための読み取り口。
Public Function GetActiveUrl() As String
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveUrl = pane.View_GetSource()
End Function

Public Function GetActiveCanGoBack() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoBack = pane.View_GetCanGoBack()
End Function

Public Function GetActiveCanGoForward() As Boolean
    Dim pane As Wv2Pane
    Set pane = ActivePane
    If pane Is Nothing Then Exit Function
    If Not pane.IsReady Then Exit Function
    GetActiveCanGoForward = pane.View_GetCanGoForward()
End Function


' ============================================================
' NormalizeUrl (第9.12a、アドレスバー入力のスキーム補完、案3-A)
'
'   ユーザーがアドレスバーに打った文字列を Navigate 可能な URL に整える。
'     "example.com"        → "https://example.com"
'     "https://example.com" → そのまま
'     "about:blank"        → そのまま (スキーム有りとみなす)
'
'   ★検索語フォールバックはスコープ外 (案3-A)★ 「検索したい言葉」のような
'     入力を検索エンジンに投げる機能は URL エンコードが必要になるため次段階に回す。
'     現状はスキームを補って Navigate するだけで、解決できなければブラウザの
'     エラーページが出る。
' ============================================================
Private Function NormalizeUrl(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function

    ' 既にスキームがあるか ("://" を含む、または "about:" 等のスキーム形式)
    If InStr(t, "://") > 0 Then
        NormalizeUrl = t
        Exit Function
    End If

    ' about: / data: / file: 等のスキーム付き (コロンが空白より前にある)
    Dim colonPos As Long
    colonPos = InStr(t, ":")
    If colonPos > 1 Then
        Dim head As String
        head = Left$(t, colonPos - 1)
        ' スキーム名は英字のみとみなす ("example.com:8080" を誤判定しないように
        ' 数字を含むものは除外)
        If Wv2Url.IsAlphaOnly(head) Then
            NormalizeUrl = t
            Exit Function
        End If
    End If

    ' ここまでで「スキーム有り」は除外済み。残りを URL / 検索語 に振り分ける (案A、9.14)。
    '   ・ドットを含み かつ 空白を含まない → ホスト名とみなして https:// を前置 (URL)
    '   ・それ以外 (空白を含む / ドット無し) → 検索語とみなして Google 検索へ
    ' Chrome/Edge のアドレスバーとほぼ同じ判定。"localhost" 単体は検索に落ちるが実用上稀。
    If InStr(t, ".") > 0 And InStr(t, " ") = 0 Then
        NormalizeUrl = "https://" & t
    Else
        NormalizeUrl = BuildSearchUrl(t)
    End If
End Function


' ============================================================
' Debug_NormalizeUrl (第9.14、ロジック検証口)
'
'   NormalizeUrl は Private なので、外から結果だけを確かめるための薄い公開ラッパ。
'   NavigateActive が内部で使うのと同一の NormalizeUrl をそのまま通すので、
'   検証ロジックが本番ロジックと二重管理にならない。
'   (Wv2Thunks.bas の Test_9_14_SearchFallback から呼ばれる。9.13 で残した
'    低レベルテストと同じ「生きた検証資産」の位置づけ。)
' ============================================================
Public Function Debug_NormalizeUrl(ByVal s As String) As String
    Debug_NormalizeUrl = NormalizeUrl(s)
End Function


' ============================================================
' BuildSearchUrl (第9.14 → 第9.16、検索語フォールバック、案A/案E-1/案P-1)
'
'   検索語を検索エンジンの URL に変換する。
'     "日本 天気"  → m_searchTemplate & "%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97"
'
'   ★第9.16: エンジンは m_searchTemplate で差し替え可能★ (案1-A)
'     9.14 では "https://www.google.com/search?q=" をハードコードしていたが、
'     モジュール変数 m_searchTemplate を前置テンプレートとして使うようにした。
'     既定は Google (Class_Initialize でセット)。Property SearchEngine で変更する。
' ============================================================
Private Function BuildSearchUrl(ByVal terms As String) As String
    BuildSearchUrl = m_searchTemplate & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' SearchEngine (第9.16、検索エンジンの検索 URL テンプレート、案1-A + 案2-A)
'
'   検索語フォールバック (NormalizeUrl の第3分岐) で使う検索 URL の前置テンプレート
'   を読み書きする。末尾にエンコード済みの検索語が連結される。
'     例: Get → "https://www.google.com/search?q="
'         Let "https://www.bing.com/search?q=" で Bing に切り替え
'
'   ★空ガード★ 空文字 (または空白のみ) を渡されたら既定の Google テンプレートに
'     フォールバックする。URL としての妥当性検証はしない (テンプレートは開発者が
'     入れる信頼できる値であり、過剰な検証はしない方針)。
' ============================================================
' ============================================================
' ============================================================
' NewWindowMode (第9.27、NewWindowRequested の捌き方を選ぶ 3 値スイッチ)
'
'   0 = 現行動作 (LEGACY)   : args.get_Uri で URL だけ取り、put_Handled(TRUE) して
'                             自前の新タブで Navigate し直す。Navigate は★常に GET★
'                             なので、POST リンクは 405 になる (これが 9.26 以前の姿)。
'   1 = ポップアップ委譲     : put_Handled を立てず素通り。ランタイムが自前ウィンドウを
'                             開いて元のリクエストを継続する (仕様事実 23 で実証済み)。
'                             POST は保たれるがタブバーに乗らず、閉じるのも手動。
'   2 = プリウォーム委譲     : ★9.27 の本命★ 温めてある予備タブの CoreWebView2 を
'                             args.put_NewWindow (vtable 4) に渡す。ランタイムが
'                             その WebView に元のリクエストを載せるので POST が保たれ、
'                             かつ自前のタブに収まる。
'
'   ★範囲外の値は 0 に丸める★ 実機からイミディエイトで叩くスイッチなので、
'     打ち間違いで未定義の挙動に落ちないよう安全側に倒す。
'   ★全タブ共通★ Wv2Pane 側は毎回 m_browser.NewWindowMode を読むので、切り替えは
'     起動中でも即時に効く (タブを開き直す必要はない)。
'   ★既定値 = 2 (第9.28 で 0 から変更)★ Class_Initialize を参照。起動直後の
'     PREWARM_DELAY_MS の間だけ予備が未生成で、その間の新規ウィンドウ要求は
'     現行動作にフォールバックする (★印つきのログが出る)。
' ============================================================
Public Property Get NewWindowMode() As Long
    NewWindowMode = m_newWindowMode
End Property

Public Property Let NewWindowMode(ByVal value As Long)
    Dim v As Long
    v = value
    If v < 0 Or v > 2 Then
        Debug.Print "Wv2Browser.NewWindowMode: 範囲外の値 " & value & " を 0 に丸めた"
        v = 0
    End If
    m_newWindowMode = v

    Select Case v
        Case 0
            Debug.Print "Wv2Browser.NewWindowMode = 0 (現行動作: put_Handled + 自前の新タブ。POST は GET に化ける)"
        Case 1
            Debug.Print "Wv2Browser.NewWindowMode = 1 (ポップアップ委譲: ランタイム任せ。POST は保たれるがタブに乗らない)"
        Case 2
            Debug.Print "Wv2Browser.NewWindowMode = 2 (プリウォーム委譲: put_NewWindow で予備タブに載せる)"
            Debug.Print "    予備タブの状態: " & Debug_PrewarmStatus()
            If Not PrewarmAvailable Then
                Debug.Print "    ★予備がまだ無いので生成を予約します (温まるまでは現行動作にフォールバック)★"
                RequestPrewarmAsync
            End If
    End Select
End Property


' ============================================================
' SetHandledWithNewWindow (第9.27、論点5 の保険スイッチ)
'
'   put_NewWindow に予備タブの CoreWebView2 をセットしたとき、続けて
'   put_Handled(TRUE) も立てるかどうか。公式サンプルは両方立てる形なので既定 True。
'
'   ★何のための保険か★ 「NewWindow をセットしたら Handled は不要 / むしろ立てると
'     遷移が破棄される」という可能性が理屈の上では残る。万一 予備タブが白いままだったり
'     二重にウィンドウが開いたりしたら、これを False にして再試行すれば
'     ★コードを直さずに★ どちらが正しいか切り分けられる。
' ============================================================
Public Property Get SetHandledWithNewWindow() As Boolean
    SetHandledWithNewWindow = m_setHandledWithNewWindow
End Property

Public Property Let SetHandledWithNewWindow(ByVal value As Boolean)
    m_setHandledWithNewWindow = value
    If value Then
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = True (put_NewWindow の後に put_Handled も立てる)"
    Else
        Debug.Print "Wv2Browser.SetHandledWithNewWindow = False (put_NewWindow だけ撃つ)"
    End If
End Property


' ============================================================
' PopupDelegateMode (第9.26 → 第9.27 で互換ラッパ化)
'
'   9.26 で入れた Boolean のトグル。9.27 で NewWindowMode (3 値) に統合したため、
'   ★内部で NewWindowMode に委譲するだけの薄いラッパ★ になった。
'   これにより Test_9_26_Popup_On / _Off が 9.27 でもそのまま動く。
'
'   Get : NewWindowMode = 1 (ポップアップ委譲) のとき True
'   Let : True → NewWindowMode = 1 / False → NewWindowMode = 0
'         ※ False を入れると ★プリウォーム委譲 (2) も解除される★ ことに注意。
'            9.27 の本命に戻すときは NewWindowMode = 2 を明示すること。
' ============================================================
Public Property Get PopupDelegateMode() As Boolean
    PopupDelegateMode = (m_newWindowMode = 1)
End Property

Public Property Let PopupDelegateMode(ByVal value As Boolean)
    If value Then
        NewWindowMode = 1
    Else
        NewWindowMode = 0
    End If
End Property


Public Property Get SearchEngine() As String
    SearchEngine = m_searchTemplate
End Property

Public Property Let SearchEngine(ByVal template As String)
    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then
        m_searchTemplate = DEFAULT_SEARCH_TEMPLATE
        Debug.Print "Wv2Browser.SearchEngine: 空テンプレートのため既定 (Google) にフォールバック"
    Else
        m_searchTemplate = t
        Wv2Log.LogD "Wv2Browser.SearchEngine: 検索テンプレートを設定 = " & t
    End If

    ' --- K-2: 設定画面を表示中の全タブへ反映を予約する ---
    RequestSettingsSyncAsync
End Property


' ============================================================
' Debug_SearchEngine (第9.16、ロジック検証口)
'
'   現在の m_searchTemplate を返す薄い公開ラッパ。BuildSearchUrl は Private なので、
'   検証側 (Test_9_16_SearchEngine) がエンジン設定の反映を確かめるために使う。
'   本番 BuildSearchUrl が使うのと同一の m_searchTemplate をそのまま返すので、
'   検証ロジックが二重管理にならない (Debug_NormalizeUrl と同じ位置づけ)。
' ============================================================
Public Function Debug_SearchEngine() As String
    Debug_SearchEngine = m_searchTemplate
End Function



' ============================================================
' ResizeAll ? 全タブの Bounds を指定矩形に再発行する
'   第9.10c-1 段階で追加。UserForm_Resize から 1 本で呼ぶ (論点⑧ a)。
'
'   不可視タブにも Bounds を再発行しておく理由: 後でアクティブ化された時に
'   サイズがずれないようにするため。可視/不可視に関わらず全タブ更新する。
' ============================================================
Public Sub ResizeAll(ByRef rc As RECT)
    ' --- 第9.11c: 上位が適用したコンテンツ矩形をキャッシュ (案A) ---
    '   以降の AddTab はこの矩形を初期 Bounds に使い、タブバーと重ならない。
    m_lastContentRect = rc
    m_hasContentRect = True

    Dim i As Long
    Dim pane As Wv2Pane
    Dim hr As Long
    For i = 1 To m_tabs.Count
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
    If index < 1 Or index > m_tabs.Count Then
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
' RequestNewTabAsync (第9.10e、案β: NewWindowRequested の遅延新タブ生成)
'
'   Wv2Pane.View_OnNewWindowRequested から呼ばれる。ここで直接 AddTabWithUrl すると
'   デッドロックする (仕様事実 14) ため、URL を保留キューに積んでワンショット
'   SafeTimer を起動し、ハンドラ完了後に m_newTabTimer_TimerCall で実処理する。
'
'   ★要求ごとに New SafeTimer する理由★ SafeTimer.StartTimer は
'   「uElapse = m_uElapse なら無視」し、停止 (StartTimer 0) 後も m_uElapse が残る。
'   同じインスタンスを使い回すと同間隔の再起動が無視されてしまうので、要求ごとに
'   新インスタンス (m_uElapse=0) を作って確実に起動させる。連続要求時は前の
'   タイマーが破棄されても保留キューには全 URL が残るので取りこぼさない。
' ============================================================
Public Sub RequestNewTabAsync(ByVal url As String)
    If m_pendingUrls Is Nothing Then Set m_pendingUrls = New Collection
    m_pendingUrls.Add url

    Set m_newTabTimer = New SafeTimer      ' 要求ごとに新インスタンス (確実に起動)
    m_newTabTimer.StartTimer 1             ' 1ms 後 (ハンドラ完了後) に TimerCall

    Debug.Print "Wv2Browser.RequestNewTabAsync: 保留 URL 追加 (" & url & _
                ")。キュー数=" & m_pendingUrls.Count & "、ワンショット Timer 起動"
End Sub


' ============================================================
' RequestSettingsTabAsync (第9.18、設定タブの遅延生成入口)
'
'   TabBar の [歯車] ボタン → OnPaneWebMessage "settings" から呼ばれる。設定タブを
'   開く処理は内部で AddTab → WaitPaneReady を通るため、WebMessage ハンドラ内
'   から同期で呼ぶと 仕様事実 14 のデッドロックに嵌まる。そこで newtab と同型に、
'   保留フラグを立ててワンショット SafeTimer を起動し即座に返す。ハンドラ完了後に
'   m_settingsTimer_TimerCall で実際の設定タブ生成が走る。
'
'   ★要求ごとに New SafeTimer する理由は RequestNewTabAsync と同じ★
'     StartTimer は「同間隔の再起動を無視」し停止後も m_uElapse が残るため、
'     使い回すと再起動が無視される。要求ごとに新インスタンスで確実に起動する。
'   ★多重要求は無害★ 既に保留中でも新タイマーを張り直すだけ。TimerCall 側で
'     フラグを見て 1 回だけ開く (連打で設定タブが乱立しない最低限のガード)。
' ============================================================
Public Sub RequestSettingsTabAsync()
    m_settingsPending = True

    Set m_settingsTimer = New SafeTimer    ' 要求ごとに新インスタンス (確実に起動)
    m_settingsTimer.StartTimer 1           ' 1ms 後 (ハンドラ完了後) に TimerCall

    Debug.Print "Wv2Browser.RequestSettingsTabAsync: 設定タブ生成を予約 (ワンショット Timer 起動)"
End Sub


' ============================================================
' m_settingsTimer_TimerCall (第9.18、設定タブ生成のワンショット発火受け口)
'
'   WebMessage ハンドラは既に return 済みのタイミングで呼ばれる。ここで
'   OpenSettingsTab を呼べば AddTab の WaitPaneReady が正常に完了する
'   (newtab の TimerCall と同じ構造)。
' ============================================================
Private Sub m_settingsTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' ワンショット化: 繰り返し発火を止める
    If Not m_settingsTimer Is Nothing Then m_settingsTimer.StartTimer 0

    If Not m_settingsPending Then Exit Sub
    m_settingsPending = False

    Debug.Print "Wv2Browser.TimerCall(settings): 設定タブを開く"
    OpenSettingsTab
End Sub


' ============================================================
' IsSettingsView (K-2)
'
'   その Pane が ★今★ 設定画面を表示しているかを、ページ自身に尋ねる。
'   設定画面の HTML には window.__wv2Settings = true が埋めてある。
'
'   ★戻り値の読み方は Elem_CheckStale と同じ流儀★ オブジェクトを返させて
'     InStr で読む。EvalSync が返すのは値の JSON テキストなので、
'     真偽値をそのまま比較するより取り違えが起きにくい。
'
'   ★判定できないときは False (安全側)★ 読み込み中・Close 済み・in-callback は
'     すべて「設定画面ではない」と扱う。取りこぼしても新しい設定タブが開くだけで、
'     誤って別ページのタブを掴むより害が小さい (設計原則93 の LastEvalOk で区別)。
' ============================================================
Private Function IsSettingsView(ByVal pane As Wv2Pane) As Boolean
    If pane Is Nothing Then Exit Function

    Dim js As String
    Dim res As String
    js = "(function(){try{return {s:(window.__wv2Settings===true)?'yes':'no'};}" & _
         "catch(e){return {s:'no'};}})()"

    res = pane.EvalSync(js, 2)

    If Not pane.LastEvalOk Then
        Wv2Log.LogD "Wv2Browser.IsSettingsView: 判定できないので設定画面ではないと扱う (" & _
                    pane.LastEvalError & ")"
        Exit Function
    End If

    IsSettingsView = (InStr(1, res, """s"":""yes""") > 0)
End Function


' ============================================================
' RequestSettingsSyncAsync / m_syncTimer_TimerCall (K-2)
'
'   ★なぜ一拍置くのか★ SetEngine / SetCustomEngine は hostObjects 経由
'   (JS→VBA) で呼ばれる。仕様事実27 により、その間 ★呼び出し元ペインの JS は
'   完全にブロックされている★。ここから EvalSync を撃つと完了通知が返らず
'   タイムアウトまで固まる。SafeTimer でハンドラの外へ出してから撃つ
'   (RequestSettingsTabAsync / RequestNewTabAsync と同じ型)。
' ============================================================
Public Sub RequestSettingsSyncAsync()
    m_syncPending = True

    Set m_syncTimer = New SafeTimer
    m_syncTimer.StartTimer 1
End Sub


Private Sub m_syncTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' ワンショット化: 繰り返し発火を止める
    If Not m_syncTimer Is Nothing Then m_syncTimer.StartTimer 0

    If Not m_syncPending Then Exit Sub
    m_syncPending = False

    SyncSettingsViewsNow
End Sub


' ============================================================
' SyncSettingsViewsNow (K-2、設定ビューの全体同期)
'
'   設定画面を表示中の全タブへ、現在のエンジン名を押し込む。
'   ★設計原則60 (VBA→JS は全体同期)★ 差分を送らず、現在値をそのまま渡す。
'
'   これが要るのは、判定を正確にしても設定ビューが 2 枚並びうるため。
'   戻る操作は OpenSettingsTab を通らないので、開くときの判定では塞げない。
'   並んでも食い違わないようにするのがこの同期の役目。
'
'   ★カスタム入力欄の中身は上書きしない★ ユーザーが入力途中かもしれないため。
'     選択状態とプレビューだけを合わせる。
' ============================================================
Private Sub SyncSettingsViewsNow()
    If m_tabs Is Nothing Then Exit Sub
    If m_tabs.Count = 0 Then Exit Sub

    ' ブリッジが無いなら設定画面は一度も開かれていない
    If m_settingsBridge Is Nothing Then Exit Sub

    Dim syncName As String
    syncName = m_settingsBridge.GetEngine()
    If Len(syncName) = 0 Then Exit Sub

    Dim syncIdx As Long
    Dim syncPane As Wv2Pane
    Dim syncJs As String
    Dim res As String
    Dim hit As Long

    For syncIdx = 1 To m_tabs.Count
        Set syncPane = m_tabs.Item(syncIdx)
        If IsSettingsView(syncPane) Then
            syncJs = "(function(){try{return {r:applyFromHost(" & _
                     syncPane.JsQuote(syncName) & ")};}catch(e){return {r:'err'};}})()"
            res = syncPane.EvalSync(syncJs, 2)
            If syncPane.LastEvalOk Then
                If InStr(1, res, """r"":""ok""") > 0 Then
                    hit = hit + 1
                Else
                    Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                                " の適用が失敗 (" & res & ")"
                End If
            Else
                Wv2Log.LogW "Wv2Browser.SyncSettingsViewsNow: index " & syncIdx & _
                            " へ届かなかった (" & syncPane.LastEvalError & ")"
            End If
        End If
    Next syncIdx

    If hit > 0 Then
        Wv2Log.LogD "Wv2Browser.SyncSettingsViewsNow: 設定ビュー " & hit & _
                    " 枚へ " & syncName & " を反映した"
    End If
End Sub

' ============================================================
' OpenSettingsTab (第9.18、設定タブを開いてブリッジを attach する)
'
'   手順 (順序A: attach してから Navigate):
'     1. AddTab で空タブ (Pane + Controller) を生成する。
'     2. Wv2SettingsBridge を用意し、自分 (Browser) を BindBrowser で結び付ける。
'     3. その Pane に View_AddHostObjectToScript("Settings", bridge) で attach。
'        ★NavigateToString の前に attach する★ 読み込まれた JS が最初から
'        hostObjects.Settings を使える、という仮説 (順序A)。
'     4. BuildSettingsHtml で設定画面 HTML を組み、View_NavigateToString で表示。
'     5. 生成したタブをアクティブにする (ActivateTab)。
'
'   ★ブリッジの保持★ attach したブリッジは設定タブが生きている間 WebView2 から
'     参照される。Browser も m_settingsBridge に 1 本持ち、確実に生存させる
'     (ローカル変数だけだと本 Sub を抜けた瞬間に解放されうる)。設定タブを複数
'     開いた場合は最後のブリッジで上書きされるが、各タブの WebView2 が自分の
'     attach 時点の参照を保持するため、既存タブの動作には影響しない。
' ============================================================
Public Sub OpenSettingsTab()
    ' --- 第9.19: 既に設定タブがあれば、それをアクティブにするだけ ---
    '   m_tabs を頭から走査して IsSettings = True の Pane を探す。見つかれば
    '   その index をアクティブにして Exit Sub (新規生成しない)。閉じられた
    '   設定タブは m_tabs から外れているので、この走査で自然に見つからない。
    ' --- K-2: フラグではなく ★ページ自身に聞く★ ---
    '   IsSettings は遷移しても False に戻らず、元・設定タブを掴んでいた。
    '   EvalSync で window.__wv2Settings を尋ねれば、遷移経路によらず
    '   「今そのタブが設定画面か」が分かる (戻るで復帰したタブも拾える)。
    Dim j As Long
    For j = 1 To m_tabs.Count
        If IsSettingsView(m_tabs.Item(j)) Then
            Wv2Log.LogD "Wv2Browser.OpenSettingsTab: 設定画面を表示中のタブ(index " & j & _
                        ")を見つけたのでアクティブ化"
            ActivateTab j
            Exit Sub
        End If
    Next j

    Dim pane As Wv2Pane
    Set pane = AddTab()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddTab 失敗"
        Exit Sub
    End If

    ' --- 第9.19: 生成した Pane を設定タブとして印を付ける ---
    '   以後の [歯車] クリックはこのフラグを見て重複を防ぐ。
    pane.IsSettings = True

    ' ブリッジを用意して Browser を結び付ける
    Set m_settingsBridge = New Wv2SettingsBridge
    m_settingsBridge.BindBrowser Me

    ' ★順序A: Navigate の前に attach★
    Dim hr As Long
    hr = pane.View_AddHostObjectToScript("Settings", m_settingsBridge)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: AddHostObjectToScript 失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Browser.OpenSettingsTab: Settings ブリッジを attach OK"
    End If

    ' 設定画面 HTML を表示
    hr = pane.View_NavigateToString(BuildSettingsHtml())
    If hr <> 0 Then
        Debug.Print "Wv2Browser.OpenSettingsTab: NavigateToString 失敗 hr=&H" & Hex(hr)
    End If

    ' 生成したタブをアクティブにする (新 index を逆引き)
    Dim i As Long
    For i = 1 To m_tabs.Count
        If m_tabs.Item(i) Is pane Then
            ActivateTab i
            Exit For
        End If
    Next i
End Sub


' ============================================================
' BuildSettingsHtml (第9.18、設定画面の HTML を組む)
'
'   検索エンジンを選ぶカード UI。hostObjects.sync.Settings 経由で VBA を直接
'   呼ぶ (postMessage コマンドは使わない):
'     ・起動時          : Settings.EngineList() でカード定義を取得して描画。
'                         Settings.GetEngine() で現在エンジンを取得し選択強調。
'     ・カードクリック  : Settings.SetEngine(id) で切替。戻り値で選択確定。
'     ・カード hover    : Settings.PreviewUrlFor(id, サンプル語) で生成 URL を
'                         ライブ表示 (同期呼び出しの妙味)。
'
'   ★日本語は実文字で書く (CRLF/UTF-8/BOMなしは出力時に Python で担保)★
'   ★hostObjects.sync は WebView2 のホストオブジェクト同期呼び出し★
'     戻り値がそのまま JS の戻り値になる (旧実装で確認済み)。
' ============================================================
Private Function BuildSettingsHtml() As String
    Dim s As String
    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>設定</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    s = s & "  html,body{height:100%;font-family:'Segoe UI','Meiryo',sans-serif;" & _
            "color:#e8eaed;background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);}" & vbLf
    s = s & "  .wrap{max-width:840px;margin:0 auto;padding:48px 28px 64px;}" & vbLf
    s = s & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    s = s & "  h1{font-size:26px;font-weight:650;margin:6px 0 4px;}" & vbLf
    s = s & "  .lead{color:#9aa7bd;font-size:13.5px;margin-bottom:28px;}" & vbLf
    s = s & "  .grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:16px;}" & vbLf
    s = s & "  .card{position:relative;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:18px 18px 16px;cursor:pointer;background:rgba(255,255,255,.03);" & _
            "transition:transform .12s ease,border-color .12s ease,background .12s ease;overflow:hidden;}" & vbLf
    s = s & "  .card:hover{transform:translateY(-2px);border-color:rgba(120,160,255,.5);background:rgba(255,255,255,.06);}" & vbLf
    s = s & "  .card.sel{border-color:#6ea8fe;background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35),0 10px 30px -12px rgba(110,168,254,.5);}" & vbLf
    s = s & "  .card .dot{position:absolute;top:16px;right:16px;width:16px;height:16px;border-radius:50%;" & _
            "border:2px solid rgba(255,255,255,.3);}" & vbLf
    s = s & "  .card.sel .dot{border-color:#6ea8fe;background:#6ea8fe;box-shadow:0 0 10px #6ea8fe;}" & vbLf
    s = s & "  .card .name{font-size:17px;font-weight:620;margin-bottom:2px;}" & vbLf
    s = s & "  .card .host{font-size:11.5px;color:#8ea2c8;word-break:break-all;}" & vbLf
    s = s & "  .badge{display:inline-block;margin-top:10px;font-size:10.5px;letter-spacing:.08em;" & _
            "color:#6ea8fe;border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;opacity:0;" & _
            "transition:opacity .12s ease;}" & vbLf
    s = s & "  .card.sel .badge{opacity:1;}" & vbLf
    s = s & "  .preview{margin-top:26px;border-top:1px solid rgba(255,255,255,.08);padding-top:16px;}" & vbLf
    s = s & "  .preview .plabel{font-size:11px;letter-spacing:.14em;color:#8ea2c8;text-transform:uppercase;margin-bottom:6px;}" & vbLf
    s = s & "  .preview .purl{font-family:'Consolas','Courier New',monospace;font-size:12.5px;color:#cfe0ff;" & _
            "background:rgba(0,0,0,.25);border-radius:8px;padding:10px 12px;word-break:break-all;min-height:38px;}" & vbLf
    s = s & "  .sample{color:#7d8aa0;}" & vbLf
    ' --- 第9.29: カスタム入力欄 (プリセット外の任意テンプレート) ---
    s = s & "  .custom{margin-top:22px;border:1px solid rgba(255,255,255,.08);border-radius:14px;" & _
            "padding:16px 18px;background:rgba(255,255,255,.03);}" & vbLf
    s = s & "  .custom.sel{border-color:#6ea8fe;" & _
            "background:linear-gradient(180deg,rgba(110,168,254,.16),rgba(110,168,254,.05));" & _
            "box-shadow:0 0 0 1px rgba(110,168,254,.35);}" & vbLf
    s = s & "  .chead{display:flex;align-items:center;gap:10px;}" & vbLf
    s = s & "  .cname{font-size:15px;font-weight:620;}" & vbLf
    s = s & "  .cbadge{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;" & _
            "opacity:0;transition:opacity .12s ease;}" & vbLf
    s = s & "  .custom.sel .cbadge{opacity:1;}" & vbLf
    s = s & "  .chint{font-size:11.5px;color:#8ea2c8;margin:6px 0 12px;}" & vbLf
    s = s & "  .crow{display:flex;gap:10px;}" & vbLf
    s = s & "  .crow input{flex:1;min-width:0;font-family:'Consolas','Courier New',monospace;" & _
            "font-size:12.5px;color:#e8eaed;background:rgba(0,0,0,.25);" & _
            "border:1px solid rgba(255,255,255,.12);border-radius:8px;padding:9px 11px;outline:none;}" & vbLf
    s = s & "  .crow input:focus{border-color:#6ea8fe;}" & vbLf
    s = s & "  .crow button{font-size:13px;color:#0e121b;background:#6ea8fe;border:0;" & _
            "border-radius:8px;padding:9px 18px;cursor:pointer;font-weight:620;}" & vbLf
    s = s & "  .crow button:hover{background:#8fbcff;}" & vbLf
    s = s & "  .cmsg{margin-top:10px;font-family:'Consolas','Courier New',monospace;font-size:12px;" & _
            "color:#a9c6ff;word-break:break-all;min-height:16px;}" & vbLf
    s = s & "  .cmsg.err{color:#ff9d9d;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "  <div class=""eyebrow"">Search Engine</div>" & vbLf
    s = s & "  <h1>検索エンジンを選ぶ</h1>" & vbLf
    s = s & "  <div class=""lead"">アドレスバーに URL でない言葉を入力したとき、どの検索エンジンに送るかを選びます。" & _
            "カードにマウスを乗せると、生成される検索 URL をその場で確認できます。</div>" & vbLf
    s = s & "  <div class=""grid"" id=""grid""></div>" & vbLf
    ' --- 第9.29: カスタム入力欄。form は使わない (Enter で submit されるため) ---
    s = s & "  <div class=""custom"" id=""custom"">" & vbLf
    s = s & "    <div class=""chead""><span class=""cname"">カスタム (プリセット以外)</span>" & _
            "<span class=""cbadge"">使用中</span></div>" & vbLf
    s = s & "    <div class=""chint"">検索語を末尾に付け足す形の URL を入力します。" & _
            "例: https://example.com/search?q=</div>" & vbLf
    s = s & "    <div class=""crow"">" & vbLf
    s = s & "      <input id=""ctpl"" type=""text"" spellcheck=""false"" " & _
            "placeholder=""https://example.com/search?q="">" & vbLf
    s = s & "      <button id=""capply"">適用</button>" & vbLf
    s = s & "    </div>" & vbLf
    s = s & "    <div class=""cmsg"" id=""cmsg""></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "  <div class=""preview"">" & vbLf
    s = s & "    <div class=""plabel"">Preview</div>" & vbLf
    s = s & "    <div class=""purl"" id=""purl""><span class=""sample"">現在のエンジンでのプレビューを表示します</span></div>" & vbLf
    s = s & "  </div>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<script>" & vbLf
    ' --- hostObjects.sync.Settings への参照。無い環境でも落ちないようガード ---
    s = s & "  var S = (window.chrome && window.chrome.webview && window.chrome.webview.hostObjects" & _
            " && window.chrome.webview.hostObjects.sync) ? window.chrome.webview.hostObjects.sync.Settings : null;" & vbLf
    s = s & "  var SAMPLE = '日本 天気';" & vbLf
    s = s & "  var current = '';" & vbLf
    ' --- 第9.19c: いまカーソルが乗っているカードの id (未 hover は null) ---
    s = s & "  var hoverId = null;" & vbLf
    ' --- K-2: このページが設定画面であることの目印 ---
    '   OpenSettingsTab が EvalSync で尋ねる。★早い位置に置く★
    '   後続の JS が落ちても、設定画面だと名乗れるようにしておく
    '   (名乗れないと歯車で設定タブが二重に開く)。
    s = s & "  window.__wv2Settings = true;" & vbLf
    ' --- 現在エンジンとカード定義を VBA から同期取得 ---
    s = s & "  function loadState(){" & vbLf
    s = s & "    if(!S){ document.getElementById('grid').textContent='ホストオブジェクト未接続'; return null; }" & vbLf
    s = s & "    try{ current = S.GetEngine(); }catch(e){ current=''; }" & vbLf
    s = s & "    try{ return JSON.parse(S.EngineList()); }catch(e){ return []; }" & vbLf
    s = s & "  }" & vbLf
    ' --- ホスト名を URL テンプレートから抜き出す (見た目用) ---
    s = s & "  function hostOf(tpl){ try{ return (tpl.split('/')[2]||tpl); }catch(e){ return tpl; } }" & vbLf
    s = s & "  function markSel(){" & vbLf
    s = s & "    var cs=document.querySelectorAll('.card');" & vbLf
    s = s & "    for(var i=0;i<cs.length;i++){ cs[i].classList.toggle('sel', cs[i].getAttribute('data-id')===current); }" & vbLf
    ' --- 第9.29: カスタムセクションの選択強調も同じ 1 箇所から更新する ---
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function showPreview(id){" & vbLf
    s = s & "    var el=document.getElementById('purl'); if(!S){ return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlFor(id, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    el.innerHTML = '<span class=""sample"">' + SAMPLE + '</span> &nbsp;&rarr;&nbsp; ' + (u||'(取得失敗)');" & vbLf
    s = s & "  }" & vbLf
    ' --- 第9.19b: 現在選択中エンジンのプレビューを出す (未 hover 時/外れた時の既定) ---
    s = s & "  function showCurrentPreview(){" & vbLf
    s = s & "    if(!S){ return; }" & vbLf
    s = s & "    if(current){ showPreview(current); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function build(list){" & vbLf
    s = s & "    var g=document.getElementById('grid'); g.innerHTML='';" & vbLf
    s = s & "    for(var i=0;i<list.length;i++){ (function(item){" & vbLf
    s = s & "      var c=document.createElement('div'); c.className='card'; c.setAttribute('data-id',item.id);" & vbLf
    s = s & "      c.innerHTML='<span class=""dot""></span>' +" & vbLf
    s = s & "        '<div class=""name"">'+item.label+'</div>' +" & vbLf
    s = s & "        '<div class=""host"">'+hostOf(item.template)+'</div>' +" & vbLf
    s = s & "        '<div class=""badge"">選択中</div>';" & vbLf
    ' --- 第9.19c: hover 中は hoverId を立て、そのカードのプレビューを出す ---
    s = s & "      c.onmouseenter=function(){ hoverId=item.id; showPreview(item.id); };" & vbLf
    ' --- 第9.19c: 外れたら、他カードに乗り移っていない場合のみ選択中へ戻す ---
    s = s & "      c.onmouseleave=function(){ hoverId=null; setTimeout(function(){ if(hoverId===null){ showCurrentPreview(); } },0); };" & vbLf
    s = s & "      c.onclick=function(){" & vbLf
    s = s & "        if(!S) return;" & vbLf
    s = s & "        var res=''; try{ res=S.SetEngine(item.id); }catch(e){ res=''; }" & vbLf
    s = s & "        if(res){ current=res; } else { current=item.id; }" & vbLf
    s = s & "        markSel(); showPreview(item.id);" & vbLf
    s = s & "      };" & vbLf
    s = s & "      g.appendChild(c);" & vbLf
    s = s & "    })(list[i]); }" & vbLf
    s = s & "    markSel();" & vbLf
    ' --- 第9.19c: 開いた直後 (未 hover) も選択中エンジンのプレビューを表示 ---
    '   (カードから外れたときの戻しは各カードの mouseleave が担う。#grid 全体の
    '    mouseleave は使わない。カード間の隙間で発火せず直感に反したため。)
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "  }" & vbLf
    ' --- 第9.29: カスタム入力欄の JS (属性値・文字列はシングルクォート) ---
    s = s & "  function markCustom(){" & vbLf
    s = s & "    var box=document.getElementById('custom'); if(!box){ return; }" & vbLf
    s = s & "    box.classList.toggle('sel', current==='custom');" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function setMsg(text, isErr){" & vbLf
    s = s & "    var m=document.getElementById('cmsg'); if(!m){ return; }" & vbLf
    s = s & "    m.textContent = text || '';" & vbLf
    s = s & "    m.className = isErr ? 'cmsg err' : 'cmsg';" & vbLf
    s = s & "  }" & vbLf
    ' --- 適用前のライブプレビュー。エンコードは VBA (本番と同一) に任せる ---
    s = s & "  function customPreview(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var v=ct.value;" & vbLf
    s = s & "    if(!v){ setMsg('', false); return; }" & vbLf
    s = s & "    var u=''; try{ u=S.PreviewUrlForTemplate(v, SAMPLE); }catch(e){ u=''; }" & vbLf
    s = s & "    if(u){ setMsg(SAMPLE + '  →  ' + u, false); }" & vbLf
    s = s & "    else { setMsg('http:// または https:// で始まる URL を入力してください', true); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function applyCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct||!S){ return; }" & vbLf
    s = s & "    var res=''; try{ res=S.SetCustomEngine(ct.value); }catch(e){ res=''; }" & vbLf
    s = s & "    if(!res){ setMsg('適用できませんでした。http:// または https:// で始まる URL を" & _
            "入力してください', true); return; }" & vbLf
    s = s & "    current=res; markSel(); showCurrentPreview();" & vbLf
    s = s & "    if(res==='custom'){ setMsg('適用しました (カスタム)', false); }" & vbLf
    s = s & "    else { setMsg('入力がプリセット (' + res + ') と同じだったので、" & _
            "そちらを選択しました', false); }" & vbLf
    s = s & "  }" & vbLf
    s = s & "  function initCustom(){" & vbLf
    s = s & "    var ct=document.getElementById('ctpl'); if(!ct){ return; }" & vbLf
    s = s & "    if(S){ try{ ct.value = S.GetCustomTemplate(); }catch(e){ } }" & vbLf
    s = s & "    ct.oninput = function(){ customPreview(); };" & vbLf
    s = s & "    ct.onkeydown = function(e){ if(e.key==='Enter'){ applyCustom(); } };" & vbLf
    s = s & "    var btn=document.getElementById('capply');" & vbLf
    s = s & "    if(btn){ btn.onclick=function(){ applyCustom(); }; }" & vbLf
    s = s & "    markCustom();" & vbLf
    s = s & "  }" & vbLf
    ' --- K-2: ホストからの全体同期の受け口 ---
    '   別の設定タブでエンジンが変わったとき、VBA がこれを呼んで現在値を渡す。
    '   カスタム入力欄 (ctpl) は入力途中かもしれないので触らない。
    s = s & "  function applyFromHost(name){" & vbLf
    s = s & "    if(!name){ return 'ng'; }" & vbLf
    s = s & "    current = name;" & vbLf
    s = s & "    markSel();" & vbLf
    s = s & "    showCurrentPreview();" & vbLf
    s = s & "    return 'ok';" & vbLf
    s = s & "  }" & vbLf
    s = s & "  var list=loadState(); if(list){ build(list); }" & vbLf
    s = s & "  initCustom();" & vbLf
    s = s & "</script>" & vbLf
    s = s & "</body></html>"
    BuildSettingsHtml = s
End Function


' ============================================================
' ============================================================
' プリウォーム機構 (第9.27) ― 予備タブを 1 枚温めておく
'
'   ★なぜ温めるのか★ NewWindowRequested のハンドラの中で新しい Controller を
'     生成すると、その完了 (ControllerCompleted) はハンドラが return するまで
'     配送されないためデッドロックする (仕様事実 14)。Deferral を使えば回避できるが、
'     args / deferral の COM 参照をイベント外へ持ち越す自前管理が必要になる。
'     ★あらかじめ作っておけば、イベント中は put_NewWindow を撃つだけで済む★ ―― これが
'     プリウォーム方式の全て。イベント内で「待つ」処理が 1 つも無くなる。
'
'   ★予備タブは m_tabs に入れない★ したがって
'     ・タブバーに出ない (TabAdded を撃たないので当然)
'     ・TabCount / ActivateTab / CloseTab / MoveTab の対象外
'     ・NotifyTitleChanged / NotifyUrlChanged は「見当たらない (無視)」で素通りする
'   採用 (AdoptPrewarmedNow) の瞬間に初めて m_tabs へ入り、普通のタブになる。
'
'   ★解放の責任★ m_tabs に入っていないので Class_Terminate のタブ解放ループが
'     拾わない。Shutdown で明示的に ClearBrowserRef → Set Nothing する (設計原則 58)。
' ============================================================

' --- 予備タブが「今すぐ使える」かどうか ---
'   生成済み / Controller Ready / View 取得済み / まだ Claim されていない、の 4 条件。
Public Property Get PrewarmAvailable() As Boolean
    If m_prewarmPane Is Nothing Then Exit Property
    If m_prewarmClaimed Then Exit Property
    If Not m_prewarmPane.IsReady Then Exit Property
    PrewarmAvailable = (m_prewarmPane.ViewPtr <> 0)
End Property


' ============================================================
' ClaimPrewarmViewPtr (第9.27、予備タブの View を「予約付き」で取り出す)
'
'   Wv2Pane.View_OnNewWindowRequested から呼ばれる。使える予備があればその ViewPtr を
'   返し、同時に m_prewarmClaimed を立てて★二重取得を封じる★。
'
'   ★なぜ Claim が要るか★ put_NewWindow を撃ってから採用 Timer が走るまでの間に
'     別の NewWindowRequested が来ると、同じ WebView に 2 つの遷移を載せてしまう。
'     Claim 済みなら 0 を返し、2 本目は現行動作へフォールバックさせる。
'
'   戻り値 0 = 使えない (未生成 / 未 Ready / 既に Claim 済み)。
' ============================================================
Public Function ClaimPrewarmViewPtr() As LongPtr
    If Not PrewarmAvailable Then
        Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: 予備タブが使えない → " & Debug_PrewarmStatus()
        Exit Function
    End If
    ' ★関数名を値として読まない★ VBA では Function 内で関数名を右辺に書くと
    '   再帰呼び出しになる。一時変数に取ってからログに出す。
    Dim pv As LongPtr
    pv = m_prewarmPane.ViewPtr
    m_prewarmClaimed = True
    ClaimPrewarmViewPtr = pv
    Debug.Print "Wv2Browser.ClaimPrewarmViewPtr: 予備タブを予約 (ViewPtr=" & pv & ")"
End Function


' --- Claim を取り消す (put_NewWindow が失敗したときの巻き戻し) ---
Public Sub ReleasePrewarmClaim()
    If Not m_prewarmClaimed Then Exit Sub
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.ReleasePrewarmClaim: 予約を取り消した (予備タブは再利用可能)"
End Sub


' ============================================================
' RequestPrewarmAsync / m_prewarmTimer_TimerCall (第9.27、予備タブ生成の予約と発火)
'
'   ★遅延して生成する理由 (論点1 案b)★ Init 直後は Environment → 最初のタブ →
'     TabBar → NavBar と Controller 生成が連続する最も混雑した区間で、仕様事実 20 の
'     窓が開いている。ここに 4 枚目を差し込まず、静まってから温める。
'
'   ★再入ガード★ CreatePaneCore の WaitPaneReady は DoEvents を回すので、起動中に
'     この TimerCall が割り込むことがある。生成中なら★今回は見送る★ (タイマーの
'     張り直しはしない ― 実行中の SafeTimer を差し替える危険を避けるため)。
'     見送っても AddTab 末尾の EnsurePrewarmScheduled が次の機会に予約し直すので、
'     起動シーケンスが終われば必ず温まる。
' ============================================================
Public Sub RequestPrewarmAsync()
    If Not m_prewarmPane Is Nothing Then
        ' ★第9.28: ログの誤解を解消★ 「使える予備がある」のと「Claim 済み (採用待ち)」
        '   とでは意味がまるで違うのに、9.27 ではどちらも同じ文言だった。
        '   ★どちらの場合も予約はしない (動作は 9.27 と同一)★ Claim 済みなら 1ms 後の
        '   採用 Timer が AdoptPrewarmedNow → CreatePrewarmPaneNow の順で必ず補充する。
        '   ここで別の Timer を張ると、1.5 秒後に「既に予備タブがある (生成をスキップ)」
        '   という無駄なログが 1 行増えるだけで、補充が早まるわけではない。
        If m_prewarmClaimed Then
            Debug.Print "Wv2Browser.RequestPrewarmAsync: 予備タブは Claim 済み (採用待ち)。" & _
                        "採用直後に自動で補充されるので、ここでは予約しない"
        Else
            Debug.Print "Wv2Browser.RequestPrewarmAsync: 使える予備タブが既にあるので何もしない"
        End If
        Exit Sub
    End If

    m_prewarmPending = True
    Set m_prewarmTimer = New SafeTimer     ' 要求ごとに新インスタンス (確実に起動)
    m_prewarmTimer.StartTimer PREWARM_DELAY_MS

    Debug.Print "Wv2Browser.RequestPrewarmAsync: 予備タブの生成を予約 (" & _
                PREWARM_DELAY_MS & "ms 後)"
End Sub


' --- 予備が無く、予約もされていなければ予約する (AddTab 末尾から呼ばれる) ---
'   ★第9.29: 入れ子のときは張らない★ 予備タブの生成中 (CreatePrewarmPaneNow →
'   CreatePaneCore の WaitPaneReady が DoEvents を回している最中) に、予備切れ
'   フォールバックの AddTab が割り込んでここへ来ることがある。そのとき
'   m_prewarmPane はまだ Nothing (外側が代入前) なので予約が張られてしまい、
'   1.5 秒後に「既に予備タブがある (生成をスキップ)」という無駄ログが出ていた。
'   外側が完了すれば予備は 1 枚になるので、ここで張る必要はない。
Private Sub EnsurePrewarmScheduled()
    If Not m_prewarmPane Is Nothing Then Exit Sub
    If m_prewarmPending Then Exit Sub

    If m_creatingPane > 0 Then
        Debug.Print "Wv2Browser.EnsurePrewarmScheduled: 予備タブの生成中 (入れ子) なので" & _
                    "予約しない (外側の完了で 1 枚になる)"
        Exit Sub
    End If

    RequestPrewarmAsync
End Sub


Private Sub m_prewarmTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' ワンショット化: 繰り返し発火を止める
    If Not m_prewarmTimer Is Nothing Then m_prewarmTimer.StartTimer 0

    If Not m_prewarmPending Then Exit Sub

    ' ★再入ガード★ Pane 生成中 (WaitPaneReady の DoEvents 中) に割り込んだ場合は見送る
    If m_creatingPane > 0 Then
        m_prewarmPending = False
        Debug.Print "Wv2Browser.TimerCall(prewarm): Pane 生成中だったので今回は見送り " & _
                    "(AddTab 完了時に予約し直される)"
        Exit Sub
    End If

    m_prewarmPending = False
    CreatePrewarmPaneNow
End Sub


' ============================================================
' CreatePrewarmPaneNow (第9.27、予備タブの実生成)
'
'   CreatePaneCore で Pane を作り、★Ctrl_PutIsVisible False で裏に伏せる★ (論点3)。
'   Bounds は CreatePaneCore が既存タブと同じ矩形で入れてあるので、採用された瞬間に
'   ActivateTab が可視化するだけで正しい位置に出る (可視制御の一元化を崩さない)。
'
'   ★Navigate しない★ 予備は about:blank のまま。put_NewWindow に渡した後、
'     ランタイムが元のリクエスト (POST ならボディごと) をこの WebView に載せる。
' ============================================================
Private Sub CreatePrewarmPaneNow()
    If Not m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: 既に予備タブがある (生成をスキップ)"
        Exit Sub
    End If
    If m_state <> Bs_EnvReady Then
        Debug.Print "Wv2Browser.Prewarm: Environment が未 Ready なので生成しない (state=" & m_state & ")"
        Exit Sub
    End If

    Debug.Print "Wv2Browser.Prewarm: 予備タブの生成を開始"
    Dim pane As Wv2Pane
    Set pane = CreatePaneCore()
    If pane Is Nothing Then
        Debug.Print "Wv2Browser.Prewarm: ★予備タブの生成に失敗★ (次の AddTab / 切替時に再試行される)"
        Exit Sub
    End If

    Dim hr As Long
    hr = pane.Ctrl_PutIsVisible(False)     ' 裏に伏せる (タブとして見えないように)
    If hr <> 0 Then
        Debug.Print "Wv2Browser.Prewarm: Ctrl_PutIsVisible(False) 失敗 hr=&H" & Hex(hr)
    End If

    Set m_prewarmPane = pane
    m_prewarmClaimed = False
    Debug.Print "Wv2Browser.Prewarm: 予備タブを温めた (ViewPtr=" & pane.ViewPtr & _
                ", 表示中のタブ数=" & m_tabs.Count & " ← 予備は含まない)"
End Sub


' ============================================================
' RequestAdoptPrewarmedAsync / m_adoptTimer_TimerCall (第9.27、採用と補充)
'
'   Wv2Pane が put_NewWindow を撃った直後にイベント内から呼ぶ。ハンドラは即 return し、
'   その後の TimerCall で
'     (1) 採用 : 予備を m_tabs へ入れ、TabAdded → ActivateTab で普通のタブにする
'     (2) 補充 : 次の予備を作る
'   の順に処理する (論点6: 採用が先。補充は Controller 生成待ちで重いので後回しにし、
'   ユーザーには先に画面を見せる)。
'
'   ★この TimerCall の中では待ってよい★ NewWindowRequested ハンドラは既に return
'     済みなので、CreatePaneCore の WaitPaneReady が正常に完了する (仕様事実 14 の
'     デッドロック条件を外れている)。m_newTabTimer_TimerCall と同じ構造。
' ============================================================
Public Sub RequestAdoptPrewarmedAsync(ByVal url As String)
    m_adoptPending = True
    m_adoptUrl = url

    Set m_adoptTimer = New SafeTimer      ' 要求ごとに新インスタンス (確実に起動)
    m_adoptTimer.StartTimer 1             ' 1ms 後 (ハンドラ完了後) に TimerCall

    Debug.Print "Wv2Browser.RequestAdoptPrewarmedAsync: 予備タブの採用を予約 (" & url & ")"
End Sub


Private Sub m_adoptTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' ワンショット化: 繰り返し発火を止める
    If Not m_adoptTimer Is Nothing Then m_adoptTimer.StartTimer 0

    If Not m_adoptPending Then Exit Sub
    m_adoptPending = False

    AdoptPrewarmedNow          ' (1) 採用 ― 先に画面を出す
    CreatePrewarmPaneNow       ' (2) 補充 ― 次の予備を温める (重いので後)
End Sub


' ============================================================
' AdoptPrewarmedNow (第9.27、予備タブを普通のタブに昇格させる)
'
'   ★先に m_prewarmPane を手放す★ m_tabs.Add より前に Nothing にしておくことで、
'     この直後の補充 (CreatePrewarmPaneNow) が「既に予備がある」と誤判定しない。
'   ★可視化は ActivateTab に任せる★ CreatePrewarmPaneNow で伏せた可視状態は
'     ActivateTab が起こす。可視制御のロジックが 1 箇所に集まったままになる。
' ============================================================
Private Sub AdoptPrewarmedNow()
    If m_prewarmPane Is Nothing Then
        Debug.Print "Wv2Browser.Adopt: ★予備タブが無い★ (採用をスキップ。put_NewWindow は" & _
                    "撃たれているので遷移は宙に浮く)"
        m_prewarmClaimed = False
        Exit Sub
    End If

    Dim pane As Wv2Pane
    Set pane = m_prewarmPane
    Set m_prewarmPane = Nothing
    m_prewarmClaimed = False

    m_tabs.Add pane
    Debug.Print "Wv2Browser.Adopt: 予備タブをタブとして採用 (現在のタブ数=" & m_tabs.Count & _
                ", 要求 URL=" & m_adoptUrl & ")"

    RaiseEvent TabAdded(m_tabs.Count)
    ActivateTab m_tabs.Count
End Sub


' --- 検証用: 予備タブの状態を 1 行の文字列で返す ---
Public Function Debug_PrewarmStatus() As String
    If m_prewarmPane Is Nothing Then
        Debug_PrewarmStatus = "予備なし (pending=" & CStr(m_prewarmPending) & ")"
        Exit Function
    End If
    Debug_PrewarmStatus = "予備あり ViewPtr=" & m_prewarmPane.ViewPtr & _
                          " / IsReady=" & CStr(m_prewarmPane.IsReady) & _
                          " / Claimed=" & CStr(m_prewarmClaimed) & _
                          " / 使用可=" & CStr(PrewarmAvailable)
End Function


' ============================================================
' m_newTabTimer_TimerCall (第9.10e、SafeTimer のワンショット発火受け口)
'
'   NewWindowRequested ハンドラは既に return 済みのタイミングで呼ばれる。
'   ここで AddTabWithUrl すれば、新 Controller の ControllerCompleted が正常に
'   配送されるためデッドロックしない。
'
'   まず StartTimer 0 で繰り返し発火を止め (ワンショット化)、保留キューを順に処理。
' ============================================================
Private Sub m_newTabTimer_TimerCall(ByVal wTime As Long, ByRef varData As Variant)
    ' ワンショット化: 繰り返し発火を止める
    If Not m_newTabTimer Is Nothing Then m_newTabTimer.StartTimer 0

    If m_pendingUrls Is Nothing Then Exit Sub
    Debug.Print "Wv2Browser.TimerCall: 保留 URL を処理 (キュー数=" & m_pendingUrls.Count & ")"

    Do While m_pendingUrls.Count > 0
        Dim url As String
        url = m_pendingUrls(1)
        m_pendingUrls.Remove 1

        Dim pane As Wv2Pane
        Set pane = AddTabWithUrl(url)
        If pane Is Nothing Then
            Debug.Print "  RequestNewTab: AddTabWithUrl 失敗 (" & url & ")"
        Else
            Debug.Print "  RequestNewTab: 新タブ生成 OK (" & url & ", ViewPtr=" & _
                        pane.ViewPtr & ")"
        End If
    Loop
End Sub


' ============================================================
' Shutdown (第9.10e、循環参照を断ち切ってから解放する)
'
'   第9.10e で Wv2Pane が Wv2Browser への後方参照 (m_browser) を持つように
'   なったため、Browser.m_tabs ⇔ Pane.m_browser の循環参照が生じる。
'   VBA は参照カウント GC なので、UserForm が単に Set m_browser = Nothing しても
'   循環参照でどちらも解放されない (Class_Terminate も呼ばれない)。
'
'   そこで UserForm_Terminate から本 Shutdown を明示的に呼び、各 Pane の後方参照
'   (ClearBrowserRef) を切ってから解放する。これにより Browser→Pane の一方向参照
'   だけになり、続く Set m_browser = Nothing で Class_Terminate が正常に走る
'   (全 Pane → Environment の順で解放)。
'
'   ★UserForm 側の実装コストはこの Shutdown 呼び出し 1 行のみ★ (案Y の狙い)。
'   冪等 (複数回呼んでも安全)。
' ============================================================
Public Sub Shutdown()
    ' 第9.10e (Timer): 保留タイマーを止めて解放する
    If Not m_newTabTimer Is Nothing Then
        m_newTabTimer.StartTimer 0
        Set m_newTabTimer = Nothing
    End If

    ' 第9.18: 設定タブ用タイマーとブリッジを解放する
    If Not m_settingsTimer Is Nothing Then
        m_settingsTimer.StartTimer 0
        Set m_settingsTimer = Nothing
    End If
    m_settingsPending = False
    Set m_settingsBridge = Nothing

    ' 第9.27: プリウォーム機構を止めて予備タブを解放する
    '   ★予備は m_tabs に入っていない★ ので、下の全タブループでは拾われない。
    '   ここで後方参照を切ってから Nothing にしないと循環参照で解放されない (設計原則 58)。
    If Not m_prewarmTimer Is Nothing Then
        m_prewarmTimer.StartTimer 0
        Set m_prewarmTimer = Nothing
    End If
    If Not m_adoptTimer Is Nothing Then
        m_adoptTimer.StartTimer 0
        Set m_adoptTimer = Nothing
    End If
    m_prewarmPending = False
    m_adoptPending = False
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        m_prewarmClaimed = False
        Debug.Print "Wv2Browser.Shutdown: 予備タブ (プリウォーム) を解放"
    End If

    If m_tabs Is Nothing Then Exit Sub
    Dim pane As Wv2Pane
    For Each pane In m_tabs
        If Not pane Is Nothing Then pane.ClearBrowserRef
    Next pane
    Debug.Print "Wv2Browser.Shutdown: 全 " & m_tabs.Count & " タブの後方参照を切断"
End Sub


' ============================================================
' Class_Terminate ? 全タブ解放 → Environment 解放 (設計原則 43: 子から親へ)
' ============================================================
Private Sub Class_Terminate()
    Debug.Print "Wv2Browser.Terminate 開始 (タブ数=" & m_tabs.Count & ")"

    ' --- 0. 第9.27: 予備タブの保険解放 ---
    '   通常は Shutdown が先に片付けている。Shutdown を通らずここへ来た場合の保険。
    If Not m_prewarmPane Is Nothing Then
        m_prewarmPane.ClearBrowserRef
        Set m_prewarmPane = Nothing
        Debug.Print "Wv2Browser.Terminate: 予備タブ (プリウォーム) を解放"
    End If

    ' --- 1. 全 Pane を解放 (各 Pane の Class_Terminate が View 系を順に解放) ---
    '   Collection から取り除くと参照が切れ、参照カウント 0 で
    '   Wv2Pane.Class_Terminate が走る。末尾から順に取り除く。
    Dim i As Long
    For i = m_tabs.Count To 1 Step -1
        m_tabs.Remove i
    Next i
    Set m_tabs = Nothing

    ' --- 2. Environment を解放 (全 Pane より後 = 親を最後に) ---
    Set m_env = Nothing

    m_hwndHost = 0
    m_state = Bs_Created
    Debug.Print "Wv2Browser.Terminate 完了"
End Sub

