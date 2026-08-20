Attribute VB_Name = "Wv2Tests"
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  D-2 段階 (要素レジストリ + Wv2Element の検証) ---
'
'   D-2 の追加事項:
'     Test_D2_Find  … 自前の検証ページを開き、GetElementById / QuerySelector と
'                     Wv2Element の読み取り 8 メンバを期待値付きで一括検証する。
'     Test_D2_Stale … 要素を掴んだままページを遷移させ、世代が変わって
'                     IsStale が True になること、読み取りが stale で失敗すること、
'                     新しいページで取り直せることを確認する。
'     Test_D2_Help  … 上 2 つの手順と、見るべきログの説明。
'
'   ★検証ページは自前 HTML (論点8 案b)★
'     外部サイトは落ちるうえ DOM が予告なく変わる (設計原則75、第9.28 の教訓)。
'     NavigateToString で流し込むので外部依存もファイル I/O も無い。
'     網羅しているもの: id / class / 任意属性 / 日本語 / 記号 / 改行 /
'     タグを含む innerHTML (仕様事実30) / input・textarea・select の value /
'     空要素 / 属性を持たない要素。
'
'   既存の Test_* は 1 つも変更していない。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  D-1 段階 (EvalSync の検証) ---
'
'   ★段階番号の体系が変わりました★ 第9.32b までの 9.xx は「ブラウザとしての機能
'     整備」という軸の里程標。D-1 からは軸ごとの記号 (D = DOM / ページ自動制御)。
'
'   D-1 の追加事項:
'     Test_D1_Eval  … EvalSync の一括検証。型・日本語・エスケープ・例外・構文エラー・
'                     タイムアウト・タイムアウト後の回復までを 1 発で流す。
'     Test_D1_Guard … 論点7 の in-callback ガードの発火と ResetCallbackGuard の確認。
'     Test_D1_Help  … 上 2 つの手順と、見るべきログの説明。
'
'   既存の Test_* は 1 つも変更していない。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.32b 段階 (ホイール横スクロールの検証手順) ---
'
'   第9.32b の追加事項:
'     ★Test_9_32b_Help (実機手順)★ JS のみの変更なので純ロジックのテストは無い。
'       (1) ホイールでタブ列が横スクロールするか
'       (2) タイトル更新の同期でスクロール位置が左端へ戻らないか
'       (3) + で足した新しいタブが自動で見える位置に来るか
'       (4) 回帰: 切替の体感 / D&D / 閉じる
'
'   既存の Test_* は 1 つも変更していない。
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.32 段階 (タブ切替の体感とスクロールバーの検証手順) ---
'
'   第9.32 の追加事項:
'     ★Test_9_32_Help (実機手順)★
'       第9.32 も JS と CSS のみの変更なので★純ロジックのテストは存在しない★
'       (目視と体感で判定する)。見るもの:
'         (1) クリックした瞬間にタブの色が変わり、本体の切替が後から追いつくか
'         (2) タブを増やしても横スクロールバーが出ず、タブが潰れないか
'         (3) ホイールでタブ列を横スクロールできるか
'         (4) 回帰: D&D 並べ替え / 閉じる / 設定タブ が壊れていないか
'
'   既存の Test_* は 1 つも変更していない。
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.31 段階 (タブバーの見た目の検証手順) ---
'
'   第9.31 の追加事項:
'     ★Test_9_31_Help (実機手順)★
'       第9.31 は CSS のみの変更なので★純ロジックのテストは存在しない★
'       (目視で判定する)。手順は 2 部構成:
'         第 1 部 = (A-1) ホバー背景 / (A-2) 閉じるボタンの表示制御 の目視確認
'         第 2 部 = ★第9.30 の宿題 1 = v0_5_2 の通し動作確認★
'                   POST リンクの新タブ展開 / プリウォーム / ドラッグ並べ替え。
'                   第9.30 でクリーンビルドした v0_5_2 は「コンパイル通過」と
'                   「設定タブからの検索エンジン切替」しか確認していないため、
'                   タブバーを触るこの回で残りをまとめて踏む。
'
'   既存の Test_* は 1 つも変更していない。
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.29 段階 (カスタム検索エンジンの検証) ---
'
'   第9.29 の追加事項:
'     ★Test_9_29_Custom (純ロジック)★
'       Wv2SettingsBridge の SetCustomEngine / GetCustomTemplate /
'       PreviewUrlForTemplate と、custom の保存・読込 (engine + template の
'       2 行) を照合する。WebView2 は起動しない。
'       実ファイル (%APPDATA%\Wv2Browser\settings.txt) を触るので、冒頭で
'       ★engine と template を対で★退避し、末尾で書き戻す。
'
'     ★Test_9_29_Help (実機手順)★
'       設定タブのカスタム入力欄 → 適用 → 検索 → 再起動で復元、まで。
'
'     ★Test_9_20_Persistence の退避/復元を custom 対応にした★
'       第9.29 で settings.txt が 2 キーになったため、Debug_SaveEngineName で
'       書き戻すと「元が custom だった人の template 行を消してしまう」。
'       Debug_SaveSettings (engine + template) で退避・復元するよう変更。
'
' --- Wv2Tests.bas  第9.28b 段階 (POST プローブの送信先を 3 系統化 / 本命を httpbingo.org へ) ---
'
'   第9.28 の実機検証で httpbin.org が★サーバ側の都合で 503 を返し続け★、本丸の
'   再確認ができなかった。素の GET (アドレスバー直打ち) でも 503 だったので、
'   健全なら 405 が返るはずのところがサーバダウンで潰れていたと確定した。
'   httpbin.org は無料の公開サービスで、周期的な 503 の報告が古くからある。
'   → ★検証基盤を単一の外部サービスに依存させない★ ように送信先を 3 つに増やす。
'
'   ★変更点 (BuildPostProbeHtml と説明文のみ。本番ファイルは 1 行も触らない)★
'     ボタン1 (本命 / POST + _blank)  : httpbin.org → ★httpbingo.org★
'     ボタン2 (サニティ / POST 同一タブ): httpbin.org → ★httpbingo.org★
'     ボタン3 (対照 / GET + _blank)   : httpbin.org/get → ★httpbingo.org/get★
'                                        (本命と同じサーバに揃えて変数を 1 つに絞る)
'     ボタン4 (予備1)                 : postman-echo.com → ★httpbin.org★ (旧本命を降格)
'     ボタン5 (予備2)                 : ★新設★ postman-echo.com
'
'   ★httpbingo.org を本命にした理由と注意★
'     ・GET で叩くと 405 Method Not Allowed を返す (実機確認済み)。職場の Tomcat と
'       同じ構図がそのまま成立する。
'     ・POST の応答 JSON に ★method フィールド★ があり、GET に化けていないことを
'       画面上で直接確認できる (httpbin には無い利点)。
'     ・★注意★ httpbingo のエラーページには title 要素が無く、Edge が URL を
'       タイトル代用にする。そのため ★イミディエイトのログでは合否が読めない★
'       (httpbin は 405 METHOD NOT ALLOWED をタイトルに出していた)。
'       判定は必ず★画面表示★で行うこと。
'
'   ★JS はゼロのまま★ 素の form だけで組む方針は維持 (VBA 文字列内の JS は事故のもと)。
'
' --- Wv2Tests.bas  第9.28 段階 (プリウォーム委譲の既定化に伴う検証手順の更新) ---
'
'   第9.28 の追加/変更:
'     ★Test_9_28_Help (新規)★
'       既定が 2 (プリウォーム委譲) になった状態での実機手順。★モードを切り替えずに★
'       POST リンクが自前タブで 200 になることを確認するのが主目的。
'     ★Test_9_27_Help / Test_9_27_Status の文言を既定 2 前提に修正★
'       9.27 の手順は「既定 0 → 対照を取ってから 2 に切替」を前提に書かれていた。
'       既定が 2 になったので、対照 (405) を見るには先に Test_9_27_Mode_Legacy を
'       打つ必要がある。その一点だけを直した。
'
'   ★Test_9_27_Mode_Legacy / _Popup / _Prewarm / _Status / _Handled_On / _Off は
'     9.27 のまま無変更★ 既定が変わっただけで、口の意味は変わっていない。
'
' --- Wv2Tests.bas  第9.27 段階 (本丸: put_NewWindow によるプリウォーム委譲の検証) ---
'
'   第9.27 の追加:
'     ★Test_9_27_Mode_Legacy / _Popup / _Prewarm (モード切替)★
'       Wv2Browser.NewWindowMode を 0 / 1 / 2 に切り替える薄い口。起動中でも即時に効く。
'     ★Test_9_27_Status (状態ダンプ)★
'       予備タブ (プリウォーム) が温まっているかを 1 行で確認する。
'     ★Test_9_27_Handled_On / _Off (論点5 の保険スイッチ)★
'       put_NewWindow の後に put_Handled(TRUE) も立てるかどうかを切り替える。
'       予備タブが白いまま / 二重に開く 等の症状が出たときの切り分け用。
'     ★Test_9_27_Help (実機手順)★
'       9.26b の POST プローブページ (Test_9_26_PostProbe) をそのまま流用する。
'       ★委譲 OFF (モード 0) で 405 / モード 2 で 200 + form エコー★ が合格条件。
'
'   ★新しい検証ページは作らない★ 9.26b の postprobe.html が
'     「GET なら 405 / POST なら 200 + ボディをエコー」という職場と同じ構図を
'     そのまま持っているので、本丸の合否判定にも過不足なく使える。
'
' --- Wv2Tests.bas  第9.26b 段階 (POST プローブページの追加) ---
'
'   第9.26 の実機検証で分かったこと:
'     プローブの機構 (トグル / 分岐 / ランタイム委譲 / 参照リークなし) はすべて成立した。
'     しかし試したサイトのポップアップが★素の GET★ だったため、委譲 OFF でも自前タブで
'     正常に開いてしまい (isSuccess: True)、405 の再現になっていなかった。
'     つまり「委譲すれば POST が保たれるのか」という本命の判定材料は未取得。
'
'   ★この段の狙い★ 職場に行かずに 405 の構図を自宅で合成して本命を判定する。
'     送信先を https://httpbin.org/post にすると
'       ・GET  で叩く → 405 METHOD NOT ALLOWED   (= 職場の Tomcat とまったく同じ構図)
'       ・POST で叩く → 200 + 送信ボディを JSON でエコー
'     となるので、1 つのボタンで「405 の再現」と「ボディが運ばれた直接証拠」の両方が
'     手に入る。委譲 OFF で 405 / ON で 200 になれば パターンA が確定する。
'
'   ★追加したもの (すべて Wv2Tests.bas 内で完結。本番ファイルは 1 行も触らない)★
'     ・Test_9_26_PostProbe      : 検証ページを書き出して新タブで開く
'     ・Test_9_26_PostProbe_Help : 手順と結果の読み方
'     ・BuildPostProbeHtml       : 検証ページの HTML (private)
'     ・WriteUtf8NoBom           : UTF-8 (BOM なし) 書き出しヘルパ (private)
'
'   ★設計判断のメモ★
'     ・開き方は仮想ホストマッピング (https://appassets.postprobe/postprobe.html)。
'       NavigateToString は origin が "null" になり Referer やフォーム送信の条件が
'       本番と変わってしまうので、検証の純度を落とさないため採らない。
'       既存の Wv2Browser.AddTabWithUrlForSpa にそのまま乗る。
'     ・書き出しは ADODB.Stream で UTF-8 (BOM なし)。UserForm1.WriteSpaAppFolder は
'       Print # による ANSI (CP932) 書き出しなので、meta charset=UTF-8 の HTML に
'       日本語を入れると化ける。Wv2TabBar/Wv2NavBar で実績のある方式に揃えた。
'       (ADODB.Stream は遅延バインドなので参照設定ゼロの方針には抵触しない)
'     ・★JS はゼロ★ 素の <form> だけで組んだ。VBA 文字列内の JS は事故のもとなので、
'       HTML だけで足りる要件にわざわざ JS を持ち込まない。
'
' --- Wv2Tests.bas  第9.26 段階 (NewWindowRequested 委譲プローブ) ---
'
'   第9.26 の追加:
'     ★Test_9_26_Popup_On / Test_9_26_Popup_Off (実機スイッチ)★
'       起動中の Wv2Browser の PopupDelegateMode を切り替えるだけの薄い口。
'       UserForm1.CurrentBrowser (第9.17 の足場) を経由するので UserForm1.frm は無変更。
'     ★Test_9_26_Popup_Help (実機手順)★
'       職場の基幹システムで POST リンクを踏み、405 が消えるかを見る手順と、
'       結果の読み方 (パターンA/B) を出力する。
'
'   ★この回は純ロジック検証が無い★ 見たいのは「WebView2 ランタイムが遷移を継続するか」
'     という実機の挙動そのものなので、WebView2 を起動しない検証にはできない
'     (9.21?9.25 と同じ判断)。
'
' --- Wv2Tests.bas  第9.25b 段階 (検証プローブの撤去 / 引数ゼロを恒久確定) ---
'
'   第9.25 の検証が「パターンA = 安全」で決着したため、Test_9_25_Probe_Help を削除した。
'   プローブ本体 (Wv2NavBar の HostProbe / HostProbeMark と JS の probeParamless) も
'   同時に撤去済みなので、手順 Sub だけ残しても実行できるものが無いため。
'
'   ★結論 (仕様事実 22。詳細は開発メモ)★
'     hostObjects の sync プロキシは、メンバー取得の時点では VBA 側を呼び出さない。
'     呼び出しは実際の () まで遅延される。したがって引数ゼロの Public Function も
'     そのまま公開してよく、9.24 のダミー引数 reserved は不要だった。
'     実測ログ: (1) 取得の直前 → (2) 取得の直後 → ★実行★ → (3) f() 成功 →
'     ★実行★ → (5) 直呼び成功。(1)-(2) 間に実行が無く、通算 2 回ちょうど。

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.25 段階 (検証: 引数ゼロ Host メソッドの可否 + reserved 削除) ---
'
'   第9.25 の追加:
'     ★Test_9_25_Probe_Help (実機手順)★
'       9.24 の宿題「ダミー引数 reserved は本当に必要だったのか」を潰す回。
'       NavBar から reserved を削除して引数ゼロに戻し、同時に副作用ゼロの検証専用
'       プローブ (HostProbe / HostProbeMark) を仕込んだ (案C)。
'
'       ★この回は起動するだけで機構の答えが出る★ プローブは NavBar の JS が
'       ready 送信の 500ms 後に自走するので、たーぼーさんは普通に起動して
'       [PROBE] 行の並びを読むだけでよい。そのうえで ← → ? を 1 回ずつ押して
'       実物が二重実行にならないかを確認する。

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.24 段階 (NavBar の hostObjects 化 + Host 一元化 検証) ---
'
'   第9.24 の追加:
'     ★Test_9_24_HostNavBar_Help (実機手順)★
'       NavBar の back/forward/reload/navigate を hostObjects 経路へ移し、処理を
'       HostBack/HostForward/HostReload/HostNavigate へ一元化した回の実機手順。
'       実 URL Navigate + JS 実行 + hostObjects 経路の成立を見るため、純ロジック
'       検証は無い (9.21?9.23 と同じ判断)。
'
'       ★この回の判定は 3 本立て★
'         (1) 4 コマンドが host 経由になったか (HostXxx ログが出て、OnPaneWebMessage
'             の当該 cmd ログが出ない)
'         (2) ★二重実行が起きていないか (論点4b)★ ← を 1 回押して HostBack が
'             1 行だけか。2 行出たり、OnPaneWebMessage の back が続いたりしたら
'             ダミー引数 reserved でも塞げていないということ (要報告)
'         (3) ★文字列引数が無傷か (仕様事実 21 候補)★ 日本語の検索語・記号入り
'             URL が HostNavigate のログに化けずに出るか
'
'       回帰確認は ready ハンドシェイク / Escape の編集破棄 / 編集中フラグ /
'       TabBar 側 (9.21?9.23) が無傷であること。

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.23 段階 (整理: activate/close/reorder の Host 一元化 検証) ---
'
'   第9.23 の追加:
'     ★Test_9_23_Consolidation_Help (実機手順)★
'       9.23 は OnPaneWebMessage の activate/close/reorder ケースの中身を
'       HostActivate/HostClose/HostReorder への呼び捨て委譲に一元化した整理
'       ステージ (案C-1)。JS 側ハイブリッドは無変更なので、通常運用では
'       hostObjects 経路が成立し Case (=フォールバック経路) はそもそも通らない。
'       よって実機で見るべきは「回帰していないこと」= activate/close/reorder が
'       従来どおり host 経由 (HostXxx ログ) で動き UX が不変であること、および
'       ready/newtab/settings が無傷であること。純ロジック検証は WebView2 実起動が
'       必須のため無し (9.21/9.22 と同じ判断)。
'

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.22 段階 (試験: close / reorder の hostObjects 化 検証) ---
'
'   第9.22 の追加:
'     ★Test_9_22_HostCloseReorder_Help (実機手順)★
'       TabBar の close / reorder を hostObjects.sync.TabBar.HostClose /
'       HostReorder 経由に移した回の実機手順。close の × / reorder のドラッグで
'       host経由ログが出て OnPaneWebMessage が出ないこと、activate (9.21) も
'       引き続き host経由であること、ready/newtab/settings は postMessage の
'       ままであること (回帰なし) を確認する。実 URL Navigate + JS 実行 +
'       hostObjects 経路の成立は実起動必須のため純ロジック検証は無し。
'
' --- Wv2Tests.bas  第9.21 段階 (試験: activate の hostObjects 化 検証) ---
'
'   第9.21 の追加:
'     ★Test_9_21_HostActivate_Help (実機手順)★
'       TabBar の activate を hostObjects.sync.TabBar.HostActivate 経由に
'       試験移行した結果を実機で確かめる手順。純ロジック検証は無い (実 URL
'       Navigate + JS 実行 + hostObjects 経路の成立を見るため、WebView2 の
'       実起動が必須)。タブを複数開いてクリックで切替え、イミディエイトに
'       『HostActivate: index=N (host経由)』が出て、かつ従来の
'       『OnPaneWebMessage: msg={"cmd":"activate"...}』が出ないことを確認する。
'
' --- Wv2Tests.bas  第9.20 段階 (検索エンジン選択の永続化 検証) ---
'
'   第9.20 の追加:
'     ★Test_9_20_Persistence (純ロジック)★
'       Wv2SettingsBridge の Debug_SaveEngineName / LoadEngineName /
'       LoadEngineTemplate を直接叩き、保存→読込の往復と、ファイル
'       が無いとき/未知エンジン名のフォールバックを照合する。
'       WebView2 は起動しない。実ファイル (%APPDATA%\Wv2Browser\
'       settings.txt) を触るので、冒頭で現在の保存値を退避し、
'       末尾で書き戻してユーザー設定を壊さないようにする。
'
'     ★Test_9_20_Persistence_Help (実機手順)★
'       起動→設定でエンジン変更→Excel 完全終了→再起動で検索が
'       前回エンジンで始まることを確かめる (復元は Class_Initialize
'       の実行を伴うので実機手順で確認)。
'
' --- Wv2Tests.bas  第9.19 段階 (設定タブの重複防止 検証手順) ---
'
'   第9.19 の追加事項:
'     ★Test_9_19_SettingsTabDedup_Help を追加★
'       設定タブの重複防止 (? 連打で設定タブが増えない) は WebView2 の実際の
'       タブ生成・アクティブ化を伴うため、純ロジック検証ではなく実機手順で
'       確認する。イミディエイトで Test_9_19_SettingsTabDedup_Help と打つと
'       手順が出る。ロジックの要 (IsSettings フラグ立て + m_tabs 走査による
'       重複回避) は Wv2Pane / Wv2Browser のコードレビューで担保する。
'
' --- Wv2Tests.bas  第9.18 段階 (Wv2SettingsBridge の純ロジック検証) ---
'
'   第9.18 の追加: Test_9_18_BridgeLogic を追加。Wv2SettingsBridge を new し、
'   実 Wv2Browser を BindBrowser で結び付けてから、設定画面の JS が呼ぶのと同じ
'   メソッド (SetEngine / GetEngine / PreviewUrl / PreviewUrlFor / EngineList) を
'   直接叩いて、名前?テンプレート解決・プレビュー URL 生成・副作用なしを照合する。
'   WebView2 は起動しない (ブリッジ + Browser の Debug_* を叩くだけ)。
'   ※hostObjects 経路が実機で通るか (順序A) の確認は Help の実機手順で行う。
'   汎用ヘルパー CheckEq (文字列一致) / CheckBool (真偽) を追加。
'
' --- Wv2Tests.bas  第9.17 段階 (検索エンジンのプリセット解決を検証) ---
'
'   第9.17 の追加: Test_9_17_SearchPresets を追加。UserForm1.UseSearchEngine が
'   プリセット名 ("google"/"bing"/"duckduckgo"/"yahoo") を解決する先の
'   テンプレート文字列を、Wv2Browser に直接 Let したときに検索 URL がそのエンジンへ
'   正しく追従するかを純ロジックで照合する。UserForm を起動せず (m_browser 非依存)、
'   Browser を new して Debug_NormalizeUrl / Debug_SearchEngine を叩くだけ。
'   ※UseSearchEngine 自体 (名前→テンプレート解決 + 未起動ガード) の確認は実機
'     (CurrentBrowser 経由) で行う。ここではプリセット値の正しさを担保する。
'
' --- Wv2Tests.bas  第9.16 段階 (整理: フィーチャー検証 Sub の集約) ---
'
'   WebView2 を起動しない「純ロジック検証 Sub」を集める標準モジュール。
'   Wv2Browser の公開検証口 (Debug_NormalizeUrl / Debug_SearchEngine /
'   Debug_TabOrderTitles / Debug_SeedDummyTabs / Debug_ActiveIndex / MoveTab)
'   を叩くだけで、サンク基盤や WebView2 本体には一切触れない。
'
'   ★なぜ分けたか (案1-A)★
'     これらの検証はもともと Wv2Thunks.bas に溜まっていたが、Wv2Thunks は
'     「機械語サンク・vtable・メモリプリミティブ」の心臓部モジュールであり、
'     サンクと無関係なフィーチャー検証が同居するのは責任範囲の侵食だった
'     (Wv2Json を切り出した「関心事で分ける」原則との不整合)。そこで検証専用の
'     受け皿としてこのモジュールを新設し、Test_9_14 以降をここへ集約した。
'     以後の Test_9_17, Test_9_18… もここに足す (ファイルはもう増やさない)。
'
'   ★移植性★ このモジュールは検証コード専用なので、本番移植時に「テストが
'     要らなければ Wv2Tests.bas はインポートしない」判断ができる。本番コードと
'     テストコードが物理的に分離するため、むしろ移植性は上がる。
'
'   ★ここに置くもの / 置かないもの★
'     置く  : Wv2Browser の Debug_* を叩くだけの純ロジック検証 (WebView2 起動不要)。
'     置かない: サンク基盤自体の検証 (Test_VirtualAllocFree_Roundtrip 等) は
'              Wv2Thunks の Private と密結合するため Wv2Thunks に残す。
'              実機起動を伴う検証 (StartWebView2_Full 等) は UserForm1 に残す
'              (UserForm のインスタンス状態 m_browser 等を握るため)。
''''''''''''''''''''''''''''''''''

Option Explicit


' ============================================================
' Test_9_14_SearchFallback  (第9.14 検索語フォールバックのロジック検証)
'
'   ★イミディエイトウィンドウで  Test_9_14_SearchFallback  と打つだけ★
'
'   WebView2 は起動しない。Wv2Browser を new して Debug_NormalizeUrl を叩き、
'   入力が「URL」と「Google 検索」に正しく振り分けられるか、および日本語・記号の
'   percent-encoding が正しいかを一括で確認する。
'   期待値と一致すれば [OK]、違えば [NG] を Debug.Print する。末尾に合否サマリ。
' ============================================================
Public Sub Test_9_14_SearchFallback()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' ※ Wv2Browser の Initialize は不要 (Debug_NormalizeUrl は NormalizeUrl を
    '    呼ぶだけで、タブ・Environment・ウィンドウに一切触れないため)。

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_14_SearchFallback 開始 ===="

    ' --- (1) スキーム有り → そのまま ---
    Check b, "https://example.com", "https://example.com", total, pass
    Check b, "http://example.com/a?b=1", "http://example.com/a?b=1", total, pass
    Check b, "about:blank", "about:blank", total, pass
    Check b, "file:///C:/tmp/x.html", "file:///C:/tmp/x.html", total, pass

    ' --- (2) ドット有り・空白なし → URL (https:// 前置) ---
    Check b, "example.com", "https://example.com", total, pass
    Check b, "www.google.co.jp", "https://www.google.co.jp", total, pass
    Check b, "192.168.0.1", "https://192.168.0.1", total, pass
    Check b, "localhost:8080", "localhost:8080", total, pass          ' コロン前が英字のみ→スキーム扱いで素通し(Edge が http:// を補う)

    ' --- (3) 空白あり or ドットなし → Google 検索 ---
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, "localhost", "https://www.google.com/search?q=localhost", total, pass
    Check b, "excel vba webview2", "https://www.google.com/search?q=excel%20vba%20webview2", total, pass

    ' --- (4) 日本語・記号のエンコード (encodeURIComponent 相当) ---
    Check b, "日本 天気", "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass
    Check b, "C++ 入門", "https://www.google.com/search?q=C%2B%2B%20%E5%85%A5%E9%96%80", total, pass
    Check b, "100% pure", "https://www.google.com/search?q=100%25%20pure", total, pass
    Check b, "a.b-c_d~e f", "https://www.google.com/search?q=a.b-c_d~e%20f", total, pass  ' 空白ありで検索、-_.~ は素通し

    ' --- (5) 前後空白トリム ---
    Check b, "  example.com  ", "https://example.com", total, pass
    Check b, "", "", total, pass   ' 空入力は空 (NavigateActive 側で無視される)

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]") & " ===="

    Set b = Nothing
End Sub

' 検証ヘルパー: 期待値と比較して Debug.Print (Test_9_14_SearchFallback 専用)
Private Sub Check(ByRef b As Wv2Browser, ByVal inp As String, _
                  ByVal expected As String, ByRef total As Long, ByRef pass As Long)
    total = total + 1
    Dim got As String
    got = b.Debug_NormalizeUrl(inp)
    If got = expected Then
        pass = pass + 1
        Debug.Print "[OK] """ & inp & """ -> " & got
    Else
        Debug.Print "[NG] """ & inp & """" & vbCrLf & _
                    "      got     = " & got & vbCrLf & _
                    "      expected= " & expected
    End If
End Sub

' ============================================================
' Test_9_15_Reorder  (第9.15 タブ並べ替えのロジック検証)
'
'   ★イミディエイトウィンドウで  Test_9_15_Reorder  と打つだけ★
'
'   WebView2 は起動しない。Wv2Browser を new し、Debug_SeedDummyTabs で
'   タグ付きダミータブ ("A","B","C",...) を仕込んでから MoveTab を色々な
'   from/to で叩き、並べ替え後のタブ順 (Debug_TabOrderTitles) と
'   アクティブ index の追従 (Debug_ActiveIndex) を期待値と照合する。
'   末尾に合否サマリ。
'
'   ★アクティブ追従の確認★ seed 直後は activeIndex=1 (=タブ"A")。以降 MoveTab
'   を重ねても「"A" がいる位置」に activeIndex が追従することを確認する
'   (MoveTab はオブジェクト参照でアクティブを逆引きするため)。
'
'   ※ ダミー Pane は View を持たないため、MoveTab 内の ActivateTab が
'     PutIsVisible 失敗ログを出すが、順序・activeIndex 追従には影響しない。
' ============================================================
Public Sub Test_9_15_Reorder()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_15_Reorder 開始 ===="

    Dim b As Wv2Browser

    ' --- ケース1: [A,B,C,D] で A(1) を末尾(4)へ → [B,C,D,A], active は A に追従=4 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 1, 4
    CheckOrder b, "case1 A->末尾", "B|C|D|A", 4, total, pass

    ' --- ケース2: [A,B,C,D] で D(4) を先頭(1)へ → [D,A,B,C], active(A) は 2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 4, 1
    CheckOrder b, "case2 D->先頭", "D|A|B|C", 2, total, pass

    ' --- ケース3: [A,B,C,D] で B(2) を C の後ろ(3)へ → [A,C,B,D], active(A) は 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 2, 3
    CheckOrder b, "case3 B->3", "A|C|B|D", 1, total, pass

    ' --- ケース4: [A,B,C,D] で C(3) を B の前(2)へ → [A,C,B,D], active(A) は 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 3, 2
    CheckOrder b, "case4 C->2", "A|C|B|D", 1, total, pass

    ' --- ケース5: from==to は無変化 (成功扱い) → [A,B,C,D], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 4
    b.MoveTab 2, 2
    CheckOrder b, "case5 同一", "A|B|C|D", 1, total, pass

    ' --- ケース6: 範囲外 from=0 は無変化 → [A,B,C], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 0, 2
    CheckOrder b, "case6 範囲外from", "A|B|C", 1, total, pass

    ' --- ケース7: 範囲外 to=99 は無変化 → [A,B,C], active 1 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 2, 99
    CheckOrder b, "case7 範囲外to", "A|B|C", 1, total, pass

    ' --- ケース8: アクティブ自身を動かす。[A,B,C] active=1(A)、A(1)を3へ ---
    '   → [B,C,A], active は A に追従して 3 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 3
    b.MoveTab 1, 3
    CheckOrder b, "case8 active移動", "B|C|A", 3, total, pass

    ' --- ケース9: 2タブ入れ替え [A,B] → B(2)を1へ → [B,A], active(A)=2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 2
    b.MoveTab 2, 1
    CheckOrder b, "case9 2タブ入替", "B|A", 2, total, pass

    ' --- ケース10: 5タブ多段。[A,B,C,D,E] で C(3)を1へ → [C,A,B,D,E], active(A)=2 ---
    Set b = New Wv2Browser
    b.Debug_SeedDummyTabs 5
    b.MoveTab 3, 1
    CheckOrder b, "case10 C->先頭", "C|A|B|D|E", 2, total, pass

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ------------------------------------------------------------
' CheckOrder ? MoveTab 後のタブ順と activeIndex を期待値と照合する (第9.15)
'   b            : 対象 Browser
'   label        : ケース名 (ログ表示用)
'   expectOrder  : 期待するタブ順 ("B|C|D|A" のような "|" 連結タグ)
'   expectActive : 期待する activeIndex
' ------------------------------------------------------------
Private Sub CheckOrder(ByRef b As Wv2Browser, ByVal label As String, _
                       ByVal expectOrder As String, ByVal expectActive As Long, _
                       ByRef total As Long, ByRef pass As Long)
    total = total + 1
    Dim gotOrder As String
    Dim gotActive As Long
    gotOrder = b.Debug_TabOrderTitles()
    gotActive = b.Debug_ActiveIndex()
    If gotOrder = expectOrder And gotActive = expectActive Then
        pass = pass + 1
        Debug.Print "[OK] " & label & " -> 順=" & gotOrder & " active=" & gotActive
    Else
        Debug.Print "[NG] " & label & vbCrLf & _
                    "      got     : 順=" & gotOrder & " active=" & gotActive & vbCrLf & _
                    "      expected: 順=" & expectOrder & " active=" & expectActive
    End If
End Sub

' ============================================================
' Test_9_16_SearchEngine  (第9.16 検索エンジン切替のロジック検証)
'
'   ★イミディエイトウィンドウで  Test_9_16_SearchEngine  と打つだけ★
'
'   WebView2 は起動しない。Wv2Browser を new し、SearchEngine (Property Let) で
'   検索テンプレートを切り替えながら、検索語入力が各エンジンの検索 URL に正しく
'   変換されるか (Debug_NormalizeUrl の第3分岐) を Check で照合する。
'   Debug_SearchEngine で現在のテンプレートも確認する。末尾に合否サマリ。
'
'   ★確認する軸★
'     (1) 既定 (無設定) は Google のまま (9.14 と同一挙動)。
'     (2) Bing / DuckDuckGo に切り替えると検索 URL がそのエンジンに追従する。
'     (3) 空文字を Let すると既定 (Google) にフォールバックする (空ガード)。
'     (4) エンジン切替後も日本語エンコード (UTF-8 percent-encoding) が正しい。
'     (5) URL / スキーム入力は分岐 (1)(2) に落ちるのでエンジン設定の影響を受けない。
' ============================================================
Public Sub Test_9_16_SearchEngine()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' ※ Init 不要 (Debug_NormalizeUrl / SearchEngine は NormalizeUrl と
    '    m_searchTemplate に触れるだけで、タブ・Environment に一切触れないため)。

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_16_SearchEngine 開始 ===="

    ' --- (1) 既定は Google (Class_Initialize でセット済み) ---
    Debug.Print "  現在のエンジン = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, "日本 天気", "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- (2) Bing に切り替え ---
    b.SearchEngine = "https://www.bing.com/search?q="
    Debug.Print "  現在のエンジン = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.bing.com/search?q=hello%20world", total, pass
    Check b, "excel vba", "https://www.bing.com/search?q=excel%20vba", total, pass
    ' 切替後も日本語エンコードが正しいこと
    Check b, "日本 天気", "https://www.bing.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass
    ' Debug_SearchEngine が現在値を返すこと
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.bing.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] Debug_SearchEngine = Bing テンプレート"
    Else
        Debug.Print "[NG] Debug_SearchEngine = " & b.Debug_SearchEngine()
    End If

    ' --- (3) DuckDuckGo に切り替え ---
    b.SearchEngine = "https://duckduckgo.com/?q="
    Check b, "privacy", "https://duckduckgo.com/?q=privacy", total, pass

    ' --- (4) 空文字を Let → 既定 (Google) にフォールバック (空ガード) ---
    b.SearchEngine = ""
    Debug.Print "  空 Let 後のエンジン = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.google.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] 空 Let で Google にフォールバック"
    Else
        Debug.Print "[NG] 空 Let 後 = " & b.Debug_SearchEngine()
    End If

    ' --- (5) 空白のみを Let → これも空ガードで Google に戻る ---
    b.SearchEngine = "https://www.bing.com/search?q="   ' 一旦 Bing に
    b.SearchEngine = "   "                              ' 空白のみ → Google に戻るはず
    total = total + 1
    If b.Debug_SearchEngine() = "https://www.google.com/search?q=" Then
        pass = pass + 1
        Debug.Print "[OK] 空白のみ Let で Google にフォールバック"
    Else
        Debug.Print "[NG] 空白のみ Let 後 = " & b.Debug_SearchEngine()
    End If

    ' --- (6) エンジン設定は URL / スキーム入力に影響しない (第1・第2分岐) ---
    b.SearchEngine = "https://duckduckgo.com/?q="   ' 検索は DDG に設定
    Check b, "example.com", "https://example.com", total, pass          ' ドット有り→URL (影響なし)
    Check b, "https://example.com", "https://example.com", total, pass  ' スキーム有り→素通し (影響なし)
    Check b, "about:blank", "about:blank", total, pass                  ' スキーム有り→素通し (影響なし)

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_16_SearchEngine_Help  (第9.16 検証手順のヘルプ)
'
'   ★イミディエイトウィンドウで  Test_9_16_SearchEngine_Help  と打つと手順が出る★
' ============================================================
Public Sub Test_9_16_SearchEngine_Help()
    Debug.Print "==== 第9.16 検索エンジン切替 検証手順 ===="
    Debug.Print ""
    Debug.Print "【A. ロジック検証 (WebView2 起動不要)】"
    Debug.Print "  イミディエイトに次を打つだけ:"
    Debug.Print "    Test_9_16_SearchEngine"
    Debug.Print "  → 既定Google / Bing / DuckDuckGo 切替、空ガード、日本語エンコード、"
    Debug.Print "     URL入力への非影響 を一括照合。末尾 [ALL OK] が出れば合格。"
    Debug.Print ""
    Debug.Print "【B. 実機 GUI 検証 (WebView2 起動)】"
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) アドレスバーに  日本 天気  と打って Enter。"
    Debug.Print "     → 既定のまま Google 検索結果が出ること。"
    Debug.Print "  3) イミディエイトで検索エンジンを Bing に切り替える:"
    Debug.Print "       Set gBrowser = <起動中の Wv2Browser 参照>   ' 起動 Sub が公開する参照"
    Debug.Print "       gBrowser.SearchEngine = ""https://www.bing.com/search?q="""
    Debug.Print "     ※ 起動中の Browser 参照の取り出し方は起動 Sub の実装による。"
    Debug.Print "       参照を持っていない場合は BuildSearchUrl の既定値 (DEFAULT_"
    Debug.Print "       SEARCH_TEMPLATE) を書き換えて再起動する方法でも確認できる。"
    Debug.Print "  4) 再び  日本 天気  と打って Enter。"
    Debug.Print "     → 今度は Bing の検索結果が出れば実機でも切替成功。"
    Debug.Print ""
    Debug.Print "  ※ B の 3) でグローバル参照の口が無い場合は、A のロジック検証で"
    Debug.Print "     切替が正しいことは担保できる。実機は既定 Google の動作確認でよい。"
End Sub

' ============================================================
' Test_9_17_SearchPresets  (第9.17 検索エンジンのプリセット解決を検証)
'
'   ★イミディエイトウィンドウで  Test_9_17_SearchPresets  と打つだけ★
'
'   WebView2 は起動しない。UserForm1.UseSearchEngine が名前を解決する先の
'   プリセットテンプレート (SE_GOOGLE 等と同一文字列) を Wv2Browser に直接 Let し、
'   検索語入力がそのエンジンの検索 URL に変換されるか (Debug_NormalizeUrl の第3
'   分岐) と、Debug_SearchEngine が設定値を返すかを Check で照合する。
'
'   ★ここで担保するもの★ 各プリセットの URL 前置テンプレートが正しく、そのまま
'   Let すれば正しい検索 URL が生成されること。UseSearchEngine の名前解決
'   ("bing" → SE_BING) と未起動ガードは実機 (CurrentBrowser 経由) で確認する。
'
'   ※プリセット文字列は UserForm1 の Private Const と重複するが、UserForm を
'   起動せずに検証するための意図的な複製 (テスト側に期待値を直書きするのと同じ)。
' ============================================================
Public Sub Test_9_17_SearchPresets()
    Dim b As Wv2Browser
    Set b = New Wv2Browser
    ' ※ Init 不要 (Test_9_16 と同じ。NormalizeUrl / m_searchTemplate にしか触れない)

    Dim total As Long, pass As Long
    total = 0: pass = 0

    Debug.Print "==== Test_9_17_SearchPresets 開始 ===="

    ' --- google プリセット ---
    b.SearchEngine = "https://www.google.com/search?q="
    Debug.Print "  google  = " & b.Debug_SearchEngine()
    Check b, "hello world", "https://www.google.com/search?q=hello%20world", total, pass
    Check b, "日本 天気", "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- bing プリセット ---
    b.SearchEngine = "https://www.bing.com/search?q="
    Debug.Print "  bing    = " & b.Debug_SearchEngine()
    Check b, "excel vba", "https://www.bing.com/search?q=excel%20vba", total, pass
    Check b, "日本 天気", "https://www.bing.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- duckduckgo プリセット (パスが /?q= 型でも正しく連結されること) ---
    b.SearchEngine = "https://duckduckgo.com/?q="
    Debug.Print "  ddg     = " & b.Debug_SearchEngine()
    Check b, "privacy", "https://duckduckgo.com/?q=privacy", total, pass
    Check b, "日本 天気", "https://duckduckgo.com/?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- yahoo プリセット (?p= 型のクエリキーでも前置テンプレートで表現できること) ---
    b.SearchEngine = "https://search.yahoo.com/search?p="
    Debug.Print "  yahoo   = " & b.Debug_SearchEngine()
    Check b, "excel vba", "https://search.yahoo.com/search?p=excel%20vba", total, pass

    ' --- プリセット切替は URL / スキーム入力に影響しない (第1・第2分岐) ---
    b.SearchEngine = "https://search.yahoo.com/search?p="   ' 検索は Yahoo に設定
    Check b, "example.com", "https://example.com", total, pass          ' ドット有り→URL
    Check b, "https://example.com", "https://example.com", total, pass  ' スキーム有り→素通し

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_17_SearchPresets_Help  (第9.17 検証手順のヘルプ)
'
'   ★イミディエイトウィンドウで  Test_9_17_SearchPresets_Help  と打つと手順が出る★
' ============================================================
Public Sub Test_9_17_SearchPresets_Help()
    Debug.Print "==== 第9.17 検索エンジン切替口 + CurrentBrowser 公開 検証手順 ===="
    Debug.Print ""
    Debug.Print "【A. ロジック検証 (WebView2 起動不要)】"
    Debug.Print "  イミディエイトに次を打つだけ:"
    Debug.Print "    Test_9_17_SearchPresets"
    Debug.Print "  → google / bing / duckduckgo / yahoo の各プリセットで検索 URL が"
    Debug.Print "     正しく生成され、日本語エンコードも保たれること、URL 入力に影響"
    Debug.Print "     しないことを一括照合。末尾 [ALL OK] が出れば合格。"
    Debug.Print ""
    Debug.Print "【B. 実機 GUI 検証 (WebView2 起動)】 ★9.17 で参照の口ができた★"
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) アドレスバーに  日本 天気  と打って Enter。"
    Debug.Print "     → 既定のまま Google 検索結果が出ること。"
    Debug.Print "  3) イミディエイトで検索エンジンを Bing に切り替える (どちらでも可):"
    Debug.Print "       方法1 (プリセット名):  UserForm1.UseSearchEngine ""bing"""
    Debug.Print "       方法2 (参照を掴む)  :  Set b = UserForm1.CurrentBrowser"
    Debug.Print "                              b.SearchEngine = ""https://www.bing.com/search?q="""
    Debug.Print "  4) 再び  日本 天気  と打って Enter。"
    Debug.Print "     → 今度は Bing の検索結果が出れば実機でも切替成功。"
    Debug.Print "  5) 名前解決の確認:  UserForm1.UseSearchEngine ""duckduckgo""  → DDG、"
    Debug.Print "       UserForm1.UseSearchEngine ""nonsense""  → 警告が出て Google に戻る。"
    Debug.Print "  6) 未起動ガード:  StartWebView2_Full を実行する前に"
    Debug.Print "       UserForm1.UseSearchEngine ""bing""  → 『Browser が未起動』の注意が出る。"
End Sub


' ============================================================
' Test_9_18_BridgeLogic  (第9.18 Wv2SettingsBridge の純ロジック検証)
'
'   ★イミディエイトウィンドウで  Test_9_18_BridgeLogic  と打つだけ★
'
'   WebView2 は起動しない。Wv2SettingsBridge を new し、実 Wv2Browser を
'   BindBrowser で結び付けてから、設定画面の JS が呼ぶのと同じメソッド
'   (SetEngine / GetEngine / PreviewUrl / PreviewUrlFor / EngineList) を
'   直接叩いて、名前?テンプレートの解決とプレビュー URL 生成を照合する。
'
'   ★これは何を保証するか★ 設定画面の JS は hostObjects.sync.Settings 経由で
'     これらのメソッドを呼ぶだけなので、メソッドの戻り値がここで正しければ、
'     残る不確実性は「hostObjects 経路が実機で通るか (順序A)」だけに絞られる。
'     その実機疎通は Test_9_18_BridgeLogic_Help の手順で確認する。
' ============================================================
Public Sub Test_9_18_BridgeLogic()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_18_BridgeLogic (Wv2SettingsBridge 純ロジック) ===="

    Dim b As Wv2Browser
    Set b = New Wv2Browser          ' 既定 Google で初期化される

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge
    br.BindBrowser b

    ' --- 初期状態: 既定は google ---
    CheckEq "GetEngine 初期", br.GetEngine(), "google", total, pass

    ' --- SetEngine の戻り値と GetEngine の追従 ---
    CheckEq "SetEngine bing 戻り値", br.SetEngine("bing"), "bing", total, pass
    CheckEq "GetEngine after bing", br.GetEngine(), "bing", total, pass

    CheckEq "SetEngine DDG 別名 ddg", br.SetEngine("ddg"), "duckduckgo", total, pass
    CheckEq "GetEngine after ddg", br.GetEngine(), "duckduckgo", total, pass

    ' --- 大小・前後空白の正規化 ---
    CheckEq "SetEngine '  YAHOO '", br.SetEngine("  YAHOO "), "yahoo", total, pass

    ' --- 未知名は Google に落ちる ---
    CheckEq "SetEngine 未知名", br.SetEngine("nonsense"), "google", total, pass
    CheckEq "GetEngine after 未知名", br.GetEngine(), "google", total, pass

    ' --- PreviewUrl は「いまのエンジン (google)」で生成、日本語エンコード ---
    CheckEq "PreviewUrl (google)", br.PreviewUrl("日本 天気"), _
            "https://www.google.com/search?q=%E6%97%A5%E6%9C%AC%20%E5%A4%A9%E6%B0%97", total, pass

    ' --- PreviewUrlFor は「指定エンジン」で生成 (現在エンジンを変えない) ---
    CheckEq "PreviewUrlFor bing", br.PreviewUrlFor("bing", "excel vba"), _
            "https://www.bing.com/search?q=excel%20vba", total, pass
    CheckEq "PreviewUrlFor yahoo", br.PreviewUrlFor("yahoo", "privacy"), _
            "https://search.yahoo.com/search?p=privacy", total, pass
    ' PreviewUrlFor 後も現在エンジンは google のまま (副作用がないこと)
    CheckEq "PreviewUrlFor 後の現在", br.GetEngine(), "google", total, pass

    ' --- PreviewUrl は URL 入力を検索にしない (第1・第2分岐は素通し) ---
    CheckEq "PreviewUrl URL入力", br.PreviewUrl("example.com"), "https://example.com", total, pass

    ' --- EngineList が 4 エンジンを含む JSON を返す ---
    Dim js As String
    js = br.EngineList()
    CheckBool "EngineList に google", (InStr(js, """id"":""google""") > 0), total, pass
    CheckBool "EngineList に bing", (InStr(js, """id"":""bing""") > 0), total, pass
    CheckBool "EngineList に duckduckgo", (InStr(js, """id"":""duckduckgo""") > 0), total, pass
    CheckBool "EngineList に yahoo", (InStr(js, """id"":""yahoo""") > 0), total, pass

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' --- 第9.18: 文字列一致チェック (汎用) ---
Private Sub CheckEq(ByVal label As String, ByVal got As String, _
                    ByVal expected As String, ByRef total As Long, ByRef pass As Long)
    total = total + 1
    If got = expected Then
        pass = pass + 1
        Debug.Print "[OK] " & label & " -> " & got
    Else
        Debug.Print "[NG] " & label & vbCrLf & _
                    "      got     = " & got & vbCrLf & _
                    "      expected= " & expected
    End If
End Sub


' --- 第9.18: 真偽チェック (汎用) ---
Private Sub CheckBool(ByVal label As String, ByVal cond As Boolean, _
                      ByRef total As Long, ByRef pass As Long)
    total = total + 1
    If cond Then
        pass = pass + 1
        Debug.Print "[OK] " & label
    Else
        Debug.Print "[NG] " & label & "  (条件が False)"
    End If
End Sub


' ============================================================
' Test_9_18_BridgeLogic_Help  (第9.18 検証手順のヘルプ)
'
'   ★イミディエイトウィンドウで  Test_9_18_BridgeLogic_Help  と打つと手順が出る★
' ============================================================
Public Sub Test_9_18_BridgeLogic_Help()
    Debug.Print "==== 第9.18 設定タブ + AddHostObjectToScript 検証手順 ===="
    Debug.Print ""
    Debug.Print "【A. ブリッジ純ロジック検証 (WebView2 起動不要)】"
    Debug.Print "  イミディエイトに次を打つだけ:"
    Debug.Print "    Test_9_18_BridgeLogic"
    Debug.Print "  → Wv2SettingsBridge の SetEngine/GetEngine/PreviewUrl(For)/EngineList を"
    Debug.Print "     直接叩き、名前解決・プレビュー URL・副作用なしを一括照合。"
    Debug.Print "     末尾 [ALL OK] が出れば、残る不確実性は hostObjects 実機疎通だけ。"
    Debug.Print ""
    Debug.Print "【B. 実機 hostObjects 疎通 + 設定画面 (WebView2 起動)】 ★9.18 の本命★"
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) タブバー右端の ? (歯車) ボタンをクリックする。"
    Debug.Print "     → 新しいタブが開き、検索エンジン選択のカード画面が出れば設定タブ生成 OK。"
    Debug.Print "  3) ★順序A の実機判定★ カードが 4 枚 (Google/Bing/DuckDuckGo/Yahoo) 描かれ、"
    Debug.Print "     現在エンジン (既定 Google) のカードが青く選択強調されていること。"
    Debug.Print "     → これが出れば『Navigate 前 attach で JS 読み込み時に hostObjects が"
    Debug.Print "        使える (順序A)』が実機で成立。カード枠に『ホストオブジェクト未接続』"
    Debug.Print "        と出たら順序A 失敗 → 順序B (ready ハンドシェイク) へ切替を検討。"
    Debug.Print "  4) カードにマウスを乗せる → 下の Preview 欄に『日本 天気 → 生成URL』が"
    Debug.Print "     ライブ表示されること (PreviewUrlFor の同期呼び出し)。"
    Debug.Print "  5) Bing のカードをクリック → Bing カードが選択強調に変わること。"
    Debug.Print "  6) 設定タブとは別の通常タブに移り、アドレスバーに  日本 天気  と打つ。"
    Debug.Print "     → Bing の検索結果が出れば、設定画面での切替が本番に反映されている。"
    Debug.Print "  7) 設定タブは × で普通に閉じられること (通常タブとして振る舞う)。"
    Debug.Print ""
    Debug.Print "  ※ 通常タブが無い状態で ? を押しても設定タブは開く (最後の1タブでも可)。"
End Sub


' ============================================================
' Test_9_19_SettingsTabDedup_Help  (第9.19 設定タブ重複防止の検証手順)
'
'   ★イミディエイトウィンドウで  Test_9_19_SettingsTabDedup_Help  と打つと手順が出る★
'
'   ロジックの要点 (コードレビューで担保):
'     ・OpenSettingsTab の冒頭で m_tabs を走査し、IsSettings=True の Pane が
'       あればそれを ActivateTab して Exit Sub (新規生成しない)。
'     ・生成時は pane.IsSettings = True を立てる。
'     ・閉じられた設定タブは m_tabs から外れるので走査で見つからず、次の ? で
'       新規に開き直せる (生存確認が走査と一体)。
' ============================================================
Public Sub Test_9_19_SettingsTabDedup_Help()
    Debug.Print "==== 第9.19 設定タブ重複防止 検証手順 (WebView2 起動) ===="
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) タブバー右端の ? (歯車) を 1 回押す。"
    Debug.Print "     → 設定タブが 1 枚開き、アクティブになること。"
    Debug.Print "  3) ★重複防止の本命★ 続けて ? をもう 2?3 回連打する。"
    Debug.Print "     → 設定タブが増えず、既存の設定タブがアクティブになるだけであること。"
    Debug.Print "       (イミディエイトに『既存の設定タブ(index N)をアクティブ化』が出る)"
    Debug.Print "  4) 別の通常タブをクリックして設定タブから離れる。"
    Debug.Print "     → その状態で ? を押すと、既存の設定タブへ切り替わる (新規は開かない)。"
    Debug.Print "  5) ★閉じてから開き直し★ 設定タブの × を押して閉じる。"
    Debug.Print "     → その後もう一度 ? を押すと、設定タブが新規に 1 枚開くこと。"
    Debug.Print "       (閉じたら m_tabs から外れるので、走査で見つからず開き直せる)"
    Debug.Print "  6) 設定画面のカード操作 (エンジン選択・hover プレビュー・切替の本番反映) が"
    Debug.Print "     9.18 と同じく動くこと (重複防止で設定機能が壊れていないことの確認)。"
    Debug.Print ""
    Debug.Print "  ※ 期待挙動まとめ: 設定タブは常に高々 1 枚。? は「無ければ開く/あれば移動」。"
End Sub


' ============================================================
' Test_9_20_Persistence  (第9.20 検索エンジン永続化の純ロジック検証)
'
'   ★イミディエイトで  Test_9_20_Persistence  と打つだけ★
'
'   Wv2SettingsBridge の保存/読込を直接叩く。WebView2 は起動しない。
'   実ファイルを触るため、先頭で現在の保存値を退避し、末尾で
'   書き戻してユーザーの実設定を壊さない。
' ============================================================
Public Sub Test_9_20_Persistence()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_20_Persistence (永続化 純ロジック) ===="

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge

    ' --- 現在の保存値を退避 (テスト後に戻す) ---
    '   ★第9.29: engine だけでなく template も対で退避する★ 元が custom だった
    '     場合に engine 行だけ書き戻すと template 行が消えてしまうため。
    Dim savedBefore As String, savedTplBefore As String
    savedBefore = br.LoadEngineName()
    savedTplBefore = br.LoadCustomTemplate()
    Debug.Print "  (退避) 現在の保存値 = '" & savedBefore & "'" & _
                IIf(Len(savedTplBefore) > 0, " / template = '" & savedTplBefore & "'", "")

    ' --- 保存→読込の往復 (bing) ---
    br.Debug_SaveEngineName "bing"
    CheckEq "保存 bing → LoadEngineName", br.LoadEngineName(), "bing", total, pass
    CheckEq "保存 bing → LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://www.bing.com/search?q=", total, pass

    ' --- 別エンジンで上書き (duckduckgo) ---
    br.Debug_SaveEngineName "duckduckgo"
    CheckEq "上書き ddg → LoadEngineName", br.LoadEngineName(), "duckduckgo", total, pass
    CheckEq "上書き ddg → LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://duckduckgo.com/?q=", total, pass

    ' --- 未知名を保存したら LoadEngineTemplate は "" に落ちる ---
    '   (LoadEngineName は生値を返すが、テンプレート解決で "" になる)
    br.Debug_SaveEngineName "nonsense"
    CheckEq "未知名 → LoadEngineName (生値)", br.LoadEngineName(), "nonsense", total, pass
    CheckEq "未知名 → LoadEngineTemplate (空)", br.LoadEngineTemplate(), "", total, pass

    ' --- 退避値を書き戻す (ユーザー設定の復元) ---
    If Len(savedBefore) > 0 Then
        br.Debug_SaveSettings savedBefore, savedTplBefore
        Debug.Print "  (復元) 保存値を '" & savedBefore & "' に戻した。"
    Else
        Debug.Print "  (復元) 元々保存は空だったので、ファイルにはテスト最終値"
        Debug.Print "         (nonsense) が残るが、LoadEngineTemplate は '' に落ちるので"
        Debug.Print "         次回起動は既定 (Google) で始まり実害なし。気になれば"
        Debug.Print "         設定画面で任意のエンジンを選べば上書きされる。"
    End If

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_20_Persistence_Help  (第9.20 永続化の実機検証手順)
'
'   ★イミディエイトで  Test_9_20_Persistence_Help  と打つと手順が出る★
'
'   復元は Wv2Browser.Class_Initialize の実行を伴うため、純ロジックでは
'   なく実機手順で確認する (Excel を完全に終了→再起動が鍵)。
' ============================================================
Public Sub Test_9_20_Persistence_Help()
    Debug.Print "==== 第9.20 検索エンジン選択の永続化 実機検証手順 ===="
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) ? を押して設定タブを開き、Bing のカードをクリックする。"
    Debug.Print "     → イミディエイトに次の 2 行が出ること:"
    Debug.Print "         Wv2SettingsBridge.SetEngine: 'bing' → 'bing'"
    Debug.Print "         Wv2SettingsBridge.SaveSettingsFile: 保存 'bing' -> ...\settings.txt"
    Debug.Print "  3) 別の通常タブでアドレスバーに   日本 天気   と打ち、Bing の検索結果が"
    Debug.Print "     出ることを確認 (ここまでは 9.18 と同じ)。"
    Debug.Print ""
    Debug.Print "  --- ★ここからが 9.20 の本命★ ---"
    Debug.Print "  4) Excel を 完全に終了 する (ブックを閉じるだけでなく Excel ごと)。"
    Debug.Print "  5) Excel を再起動し、ブックを開いて StartWebView2_Full を実行する。"
    Debug.Print "     → イミディエイトに次の 1 行が出ること (復元の証):"
    Debug.Print "         Wv2Browser.Class_Initialize: 保存済みエンジンを復元 = https://www.bing.com/search?q="
    Debug.Print "  6) 設定タブを開かずに、いきなり通常タブで   日本 天気   と打つ。"
    Debug.Print "     → 前回選んだ Bing の検索結果がすぐに出れば永続化成功。"
    Debug.Print "  7) (任意) 設定タブを開くと Bing カードが選択強調になっていることも確認。"
    Debug.Print ""
    Debug.Print "  ※ 保存先: %APPDATA%\Wv2Browser\settings.txt (中身は  engine=bing  の 1 行)。"
    Debug.Print "     メモ帳や手作業で中身を見たいときはエクスプローラーで開ける。"
    Debug.Print "  ※ ファイルを手で削除して再起動すれば、既定 (Google) に戻ることも確かめられる。"
End Sub


' ============================================================
' Test_9_21_HostActivate_Help  (第9.21 activate の hostObjects 化 実機手順)
'
'   ★イミディエイトで  Test_9_21_HostActivate_Help  と打つと手順が出る★
'
'   実 URL (appassets.tabbar) への View_Navigate でも仕様事実 18 が効くか
'   (= Navigate 前 attach で JS 初回から hostObjects を掴めるか) を実機で見る。
' ============================================================
Public Sub Test_9_21_HostActivate_Help()
    Debug.Print "==== 第9.21 activate の hostObjects 化 実機検証手順 ===="
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "     → 起動ログに次の 1 行が出ること (attach 成功の証):"
    Debug.Print "         Wv2TabBar.Init: TabBar ブリッジを attach OK (activate は hostObjects 優先)"
    Debug.Print "       ※ 代わりに『AddHostObjectToScript(TabBar) 失敗』が出た場合は"
    Debug.Print "         attach が通っておらず、以降 activate は postMessage に落ちる。"
    Debug.Print "  2) ＋ ボタンでタブを 3?4 枚に増やす。"
    Debug.Print "  3) 別のタブ (今アクティブでないタブ) の本体をクリックして切替える。"
    Debug.Print ""
    Debug.Print "  --- ★ここが 9.21 の判定★ ---"
    Debug.Print "  4) クリックのたびにイミディエイトへ次の 1 行が出れば hostObjects 経路成立:"
    Debug.Print "         Wv2TabBar.HostActivate: index=N (host経由)"
    Debug.Print "     かつ、従来の postMessage 経路のログ"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""activate"",...}"
    Debug.Print "     が activate では出ない ことを確認する (close/newtab/reorder/settings"
    Debug.Print "     では従来どおり OnPaneWebMessage が出る = それらは postMessage のまま)。"
    Debug.Print "  5) 切替のたびにアクティブ強調とナビバー (URL 欄・戻る/進む) が"
    Debug.Print "     そのタブに追従することを確認 (全体同期は不変なので触り心地は同じはず)。"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (9.18/9.19/9.20 に影響がないこと) ---"
    Debug.Print "  6) × でタブを閉じる / ＋ で開く / ドラッグで並べ替える がすべて従来どおり動く。"
    Debug.Print "  7) ? で設定タブが開き、エンジンを変えると検索に反映され、"
    Debug.Print "     Excel 再起動後もそのエンジンが持ち越される (9.20 の永続化) ことを確認。"
    Debug.Print ""
    Debug.Print "  ※ ブラウザの DevTools コンソールを開ける場合は、タブクリック時に"
    Debug.Print "     [host] activate N が出る (postMessage 落ち時は [postmsg] activate N)。"
    Debug.Print "     VBA 側ログだけでも host経由 か postMessage経由 かは判別できる。"
End Sub


' ============================================================
' Test_9_22_HostCloseReorder_Help  (第9.22 close/reorder の hostObjects 化 実機手順)
'
'   ★イミディエイトで  Test_9_22_HostCloseReorder_Help  と打つと手順が出る★
'
'   activate に続き close / reorder も hostObjects 直呼びへ移した。実 URL
'   (appassets.tabbar) でも仕様事実 19 (Navigate 前 attach) の射程が close/reorder
'   に及ぶこと、および activate/ready/newtab/settings が無傷なことを実機で見る。
' ============================================================
Public Sub Test_9_22_HostCloseReorder_Help()
    Debug.Print "==== 第9.22 close / reorder の hostObjects 化 実機検証手順 ===="
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "     → 起動ログに 9.21 と同じ attach 成功行が出ること:"
    Debug.Print "         Wv2TabBar.Init: TabBar ブリッジを attach OK (activate は hostObjects 優先)"
    Debug.Print "       ※ attach は TabBar 全体 (activate/close/reorder) で共有する 1 回の登録。"
    Debug.Print "  2) ＋ ボタンでタブを 4?5 枚に増やす。"
    Debug.Print ""
    Debug.Print "  --- ★ここが 9.22 の判定 (close)★ ---"
    Debug.Print "  3) 適当なタブの × をクリックして閉じる。次の 1 行が出れば host経由成立:"
    Debug.Print "         Wv2TabBar.HostClose: index=N (host経由)"
    Debug.Print "     かつ、従来の postMessage 経路のログ"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""close"",...}"
    Debug.Print "     が close では出ない ことを確認する。閉じた後のアクティブ補正 (残った"
    Debug.Print "     タブへ強調が移る) とナビバー追従が従来どおりであることも確認。"
    Debug.Print ""
    Debug.Print "  --- ★ここが 9.22 の判定 (reorder)★ ---"
    Debug.Print "  4) タブをドラッグして並べ替える。次の 1 行が出れば host経由成立:"
    Debug.Print "         Wv2TabBar.HostReorder: from=A to=B (host経由)"
    Debug.Print "     かつ、従来の postMessage 経路のログ"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""reorder"",...}"
    Debug.Print "     が reorder では出ない ことを確認する。並べ替え後の順序で"
    Debug.Print "     タブバーが再描画され、掴んだタブがアクティブ強調のまま追従すること。"
    Debug.Print ""
    Debug.Print "  --- 9.21 の activate が引き続き host経由であること ---"
    Debug.Print "  5) 別タブの本体をクリックして切替える。従来どおり次が出ること:"
    Debug.Print "         Wv2TabBar.HostActivate: index=N (host経由)"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (ready/newtab/settings は postMessage のまま) ---"
    Debug.Print "  6) ＋ で新規タブ → OnPaneWebMessage: msg={""cmd"":""newtab""} が出る"
    Debug.Print "     (SafeTimer 経由の従来動作。host経由ではない=正しい)。"
    Debug.Print "  7) ? で設定タブ → OnPaneWebMessage: msg={""cmd"":""settings""} が出て"
    Debug.Print "     設定タブが開く。エンジン変更→検索反映→再起動持ち越し (9.20) も従来どおり。"
    Debug.Print "  8) 起動直後の初期同期で OnPaneWebMessage: msg={""cmd"":""ready""} が出る"
    Debug.Print "     (ready は通知なので postMessage のまま=正しい)。"
    Debug.Print ""
    Debug.Print "  ※ 期待: activate/close/reorder は host経由 (HostXxx ログ)、"
    Debug.Print "     ready/newtab/settings は postMessage経由 (OnPaneWebMessage ログ)。"
    Debug.Print "     この住み分けが 9.22 の狙い。DevTools が開けるなら × クリックで"
    Debug.Print "     [host] close N、ドラッグで [host] reorder A->B が出る。"
End Sub


' ============================================================
' Test_9_23_Consolidation_Help  (第9.23 activate/close/reorder の Host 一元化 実機手順)
'
'   ★イミディエイトで  Test_9_23_Consolidation_Help  と打つと手順が出る★
'
'   9.23 は「処理の置き場所」だけを変えた整理ステージ。activate/close/reorder の
'   ActivateTab/CloseTab/MoveTab 呼び出しと失敗ログを Host メソッド 1 箇所へ集約し、
'   OnPaneWebMessage の Case はフォールバック時に Host メソッドへ委譲する薄い
'   アダプタになった。機能は 9.22 と完全に同一のはずで、狙いは回帰ゼロの確認。
' ============================================================
Public Sub Test_9_23_Consolidation_Help()
    Debug.Print "==== 第9.23 activate/close/reorder の Host 一元化 実機検証手順 ===="
    Debug.Print ""
    Debug.Print "  【前提】9.23 は処理の置き場所を変えただけ。機能は 9.22 と同一。"
    Debug.Print "     通常運用では hostObjects が成立するため、Case activate/close/reorder"
    Debug.Print "     (=フォールバック経路) は通常は通らない。よって主眼は回帰ゼロ確認。"
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "     → 起動ログに 9.21/9.22 と同じ attach 成功行が出ること:"
    Debug.Print "         Wv2TabBar.Init: TabBar ブリッジを attach OK (activate は hostObjects 優先)"
    Debug.Print "  2) ＋ ボタンでタブを 4?5 枚に増やす。"
    Debug.Print ""
    Debug.Print "  --- ★回帰確認: 3 コマンドが従来どおり host 経由★ ---"
    Debug.Print "  3) 別タブの本体クリック → Wv2TabBar.HostActivate: index=N (host経由)"
    Debug.Print "  4) タブの × クリック    → Wv2TabBar.HostClose: index=N (host経由)"
    Debug.Print "  5) タブをドラッグ並べ替え → Wv2TabBar.HostReorder: from=A to=B (host経由)"
    Debug.Print "     いずれも従来どおり動き、UX (アクティブ補正/再描画/ナビバー追従) が"
    Debug.Print "     9.22 と一切変わらないこと。この 3 つで OnPaneWebMessage の"
    Debug.Print "     activate/close/reorder ログが出ない のも 9.22 と同じ (host 経由のため)。"
    Debug.Print ""
    Debug.Print "  --- 回帰確認: ready/newtab/settings は postMessage のまま ---"
    Debug.Print "  6) ＋ で新規タブ → OnPaneWebMessage: msg={""cmd"":""newtab""} (SafeTimer 経由)"
    Debug.Print "  7) ? で設定タブ → OnPaneWebMessage: msg={""cmd"":""settings""} → 設定タブが開く"
    Debug.Print "  8) 起動直後の初期同期 → OnPaneWebMessage: msg={""cmd"":""ready""}"
    Debug.Print ""
    Debug.Print "  --- (任意) フォールバック経路の生存確認 ---"
    Debug.Print "  9) 通常環境ではフォールバックは踏めない。もし DevTools で"
    Debug.Print "     window.chrome.webview.hostObjects を一時的に無効化できる場合のみ、"
    Debug.Print "     × クリックで次の 2 行が続けて出れば委譲が生きている:"
    Debug.Print "         Wv2TabBar.OnPaneWebMessage: msg={""cmd"":""close"",...}"
    Debug.Print "         → close(index=N) [postmsg fallback → HostClose]"
    Debug.Print "         Wv2TabBar.HostClose: index=N (host経由)   ※Host に委譲された証跡"
    Debug.Print "     (無効化できない環境ではこの手順は省略してよい。9.23 の本質は回帰ゼロ。)"
    Debug.Print ""
    Debug.Print "  ※ 期待: 見た目・触り心地は 9.22 と完全同一。ログの住み分け"
    Debug.Print "     (activate/close/reorder=HostXxx、ready/newtab/settings=OnPaneWebMessage)"
    Debug.Print "     も 9.22 と同一。何も変わって見えなければ 9.23 は成功。"
End Sub



' ============================================================
' Test_9_24_HostNavBar_Help  (第9.24 NavBar の hostObjects 化 実機手順)
'
'   ★イミディエイトで  Test_9_24_HostNavBar_Help  と打つと手順が出る★
'
'   9.24 は TabBar (9.21?9.23) で確立した型を NavBar へ同型展開した回。
'   JS はハイブリッド (hostObjects 優先・postMessage フォールバック)、VBA は
'   処理を Host メソッド 4 本へ一元化、Select Case は薄いアダプタ。
' ============================================================
Public Sub Test_9_24_HostNavBar_Help()
    Debug.Print "==== 第9.24 NavBar の hostObjects 化 実機検証手順 ===="
    Debug.Print ""
    Debug.Print "  【前提】StartWebView2_Full で通常起動する。"
    Debug.Print "     → 起動ログに TabBar/NavBar 両方の attach 行が出ること:"
    Debug.Print "         Wv2TabBar.Init: TabBar ブリッジを attach OK (...)"
    Debug.Print "         Wv2NavBar.Init: NavBar ブリッジを attach OK (back/forward/reload/navigate は hostObjects 優先)"
    Debug.Print "       ★NavBar 側の行が出なければ以降は全部 postMessage 経路になる"
    Debug.Print "         (機能は動くが 9.24 の狙いは未達)。hr をメモして報告のこと。"
    Debug.Print ""
    Debug.Print "  --- ★判定1: navigate が host 経由か (文字列引数の初成立)★ ---"
    Debug.Print "  1) URL 欄に  example.com  と打って Enter。次の 1 行が出れば成立:"
    Debug.Print "         Wv2NavBar.HostNavigate: url=example.com (host経由)"
    Debug.Print "     かつ、従来の postMessage 経路のログ"
    Debug.Print "         Wv2NavBar.OnPaneWebMessage: msg={""cmd"":""navigate"",...}"
    Debug.Print "     が出ない こと。スキーム補完 (https://) が効いて遷移することも確認。"
    Debug.Print ""
    Debug.Print "  --- ★判定2: back/forward/reload が host 経由か + 二重実行チェック★ ---"
    Debug.Print "  2) 同じタブでもう 1 ページ遷移してから ← を『1 回だけ』クリック。"
    Debug.Print "     期待は次の 1 行『だけ』:"
    Debug.Print "         Wv2NavBar.HostBack: (host経由)"
    Debug.Print "     ★★ここが今回いちばん大事★★ 次のどれかが起きたら二重実行:"
    Debug.Print "         ・HostBack の行が 2 行出る"
    Debug.Print "         ・HostBack の後に OnPaneWebMessage: msg={""cmd"":""back""} が続く"
    Debug.Print "         ・1 回のクリックで 2 ページ分戻ってしまう"
    Debug.Print "       (ダミー引数 reserved で塞いだつもりの現象。出たら即報告のこと)"
    Debug.Print "  3) → を 1 回クリック → Wv2NavBar.HostForward: (host経由) が 1 行だけ。"
    Debug.Print "  4) 再読み込みボタンを 1 回クリック → Wv2NavBar.HostReload: (host経由) が"
    Debug.Print "     1 行だけ。ページが 1 回だけリロードされること。"
    Debug.Print ""
    Debug.Print "  --- ★判定3: 文字列引数が無傷か (仕様事実 21 候補)★ ---"
    Debug.Print "  5) URL 欄に日本語を打って Enter (例: 北海道 天気)。"
    Debug.Print "     → HostNavigate のログに日本語がそのまま出ること (文字化け・欠落なし)。"
    Debug.Print "         Wv2NavBar.HostNavigate: url=北海道 天気 (host経由)"
    Debug.Print "       設定中の検索エンジンで検索されること (9.14/9.16?9.20 の経路)。"
    Debug.Print "  6) 記号入り URL も試す (例: https://www.google.com/search?q=a&b=c#x )。"
    Debug.Print "     → ? & # がログでそのまま出て、そのページへ遷移すること。"
    Debug.Print "     ※ postMessage 経路では JSON エスケープを通っていた箇所。hostObjects"
    Debug.Print "        では BSTR 直渡しになるので、むしろ安全になっているはず。"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (触っていない経路) ---"
    Debug.Print "  7) 起動直後に OnPaneWebMessage: msg={""cmd"":""ready""} が出て URL 欄と"
    Debug.Print "     ボタン活性が初期同期されること (ready は postMessage のまま=正しい)。"
    Debug.Print "  8) URL 欄に何か打ってから Escape → 打った文字が消えて元の URL に戻る"
    Debug.Print "     (lastUrl からの即復元。9.12d の挙動が無傷か)。"
    Debug.Print "  9) 編集中に裏で遷移が起きても入力中の文字列が消えないこと (editing フラグ)。"
    Debug.Print " 10) タブ切替 / × / ドラッグ並べ替えが 9.23 のまま動くこと"
    Debug.Print "     (HostActivate / HostClose / HostReorder のログ)。ナビバーの URL 欄が"
    Debug.Print "     タブ切替に追従すること (ActiveChanged → PushNavSyncToJs)。"
    Debug.Print ""
    Debug.Print "  --- (任意) フォールバック経路の生存確認 ---"
    Debug.Print " 11) 通常環境では踏めない。DevTools で hostObjects を一時無効化できる場合のみ、"
    Debug.Print "     ← クリックで次の 2 行が続けて出れば委譲が生きている:"
    Debug.Print "         Wv2NavBar.OnPaneWebMessage: msg={""cmd"":""back""}"
    Debug.Print "         → back [postmsg fallback → HostBack]"
    Debug.Print "         Wv2NavBar.HostBack: (host経由)   ※Host に委譲された証跡"
    Debug.Print ""
    Debug.Print "  ※ 期待: 見た目・触り心地は 9.23 と完全同一。違いはログの住み分けだけ"
    Debug.Print "     (back/forward/reload/navigate=HostXxx、ready=OnPaneWebMessage)。"
End Sub
' ============================================================
' Test_9_26_Popup_On / Test_9_26_Popup_Off (第9.26、実機スイッチ)
'
'   起動中の Browser の委譲モードを切り替える。StartWebView2_Full の後に
'   イミディエイトで Sub 名を打つだけでよい。
'   ・On  : NewWindowRequested をランタイムへ委譲 (ポップアップで開く)
'   ・Off : 現行動作 (put_Handled + 自前の新タブ)
'   タブを開き直す必要はない。次に踏んだリンクから即座に効く。
' ============================================================
Public Sub Test_9_26_Popup_On()
    SetPopupDelegate True
End Sub

Public Sub Test_9_26_Popup_Off()
    SetPopupDelegate False
End Sub

Private Sub SetPopupDelegate(ByVal onOff As Boolean)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.26] Browser が起動していません。先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If
    b.PopupDelegateMode = onOff
    Debug.Print "[9.26] 現在の PopupDelegateMode = " & b.PopupDelegateMode
End Sub


' ============================================================
' Test_9_26_Popup_Help (第9.26、実機手順)
'
'   POST リンクの新タブ展開が 405 になる件について、「ランタイムに委譲すれば
'   POST は保たれるのか」だけを確かめる回。実装の変更は分岐 1 個で、恒久策では
'   ないことに注意 (ポップアップは自前タブ UI の外に出る)。
' ============================================================
Public Sub Test_9_26_Popup_Help()
    Debug.Print "==== 第9.26 実機手順 (NewWindowRequested 委譲プローブ) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    put_Handled を立てずに WebView2 ランタイムへ委譲すれば、"
    Debug.Print "    元の POST (メソッド + ボディ + Referer) が保たれて 405 が消えるのか。"
    Debug.Print "    → 消えるなら、次段の本丸 (args.put_NewWindow で別タブの CoreWebView2 を"
    Debug.Print "       渡す方式) が効くことがほぼ確定する。消えないなら原因は POST ボディ"
    Debug.Print "       以外 (Referer/Cookie/セッション) なので、本丸の前に方向転換する。"
    Debug.Print ""
    Debug.Print "  --- 手順 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "  2) 基幹システムにログインする (この時点ではまだ現行動作 = 委譲 OFF)"
    Debug.Print "  3) ★対照★ 委譲 OFF のまま『アポイント管理』をクリック"
    Debug.Print "       → 期待: 自前の新タブが開き、HTTP 405 が表示される (現状の再現)"
    Debug.Print "       → ログ: Wv2Pane.View_OnNewWindowRequested → put_Handled OK"
    Debug.Print "                → 新タブ生成を Timer に依頼 (RequestNewTabAsync)"
    Debug.Print "  4) 405 のタブを閉じ、元の画面に戻る"
    Debug.Print "  5) Test_9_26_Popup_On            ' 委譲モードへ"
    Debug.Print "  6) もう一度『アポイント管理』をクリック"
    Debug.Print "       → ログ: [9.26 委譲モード] put_Handled を立てずにランタイムへ委譲する"
    Debug.Print "          ※ put_Handled OK / RequestNewTabAsync の行は出ない (出たら報告)"
    Debug.Print "  7) 開いたポップアップウィンドウの中身を確認する"
    Debug.Print "  8) 確認できたらポップアップを × で閉じ、Test_9_26_Popup_Off で元に戻す"
    Debug.Print ""
    Debug.Print "  --- 結果の読み方 ---"
    Debug.Print "  ★パターンA: ポップアップに業務画面が正常に出た★"
    Debug.Print "      = ランタイム委譲で POST は保たれる。原因の見立ては正しかった。"
    Debug.Print "        次段は本丸 (put_NewWindow) へ。UX (自前タブに収める) の設計に集中できる。"
    Debug.Print "  ★パターンB: ポップアップでも 405 のまま★"
    Debug.Print "      = POST ボディ以外の要因 (Referer / Cookie / セッション / ユーザーエージェント)"
    Debug.Print "        を疑う段に移る。本丸を作っても解決しない可能性が高いので、着手前に判明して"
    Debug.Print "        得をした状態。405 ページの本文 (Tomcat のメッセージ全文) を控えること。"
    Debug.Print "  ★パターンC: ポップアップが開かない / 別の症状★"
    Debug.Print "      = ログをそのまま共有してください (put_Handled の行が出ていないかが分かれ目)。"
    Debug.Print ""
    Debug.Print "  --- 注意 ---"
    Debug.Print "  ・委譲モードは検証用。ONのままだと全ての target=_blank リンクがポップアップに"
    Debug.Print "    なり、タブバーに乗りません。確認が済んだら Test_9_26_Popup_Off に戻すこと。"
    Debug.Print "  ・ポップアップはこちらの Controller 管理外なので、閉じるのは手動です。"
    Debug.Print "  ・★仕様事実 20★ WebView2 のイベントバーストが静まるまで (全タブのロード完了 +"
    Debug.Print "    イミディエイトのログが止まるまで) ブレーク/ステップ実行はしないこと。"
    Debug.Print "    この回はイベントハンドラ内の分岐を見るので、判断材料は Debug.Print のログのみ。"
End Sub
' ============================================================
' Test_9_26_PostProbe (第9.26b、POST プローブページを開く)
'
'   %TEMP%\Wv2PostProbe\postprobe.html を書き出し、仮想ホスト
'   https://appassets.postprobe/postprobe.html として新タブで開く。
'   StartWebView2_Full の後にイミディエイトで Sub 名を打つだけでよい。
'
'   ページには 4 つのボタンがある (詳細は Test_9_26_PostProbe_Help)。
' ============================================================
Public Sub Test_9_26_PostProbe()
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.26b] Browser が起動していません。先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Dim folderPath As String
    folderPath = Environ$("TEMP")
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & "Wv2PostProbe"

    If Not WriteUtf8NoBom(folderPath, "postprobe.html", BuildPostProbeHtml()) Then
        Debug.Print "[9.26b] HTML の書き出しに失敗しました。中止します。"
        Exit Sub
    End If

    Dim pane As Wv2Pane
    Set pane = b.AddTabWithUrlForSpa("appassets.postprobe", folderPath, "postprobe.html")
    If pane Is Nothing Then
        Debug.Print "[9.26b] タブの生成に失敗しました。"
        Exit Sub
    End If

    Debug.Print "[9.26b] プローブページを開きました。現在の NewWindowMode = " & b.NewWindowMode
    Debug.Print "        (0=現行/GET横取り  1=ポップアップ委譲  2=プリウォーム委譲)"
    Debug.Print "        手順は Test_9_28_Help (最新) / Test_9_26_PostProbe_Help (ボタンの役割) を参照。"
    Debug.Print "        第9.28b: 本命の送信先は httpbingo.org。落ちていたらボタン4 (httpbin) /"
    Debug.Print "        ボタン5 (postman-echo) を使うこと。★httpbingo は 405 をタイトルに出さない★ ので、"
    Debug.Print "        合否はログではなく画面表示で判定すること。"
End Sub


' ============================================================
' Test_9_26_PostProbe_Help (第9.26b、実機手順)
' ============================================================
Public Sub Test_9_26_PostProbe_Help()
    Debug.Print "==== 第9.26b 実機手順 (POST プローブで委譲モードを判定する) ===="
    Debug.Print ""
    Debug.Print "  【なぜこの段が要るか】"
    Debug.Print "    9.26 で試したサイトのポップアップは素の GET だったため、委譲 OFF でも"
    Debug.Print "    自前タブで正常に開いてしまい 405 の再現になっていなかった。そこで"
    Debug.Print "    httpbin.org を使って職場と同じ構図 (GET なら 405 / POST なら 200) を"
    Debug.Print "    自宅で合成し、本命の判定を取る。"
    Debug.Print ""
    Debug.Print "  --- 手順 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "  2) Test_9_26_PostProbe            ' プローブページが新タブで開く"
    Debug.Print "  3) ★対照 (委譲 OFF のまま)★ ボタン1 (POST + 別ウィンドウ) を押す"
    Debug.Print "       → 期待: 自前の新タブが開き、405 METHOD NOT ALLOWED が出る"
    Debug.Print "       → ログ: put_Handled OK → RequestNewTabAsync"
    Debug.Print "       ★ここで 405 が出れば、職場の症状が自宅で再現できたことになる★"
    Debug.Print "  4) その 405 タブを閉じ、プローブページのタブに戻る"
    Debug.Print "  5) Test_9_26_Popup_On             ' 委譲モードへ"
    Debug.Print "  6) もう一度ボタン1を押す"
    Debug.Print "       → ログ: [9.26 委譲モード] put_Handled を立てずにランタイムへ委譲する"
    Debug.Print "       → 期待: ポップアップに 200 の JSON が出る"
    Debug.Print "  7) ポップアップを × で閉じ、Test_9_26_Popup_Off で元に戻す"
    Debug.Print ""
    Debug.Print "  --- ボタンの役割 ---"
    Debug.Print "  ボタン1 [POST + 別ウィンドウ] ★本命★"
    Debug.Print "      委譲 OFF なら 405、ON なら 200 になるはず。ここだけ挙動が変わる。"
    Debug.Print "  ボタン2 [POST + 同じタブ]     サニティチェック"
    Debug.Print "      NewWindowRequested を通らない経路。どちらのモードでも 200 のはず。"
    Debug.Print "      ここが 200 なら『この WebView2 で POST 自体は普通に通る』が確定する。"
    Debug.Print "  ボタン3 [GET + 別ウィンドウ]  対照"
    Debug.Print "      どちらのモードでも 200 のはず。ポップアップ経路自体は壊れていない証拠。"
    Debug.Print "      (9.26 で見た『GET のポップアップは OFF でも開く』の再現)"
    Debug.Print "  ボタン4 [POST + 別ウィンドウ / 予備の送信先 その1]"
    Debug.Print "      httpbingo.org が落ちているときの逃げ道 (httpbin.org)。役割はボタン1と同じ。"
    Debug.Print "  ボタン5 [POST + 別ウィンドウ / 予備の送信先 その2]  ※第9.28b で新設"
    Debug.Print "      さらにその逃げ道 (postman-echo.com)。役割はボタン1と同じ。"
    Debug.Print ""
    Debug.Print "  --- 結果の読み方 ---"
    Debug.Print "  ★パターンA: ボタン1が OFF で 405 / ON で 200★"
    Debug.Print "      = ランタイム委譲で POST は保たれる。見立ては正しかった。"
    Debug.Print "        次段は本丸 (put_NewWindow で自前タブに収める) へ進んでよい。"
    Debug.Print "        さらに 200 の JSON 内 form に probe / nihongo / memo の値が見えれば、"
    Debug.Print "        ボディが欠落せず運ばれたことの直接証拠になる (日本語も確認できる)。"
    Debug.Print "  ★パターンB: ON でも 405 のまま★"
    Debug.Print "      = ランタイムに委譲しても POST が保たれていない。本丸を作っても"
    Debug.Print "        解決しない可能性が高いので、着手前に判明して得をした状態。"
    Debug.Print "  ★パターンC: OFF でも 405 が出ない★"
    Debug.Print "      = 405 の再現に失敗している。送信先が変わった可能性があるので、"
    Debug.Print "        表示された内容 (ステータス・本文) をそのまま共有してください。"
    Debug.Print "  ★ボタン2が 200 にならない場合★"
    Debug.Print "      = そもそも POST が通っていない。原因が別の層にあるので要相談。"
    Debug.Print ""
    Debug.Print "  --- 注意 ---"
    Debug.Print "  ・外部ネットワーク (httpbingo.org / httpbin.org / postman-echo.com) に"
    Debug.Print "    出られる環境で実行すること。1 つが 503 でも残り 2 つで検証を続けられる。"
    Debug.Print "  ・★第9.28b の注意★ httpbingo のエラーページには title が無いため、Edge が"
    Debug.Print "    URL をタイトル代用にする。ログの タイトル: 行では合否が読めないので、"
    Debug.Print "    必ず画面表示 (405 の文字 か 200 の JSON か) で判定すること。"
    Debug.Print "  ・確認が済んだら Test_9_26_Popup_Off で必ず委譲モードを戻すこと。"
    Debug.Print "  ・★仕様事実 20★ イベントバーストが静まるまでブレーク/ステップ実行はしない。"
    Debug.Print "    判断材料は Debug.Print のログと画面表示のみで押し切ること。"
End Sub


' ============================================================
' BuildPostProbeHtml (第9.26b、private)
'
'   検証ページの HTML を組み立てる。★JS は一切使わない★ 素の <form> だけで
'   足りる要件なので、VBA 文字列内に JS を持ち込まない (事故の芽を作らない)。
'
'   隠しフィールド nihongo に全角スペースを含む日本語を入れてあるので、
'   POST が保たれた場合は httpbin のエコー JSON にその文字列がそのまま現れる。
'   = ボディが欠落なく運ばれたことの直接証拠になる。
' ============================================================
Private Function BuildPostProbeHtml() As String
    Dim s As String
    ' 第9.28b: 送信先を 3 系統に。本命 = httpbingo.org (httpbin の Go 実装)。
    '   GET なら 405 / POST なら 200 + ボディをエコー、という構図は httpbin と同じ。
    Dim ep As String
    Dim epGet As String
    Dim ep2 As String
    Dim ep3 As String
    ep = "https://httpbingo.org/post"
    epGet = "https://httpbingo.org/get"
    ep2 = "https://httpbin.org/post"
    ep3 = "https://postman-echo.com/post"

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>POST プローブ (第9.26b / 第9.28b 更新)</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  *{box-sizing:border-box;margin:0;padding:0;}" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;color:#e8eaed;" & _
            "background:radial-gradient(1200px 600px at 20% -10%,#2b3550 0%,#151a26 55%,#0e121b 100%);" & _
            "min-height:100vh;padding:40px 28px;}" & vbLf
    s = s & "  .wrap{max-width:760px;margin:0 auto;}" & vbLf
    s = s & "  .eyebrow{letter-spacing:.22em;font-size:11px;color:#8ea2c8;text-transform:uppercase;}" & vbLf
    s = s & "  h1{font-size:24px;font-weight:650;margin:6px 0 6px;}" & vbLf
    s = s & "  .lead{color:#9aa7bd;font-size:13.5px;line-height:1.7;margin-bottom:24px;}" & vbLf
    s = s & "  .row{border:1px solid rgba(255,255,255,.08);border-radius:12px;padding:16px 18px;" & _
            "margin-bottom:14px;background:rgba(255,255,255,.03);}" & vbLf
    s = s & "  .row.main{border-color:rgba(110,168,254,.5);" & _
            "background:linear-gradient(180deg,rgba(110,168,254,.14),rgba(110,168,254,.04));}" & vbLf
    s = s & "  .tag{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;}" & vbLf
    s = s & "  .desc{font-size:12.5px;color:#9aa7bd;margin:8px 0 12px;line-height:1.6;}" & vbLf
    s = s & "  button{font-family:inherit;font-size:14px;font-weight:600;color:#0e121b;" & _
            "background:#6ea8fe;border:0;border-radius:8px;padding:9px 18px;cursor:pointer;}" & vbLf
    s = s & "  button:hover{background:#8fbcff;}" & vbLf
    s = s & "  .sub button{background:rgba(255,255,255,.14);color:#e8eaed;}" & vbLf
    s = s & "  .sub button:hover{background:rgba(255,255,255,.22);}" & vbLf
    s = s & "  label{display:block;font-size:11.5px;color:#8ea2c8;margin-bottom:5px;}" & vbLf
    s = s & "  input[type=text]{width:100%;font-family:inherit;font-size:13px;color:#cfe0ff;" & _
            "background:rgba(0,0,0,.28);border:1px solid rgba(255,255,255,.1);" & _
            "border-radius:8px;padding:8px 10px;margin-bottom:14px;}" & vbLf
    s = s & "  .ep{font-family:'Consolas','Courier New',monospace;font-size:11.5px;color:#7d8aa0;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<div class=""wrap"">" & vbLf
    s = s & "  <div class=""eyebrow"">POST Probe</div>" & vbLf
    s = s & "  <h1>NewWindowRequested 委譲モードの判定</h1>" & vbLf
    s = s & "  <div class=""lead"">送信先は GET で叩くと 405、POST で叩くと 200 を返します。" & _
            "職場の基幹システムとまったく同じ構図です。第9.28 で既定が プリウォーム委譲 (モード2) に" & _
            "なったので、<b>何も切り替えずにボタン1を押せば自前のタブに 200 が出る</b>のが正常です。" & _
            "対照 (405) を見たいときは先に Test_9_27_Mode_Legacy を実行してください。<br>" & _
            "本命の送信先が落ちているときは、同じ役割の予備ボタン (4・5) を使ってください。</div>" & vbLf

    ' --- 共通の隠しフィールド (各フォームに同じものを入れる) ---
    Dim fields As String
    fields = "    <input type=""hidden"" name=""probe"" value=""wv2-post-probe"">" & vbLf & _
             "    <input type=""hidden"" name=""nihongo"" value=""日本語　全角スペース入り"">" & vbLf

    ' --- ボタン1: POST + 別ウィンドウ (本命) ---
    s = s & "  <div class=""row main"">" & vbLf
    s = s & "    <span class=""tag"">本命</span>" & vbLf
    s = s & "    <div class=""desc"">POST + target=_blank。モード 0 なら 405、モード 2 なら 200 になるはず。" & _
            "ここだけモードで挙動が変わります。応答 JSON の method フィールドが POST になっていれば、" & _
            "GET に化けずにボディごと運ばれた直接証拠です。" & _
            "<br><span class=""ep"">" & ep & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <label>memo (POST ボディに載る文字列。自由に書き換え可)</label>" & vbLf
    s = s & "    <input type=""text"" name=""memo"" value=""hello from Excel VBA WebView2"">" & vbLf
    s = s & "    <button type=""submit"">ボタン1 : POST + 別ウィンドウ</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- ボタン2: POST + 同一タブ (サニティチェック) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">サニティ</span>" & vbLf
    s = s & "    <div class=""desc"">POST + 同じタブ。NewWindowRequested を通らない経路なので、" & _
            "どちらのモードでも 200 のはず。この WebView2 で POST 自体が通ることの確認です。</div>" & vbLf
    s = s & "    <form action=""" & ep & """ method=""post"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">ボタン2 : POST + 同じタブ</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- ボタン3: GET + 別ウィンドウ (対照) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">対照</span>" & vbLf
    s = s & "    <div class=""desc"">GET + target=_blank。どちらのモードでも 200 のはず。" & _
            "新規ウィンドウ経路そのものは壊れていないことの確認です。" & _
            "<br><span class=""ep"">" & epGet & "</span></div>" & vbLf
    s = s & "    <form action=""" & epGet & """ method=""get"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">ボタン3 : GET + 別ウィンドウ</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- ボタン4: POST + 別ウィンドウ (予備の送信先) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">予備</span>" & vbLf
    s = s & "    <div class=""desc"">ボタン1と同じ役割の逃げ道 その1。httpbingo.org が落ちているときに使います。" & _
            "<br><span class=""ep"">" & ep2 & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep2 & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">ボタン4 : POST + 別ウィンドウ (予備1 httpbin)</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    ' --- ボタン5: POST + 別ウィンドウ (予備の送信先 その2。第9.28b で新設) ---
    s = s & "  <div class=""row sub"">" & vbLf
    s = s & "    <span class=""tag"">予備</span>" & vbLf
    s = s & "    <div class=""desc"">ボタン1と同じ役割の逃げ道 その2。送信先を 3 つ持つことで、" & _
            "外部サービス 1 つのダウンで検証が止まるのを防ぎます (第9.28 で実際に足を取られた)。" & _
            "<br><span class=""ep"">" & ep3 & "</span></div>" & vbLf
    s = s & "    <form action=""" & ep3 & """ method=""post"" target=""_blank"">" & vbLf
    s = s & fields
    s = s & "    <button type=""submit"">ボタン5 : POST + 別ウィンドウ (予備2 postman-echo)</button>" & vbLf
    s = s & "    </form>" & vbLf
    s = s & "  </div>" & vbLf

    s = s & "</div>" & vbLf
    s = s & "</body></html>"

    BuildPostProbeHtml = s
End Function


' ============================================================
' WriteUtf8NoBom (第9.26b、private)
'
'   ADODB.Stream で UTF-8 (BOM なし) のテキストファイルを書き出す。
'   Wv2TabBar.WriteTabBarHtml と同じ手順 (テキストで書いてバイナリで読み直し、
'   先頭 3 バイトの BOM を捨ててから保存) をそのまま使う。
'
'   ★UserForm1.WriteSpaAppFolder を使わない理由★
'     あちらは Print # による ANSI (CP932) 書き出しなので、meta charset=UTF-8 を
'     宣言した HTML に日本語を入れると文字化けする。方式を TabBar/NavBar に揃える。
' ============================================================
Private Function WriteUtf8NoBom(ByVal folderPath As String, _
                                ByVal fileName As String, _
                                ByVal content As String) As Boolean
    On Error GoTo eh

    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
        Debug.Print "[9.26b] フォルダ作成 " & folderPath
    End If

    Dim fullPath As String
    fullPath = folderPath
    If Right$(fullPath, 1) <> "\" Then fullPath = fullPath & "\"
    fullPath = fullPath & fileName

    ' テキストとして UTF-8 で書き、バイナリで読み直して BOM (EF BB BF) を捨てる
    Dim bin As Object
    Set bin = CreateObject("ADODB.Stream")
    bin.Type = 2                ' adTypeText
    bin.Charset = "UTF-8"
    bin.Open
    bin.WriteText content
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

    Debug.Print "[9.26b] 書き出し OK " & fullPath & " (" & Len(content) & " 文字, UTF-8 BOMなし)"
    WriteUtf8NoBom = True
    Exit Function
eh:
    Debug.Print "[9.26b] 書き出しエラー " & Err.Number & " " & Err.Description
    WriteUtf8NoBom = False
End Function


' ============================================================
' 第9.27 ― プリウォーム委譲 (put_NewWindow) の実機スイッチと手順
' ============================================================

' ============================================================
' Test_9_27_Mode_Legacy / _Popup / _Prewarm (第9.27、モード切替)
'
'   Wv2Browser.NewWindowMode を切り替えるだけの薄い口。全タブ共通で、次に踏んだ
'   リンクから即座に効く (タブを開き直す必要はない)。
'     0 = 現行動作     : put_Handled + 自前の新タブ。Navigate は常に GET → POST は 405
'     1 = ポップアップ : ランタイム任せ。POST は保たれるがタブバーに乗らない (9.26)
'     2 = プリウォーム : 予備タブの CoreWebView2 を put_NewWindow に渡す (9.27 本命)
' ============================================================
Public Sub Test_9_27_Mode_Legacy()
    SetNewWindowMode 0
End Sub

Public Sub Test_9_27_Mode_Popup()
    SetNewWindowMode 1
End Sub

Public Sub Test_9_27_Mode_Prewarm()
    SetNewWindowMode 2
End Sub

Private Sub SetNewWindowMode(ByVal modeValue As Long)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser が起動していません。先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If
    b.NewWindowMode = modeValue
    Debug.Print "[9.27] 現在の NewWindowMode = " & b.NewWindowMode
End Sub


' ============================================================
' Test_9_27_Status (第9.27、予備タブの状態ダンプ)
'
'   モード 2 で開く前に、予備タブが温まっているかを確認するための口。
'   起動直後は「予備なし (pending=True)」で、1.5 秒ほどで「予備あり ... 使用可=True」
'   に変わる。使用可=False のまま変わらない場合はログを共有すること。
' ============================================================
Public Sub Test_9_27_Status()
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser が起動していません。先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Debug.Print "==== 第9.27 状態ダンプ ===="
    Debug.Print "  NewWindowMode            = " & b.NewWindowMode & _
                "   (0=現行 / 1=ポップアップ / 2=プリウォーム。第9.28 以降の既定は 2)"
    Debug.Print "  SetHandledWithNewWindow  = " & b.SetHandledWithNewWindow
    Debug.Print "  予備タブ                 = " & b.Debug_PrewarmStatus()
    Debug.Print "  表示中のタブ数           = " & b.TabCount & "   (予備は含まない)"
    Debug.Print "  アクティブ index         = " & b.ActiveIndex
End Sub


' ============================================================
' Test_9_27_Handled_On / _Off (第9.27、論点5 の保険スイッチ)
'
'   put_NewWindow に予備タブをセットしたとき、続けて put_Handled(TRUE) も立てるか。
'   既定は On (公式サンプルと同じ形)。予備タブが白いまま / ウィンドウが二重に開く
'   といった症状が出たら Off にして再試行し、どちらが正しいかを切り分ける。
'   ★コードを直さずに切り分けられる★ ようにするための保険。
' ============================================================
Public Sub Test_9_27_Handled_On()
    SetHandledFlag True
End Sub

Public Sub Test_9_27_Handled_Off()
    SetHandledFlag False
End Sub

Private Sub SetHandledFlag(ByVal onOff As Boolean)
    Dim b As Wv2Browser
    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[9.27] Browser が起動していません。先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If
    b.SetHandledWithNewWindow = onOff
    Debug.Print "[9.27] 現在の SetHandledWithNewWindow = " & b.SetHandledWithNewWindow
End Sub


' ============================================================
' Test_9_27_Help (第9.27、実機手順)
' ============================================================
Public Sub Test_9_27_Help()
    Debug.Print "==== 第9.27 実機手順 (本丸: put_NewWindow で予備タブに載せる) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    仕様事実 23 で「ランタイムに遷移を渡せば POST は保たれる」ことは確定した。"
    Debug.Print "    今回はそれを ★自前のタブ★ に載せられるかを見る。予備タブ (プリウォーム) の"
    Debug.Print "    CoreWebView2 を args.put_NewWindow に渡し、ポップアップではなくタブバーに"
    Debug.Print "    乗った状態で 200 + form エコーが出れば成功。"
    Debug.Print ""
    Debug.Print "  --- 手順 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "  2) 起動ログが静まってから Test_9_27_Status"
    Debug.Print "       → 期待: 予備タブ = 予備あり ViewPtr=... / IsReady=True / 使用可=True"
    Debug.Print "       ※ 起動直後は 予備なし (pending=True)。1.5 秒ほどで温まる。"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  3) Test_9_26_PostProbe            ' 9.26b の検証ページを新タブで開く"
    Debug.Print "  4) ★対照★ Test_9_27_Mode_Legacy を打ってから ボタン1 (POST + _blank) を押す"
    Debug.Print "       ※第9.28 で既定が 2 になったので、対照を取るには明示的に 0 へ戻す。"
    Debug.Print "       → 期待: 自前の新タブに 405 METHOD NOT ALLOWED (職場と同じ症状の再現)"
    Debug.Print "       → ログ: put_Handled OK → 新タブ生成を Timer に依頼"
    Debug.Print "  5) 405 のタブを閉じ、プローブページのタブに戻る"
    Debug.Print "  6) Test_9_27_Mode_Prewarm         ' 本命モードへ (第9.28 の既定に戻す)"
    Debug.Print "  7) もう一度 ボタン1 を押す"
    Debug.Print "  8) Test_9_27_Status               ' 補充されたかを確認"
    Debug.Print "       → 期待: 予備あり / 使用可=True (使った分が補充されている)"
    Debug.Print ""
    Debug.Print "  --- 期待するログ (手順 7) ---"
    Debug.Print "    Wv2Pane.View_OnNewWindowRequested: pSender=... , pArgs=..."
    Debug.Print "      → 新規ウィンドウ要求 URL: https://httpbin.org/post"
    Debug.Print "    Wv2Browser.ClaimPrewarmViewPtr: 予備タブを予約 (ViewPtr=...)"
    Debug.Print "      [9.27 プリウォーム] put_NewWindow OK (予備タブの View=...)"
    Debug.Print "        put_Handled OK (put_NewWindow と併用)"
    Debug.Print "    Wv2Browser.RequestAdoptPrewarmedAsync: 予備タブの採用を予約 (...)"
    Debug.Print "    Wv2Browser.Adopt: 予備タブをタブとして採用 (現在のタブ数=N, 要求 URL=...)"
    Debug.Print "    Wv2Browser.Prewarm: 予備タブの生成を開始 → 予備タブを温めた (ViewPtr=...)"
    Debug.Print ""
    Debug.Print "  --- 結果の読み方 ---"
    Debug.Print "  ★パターンA (成功): 新しいタブがタブバーに増え、その中に 200 の JSON が出た★"
    Debug.Print "      form に 3 フィールド / Content-Length: 172 が出ていれば POST が完全に"
    Debug.Print "      保たれている。本丸成立。UX まで含めて 405 問題が解決したことになる。"
    Debug.Print "  ★パターンB: タブは増えたが中身が真っ白 / 二重にウィンドウが開く★"
    Debug.Print "      put_Handled の併用が悪さをしている可能性。Test_9_27_Handled_Off にして"
    Debug.Print "      手順 5?7 をやり直す。これで直れば「NewWindow をセットしたら Handled は"
    Debug.Print "      立てない」が正解と確定する (論点5 の保険が効いた形)。"
    Debug.Print "  ★パターンC: ★印つきのフォールバックログが出た★"
    Debug.Print "      予備タブが使えなかったケース。Test_9_27_Status の結果と併せて共有すること。"
    Debug.Print "      (予備なし / 使用可=False のまま補充されない等、原因は状態ダンプで切り分ける)"
    Debug.Print "  ★パターンD: put_NewWindow が hr 付きで失敗した★"
    Debug.Print "      hr の値をそのまま共有してください。vtable index か引数型の問題になる。"
    Debug.Print ""
    Debug.Print "  --- 注意 ---"
    Debug.Print "  ・予備タブはタブバーに出ません (m_tabs に入れていないため)。Excel のタスク"
    Debug.Print "    マネージャで msedgewebview2.exe が 1 つ増えて見えるのは正常です。"
    Debug.Print "  ・Test_9_26_Popup_On / _Off は 9.27 でも動きます (内部でモード 1 / 0 に委譲)。"
    Debug.Print "    ただし _Off はモード 2 も解除するので、本命に戻すときは Test_9_27_Mode_Prewarm。"
End Sub


' ============================================================
' Test_9_28_Help (第9.28、実機手順)
'
'   既定が 2 (プリウォーム委譲) になった状態の確認。9.27 との最大の違いは
'   ★モード切替を一切打たずに POST リンクが 200 になること★。
' ============================================================
Public Sub Test_9_28_Help()
    Debug.Print "==== 第9.28 実機手順 (プリウォーム委譲の既定化 + 小修正 2 件) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    (A) モード切替を打たずに POST リンクが自前タブで 200 になるか"
    Debug.Print "    (B) 起動時に入れ子が起きても予備が最終的に温まるか (カウンタ化)"
    Debug.Print "    (C) 予備切れ時のログが実態に合った文言になっているか"
    Debug.Print ""
    Debug.Print "  --- 手順 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  2) 起動ログが静まってから Test_9_27_Status"
    Debug.Print "       → 期待: NewWindowMode = 2   ★ここが 0 なら既定値の変更が入っていない★"
    Debug.Print "       → 期待: 予備タブ = 予備あり ViewPtr=... / IsReady=True / 使用可=True"
    Debug.Print "  3) Test_9_26_PostProbe            ' 9.26b の検証ページを新タブで開く"
    Debug.Print "  4) ★本題★ モードを一切触らずに ボタン1 (POST + _blank) を押す"
    Debug.Print "       → 期待: 新しいタブがタブバーに増え、その中に 200 の JSON が出る"
    Debug.Print "               ★判定は画面で★ 応答 JSON の method が POST になっていて、"
    Debug.Print "               form に probe / nihongo / memo の 3 フィールドが見えれば POST 生存。"
    Debug.Print "               (httpbingo は 405 をタイトルに出さないので、ログでは判別できない)"
    Debug.Print "       → 期待するログ:"
    Debug.Print "           Wv2Browser.ClaimPrewarmViewPtr: 予備タブを予約 (ViewPtr=...)"
    Debug.Print "             [9.27 プリウォーム] put_NewWindow OK"
    Debug.Print "           Wv2Browser.Adopt: 予備タブをタブとして採用 (現在のタブ数=N, ...)"
    Debug.Print "           Wv2Browser.Prewarm: 予備タブを温めた (ViewPtr=...)"
    Debug.Print "  5) Test_9_27_Status               ' 補充されたかを確認"
    Debug.Print "       → 期待: 予備あり / 使用可=True"
    Debug.Print ""
    Debug.Print "  --- (C) の確認 (任意。ログ文言だけの修正なので余裕があれば) ---"
    Debug.Print "  6) プローブページに戻り、ボタン1 を★素早く 2 回★押す (予備切れを作る)"
    Debug.Print "       → 期待: 2 本目に ★予備タブが使えないので現行動作にフォールバック★ が出て、"
    Debug.Print "               続けて次の 1 行が出る (第9.28 で文言を変えた箇所):"
    Debug.Print "           Wv2Browser.RequestPrewarmAsync: 予備タブは Claim 済み (採用待ち)。"
    Debug.Print "           採用直後に自動で補充されるので、ここでは予約しない"
    Debug.Print "       ※9.27 では ここが 「既に予備タブがあるので何もしない」 と出て紛らわしかった。"
    Debug.Print "       ※2 本目は 405 になる (予備が 1 枚しかないため)。これは想定どおり。"
    Debug.Print "       ※連打や外部サービスの不調で 503 が返ることがある (405 と紛らわしい)。"
    Debug.Print "         切り分け方: アドレスバーに送信先 URL を直打ちして 405 が出るか見る。"
    Debug.Print "         503 ならサーバ側のダウンなので、ボタン4 / ボタン5 の予備に切り替える。"
    Debug.Print ""
    Debug.Print "  --- (B) の確認 (ログを見るだけ) ---"
    Debug.Print "  7) 起動ログに次の行があってもよい (むしろ入れ子が起きた証拠):"
    Debug.Print "       Wv2Browser.TimerCall(prewarm): Pane 生成中だったので今回は見送り"
    Debug.Print "     その場合でも AddTab 末尾の予約し直しで最終的に温まるので、手順 2 で"
    Debug.Print "     使用可=True になっていれば正常。"
    Debug.Print ""
    Debug.Print "  --- 終了時 ---"
    Debug.Print "  8) UserForm を閉じる"
    Debug.Print "       → 期待: Wv2Browser.Shutdown: 予備タブ (プリウォーム) を解放 が出る"
    Debug.Print ""
    Debug.Print "  --- 従来動作に戻したいとき ---"
    Debug.Print "  ・Test_9_27_Mode_Legacy   (モード 0 = 9.27 以前の既定。POST は 405 に戻る)"
    Debug.Print "  ・Test_9_27_Mode_Prewarm  (モード 2 = 第9.28 の既定へ復帰)"
End Sub


' ============================================================
' Test_9_29_Custom  (第9.29 カスタム検索エンジンの純ロジック検証)
'
'   ★イミディエイトで  Test_9_29_Custom  と打つだけ★
'
'   Wv2SettingsBridge のカスタム系 (SetCustomEngine / GetCustomTemplate /
'   PreviewUrlForTemplate) と、custom の永続化 (engine + template の 2 行) を
'   照合する。WebView2 は起動しない (ブリッジ + Browser の Debug_* だけ)。
'   実ファイルを触るため、先頭で engine と template を対で退避し、末尾で戻す。
' ============================================================
Public Sub Test_9_29_Custom()
    Dim total As Long, pass As Long
    total = 0: pass = 0
    Debug.Print "==== Test_9_29_Custom (カスタム検索エンジン 純ロジック) ===="

    Dim br As Wv2SettingsBridge
    Set br = New Wv2SettingsBridge

    ' --- 現在の保存値を退避 (engine + template の対) ---
    Dim savedName As String, savedTpl As String
    savedName = br.LoadEngineName()
    savedTpl = br.LoadCustomTemplate()
    Debug.Print "  (退避) engine = '" & savedName & "' / template = '" & savedTpl & "'"

    Dim b As Wv2Browser
    Set b = New Wv2Browser
    br.BindBrowser b

    Dim before As String
    before = b.Debug_SearchEngine()

    Const TPL As String = "https://example.com/find?query="

    ' --- 不正入力: 戻り値は "" で、Browser も保存も触らない ---
    CheckEq "空文字 → 戻り値は空", br.SetCustomEngine(""), "", total, pass
    CheckEq "スキーム無し → 戻り値は空", br.SetCustomEngine("example.com/?q="), "", total, pass
    CheckEq "ftp:// → 戻り値は空", br.SetCustomEngine("ftp://example.com/?q="), "", total, pass
    CheckEq "不正入力後も Browser は無変化", b.Debug_SearchEngine(), before, total, pass

    ' --- 正常な適用 ---
    CheckEq "適用 → 'custom'", br.SetCustomEngine(TPL), "custom", total, pass
    CheckEq "Browser に反映", b.Debug_SearchEngine(), TPL, total, pass
    CheckEq "GetEngine は custom", br.GetEngine(), "custom", total, pass
    CheckEq "GetCustomTemplate", br.GetCustomTemplate(), TPL, total, pass
    CheckEq "検索 URL 生成 (本番経路)", b.Debug_NormalizeUrl("hokkaido"), TPL & "hokkaido", total, pass

    ' --- 永続化 (custom は engine + template の 2 行) ---
    CheckEq "保存 → LoadEngineName", br.LoadEngineName(), "custom", total, pass
    CheckEq "保存 → LoadCustomTemplate", br.LoadCustomTemplate(), TPL, total, pass
    CheckEq "保存 → LoadEngineTemplate", br.LoadEngineTemplate(), TPL, total, pass

    ' --- 適用前プレビュー (Browser の状態を変えない) ---
    CheckEq "PreviewUrlForTemplate", _
            br.PreviewUrlForTemplate("https://foo.test/s?k=", "hokkaido"), _
            "https://foo.test/s?k=hokkaido", total, pass
    CheckEq "PreviewUrlForTemplate 不正 → 空", _
            br.PreviewUrlForTemplate("foo.test/s?k=", "hokkaido"), "", total, pass
    CheckEq "プレビューは Browser を変えない", b.Debug_SearchEngine(), TPL, total, pass

    ' --- プリセットへ戻すと template 行が消える (論点3 案a-1) ---
    CheckEq "SetEngine bing", br.SetEngine("bing"), "bing", total, pass
    CheckEq "→ LoadEngineName", br.LoadEngineName(), "bing", total, pass
    CheckEq "→ LoadCustomTemplate は空", br.LoadCustomTemplate(), "", total, pass
    CheckEq "→ LoadEngineTemplate", br.LoadEngineTemplate(), _
            "https://www.bing.com/search?q=", total, pass
    CheckEq "→ GetCustomTemplate も空", br.GetCustomTemplate(), "", total, pass

    ' --- 入力がプリセットと同一 URL なら、そのプリセット名になる ---
    CheckEq "プリセット同一 URL → 'google'", _
            br.SetCustomEngine("https://www.google.com/search?q="), "google", total, pass
    CheckEq "→ template 行は書かれない", br.LoadCustomTemplate(), "", total, pass

    ' --- 壊れた template 行は復元時に "" へ落ちる (既定 Google で起動する) ---
    br.Debug_SaveSettings "custom", "example.com/?q="
    CheckEq "壊れた template → LoadEngineTemplate は空", br.LoadEngineTemplate(), "", total, pass

    ' --- 退避値を書き戻す (ユーザー設定の復元) ---
    If Len(savedName) > 0 Then
        br.Debug_SaveSettings savedName, savedTpl
        Debug.Print "  (復元) engine = '" & savedName & "' / template = '" & savedTpl & "' に戻した。"
    Else
        br.Debug_SaveSettings "google", ""
        Debug.Print "  (復元) 元々保存が無かったので engine=google を書いた (既定と同じなので実害なし)。"
    End If

    Debug.Print "==== 結果: " & pass & " / " & total & _
                IIf(pass = total, "  [ALL OK]", "  [!! FAILED !!]")
End Sub


' ============================================================
' Test_9_29_Help  (第9.29 実機手順)
'
'   ★イミディエイトで  Test_9_29_Help  と打つと手順が出る★
' ============================================================
Public Sub Test_9_29_Help()
    Debug.Print "==== 第9.29 実機手順 (カスタム検索 URL + 予備予約の掃除) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    (A) 設定タブのカスタム入力欄から任意の検索 URL を適用できるか"
    Debug.Print "    (B) それが Excel 再起動後も復元されるか (engine=custom + template)"
    Debug.Print "    (C) 予備切れフォールバック時の無駄な予約ログが消えたか"
    Debug.Print ""
    Debug.Print "  --- 事前 (任意) ---"
    Debug.Print "  0) Test_9_29_Custom            ' 純ロジックが [ALL OK] であること"
    Debug.Print ""
    Debug.Print "  --- (A) カスタムの適用 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  2) ? を押して設定タブを開く"
    Debug.Print "       → 期待: カードの下に ★カスタム (プリセット以外)★ の枠と入力欄が出る"
    Debug.Print "  3) 入力欄に次を貼って、まだ適用は押さない:"
    Debug.Print "         https://ja.wikipedia.org/w/index.php?search="
    Debug.Print "       → 期待: 入力中に下の行へ  日本 天気 → https://ja.wikipedia.org/... "
    Debug.Print "               と★適用前のプレビュー★が出る (VBA がエンコードしている)"
    Debug.Print "  4) わざと  example.com/?q=  のようにスキーム無しで打ってみる"
    Debug.Print "       → 期待: 赤字で『http:// または https:// で始まる URL を…』が出る"
    Debug.Print "               ★この状態で [適用] を押してもログは『不正なテンプレートなので無視』"
    Debug.Print "                 だけで、検索エンジンは変わらない★"
    Debug.Print "  5) 手順 3 の URL に戻して [適用] を押す (Enter でも同じ)"
    Debug.Print "       → 期待するログ:"
    Debug.Print "           Wv2Browser.SearchEngine: 検索テンプレートを設定 = https://ja.wikipedia.org/..."
    Debug.Print "           Wv2SettingsBridge.SaveSettingsFile: 保存 'custom' / template=... -> ...\settings.txt"
    Debug.Print "           Wv2SettingsBridge.SetCustomEngine: 適用 '...' -> 'custom'"
    Debug.Print "       → 期待: カスタム枠が青く光り『使用中』バッジが出る。カードは全て非選択。"
    Debug.Print "  6) 別の通常タブでアドレスバーに   日本 天気   と打つ"
    Debug.Print "       → 期待: Wikipedia の検索結果が出る"
    Debug.Print ""
    Debug.Print "  --- (B) 永続化 ---"
    Debug.Print "  7) Excel を 完全に終了 する (ブックを閉じるだけでなく Excel ごと)"
    Debug.Print "  8) 再起動して StartWebView2_Full"
    Debug.Print "       → 期待するログ (復元の証):"
    Debug.Print "           Wv2Browser.Class_Initialize: 保存済みエンジンを復元 = https://ja.wikipedia.org/..."
    Debug.Print "  9) 設定タブを開かずに、いきなり通常タブで   日本 天気   と打つ"
    Debug.Print "       → 期待: Wikipedia の検索結果 (カスタムが復元されている)"
    Debug.Print " 10) 設定タブを開く"
    Debug.Print "       → 期待: 入力欄に前回の URL が入っていて『使用中』バッジが出ている"
    Debug.Print ""
    Debug.Print "  --- (C) 予備予約の掃除 (ログを見るだけ) ---"
    Debug.Print " 11) Test_9_26_PostProbe でプローブページを開き、ボタン1 を★素早く 2 回★押す"
    Debug.Print "       → 期待: 2 本目で ★予備タブが使えないので現行動作にフォールバック★ の後、"
    Debug.Print "               1.5 秒待っても次の行が★出ない★こと (第9.29 で消した無駄ログ):"
    Debug.Print "                 Wv2Browser.Prewarm: 既に予備タブがある (生成をスキップ)"
    Debug.Print "       → 代わりに次が出ることがある (これは正常。入れ子を検知した印):"
    Debug.Print "           Wv2Browser.EnsurePrewarmScheduled: 予備タブの生成中 (入れ子) なので予約しない"
    Debug.Print " 12) Test_9_27_Status               ' 予備が 1 枚に保たれているか"
    Debug.Print "       → 期待: 予備あり / 使用可=True"
    Debug.Print ""
    Debug.Print "  --- 元に戻したいとき ---"
    Debug.Print "  ・設定タブで Google のカードをクリックすれば戻る"
    Debug.Print "    (このとき settings.txt の template 行は消える = 仕様)"
    Debug.Print "  ・%APPDATA%\Wv2Browser\settings.txt を手で消して再起動しても既定に戻る"
End Sub
' ============================================================
' Test_9_31_Help  (第9.31 実機手順 / 目視)
'
'   ★イミディエイトで  Test_9_31_Help  と打つと手順が出る★
'
'   第9.31 は CSS のみの変更なので純ロジックのテストは無い。
'   併せて第9.30 の宿題 1 (v0_5_2 の通し動作確認) をここで消化する。
' ============================================================
Public Sub Test_9_31_Help()
    Debug.Print "==== 第9.31 実機手順 (タブのホバー + 閉じるボタンの表示制御) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    (A-1) タブにマウスを乗せると背景が明るくなるか (非アクティブのみ)"
    Debug.Print "    (A-2) 閉じるボタンが『アクティブ or ホバー時だけ』見えるか"
    Debug.Print "    (C)   ★第9.30 の宿題★ v0_5_2 が通しで正常に動くか"
    Debug.Print ""
    Debug.Print "  ※ この回は★CSS だけ★の変更です。ログの文言は第9.29 から一切変わりません。"
    Debug.Print "     判定はすべて★画面の目視★で行ってください。"
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  2) + を 2?3 回押してタブを 3?4 枚にする"
    Debug.Print "       (アクティブ 1 枚 + 非アクティブ複数、の状態を作るのが目的)"
    Debug.Print ""
    Debug.Print "  --- (A-1) ホバー背景 ---"
    Debug.Print "  3) ★非アクティブ★のタブにマウスを乗せる → 離す"
    Debug.Print "       → 期待: 乗せると背景がわずかに明るくなり (#e7e7e7 → #efefef)、"
    Debug.Print "               離すと戻る。切り替わりは★じんわり (0.12 秒)★"
    Debug.Print "       → 期待: タブの★幅も文字の位置も動かない★ (色だけが変わる)"
    Debug.Print "  4) ★アクティブ★のタブにマウスを乗せる"
    Debug.Print "       → 期待: ★何も変わらない★ (元から白 = 最も明るいので変えない仕様)"
    Debug.Print "  5) タブを切り替える (別のタブをクリック)"
    Debug.Print "       → 期待: 白へ変わるときも同じ 0.12 秒でなめらかに変わる"
    Debug.Print "               ※ここが不自然なら論点 5 を『遷移なし』に戻す判断材料になる"
    Debug.Print ""
    Debug.Print "  --- (A-2) 閉じるボタンの表示制御 ---"
    Debug.Print "  6) マウスをタブバーの外 (ページ本体など) へ完全に逃がして、タブ列を眺める"
    Debug.Print "       → 期待: ★アクティブなタブにだけ × が見えている★"
    Debug.Print "       → 期待: 非アクティブなタブには × が★見えない★"
    Debug.Print "  7) 非アクティブのタブにマウスを乗せる"
    Debug.Print "       → 期待: そのタブにだけ × が現れる (即座に。フェードはしない仕様)"
    Debug.Print "       → 期待: ★タイトルの文字が左右に動かない★"
    Debug.Print "               (visibility:hidden で場所を確保しているため。ここが動くなら"
    Debug.Print "                display:none になってしまっているので要報告)"
    Debug.Print "  8) × の上にマウスを乗せる"
    Debug.Print "       → 期待: × の背景が灰色の角丸になる (第9.29 までと同じ)"
    Debug.Print "  9) × が見えていない非アクティブタブの『× があるはずの位置』をクリックする"
    Debug.Print "       → 期待: ★タブが閉じない★ (= 見えないボタンは押せない)。"
    Debug.Print "               そのタブへ★切り替わる★のが正しい挙動。"
    Debug.Print "               ※これが閉じてしまうなら opacity:0 相当になっているので要報告"
    Debug.Print " 10) タブを 8?10 枚まで増やして、タブが細くなった状態で 6?7 を再確認"
    Debug.Print "       → 期待: min-width:80px で頭打ちになり、× の場所は常に確保されている"
    Debug.Print "       → ここでタイトルが読めなさすぎる等の不満が出たら、次段の材料にする"
    Debug.Print ""
    Debug.Print "  --- (B) 回帰: ドラッグ並べ替えが壊れていないこと ---"
    Debug.Print " 11) タブを掴んで左右に動かし、別のタブの上で放す"
    Debug.Print "       → 期待: 掴んだタブが薄くなり (opacity 0.45)、"
    Debug.Print "               放す先のタブの左右どちらかに★青い縦線★が出る"
    Debug.Print "       → 期待するログ: [host] reorder N->M (WebView2 の DevTools 側)"
    Debug.Print "       → 期待: 放すと並びが変わり、アクティブが追従する"
    Debug.Print "       ※ドラッグ中にホバー色が出ることがあるが★仕様★ (論点 7 で先送り)"
    Debug.Print ""
    Debug.Print "  --- (C) ★第9.30 の宿題 1: v0_5_2 の通し動作確認★ ---"
    Debug.Print "  ※ここは第9.31 の変更とは無関係。クリーンビルドしたブックで"
    Debug.Print "    未確認だった 3 つを、この機会にまとめて踏むための手順です。"
    Debug.Print ""
    Debug.Print " 12) Test_9_27_Status               ' 予備タブが温まっているか"
    Debug.Print "       → 期待: 予備あり / 使用可=True"
    Debug.Print " 13) Test_9_26_PostProbe            ' POST プローブページを開く"
    Debug.Print " 14) ボタン1 (httpbingo POST + _blank) を押す"
    Debug.Print "       → 期待: ★自前の新しいタブ★が開き、画面に method: POST と"
    Debug.Print "               form の中身が出る (405 でも空白でもない)"
    Debug.Print "       → ★注意★ タイトルは成功でも失敗でも httpbingo.org/post なので、"
    Debug.Print "                  判定は必ず★画面表示★で行うこと (第9.28b 由来)"
    Debug.Print "       → 503 が出たら Test_9_28_Help の切り分け手順へ (サーバ側の都合)"
    Debug.Print " 15) Test_9_27_Status               ' 予備が補充されたか"
    Debug.Print "       → 期待: 予備あり / 使用可=True (使った直後に補充されている)"
    Debug.Print " 16) 適当なページのリンクを Ctrl+クリック か、target=_blank のリンクを押す"
    Debug.Print "       → 期待: 自前のタブで開く (ポップアップウィンドウにならない)"
    Debug.Print ""
    Debug.Print "  --- 判定 ---"
    Debug.Print "  ・3?11 がすべて期待どおり → 第9.31 は合格"
    Debug.Print "  ・12?16 がすべて期待どおり → ★v0_5_2 の通し確認 (第9.30 の宿題 1) を消化★"
    Debug.Print "  ・気になった点 (色が濃い/薄い、遷移が遅い/速い、× が小さい等) は"
    Debug.Print "    論点 2・5・6 の再調整で対応できるので、遠慮なく挙げてください。"
End Sub
' ============================================================
' Test_9_32_Help  (第9.32 実機手順 / 目視・体感)
'
'   ★イミディエイトで  Test_9_32_Help  と打つと手順が出る★
'
'   第9.32 は JS と CSS のみの変更。純ロジックのテストは無い。
' ============================================================
Public Sub Test_9_32_Help()
    Debug.Print "==== 第9.32 実機手順 (タブ切替の楽観的更新 + スクロールバーの整理) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    (1) クリックした★瞬間★にタブの色が変わるか (本体の切替は後から追いつく)"
    Debug.Print "    (2) タブを増やしても横スクロールバーが出ず、タブが潰れないか"
    Debug.Print "    (3) タブ列をマウスホイールで横スクロールできるか"
    Debug.Print "    (4) 回帰: D&D 並べ替え / 閉じる / 設定タブ が壊れていないか"
    Debug.Print ""
    Debug.Print "  ※ VBA のロジックは 1 行も変えていないので★ログの文言も順序も第9.31 と同じ★です。"
    Debug.Print "     変わるのは『ログが出るより先にタブバーの色が変わる』という点だけ。"
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full   ' 通常どおり起動"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  2) + を押してタブを 4?5 枚にし、それぞれ別のページを開いておく"
    Debug.Print "       (切替に時間がかかるページの方が違いが分かりやすい)"
    Debug.Print ""
    Debug.Print "  --- (1) ★本命: タブ切替の体感★ ---"
    Debug.Print "  3) 非アクティブのタブをクリックする"
    Debug.Print "       → 期待: ★クリックとほぼ同時にタブバーの色が切り替わる★"
    Debug.Print "               (押した瞬間に白くなり、押していたタブが灰色に戻る)"
    Debug.Print "       → 期待: その★後から★ブラウザ本体の表示が切り替わる"
    Debug.Print "       → 期待: イミディエイトのログもタブバーの色変化より★後★に出る"
    Debug.Print "               Wv2TabBar.HostActivate: index=N (host経由)"
    Debug.Print "               Wv2TabBar[EVT] ActiveChanged(index=N) → PushSyncToJs"
    Debug.Print "               Wv2TabBar.PushSyncToJs: 送信 OK (...)"
    Debug.Print "  4) タブを 5?6 回続けて切り替えて、体感を第9.31 と比べる"
    Debug.Print "       → 判定: 『クリックと画面の変化がずれている』感覚が減ったか"
    Debug.Print "       ★注意★ ブラウザ本体の切替そのものは★速くなっていません★。"
    Debug.Print "               速くなったのは『タブバーが反応するまでの時間』だけです。"
    Debug.Print "               本体側がまだ遅すぎると感じるなら、ActivateTab の中身を"
    Debug.Print "               見る別の段階が必要になるので、そう報告してください。"
    Debug.Print "  5) 閉じるボタン (×) を押したときの体感も見ておく"
    Debug.Print "       → 期待: ★こちらは第9.31 と同じ (遅いまま)★ が正しい。"
    Debug.Print "               論点 3 で close には楽観的更新を入れないと決めたため。"
    Debug.Print "               気になるようなら次段で同じ手を close にも入れられます。"
    Debug.Print ""
    Debug.Print "  --- (2)(3) スクロールバーとホイール ---"
    Debug.Print "  6) + を連打してタブを 10?12 枚まで増やす"
    Debug.Print "       → 期待: ★タブ列の下に横スクロールバーが出ない★"
    Debug.Print "       → 期待: タブの高さが第9.31 のときのように★潰れない★"
    Debug.Print "               (第9.31 ではバーが高さを内側から食ってタブが上に寄っていた)"
    Debug.Print "  7) タブ列の上にマウスを置いて、ホイールを上下に回す"
    Debug.Print "       → 期待: タブ列が★左右にスクロールする★"
    Debug.Print "               (Chromium が縦ホイールを横スクロールへ変換する)"
    Debug.Print "       → 期待: ブラウザ本体のページはスクロールしない"
    Debug.Print "  8) 右端まで送って、いちばん右のタブをクリックする"
    Debug.Print "       → 期待: 手順 3 と同じく即座に色が変わり、切替も正しく効く"
    Debug.Print "       ※ ホイールが効かない場合は報告してください (論点 4 の代替案に戻します)"
    Debug.Print ""
    Debug.Print "  --- (4) 回帰 ---"
    Debug.Print "  9) タブを掴んで並べ替える"
    Debug.Print "       → 期待: 薄くなる + 青い縦線 + 放すと並びが変わる (第9.31 と同じ)"
    Debug.Print "       → 期待: ★非アクティブのタブを掴んでもアクティブにならない★"
    Debug.Print "               (論点 6 で現状維持と決定した仕様です)"
    Debug.Print " 10) × でタブを何枚か閉じる"
    Debug.Print "       → 期待: 正しく閉じ、アクティブが補正される"
    Debug.Print " 11) ? で設定タブを開き、検索エンジンのカードをクリックする"
    Debug.Print "       → 期待: 第9.29 までと同じ (選択が反映され、保存ログが出る)"
    Debug.Print " 12) 設定タブを閉じて、通常タブのアドレスバーで検索してみる"
    Debug.Print "       → 期待: 選んだエンジンで検索できる"
    Debug.Print ""
    Debug.Print "  --- 判定 ---"
    Debug.Print "  ・3?4 で体感が改善 → 第9.32 の本命は成功"
    Debug.Print "  ・6?8 が期待どおり → スクロールバーの整理も成功"
    Debug.Print "  ・9?12 が第9.31 と同じ → 回帰なし"
    Debug.Print "  ・色の変化が速すぎて『押す前に変わった』ように感じる等の違和感があれば"
    Debug.Print "    報告してください。二重 rAF を 1 回に減らす等で調整できます。"
End Sub
' ============================================================
' Test_9_32b_Help  (第9.32b 実機手順 / 目視)
'
'   ★イミディエイトで  Test_9_32b_Help  と打つと手順が出る★
' ============================================================
Public Sub Test_9_32b_Help()
    Debug.Print "==== 第9.32b 実機手順 (ホイール横スクロール + スクロール位置の維持) ===="
    Debug.Print ""
    Debug.Print "  【この回で知りたいこと】"
    Debug.Print "    (1) タブ列の上でホイールを回すと横スクロールするか"
    Debug.Print "    (2) タイトル更新の同期が来てもスクロール位置が左端へ戻らないか"
    Debug.Print "    (3) + で足した新しいタブが自動で見える位置に来るか"
    Debug.Print "    (4) 回帰: 切替の体感 / D&D / 閉じる"
    Debug.Print ""
    Debug.Print "  ※ VBA は 1 行も変えていないので★ログは第9.32 と完全に同じ★です。"
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) UserForm1.StartWebView2_Full"
    Debug.Print "       ※ ★仕様事実 20★ ログが止まるまでブレーク/ステップ実行はしないこと。"
    Debug.Print "  2) + を連打してタブを 12 枚程度まで増やす (溢れる状態を作る)"
    Debug.Print ""
    Debug.Print "  --- (1) ホイール ---"
    Debug.Print "  3) タブ列の上にマウスを置いてホイールを下へ回す"
    Debug.Print "       → 期待: タブ列が★右へスクロールする★"
    Debug.Print "       → 期待: ブラウザ本体のページは★スクロールしない★"
    Debug.Print "  4) ホイールを上へ回す → 左へ戻る"
    Debug.Print "  5) タブが 1?2 枚しか無い状態 (溢れていない状態) でもホイールを回す"
    Debug.Print "       → 期待: ★何も起きない★ (scrollWidth <= clientWidth で早期 return)"
    Debug.Print "               ここで本体ページがスクロールするなら仕様どおり (preventDefault"
    Debug.Print "               まで到達しないため)。タブバー上でのことなので実害は無い。"
    Debug.Print ""
    Debug.Print "  --- (2) ★スクロール位置が勝手に戻らないこと★ ---"
    Debug.Print "  6) 右端あたりまでスクロールした状態で放置する"
    Debug.Print "       (読み込み中のタブがあると、完了時にタイトル更新の同期が飛ぶ)"
    Debug.Print "       → 期待: ログに PushSyncToJs が出ても★表示位置が動かない★"
    Debug.Print "       → 第9.32 まではここで★左端へ戻っていた★ (今回の修正点)"
    Debug.Print "  7) 右のほうを見ている状態で、どれかのタブのタイトルが変わる操作をする"
    Debug.Print "       (例: 見えているタブでページを読み込ませる)"
    Debug.Print "       → 期待: 位置は動かない。アクティブタブの位置へ引き戻されない"
    Debug.Print ""
    Debug.Print "  --- (3) 新規タブが見えること ---"
    Debug.Print "  8) 左端まで戻した状態で + を押す"
    Debug.Print "       → 期待: 新しいタブは右端に付くが、★自動で右端まで寄って見える★"
    Debug.Print "  9) 逆に、右端を見ている状態で左のほうのタブをクリックする"
    Debug.Print "       → 期待: そのタブは既に見えているので★大きく動かない★"
    Debug.Print "               (nearest 指定なので、見えているものは寄せない)"
    Debug.Print " 10) 右端を見ている状態で、Test_9_27_Status 等で左のタブに切り替わる操作をする"
    Debug.Print "       → 期待: アクティブが変わったときは可視範囲へ寄る"
    Debug.Print ""
    Debug.Print "  --- (4) 回帰 ---"
    Debug.Print " 11) タブ切替の体感が第9.32 と同じか (クリック即座に色が変わる)"
    Debug.Print " 12) D&D 並べ替えが効くか。ドラッグ後に位置が飛ばないか"
    Debug.Print " 13) × で閉じられるか。閉じた後にスクロール位置が破綻しないか"
    Debug.Print " 14) ? で設定タブが開くか"
    Debug.Print ""
    Debug.Print "  --- 判定 ---"
    Debug.Print "  ・3?4 が効けば論点 1 は成功 (効かなければ論点 1 を (b) 3px バーに戻す)"
    Debug.Print "  ・6?7 で位置が動かなければ論点 3 の退避・復元は成功"
    Debug.Print "  ・8 で新規タブが見えれば nearest 寄せも成功"
    Debug.Print "  ・スクロール量が速すぎ/遅すぎる場合は数値 1 つで調整できます"
End Sub


' ============================================================
' ★D-1 の検証★
' ============================================================


' ============================================================
' Test_D1_Eval (D-1)
'   EvalSync を一通り叩いて結果を並べる。
'   前提: StartWebView2_Full でブラウザが起動していること。
'         アクティブタブが実ページを表示していること (about:blank でも大半は通る)。
'   ★DoEvents ループ中はブレーク/ステップ実行しないこと (仕様事実 20)★
' ============================================================
Public Sub Test_D1_Eval()
    Dim p As Wv2Pane
    Set p = UserForm1.GetActivePane

    If p Is Nothing Then
        Debug.Print "Test_D1_Eval: アクティブな Pane がありません。" & _
                    "先に StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Debug.Print ""
    Debug.Print "================ Test_D1_Eval 開始 ================"
    Debug.Print "  対象タブ: " & p.DocumentTitle
    Debug.Print "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    Debug.Print ""
    Debug.Print "  --- (1) 型ごとの戻り値 ---"

    D1Case p, "整数", "1 + 2", 5, True
    D1Case p, "小数", "1 / 4", 5, True
    D1Case p, "文字列 (引用符付きで返るのが正常)", "'abc'", 5, True
    D1Case p, "日本語 (uXXXX の復号を確認)", "'日本語テスト'", 5, True
    D1Case p, "引用符を含む文字列", "'[' + String.fromCharCode(34) + ']'", 5, True
    D1Case p, "バックスラッシュを含む文字列", "'[' + String.fromCharCode(92) + ']'", 5, True
    D1Case p, "真偽値", "1 === 1", 5, True
    D1Case p, "null", "null", 5, True
    D1Case p, "undefined (value キーが消えるケース)", "undefined", 5, True
    D1Case p, "オブジェクト", "({a:1,b:[1,2,3]})", 5, True
    D1Case p, "配列", "[1,'x',null]", 5, True
    D1Case p, "長い文字列 (300 字)", _
              "(function(){var s='';for(var i=0;i<300;i++){s+='x';}return s;})()", 5, True

    Debug.Print ""
    Debug.Print "  --- (2) 実ページの情報 ---"
    D1Case p, "document.title", "document.title", 5, True
    D1Case p, "location.href", "document.location.href", 5, True
    D1Case p, "body の文字数", "document.body.innerText.length", 5, True

    Debug.Print ""
    Debug.Print "  --- (3) 失敗系 (FAIL 表示にならず OK と出れば正常) ---"
    D1Case p, "JS 例外 (ReferenceError)", "nonexistentFunctionForTest()", 5, False
    D1Case p, "JS 例外 (throw を即時関数で包む)", _
              "(function(){throw new Error('boom');})()", 5, False
    D1Case p, "構文エラー (式が壊れている)", "1 +", 5, False

    Debug.Print ""
    Debug.Print "  --- (4) タイムアウトと回復 ---"
    Debug.Print "  ※ JS を 3 秒ブロックし、1 秒で打ち切る。数秒後に"
    Debug.Print "     ★破棄済み★ の遅延到着ログが出れば論点5 は成功。"
    D1Case p, "タイムアウト (3 秒を 1 秒で打ち切り)", _
              "(function(){var t=Date.now();while(Date.now()-t<3000){}return 1;})()", 1, False

    ' JS スレッドが空くまで待つ (この間に遅延到着のログが出る)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) < 4
        DoEvents
    Loop

    D1Case p, "タイムアウト後の回復", "1 + 1", 5, True

    Debug.Print ""
    Debug.Print "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    Debug.Print "================ Test_D1_Eval 終了 ================"
    Debug.Print ""
End Sub


' ============================================================
' D1Case (D-1 検証用の 1 ケース実行)
'   expectOk = True  … 成功するはずのケース
'   expectOk = False … 失敗するはずのケース (エラー理由が出れば OK)
' ============================================================
Private Sub D1Case(ByVal p As Wv2Pane, _
                   ByVal caseName As String, _
                   ByVal expr As String, _
                   ByVal toSec As Single, _
                   ByVal expectOk As Boolean)
    Dim r As String
    Dim mark As String

    r = p.EvalSync(expr, toSec)

    If p.LastEvalOk = expectOk Then
        mark = "  [OK  ] "
    Else
        mark = "  [FAIL] "
    End If

    If p.LastEvalOk Then
        Debug.Print mark & caseName & " → " & Left$(r, 100)
    Else
        Debug.Print mark & caseName & " → 失敗: " & p.LastEvalError
    End If
End Sub


' ============================================================
' Test_D1_Guard (D-1 論点7)
'   in-callback ガードが効くこと、ResetCallbackGuard で復帰できることを確認する。
'   実際のイベントハンドラ内から呼ぶ経路は再現しにくいので、
'   Debug_SetInCallback で同じ状態を作って確かめる。
' ============================================================
Public Sub Test_D1_Guard()
    Dim p As Wv2Pane
    Dim r As String

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Debug.Print "Test_D1_Guard: アクティブな Pane がありません。" & _
                    "先に StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Debug.Print ""
    Debug.Print "================ Test_D1_Guard 開始 ================"

    Debug.Print "  1) 通常状態 (深さ=" & p.InCallbackDepth & ") で EvalSync"
    r = p.EvalSync("1 + 1", 5)
    If p.LastEvalOk And r = "2" Then
        Debug.Print "     [OK  ] 2 が返った"
    Else
        Debug.Print "     [FAIL] r=" & r & " err=" & p.LastEvalError
    End If

    Debug.Print "  2) ハンドラ内にいる状態を作って EvalSync (拒否されるのが正常)"
    p.Debug_SetInCallback 1
    r = p.EvalSync("1 + 1", 5)
    If (Not p.LastEvalOk) And p.LastEvalError = "in-callback" Then
        Debug.Print "     [OK  ] in-callback で拒否された (固まらずに即戻った)"
    Else
        Debug.Print "     [FAIL] r=" & r & " ok=" & p.LastEvalOk & " err=" & p.LastEvalError
    End If

    Debug.Print "  3) ResetCallbackGuard で復帰させて再実行"
    p.ResetCallbackGuard
    r = p.EvalSync("1 + 1", 5)
    If p.LastEvalOk And r = "2" Then
        Debug.Print "     [OK  ] 復帰した (深さ=" & p.InCallbackDepth & ")"
    Else
        Debug.Print "     [FAIL] r=" & r & " err=" & p.LastEvalError
    End If

    Debug.Print "================ Test_D1_Guard 終了 ================"
    Debug.Print ""
End Sub


' ============================================================
' Test_D1_Help (D-1 の手順)
' ============================================================
Public Sub Test_D1_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-1 検証手順 (EvalSync = ExecuteScript の同期取得)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) StartWebView2_Full でブラウザを起動する"
    Debug.Print "  2) 適当な実ページを開く (Google などで可)"
    Debug.Print "  3) ★イベントバーストが静まるまで待つ★ (仕様事実 20)"
    Debug.Print "     イミディエイトのログが止まってから次に進むこと。"
    Debug.Print ""
    Debug.Print "  --- 実行 ---"
    Debug.Print "  4) Test_D1_Eval  … 型・日本語・例外・タイムアウトを一括で流す"
    Debug.Print "  5) Test_D1_Guard … in-callback ガードの発火を確認する"
    Debug.Print ""
    Debug.Print "  ★実行中はブレーク/ステップ実行しないこと★"
    Debug.Print "    EvalSync は DoEvents を回して待つので、その最中に止めると"
    Debug.Print "    仕様事実 20 の窓を踏む。"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D1_Eval) ---"
    Debug.Print "  ・(1) 全ケースが [OK  ] であること"
    Debug.Print "      ※文字列は★引用符付き★で返るのが正常 (D-2 で剥がす)"
    Debug.Print "      ※日本語が化けていないこと (uXXXX の復号ができている証拠)"
    Debug.Print "      ※undefined は文字列 undefined が返るのが正常"
    Debug.Print "  ・(2) document.title と location.href が実際のページと一致すること"
    Debug.Print "  ・(3) 失敗系も [OK  ] と出ること (期待どおり失敗した、の意味)"
    Debug.Print "      ReferenceError / Error: boom / syntax-error-or-null が見えるはず"
    Debug.Print "  ・(4) タイムアウトが★1 秒程度で戻る★こと (5 秒待たない)"
    Debug.Print "      その数秒後に ★破棄済み★ の遅延到着ログが出ること"
    Debug.Print "      最後の 回復 が [OK  ] であること"
    Debug.Print "  ・最後の in-callback 深さが 0 であること"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D1_Guard) ---"
    Debug.Print "  ・2) が★即座に★戻ること (固まらない)。これが論点7 の目的"
    Debug.Print "  ・3) で復帰し、深さが 0 に戻ること"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (D-1 は Wv2Pane のコールバックを包み直したため) ---"
    Debug.Print "  6) タブの追加・切替・閉じる・D&D 並べ替え"
    Debug.Print "  7) URL 入力・戻る・進む・リロード (タイトルが更新されること)"
    Debug.Print "  8) リンクの新タブ展開 (Test_9_26_PostProbe の POST リンクも)"
    Debug.Print "  9) 設定タブ (歯車) が開き、検索エンジンを変更できること"
    Debug.Print "     ※ 6?9 は View_On* を Core に分離した影響を見るためのもの。"
    Debug.Print "       どれか 1 つでも動かなければガードのラッパーを疑うこと。"
    Debug.Print ""
    Debug.Print "  --- 手で試したいとき ---"
    Debug.Print "  ?UserForm1.GetActivePane.EvalSync(""document.title"")"
    Debug.Print "  ?UserForm1.GetActivePane.EvalSync(""document.querySelectorAll('a').length"")"
    Debug.Print "  ?UserForm1.GetActivePane.LastEvalError"
    Debug.Print ""
End Sub


' ============================================================
' Test_D2_Find (D-2 本体の検証)
'   検証ページを新しいタブに開き、要素の取得と読み取りを一括で流す。
' ============================================================
Public Sub Test_D2_Find()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim el2 As Wv2Element

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "Test_D2_Find: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Debug.Print "Test_D2_Find: タブの生成に失敗しました。"
        Exit Sub
    End If

    If Not D2WaitTitle(p, "D-2 プローブ", 10) Then
        Debug.Print "Test_D2_Find: 検証ページの読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Debug.Print ""
    Debug.Print "================ Test_D2_Find 開始 ================"
    Debug.Print "  世代 (取得前は空文字が正常): [" & p.CurrentDomGen & "]"
    Debug.Print ""
    Debug.Print "  --- (1) 取得できること ---"

    Set el = p.GetElementById("ttl")
    D2Bool "GetElementById(ttl) が Nothing でない", Not (el Is Nothing)
    If el Is Nothing Then
        Debug.Print "  以降の検証は続けられません。中止します。"
        Exit Sub
    End If
    Debug.Print "        handle=" & el.Handle & " gen=" & el.Generation
    Debug.Print "        Pane 側の世代キャッシュ: " & p.CurrentDomGen
    D2Bool "取得直後は stale でない", (el.IsStale = False)

    Debug.Print ""
    Debug.Print "  --- (2) 読み取り 4 種 + 属性 ---"
    D2Eq "TagName (大文字で返る)", el, el.TagName, "H1"
    D2Eq "InnerText (日本語)", el, el.InnerText, "D-2 要素レジストリのプローブ"

    Set el = D2El(p, "box")
    D2Eq "GetAttribute(class)", el, el.GetAttribute("class"), "card"
    D2Eq "GetAttribute(data-note) 日本語属性", el, _
         el.GetAttribute("data-note"), "属性の値 (日本語)"
    D2Eq "InnerHTML (★仕様事実30 の復号★)", el, _
         el.InnerHTML, "<span class=""tag"">内側</span>テキスト"

    Set el = D2El(p, "esc")
    D2Eq "記号の混在", el, el.InnerText, "記号: < > & "" ' \ の混在"

    Set el = D2El(p, "pre")
    D2Eq "改行を含むテキスト (\n の復号)", el, el.InnerText, "1 行目" & vbLf & "2 行目"

    Debug.Print ""
    Debug.Print "  --- (3) 入力要素の value ---"
    Set el = D2El(p, "txt")
    D2Eq "input の TagName", el, el.TagName, "INPUT"
    D2Eq "input の Value", el, el.value, "初期値"
    D2Eq "input の GetAttribute(value)", el, el.GetAttribute("value"), "初期値"

    Set el = D2El(p, "area")
    D2Eq "textarea の Value", el, el.value, "テキストエリアの値"

    Set el = D2El(p, "sel")
    D2Eq "select の Value (selected の option)", el, el.value, "b"

    Debug.Print ""
    Debug.Print "  --- (4) 空・不在は★成功して空文字★になること (LastOk=True) ---"
    Set el = D2El(p, "empty")
    D2Eq "空要素の InnerText", el, el.InnerText, ""

    Set el = D2El(p, "lnk")
    D2Eq "a 要素の Value (value を持たない)", el, el.value, ""
    D2Eq "存在しない属性", el, el.GetAttribute("data-nothing"), ""
    D2Eq "href 属性", el, el.GetAttribute("href"), "https://example.com/path?x=1"

    Debug.Print ""
    Debug.Print "  --- (5) QuerySelector (★セレクタ内のシングルクォート★) ---"
    Set el2 = p.QuerySelector("input[name='q']")
    D2Bool "QuerySelector(input[name='q']) が取れる", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        D2Eq "同じ要素が取れている", el2, el2.value, "初期値"
    End If

    Set el2 = p.QuerySelector("#box .tag")
    D2Bool "子孫セレクタが効く", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        D2Eq "子孫セレクタの InnerText", el2, el2.InnerText, "内側"
    End If

    Debug.Print ""
    Debug.Print "  --- (6) 見つからない / 失敗の区別 (論点4) ---"
    Set el2 = p.QuerySelector("#nothing-here")
    D2Bool "存在しないセレクタ → Nothing", (el2 Is Nothing)
    D2Bool "  かつ LastEvalOk = True (本当に無い、の意味)", (p.LastEvalOk = True)

    Set el2 = p.QuerySelector("###")
    D2Bool "不正なセレクタ → Nothing", (el2 Is Nothing)
    D2Bool "  かつ LastEvalOk = False (失敗、の意味)", (p.LastEvalOk = False)
    Debug.Print "        LastEvalError = " & p.LastEvalError

    Debug.Print ""
    Debug.Print "  --- (7) ClearElementRegistry (論点3) ---"
    Set el = D2El(p, "ttl")
    D2Bool "掃除前は stale でない", (el.IsStale = False)
    D2Bool "ClearElementRegistry が成功する", p.ClearElementRegistry()
    D2Bool "掃除後は stale になる", (el.IsStale = True)
    Set el = p.GetElementById("ttl")
    D2Bool "掃除後も取り直せる", Not (el Is Nothing)
    If Not el Is Nothing Then
        D2Eq "取り直した要素が読める", el, el.TagName, "H1"
    End If

    Debug.Print ""
    Debug.Print "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    Debug.Print "================ Test_D2_Find 終了 ================"
    Debug.Print ""
End Sub


' ============================================================
' Test_D2_Stale (D-2 論点2 の検証: ページ遷移で世代が変わること)
' ============================================================
Public Sub Test_D2_Stale()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim genBefore As String
    Dim genAfter As String
    Dim v As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "Test_D2_Stale: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Debug.Print "Test_D2_Stale: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-2 プローブ", 10) Then
        Debug.Print "Test_D2_Stale: 検証ページの読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Debug.Print ""
    Debug.Print "================ Test_D2_Stale 開始 ================"

    Set el = p.GetElementById("ttl")
    If el Is Nothing Then
        Debug.Print "  [FAIL] 要素が取れないので中止します。"
        Exit Sub
    End If
    genBefore = el.Generation
    Debug.Print "  1) 遷移前: handle=" & el.Handle & " gen=" & genBefore
    D2Eq "     読める", el, el.TagName, "H1"
    D2Bool "     stale でない", (el.IsStale = False)

    Debug.Print "  2) 同じタブを別のページへ遷移させる"
    p.View_NavigateToString BuildD2SecondHtml()
    If Not D2WaitTitle(p, "D-2 プローブ 2 枚目", 10) Then
        Debug.Print "  [FAIL] 2 枚目の読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Debug.Print "  3) 遷移後: 古いハンドルの状態を見る"
    D2Bool "     IsStale = True になる", (el.IsStale = True)
    v = el.TagName
    D2Bool "     読み取りは空文字 + LastOk=False", (Len(v) = 0 And el.LastOk = False)
    Debug.Print "        LastError = " & el.LastError
    D2Bool "     LastError が stale であること", (el.LastError = "stale")

    Debug.Print "  4) 新しいページで取り直せる"
    Set el = p.GetElementById("second")
    D2Bool "     取得できる", Not (el Is Nothing)
    If Not el Is Nothing Then
        genAfter = el.Generation
        D2Eq "     読める", el, el.InnerText, "2 枚目のページ"
        Debug.Print "        新しい gen=" & genAfter
        D2Bool "     世代が変わっている", (genAfter <> genBefore)
    End If

    Debug.Print ""
    Debug.Print "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    Debug.Print "================ Test_D2_Stale 終了 ================"
    Debug.Print ""
End Sub


' ============================================================
' D2Eq / D2Bool (D-2 の判定ヘルパー)
'   D2Eq  … 値の一致を見る。LastOk が False なら失敗理由も出す。
'   D2Bool… 条件だけを見る。
' ============================================================
Private Sub D2Eq(ByVal label As String, _
                 ByVal el As Wv2Element, _
                 ByVal got As String, _
                 ByVal want As String)
    If got = want Then
        Debug.Print "  [OK  ] " & label
    Else
        Debug.Print "  [FAIL] " & label
        Debug.Print "         期待: [" & want & "]"
        Debug.Print "         実際: [" & got & "]"
    End If

    If el Is Nothing Then Exit Sub
    If Not el.LastOk Then
        Debug.Print "         ※ LastOk=False err=" & el.LastError
    End If
End Sub

Private Sub D2Bool(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        Debug.Print "  [OK  ] " & label
    Else
        Debug.Print "  [FAIL] " & label
    End If
End Sub


' ============================================================
' D2El (D-2: 要素取得のガード付きヘルパー)
'
'   取れなければ FAIL を出したうえで★Pane 未設定のダミー★を返す。
'   ダミーは読み取ると LastOk=False / LastError="no-pane" を返すだけなので、
'   検証が実行時エラー 91 で止まらずに最後まで流れる。
'   (Wv2Element を素で New したときの経路をそのまま利用している)
' ============================================================
Private Function D2El(ByVal p As Wv2Pane, ByVal elementId As String) As Wv2Element
    Dim e As Wv2Element

    Set e = p.GetElementById(elementId)
    If e Is Nothing Then
        Debug.Print "  [FAIL] 要素 #" & elementId & " が取得できない " & _
                    "(LastEvalOk=" & p.LastEvalOk & " err=" & p.LastEvalError & ")"
        Set e = New Wv2Element
    End If

    Set D2El = e
End Function


' ============================================================
' D2WaitTitle (D-2: ページの読み込み完了を title で待つ)
'
'   NavigateToString は非同期なので、EvalSync で document.title を読んで
'   期待するタイトルになるまで待つ。IsPageLoaded を使わないのは、
'   同じ Pane を 2 回遷移させる Test_D2_Stale で「前のページの完了」と
'   区別できないため。
'
'   ★この待ちループ中もブレーク/ステップ実行しないこと★ (仕様事実20)
' ============================================================
Private Function D2WaitTitle(ByVal p As Wv2Pane, _
                             ByVal wantTitle As String, _
                             ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim res As String
    Dim cur As String

    D2WaitTitle = False
    t0 = Timer

    Do
        DoEvents
        res = p.EvalSync("document.title", 3)
        If p.LastEvalOk Then
            cur = Wv2Json.JsonUnescape(res)
            If cur = wantTitle Then
                D2WaitTitle = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then
            Debug.Print "D2WaitTitle: タイムアウト (期待=" & wantTitle & _
                        " 実際=" & cur & " err=" & p.LastEvalError & ")"
            Exit Function
        End If
    Loop
End Function


' ============================================================
' BuildD2ProbeHtml (D-2 の検証ページ、論点8 案b)
'
'   ★静的な HTML 部分の引用符は VBA の "" で書く★ (JS ではないので問題ない)
'   ★JS は一切使っていない★ (このページは DOM を提供するだけ)
' ============================================================
Private Function BuildD2ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>D-2 プローブ</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .tag{color:#6ea8fe;font-weight:600;}" & vbLf
    s = s & "  pre{background:#0b0e15;padding:10px 12px;border-radius:8px;margin:12px 0;}" & vbLf
    s = s & "  input,textarea,select{font-size:14px;padding:6px 8px;margin:4px 0;" & _
            "background:#0b0e15;color:#e8eaed;border:1px solid rgba(255,255,255,.18);" & _
            "border-radius:6px;}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<h1 id=""ttl"">D-2 要素レジストリのプローブ</h1>" & vbLf
    s = s & "<p class=""note"">このページは Test_D2_Find / Test_D2_Stale 専用です。" & _
            "外部サイトに依存しません。</p>" & vbLf
    s = s & "<div id=""box"" class=""card"" data-note=""属性の値 (日本語)"">" & _
            "<span class=""tag"">内側</span>テキスト</div>" & vbLf
    s = s & "<p id=""esc"">記号: &lt; &gt; &amp; &quot; ' \ の混在</p>" & vbLf
    s = s & "<pre id=""pre"">1 行目" & vbLf & "2 行目</pre>" & vbLf
    s = s & "<div class=""card"">" & vbLf
    s = s & "  <input id=""txt"" type=""text"" name=""q"" value=""初期値"">" & vbLf
    s = s & "  <textarea id=""area"" rows=""2"">テキストエリアの値</textarea>" & vbLf
    s = s & "  <select id=""sel""><option value=""a"">A</option>" & _
            "<option value=""b"" selected>B</option></select>" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<p><a id=""lnk"" href=""https://example.com/path?x=1"" title=""リンク"">" & _
            "リンク</a></p>" & vbLf
    s = s & "<div id=""empty""></div>" & vbLf
    s = s & "</body></html>"

    BuildD2ProbeHtml = s
End Function


' ============================================================
' BuildD2SecondHtml (D-2: 遷移先の 2 枚目。世代が変わることを見るためだけ)
' ============================================================
Private Function BuildD2SecondHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>D-2 プローブ 2 枚目</title>" & vbLf
    s = s & "<style>body{font-family:'Segoe UI','Meiryo',sans-serif;background:#141b14;" & _
            "color:#e8eaed;padding:32px;}</style></head><body>" & vbLf
    s = s & "<h1 id=""second"">2 枚目のページ</h1>" & vbLf
    s = s & "<p>ここに来たら、1 枚目で取った Wv2Element はすべて stale になります。</p>" & vbLf
    s = s & "</body></html>"

    BuildD2SecondHtml = s
End Function


' ============================================================
' Test_D2_Help (D-2 の手順)
' ============================================================
Public Sub Test_D2_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-2 検証手順 (要素レジストリ + Wv2Element の読み取り)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) StartWebView2_Full でブラウザを起動する"
    Debug.Print "  2) ★イベントバーストが静まるまで待つ★ (仕様事実 20)"
    Debug.Print "     イミディエイトのログが止まってから次に進むこと。"
    Debug.Print "     ※検証ページは自前 HTML なので、外部サイトを開く必要はない。"
    Debug.Print ""
    Debug.Print "  --- 実行 ---"
    Debug.Print "  3) Test_D2_Find  … 取得・読み取り・不在・掃除を一括で流す"
    Debug.Print "  4) Test_D2_Stale … ページ遷移で世代が変わることを確認する"
    Debug.Print "     どちらも★新しいタブを 1 枚開く★ (実行後は手で閉じてよい)"
    Debug.Print ""
    Debug.Print "  ★実行中はブレーク/ステップ実行しないこと★"
    Debug.Print "    EvalSync が DoEvents を回して待つので、その最中に止めると"
    Debug.Print "    仕様事実 20 の窓を踏む。"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D2_Find) ---"
    Debug.Print "  ・(1)?(7) の全行が [OK  ] であること"
    Debug.Print "  ・(2) の InnerHTML が★タグごと正しく★出ること"
    Debug.Print "      仕様事実30 (< が \u003C で届く) の復号が効いている証拠。"
    Debug.Print "      ここが FAIL なら Wv2Json.JsonUnescapeAt を疑う。"
    Debug.Print "  ・(4) が [OK  ] であること = 「空」と「失敗」を区別できている"
    Debug.Print "  ・(6) の 2 つが★対で★ [OK  ] であること (論点4 の規約):"
    Debug.Print "      存在しない  → Nothing + LastEvalOk=True"
    Debug.Print "      不正な指定  → Nothing + LastEvalOk=False"
    Debug.Print "  ・最後の in-callback 深さが 0 であること"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D2_Stale) ---"
    Debug.Print "  ・3) で IsStale=True、読み取りが空文字 + LastError=stale になること"
    Debug.Print "  ・4) で新しい世代の要素が取れること (gen が前と違うこと)"
    Debug.Print "    ここが論点2 の核心 (世代は JS が発行し、遷移で必ず変わる)。"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (D-2 は Wv2Pane と Wv2Json を触ったため) ---"
    Debug.Print "  5) Test_D1_Eval  … EvalSync が壊れていないこと"
    Debug.Print "     ★JsonPickStr を Wv2Json へ移設したので、ここは必ず流すこと★"
    Debug.Print "  6) Test_D1_Guard … in-callback ガードが効くこと"
    Debug.Print "  7) タブの追加・切替・閉じる・D&D 並べ替え"
    Debug.Print "  8) 設定タブ (歯車) が開き、検索エンジンを変更できること"
    Debug.Print "     ※ Wv2Json は TabBar / NavBar / SettingsBridge も使っている。"
    Debug.Print ""
    Debug.Print "  --- 手で試したいとき ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  Set el = p.QuerySelector(""input[name='q']"")"
    Debug.Print "  ?el.TagName"
    Debug.Print "  ?el.Value"
    Debug.Print "  ?el.IsStale"
    Debug.Print "  ?p.CurrentDomGen"
    Debug.Print ""
End Sub


' ============================================================
' Test_K1_Log  (K-1 段階)
'
'   Wv2Log をイミディエイトから 1 発で検証する。
'   ブラウザを起動していなくても動く (WebView2 に依存しない)。
' ============================================================
Public Sub Test_K1_Log()
    Dim k1Total As Long
    Dim k1Pass As Long

    Debug.Print String$(64, "=")
    Debug.Print "K-1 検証: Wv2Log (デバッグログのファイル出力)"
    Debug.Print String$(64, "=")

    ' --- 準備: 新しいログを開く ---
    Wv2Log.LogLevel = LOG_DEBUG
    Wv2Log.LogEcho = True
    Wv2Log.LogStart

    Dim k1Path As String
    k1Path = Wv2Log.LogPath
    Debug.Print "  ログ: " & k1Path
    Debug.Print ""

    K1Bool "(1) LogPath が取れる", (Len(k1Path) > 0), k1Total, k1Pass
    K1Bool "(2) ファイルが実在する", (Len(Dir(k1Path)) > 0), k1Total, k1Pass

    ' --- 4 レベルを書く ---
    Wv2Log.LogE "K-1: ERROR の行"
    Wv2Log.LogW "K-1: WARN の行"
    Wv2Log.LogI "K-1: INFO の行"
    Wv2Log.LogD "K-1: DEBUG の行"

    ' --- しきい値で切れるか ---
    Wv2Log.LogLevel = LOG_WARN
    Wv2Log.LogI "K-1: この INFO は出ないはず"
    Wv2Log.LogD "K-1: この DEBUG は出ないはず"
    Wv2Log.LogLevel = LOG_DEBUG
    Wv2Log.LogD "K-1: しきい値を戻した"

    ' --- 実行時に組み立てた Unicode (ソースは CP932 のまま) ---
    Dim k1Emoji As String
    k1Emoji = ChrW$(&HD83D&) & ChrW$(&HDE00&)
    Wv2Log.LogD "K-1: サロゲートペア [" & k1Emoji & "]"
    Wv2Log.LogD "K-1: 日本語と記号 ★ ← → ① ～ ―"

    ' --- 再入ガード: 積んで、あとでまとめて流れるか ---
    Wv2Log.Debug_SetLogDepth 1
    Wv2Log.LogD "K-1: 再入中の 1 行目"
    Wv2Log.LogD "K-1: 再入中の 2 行目"
    Wv2Log.Debug_SetLogDepth 0
    Wv2Log.LogD "K-1: 再入から抜けた"

    Wv2Log.LogFlush

    ' --- 読み戻して照合 ---
    '   ★「無いことの確認」は、読み戻せて初めて意味を持つ★
    '   読めていないのに InStr(空, x) = 0 で通ってしまうと、空振りで合格する。
    '   K-1 の実機検証で実際にこれをやってしまったので、読めた場合だけ照合する。
    Dim k1Text As String

    ' (3) 実行中のまま読めるか (エディタで開けるかの確認)
    '   ★ADODB.Stream では測れない★ ADODB は共有モードに関わらず、
    '   他のハンドルが開いているファイルを開けない (K-1 の実機検証で実測)。
    '   共有モードを明示する読み手なら開けるので、VBA 自身の
    '   Open ... For Binary Access Read Shared で確かめる。
    Debug.Print ""
    K1Bool "(3) 実行中でもログを読める (共有読み取り)", _
           K1CanReadWhileOpen(k1Path), k1Total, k1Pass

    ' 本文の照合は ADODB で読む。開いたままでは開けないので先に閉じる
    Wv2Log.LogStop
    k1Text = K1ReadUtf8(k1Path)
    K1Bool "(4) ログを読み戻せた", (Len(k1Text) > 0), k1Total, k1Pass

    If Len(k1Text) = 0 Then
        Debug.Print ""
        Debug.Print "  ★ログを読み戻せないので、以降の照合は打ち切る★"
        Debug.Print "     (空文字に対する InStr は何でも 0 を返すので、"
        Debug.Print "      ここで続けると空振りで合格してしまう)"
    Else
        K1Bool "(5) ERROR の行がある", (InStr(k1Text, "K-1: ERROR の行") > 0), k1Total, k1Pass
        K1Bool "(6) WARN の行がある", (InStr(k1Text, "K-1: WARN の行") > 0), k1Total, k1Pass
        K1Bool "(7) INFO の行がある", (InStr(k1Text, "K-1: INFO の行") > 0), k1Total, k1Pass
        K1Bool "(8) DEBUG の行がある", (InStr(k1Text, "K-1: DEBUG の行") > 0), k1Total, k1Pass
        K1Bool "(9) しきい値で切った INFO が無い", (InStr(k1Text, "この INFO は出ないはず") = 0), k1Total, k1Pass
        K1Bool "(10) しきい値で切った DEBUG が無い", (InStr(k1Text, "この DEBUG は出ないはず") = 0), k1Total, k1Pass
        K1Bool "(11) サロゲートペアが無傷", (InStr(k1Text, k1Emoji) > 0), k1Total, k1Pass
        K1Bool "(12) 日本語が無傷", (InStr(k1Text, "日本語と記号 ★ ← → ① ～ ―") > 0), k1Total, k1Pass
        K1Bool "(13) 再入中の 2 行が後から流れている", _
               (InStr(k1Text, "K-1: 再入中の 1 行目") > InStr(k1Text, "K-1: 再入から抜けた")), _
               k1Total, k1Pass
    End If

    ' --- 連番が飛んでいないか (K-1 の動機そのもの) ---
    Dim k1Lines As Variant
    Dim k1Idx As Long
    Dim k1Num As Long
    Dim k1Prev As Long
    Dim k1Ok As Boolean
    ' ★順序ではなく欠番の有無を見る★
    '   再入ガードの検証で保留キューを後から流すため、ファイル上の順序は
    '   わざと入れ替わる。K-1 の動機は「流れたことに気づけない」ことなので、
    '   見るべきは連番の抜けであって並び順ではない。
    Dim k1Seq As Long
    k1Lines = Split(k1Text, vbCrLf)
    k1Prev = 0
    k1Seq = 0
    For k1Idx = LBound(k1Lines) To UBound(k1Lines)
        If Len(k1Lines(k1Idx)) >= 6 Then
            If IsNumeric(Left$(k1Lines(k1Idx), 6)) Then
                k1Num = CLng(Left$(k1Lines(k1Idx), 6))
                k1Seq = k1Seq + 1
                If k1Num > k1Prev Then k1Prev = k1Num
            End If
        End If
    Next k1Idx
    k1Ok = (k1Seq = k1Prev) And (k1Seq > 0)
    If Len(k1Text) > 0 Then
        K1Bool "(14) 連番に欠番が無い (" & k1Seq & " 行 / 最大 " & k1Prev & ")", _
               k1Ok, k1Total, k1Pass
    End If

    Debug.Print ""
    Debug.Print "  結果: " & k1Pass & " / " & k1Total & " 合格"
    If k1Pass = k1Total Then
        Debug.Print "  ★K-1 合格★"
    Else
        Debug.Print "  ★不合格あり。上の [FAIL] を見ること★"
    End If
    Debug.Print "  ログの実物: " & k1Path
    Debug.Print String$(64, "=")
End Sub


Private Sub K1Bool(ByVal label As String, ByVal cond As Boolean, _
                   ByRef k1Total As Long, ByRef k1Pass As Long)
    k1Total = k1Total + 1
    If cond Then
        k1Pass = k1Pass + 1
        Debug.Print "  [OK  ] " & label
    Else
        Debug.Print "  [FAIL] " & label
    End If
End Sub


' 開かれたままのファイルを共有モードで読めるかどうかだけを見る
'   ★ADODB.Stream はこの用途に使えない★ 共有モードに関わらず
'   他のハンドルが開いているファイルを開けない (K-1 で実測)。
Private Function K1CanReadWhileOpen(ByVal k1Path As String) As Boolean
    On Error GoTo eh
    Dim k1Handle As Long
    Dim k1Bytes() As Byte
    k1Handle = FreeFile
    Open k1Path For Binary Access Read Shared As #k1Handle
    If LOF(k1Handle) > 0 Then
        ReDim k1Bytes(0 To LOF(k1Handle) - 1)
        Get #k1Handle, 1, k1Bytes
        K1CanReadWhileOpen = True
    End If
    Close #k1Handle
    Exit Function
eh:
    On Error Resume Next
    Close #k1Handle
End Function


' UTF-8 / BOM なしのファイルを読む (検証専用。ADODB でよい)
Private Function K1ReadUtf8(ByVal k1Path As String) As String
    On Error GoTo eh
    Dim k1Stream As Object
    Set k1Stream = CreateObject("ADODB.Stream")
    k1Stream.Type = 2
    k1Stream.Charset = "UTF-8"
    k1Stream.Open
    k1Stream.LoadFromFile k1Path
    K1ReadUtf8 = k1Stream.ReadText
    k1Stream.Close
    Exit Function
eh:
    Debug.Print "K1ReadUtf8: 失敗 (" & Err.Number & ") " & Err.Description
End Function


' ============================================================
' Test_K1_Help  (K-1 段階)
' ============================================================
Public Sub Test_K1_Help()
    Debug.Print String$(64, "=")
    Debug.Print "K-1 (デバッグログのファイル出力) の使い方"
    Debug.Print String$(64, "=")
    Debug.Print ""
    Debug.Print "  1) Test_K1_Log      … 自動検証を 1 発で回す"
    Debug.Print "  2) ?Wv2Log.LogPath  … 今のログファイルの場所"
    Debug.Print "  3) Wv2Log.LogStart  … 新しいログに切り替える (検証の仕切り直し)"
    Debug.Print ""
    Debug.Print "  ★D 軸の検証手順★"
    Debug.Print "    Wv2Log.LogStart を打ってから Test_D2_Find を回すと、"
    Debug.Print "    そのテスト 1 回分だけが 1 ファイルに閉じる。"
    Debug.Print "    イミディエイトが流れても、ファイルには全部残っている。"
    Debug.Print ""
    Debug.Print "  ★行頭の 6 桁は連番★ 欠番があればログが落ちている。"
    Debug.Print "    これが K-1 を作った動機 (流れたことに気づけないのが危険)。"
    Debug.Print ""
    Debug.Print "  設定:"
    Debug.Print "    Wv2Log.LogLevel = LOG_ERROR / LOG_WARN / LOG_INFO / LOG_DEBUG"
    Debug.Print "    Wv2Log.LogEcho  = False   … イミディエイトへの併記を止める"
    Debug.Print ""
    Debug.Print "  ★実行中のログをエディタで開ける★ (Shared で開いているため)。"
    Debug.Print "    ただし ADODB.Stream だけは開けない。共有モードに関わらず"
    Debug.Print "    他のハンドルが開いているファイルを開けない仕様 (K-1 で実測)。"
    Debug.Print ""
    Debug.Print "  ログは起動ごとに 1 本。20 本を超えると古いものから消える。"
    Debug.Print "  置き場所: %APPDATA%\Wv2Browser\logs\"
    Debug.Print String$(64, "=")
End Sub


' ============================================================
' Test_K2_Help  (K-2 段階)
'
'   設定タブは WebView2 の実起動が要るので自動化しきれない。
'   画面で確かめる手順を出す。
' ============================================================
Public Sub Test_K2_Help()
    Debug.Print String$(64, "=")
    Debug.Print "K-2 検証: 設定タブが検索で潰れるバグ"
    Debug.Print String$(64, "=")
    Debug.Print ""
    Debug.Print "  ★先に Wv2Log.LogStart を打つと、この検証 1 回分が 1 ファイルに閉じる★"
    Debug.Print "    ログの場所は ?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  【1】 直った本体"
    Debug.Print "    1) ブラウザを起動する"
    Debug.Print "    2) タブバー右端の歯車を押して設定タブを開く"
    Debug.Print "    3) その設定タブがアクティブなまま、アドレスバーに 適当な検索語 を入れる"
    Debug.Print "       → 検索結果へ遷移する (これは正常。設定タブも普通のタブ)"
    Debug.Print "    4) ★もう一度 歯車 を押す★"
    Debug.Print "       → ★新しい設定タブが開けば直っている★"
    Debug.Print "       → 何も起きない / 検索結果のタブへ切り替わるだけなら直っていない"
    Debug.Print ""
    Debug.Print "  【2】 重複防止が壊れていないこと (第9.19 の回帰)"
    Debug.Print "    5) 設定タブを開いたまま歯車を 3～4 回連打する"
    Debug.Print "       → 設定タブは常に 1 枚のまま。増えないこと"
    Debug.Print ""
    Debug.Print "  【3】 アドレスバー以外の経路でも直っていること"
    Debug.Print "    6) 設定タブでリンクのあるページへ遷移し、歯車 → 新しい設定タブが開く"
    Debug.Print "    7) 設定タブでリロードしても設定画面が保たれること"
    Debug.Print ""
    Debug.Print "  【4】 ★設定ビューが 2 枚でも値が食い違わないこと (今回の新機能)★"
    Debug.Print "    8) 歯車で設定タブ A を開く"
    Debug.Print "    9) A のアドレスバーで適当なページへ遷移する"
    Debug.Print "   10) 歯車を押して設定タブ B を開く"
    Debug.Print "   11) A に切り替えて ★戻る★ を押す → A も設定画面に戻る (2 枚になる)"
    Debug.Print "   12) B に切り替えて、検索エンジンを別のものに変える"
    Debug.Print "   13) ★A に切り替える★"
    Debug.Print "       → ★A の選択表示とプレビューも、変更後のエンジンになっていること★"
    Debug.Print "       → A が古いエンジンを選択したままなら同期が効いていない"
    Debug.Print ""
    Debug.Print "  【5】 回帰 (壊していないこと)"
    Debug.Print "   14) 設定タブでエンジンを変え、通常タブで検索して反映されること"
    Debug.Print "   15) カスタム URL の入力・適用が従来どおり動くこと"
    Debug.Print "       ※ カスタム入力欄の中身は同期で上書きしない仕様 (入力途中を壊さないため)"
    Debug.Print ""
    Debug.Print "  ログに出るもの:"
    Debug.Print "    Wv2Browser.OpenSettingsTab: 設定画面を表示中のタブ(index N)を..."
    Debug.Print "    Wv2Browser.SyncSettingsViewsNow: 設定ビュー N 枚へ <engine> を反映した"
    Debug.Print "    Wv2Browser.IsSettingsView: 判定できないので... (読み込み中などで正常)"
    Debug.Print String$(64, "=")
End Sub

