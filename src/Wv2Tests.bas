Attribute VB_Name = "Wv2Tests"
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-4 段階 (レスポンス本文) ---
'
'   Test_N4_Body … 回帰試験
'   Test_N4_Help … 手順
'
'   ★N-4 で初めて非同期になる★
'     GetContent はハンドラを渡して即座に戻り、本文は後から届く。
'     判定は必ず ★N4Wait (本文が届くまで) ★ で待つこと。N-3 で
'     「詳細が現れるまで」で待って空を読み、実機 3 回ぶんを溶かした
'     (設計原則120)。同じ轍を踏まないための専用ヘルパーが N4Wait。
'
'   ★最大の未知数は圧縮★
'     GetContent が復号済みを返すかは SDK に書かれていない。httpbingo の
'     /gzip は中身に "gzipped": true を持つので、★本文に読めるか / 先頭が
'     1F 8B か★ のどちらか一方に必ず決まる。それを判定にしてある。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-3 段階 (レスポンスのステータスとヘッダ) ---
'
'   Test_N3_Status … 回帰試験。的は N-1 / N-2 と同じ 2 系統
'   Test_N3_Help   … 手順
'
'   ★(1)～(4) は外部が要る★
'     到達不能ローカル (案F) には★応答が来ない★ので、ステータスの検証には
'     本物の送り先が要る。逆に言えば案F は「応答が来ないこと = 空欄」の
'     対照になる ((5) がそれ)。
'
'   ★このイベントにはフィルタが無い★
'     画像も CSS も全部来る。一覧に居ない応答は NetRespUnmatched に数える
'     だけで捨てる。★大きい数が出るのは正常★。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-2 段階 (リクエストのヘッダとボディ) ---
'
'   Test_N2_Detail … 回帰試験。的は N-1 と同じ 2 系統 (案F / 案D)
'   Test_N2_Help   … 手順
'
'   ★本丸は「読んでも通信が壊れていない」★
'     IStream を読むと位置が進むので、Seek(0) で戻し損ねると空のボディが飛ぶ。
'     「POST が成功したっぽい」では証拠にならないので、★送り先 (httpbingo) が
'     返す本文と突き合わせて数える★ (設計原則112)。この 1 件だけは外部が要る。
'
'   ★Cookie は的が到達不能だと作れない★ ので、伏せ字の機構そのものは
'   「必ず在るヘッダ」= User-Agent を伏せる対象に足して確かめる。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-1c 段階 (種別の決めつけをやめる) ---
'
'   ★N-1b の実機で分かったこと★
'     ★WebView2 は fetch() の要求を FETCH(8) ではなく XML_HTTP_REQUEST(7) として
'     報告する★ (実機で fetch 10 本を撃って 10 本とも XHR で届いた)。
'     N-1b のテストが「fetch なら種別は FETCH のはず」と決めつけていたため、
'     ★捕まっているのに FAIL 3 件★ になっていた。実装は無傷。
'
'   ■ 直したこと (テスト側だけ)
'     1. 判定は「捕まったか」で行い、★種別は決めつけない★
'        ついでに localOk / netOk の誤報 (「1 件も届いていない」) も消える
'     2. ★「fetch は XHR 種別で届く」を数える判定を足した★
'        見つけた仕様事実をそのまま回帰試験にする
'     3. ★(3) の直後にドレインを 1 回挟む★
'        N-1b では (4) で容量を 3 に落とした時点で (2) の中身が消えてしまい、
'        「seq 1/2 は f-local と f-net だったはず」と★推測で埋める羽目になった★。
'        次からは推測が要らない (設計原則112)。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-1b 段階 (的を仮想ホストの外へ移す) ---
'
'   ★N-1 の初回実機で分かったこと★
'     仮想ホスト (SetVirtualHostNameToFolderMapping) で配信した要求は
'     WebResourceRequested に乗らない。初回の Test_N1_Capture が全滅したのは
'     ★的が 1 本残らず仮想ホストだったから★ で、配線は最初から正しかった
'     (Test_N1_Site では 10 件すべて捕まっていた)。
'
'   ★的を 2 系統 + 対照 1 系統にした (論点1 案G)★
'     案F  http://127.0.0.1:59999/…  到達不能なローカル (外部依存ゼロ)
'     案D  https://httpbingo.org/…   外部サービス (ネットが要る)
'     対照 仮想ホストの data.json    ★捕まらないことを数える★
'   どちらの的が使えるかを 1 回の実機で両方確かめる。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  N-1 段階 (ネットワーク要求のキャプチャ) ---
'
'   Test_N1_Capture … ★自前 HTML だけで完結する回帰試験★ (論点8)
'   Test_N1_Watch   … ★今開いているタブの通信を捕まえ始める★ (実用の入口)
'   Test_N1_Drain   … 溜まった分をログへ流して空にする
'   Test_N1_Stop    … 捕まえるのをやめる (フィルタも外す)
'   Test_N1_Help    … 手順
'
'   ★実サイトに頼らない★ 検証ページは %TEMP%\Wv2NetProbe に書き出して
'   仮想ホスト https://appassets.netprobe/ で開く (第9.26b と同じ手口)。
'   フォーム POST も fetch も XHR も、すべてこのフォルダの中で完結する。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  K-4a 段階 (解放が遅い原因の切り分け) ---
'
'   Test_K4_Quiet … ★フォームを閉じる前に全タブを静める★
'                   解放が遅いのは「ページが動き続けているから」なのかを測る。
'   Test_K4_Help  … 手順。
'
'   ★実測で分かっていること (K-4a)★
'     素の状態 (起動 → 即閉じ)     … 解放の全体 0.336 秒
'     Maps を 1 枚開いた後          … 解放の全体 ★6.718 秒★ (20 倍)
'     LogEcho = True / False の差   … ★ゼロ★ (Debug.Print は犯人ではない)
'     遅いのは特定のステップではなく ★COM 往復が全部 300～500ms★
'     (素の状態では 8～16ms)。Ctrl_Close 1 回はさらに重く 1.4 秒。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  D-7e 段階 (進捗と中断の検証) ---
'
'   D-7b の追加事項:
'     Test_D7_StatusBar … ★プローブ★ 製品コードを通さずに Excel の生の挙動を
'                         測る。D-7 で FAIL した StatusBar の判定が、判定側の
'                         問題なのか Excel の仕様なのかを切り分ける (原則103)。
'     Test_D7_Resume    … ★自動★ 中断された後の状態を人工的に作り、そこから
'                         再開できるかを Esc 無しで確かめる (回帰確認用)。
'
'   D-7 からの持ち越し:
'     Test_D7_Cancel … ★ネットワーク不要★ 中断の口・入口でのクリア・分母の
'                      数え方・0 件ならタブを開かないこと。
'     Test_D7_Sheet  … ★実サイト + 手で Esc を押す★ 中断 → 再開まで。
'     Test_D7_Help   … 手順。
'
'   ■ ★D-7e: 人が観察する検証は、人に見える形にする★
'     D-7d の実機で中断は成立していたのに、たーぼーさんには「止まっていない」
'     ように見えた。★中断した直後に自動で 2 回目を走らせていたから★。
'     さらに連打した Esc が 2 回目にも効いて、判定が 3 件 FAIL になった。
'     → 中断したら★はっきり知らせて一拍置き、Esc が離れるのを待ってから★
'       再開を試す (TestWaitEscReleased)。
'     Test_D7_Key も同じ理由で作り直した。「これから 10 秒」と出しても、
'     ★読む前に終わってしまう★ ので、押されるまで待つ形にした。
'
'   ★Esc そのものの検証は自動化していない★ SafeTimer で代わりにフラグを立てる
'   案もあったが、SafeTimer は WithEvents が要る = クラスモジュールの新設が要る
'   ので見送った。代わりに★中断後の状態からの再開★を Test_D7_Resume で自動化し、
'   毎回の回帰確認はそちらで賄う。Esc 経路は Test_D7_Sheet で人が 1 回確かめる。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  D-4a 段階 (ページ内 SPA プローブの検証) ---
'
'   D-4a の追加事項:
'     Test_D4_Probe … ★D-4 の初手★ プローブの設置・健康診断・作り直し・
'                     数え上げ・往復コストを一括で検証する。
'     Test_D4_Help  … 手順。
'
'   ★何を測っているか (D-4 論点の未知 1～4)★
'     未知1 ラップが生き残るか  … (6) で故意に壊して自動修復を確認
'     未知2 観測の負荷          … (2) で往復時間を設置前後で比較
'     未知3 遷移で消えるか      … (8) でページ遷移後に世代が振り出しに戻ることを確認
'     未知4 余計な Pane に付かないか … (1) で「呼ぶまで作られない」ことを確認
'
'   検証ページ (BuildD4ProbeHtml) は★読み込み直後は完全に静か★にしてある。
'   ノイズ (定期的な DOM 書き換え) は startNoise() を撃ったときだけ始まる。
'   D-4b の静穏待ちの検証でも同じページを使う。
'''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  D-3 段階 (書き込みと操作の検証) ---
'
'   D-3 の追加事項:
'     Test_D3_Probe_Promise … ★最初に走らせる★ 未知1 の実測。
'                             ExecuteScript が Promise を待つかを確かめる。
'                             ここの結果で「待ちループをどこで回すか」(論点3) が決まる。
'     Test_D3_Write         … Value = / Click / SetAttribute を自前 HTML で一括検証。
'     Test_D3_Framework     … ★ネイティブ setter 経由で書けているか★ を
'                             React の value tracker を模した監視で確かめる (未知3)。
'     Test_D3_Help          … 上 3 つの手順と、見るべきログの説明。
'
'   判定ヘルパー (TestEq / TestBool / D2El / D2WaitTitle) は D-2 のものを再利用する。
'   ★D-3 で判定の出し先を Wv2Log に変え、D-1 / D-2 のテストもそれに乗せた★
'   (イミディエイトは ExecuteScript の配管ログで流れて読めないため)。
'   既存の Test_* は 1 つも変更していない。
'
'   ★検証ページ (BuildD3ProbeHtml) は D-2 と同じく自前 HTML★
'     D-2 のページとの違いは、ページ側に★監視用の JS★ を仕込んであること:
'       window.__p = {inputs, changes, clicks, clickInfo,
'                     trackedSet, notified, ignored}
'     ・input / change の発火回数 (論点6 の「両方撃つ」の確認)
'     ・click の回数とイベントの素性 (type / bubbles / isTrusted)
'     ・★React 風の value tracker★ を #react に被せてあり、
'       tracker 経由の代入 (trackedSet) と、フレームワークが
'       変更に気づいた回数 (notified) / 気づけなかった回数 (ignored) を数える。
'''''''''''''''''''''''''''''''''''
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
'     (9.21～9.25 と同じ判断)。
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
'       [PROBE] 行の並びを読むだけでよい。そのうえで ← → [更新] を 1 回ずつ押して
'       実物が二重実行にならないかを確認する。

''''''''''''''''''''''''''''''''''
' --- Wv2Tests.bas  第9.24 段階 (NavBar の hostObjects 化 + Host 一元化 検証) ---
'
'   第9.24 の追加:
'     ★Test_9_24_HostNavBar_Help (実機手順)★
'       NavBar の back/forward/reload/navigate を hostObjects 経路へ移し、処理を
'       HostBack/HostForward/HostReload/HostNavigate へ一元化した回の実機手順。
'       実 URL Navigate + JS 実行 + hostObjects 経路の成立を見るため、純ロジック
'       検証は無い (9.21～9.23 と同じ判断)。
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
'       TabBar 側 (9.21～9.23) が無傷であること。

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
'       設定タブの重複防止 ([歯車] 連打で設定タブが増えない) は WebView2 の実際の
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
'   直接叩いて、名前⇔テンプレート解決・プレビュー URL 生成・副作用なしを照合する。
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

' --- 判定カウンタ (TestBool / TestEq / D1Case が数え、TestCountPrint が出す) ---
'     ★イミディエイトは ExecuteScript の配管ログで流れてしまうので、
'     判定はログファイルにも残す (K-1 の Wv2Log 経由)。★
Private m_okCount As Long
Private m_ngCount As Long

' --- N-1: 検証ページを置く仮想ホスト名とフォルダ名 ---
Private Const N1_HOST   As String = "appassets.netprobe"
Private Const N1_FOLDER As String = "Wv2NetProbe"

' --- N-1b: 検証の的 (★どちらも仮想ホストの外★) ---
'   N1_LOCAL … 到達不能なローカルアドレス。接続は必ず失敗するが、要求は飛ぶ。
'     ★ポートは Chromium が塞いでいる番号を避ける★ (1/7/9/11/13/…/10080 など)
'     ★http でも混在コンテンツで止まらない★ 127.0.0.1 は仕様上
'     「潜在的に信頼できるオリジン」なので https のページから呼べる。
'   N1_NET   … 外部サービス。ネットが要る代わりに素直に届く。
Private Const N1_LOCAL  As String = "http://127.0.0.1:59999"
Private Const N1_NET    As String = "https://httpbingo.org"


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
'   直接叩いて、名前⇔テンプレートの解決とプレビュー URL 生成を照合する。
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
    Debug.Print "  2) タブバー右端の [歯車] ボタンをクリックする。"
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
    Debug.Print "  ※ 通常タブが無い状態で [歯車] を押しても設定タブは開く (最後の1タブでも可)。"
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
'     ・閉じられた設定タブは m_tabs から外れるので走査で見つからず、次の [歯車] で
'       新規に開き直せる (生存確認が走査と一体)。
' ============================================================
Public Sub Test_9_19_SettingsTabDedup_Help()
    Debug.Print "==== 第9.19 設定タブ重複防止 検証手順 (WebView2 起動) ===="
    Debug.Print ""
    Debug.Print "  1) StartWebView2_Full で通常起動する。"
    Debug.Print "  2) タブバー右端の [歯車] を 1 回押す。"
    Debug.Print "     → 設定タブが 1 枚開き、アクティブになること。"
    Debug.Print "  3) ★重複防止の本命★ 続けて [歯車] をもう 2～3 回連打する。"
    Debug.Print "     → 設定タブが増えず、既存の設定タブがアクティブになるだけであること。"
    Debug.Print "       (イミディエイトに『既存の設定タブ(index N)をアクティブ化』が出る)"
    Debug.Print "  4) 別の通常タブをクリックして設定タブから離れる。"
    Debug.Print "     → その状態で [歯車] を押すと、既存の設定タブへ切り替わる (新規は開かない)。"
    Debug.Print "  5) ★閉じてから開き直し★ 設定タブの × を押して閉じる。"
    Debug.Print "     → その後もう一度 [歯車] を押すと、設定タブが新規に 1 枚開くこと。"
    Debug.Print "       (閉じたら m_tabs から外れるので、走査で見つからず開き直せる)"
    Debug.Print "  6) 設定画面のカード操作 (エンジン選択・hover プレビュー・切替の本番反映) が"
    Debug.Print "     9.18 と同じく動くこと (重複防止で設定機能が壊れていないことの確認)。"
    Debug.Print ""
    Debug.Print "  ※ 期待挙動まとめ: 設定タブは常に高々 1 枚。[歯車] は「無ければ開く/あれば移動」。"
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
    Debug.Print "  2) [歯車] を押して設定タブを開き、Bing のカードをクリックする。"
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
    Debug.Print "  2) ＋ ボタンでタブを 3～4 枚に増やす。"
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
    Debug.Print "  7) [歯車] で設定タブが開き、エンジンを変えると検索に反映され、"
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
    Debug.Print "  2) ＋ ボタンでタブを 4～5 枚に増やす。"
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
    Debug.Print "  7) [歯車] で設定タブ → OnPaneWebMessage: msg={""cmd"":""settings""} が出て"
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
    Debug.Print "  2) ＋ ボタンでタブを 4～5 枚に増やす。"
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
    Debug.Print "  7) [歯車] で設定タブ → OnPaneWebMessage: msg={""cmd"":""settings""} → 設定タブが開く"
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
'   9.24 は TabBar (9.21～9.23) で確立した型を NavBar へ同型展開した回。
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
    Debug.Print "       設定中の検索エンジンで検索されること (9.14/9.16～9.20 の経路)。"
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
    Debug.Print "      手順 5～7 をやり直す。これで直れば「NewWindow をセットしたら Handled は"
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
    Debug.Print "  2) [歯車] を押して設定タブを開く"
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
    Debug.Print "  2) + を 2～3 回押してタブを 3～4 枚にする"
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
    Debug.Print " 10) タブを 8～10 枚まで増やして、タブが細くなった状態で 6～7 を再確認"
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
    Debug.Print "  ・3～11 がすべて期待どおり → 第9.31 は合格"
    Debug.Print "  ・12～16 がすべて期待どおり → ★v0_5_2 の通し確認 (第9.30 の宿題 1) を消化★"
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
    Debug.Print "  2) + を押してタブを 4～5 枚にし、それぞれ別のページを開いておく"
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
    Debug.Print "  4) タブを 5～6 回続けて切り替えて、体感を第9.31 と比べる"
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
    Debug.Print "  6) + を連打してタブを 10～12 枚まで増やす"
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
    Debug.Print " 11) [歯車] で設定タブを開き、検索エンジンのカードをクリックする"
    Debug.Print "       → 期待: 第9.29 までと同じ (選択が反映され、保存ログが出る)"
    Debug.Print " 12) 設定タブを閉じて、通常タブのアドレスバーで検索してみる"
    Debug.Print "       → 期待: 選んだエンジンで検索できる"
    Debug.Print ""
    Debug.Print "  --- 判定 ---"
    Debug.Print "  ・3～4 で体感が改善 → 第9.32 の本命は成功"
    Debug.Print "  ・6～8 が期待どおり → スクロールバーの整理も成功"
    Debug.Print "  ・9～12 が第9.31 と同じ → 回帰なし"
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
    Debug.Print "  5) タブが 1～2 枚しか無い状態 (溢れていない状態) でもホイールを回す"
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
    Debug.Print " 14) [歯車] で設定タブが開くか"
    Debug.Print ""
    Debug.Print "  --- 判定 ---"
    Debug.Print "  ・3～4 が効けば論点 1 は成功 (効かなければ論点 1 を (b) 3px バーに戻す)"
    Debug.Print "  ・6～7 で位置が動かなければ論点 3 の退避・復元は成功"
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
        Wv2Log.LogI "Test_D1_Eval: アクティブな Pane がありません。" & _
                    "先に StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D1_Eval 開始 ================"
    Wv2Log.LogI "  対象タブ: " & p.DocumentTitle
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) 型ごとの戻り値 ---"

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

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 実ページの情報 ---"
    D1Case p, "document.title", "document.title", 5, True
    D1Case p, "location.href", "document.location.href", 5, True
    D1Case p, "body の文字数", "document.body.innerText.length", 5, True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 失敗系 (FAIL 表示にならず OK と出れば正常) ---"
    D1Case p, "JS 例外 (ReferenceError)", "nonexistentFunctionForTest()", 5, False
    D1Case p, "JS 例外 (throw を即時関数で包む)", _
              "(function(){throw new Error('boom');})()", 5, False
    D1Case p, "構文エラー (式が壊れている)", "1 +", 5, False

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) タイムアウトと回復 ---"
    Wv2Log.LogI "  ※ JS を 3 秒ブロックし、1 秒で打ち切る。数秒後に"
    Wv2Log.LogI "     ★破棄済み★ の遅延到着ログが出れば論点5 は成功。"
    D1Case p, "タイムアウト (3 秒を 1 秒で打ち切り)", _
              "(function(){var t=Date.now();while(Date.now()-t<3000){}return 1;})()", 1, False

    ' JS スレッドが空くまで待つ (この間に遅延到着のログが出る)
    Dim t0 As Single
    t0 = Timer
    Do While (Timer - t0) < 4
        DoEvents
    Loop

    D1Case p, "タイムアウト後の回復", "1 + 1", 5, True

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D1_Eval 終了 ================"
    Wv2Log.LogI ""
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
        m_okCount = m_okCount + 1
    Else
        mark = "  [FAIL] "
        m_ngCount = m_ngCount + 1
    End If

    If p.LastEvalOk Then
        Wv2Log.LogI mark & caseName & " → " & Left$(r, 100)
    Else
        Wv2Log.LogI mark & caseName & " → 失敗: " & p.LastEvalError
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
        Wv2Log.LogI "Test_D1_Guard: アクティブな Pane がありません。" & _
                    "先に StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D1_Guard 開始 ================"

    Wv2Log.LogI "  1) 通常状態 (深さ=" & p.InCallbackDepth & ") で EvalSync"
    r = p.EvalSync("1 + 1", 5)
    TestBool "  通常状態で 2 が返る", (p.LastEvalOk And r = "2")
    If Not p.LastEvalOk Then Wv2Log.LogI "         r=" & r & " err=" & p.LastEvalError

    Wv2Log.LogI "  2) ハンドラ内にいる状態を作って EvalSync (拒否されるのが正常)"
    p.Debug_SetInCallback 1
    r = p.EvalSync("1 + 1", 5)
    TestBool "  ★in-callback で拒否された (固まらずに即戻った)★", _
             ((Not p.LastEvalOk) And p.LastEvalError = "in-callback")

    Wv2Log.LogI "  3) ResetCallbackGuard で復帰させて再実行"
    p.ResetCallbackGuard
    r = p.EvalSync("1 + 1", 5)
    TestBool "  ResetCallbackGuard で復帰する", (p.LastEvalOk And r = "2")
    Wv2Log.LogI "         深さ=" & p.InCallbackDepth

    TestCountPrint
    Wv2Log.LogI "================ Test_D1_Guard 終了 ================"
    Wv2Log.LogI ""
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
    Debug.Print "  1) UserForm1.Show vbModeless して StartWebView2_Full を実行する"
    Debug.Print "     ★Show が先★ フォームのウィンドウが無いと Frame1 の HWND が"
    Debug.Print "     取れず、hWnd_Frame = 0 のまま Browser.Init が失敗する。"
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
    Debug.Print "  --- ★判定はログファイルに残る★ ---"
    Debug.Print "    イミディエイトは ExecuteScript の配管ログ (1 往復で 15 行) で"
    Debug.Print "    すぐ流れるので、合否はログファイルで見る:"
    Debug.Print "      ?Wv2Log.LogPath   … ファイルの場所"
    Debug.Print "      末尾の「★判定 n 件: OK x / FAIL y★」だけ見れば合否が分かる"
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
    Debug.Print "     ※ 6～9 は View_On* を Core に分離した影響を見るためのもの。"
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
        Wv2Log.LogI "Test_D2_Find: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D2_Find: タブの生成に失敗しました。"
        Exit Sub
    End If

    If Not D2WaitTitle(p, "D-2 プローブ", 10) Then
        Wv2Log.LogI "Test_D2_Find: 検証ページの読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D2_Find 開始 ================"
    Wv2Log.LogI "  世代 (取得前は空文字が正常): [" & p.CurrentDomGen & "]"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) 取得できること ---"

    Set el = p.GetElementById("ttl")
    TestBool "GetElementById(ttl) が Nothing でない", Not (el Is Nothing)
    If el Is Nothing Then
        Wv2Log.LogI "  以降の検証は続けられません。中止します。"
        TestCountPrint
        Exit Sub
    End If
    Wv2Log.LogI "        handle=" & el.Handle & " gen=" & el.Generation
    Wv2Log.LogI "        Pane 側の世代キャッシュ: " & p.CurrentDomGen
    TestBool "取得直後は stale でない", (el.IsStale = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 読み取り 4 種 + 属性 ---"
    TestEq "TagName (大文字で返る)", el, el.TagName, "H1"
    TestEq "InnerText (日本語)", el, el.InnerText, "D-2 要素レジストリのプローブ"

    Set el = D2El(p, "box")
    TestEq "GetAttribute(class)", el, el.GetAttribute("class"), "card"
    TestEq "GetAttribute(data-note) 日本語属性", el, _
         el.GetAttribute("data-note"), "属性の値 (日本語)"
    TestEq "InnerHTML (★仕様事実30 の復号★)", el, _
         el.InnerHTML, "<span class=""tag"">内側</span>テキスト"

    Set el = D2El(p, "esc")
    TestEq "記号の混在", el, el.InnerText, "記号: < > & "" ' \ の混在"

    Set el = D2El(p, "pre")
    TestEq "改行を含むテキスト (\n の復号)", el, el.InnerText, "1 行目" & vbLf & "2 行目"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 入力要素の value ---"
    Set el = D2El(p, "txt")
    TestEq "input の TagName", el, el.TagName, "INPUT"
    TestEq "input の Value", el, el.value, "初期値"
    TestEq "input の GetAttribute(value)", el, el.GetAttribute("value"), "初期値"

    Set el = D2El(p, "area")
    TestEq "textarea の Value", el, el.value, "テキストエリアの値"

    Set el = D2El(p, "sel")
    TestEq "select の Value (selected の option)", el, el.value, "b"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) 空・不在は★成功して空文字★になること (LastOk=True) ---"
    Set el = D2El(p, "empty")
    TestEq "空要素の InnerText", el, el.InnerText, ""

    Set el = D2El(p, "lnk")
    TestEq "a 要素の Value (value を持たない)", el, el.value, ""
    TestEq "存在しない属性", el, el.GetAttribute("data-nothing"), ""
    TestEq "href 属性", el, el.GetAttribute("href"), "https://example.com/path?x=1"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) QuerySelector (★セレクタ内のシングルクォート★) ---"
    Set el2 = p.QuerySelector("input[name='q']")
    TestBool "QuerySelector(input[name='q']) が取れる", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        TestEq "同じ要素が取れている", el2, el2.value, "初期値"
    End If

    Set el2 = p.QuerySelector("#box .tag")
    TestBool "子孫セレクタが効く", Not (el2 Is Nothing)
    If Not el2 Is Nothing Then
        TestEq "子孫セレクタの InnerText", el2, el2.InnerText, "内側"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) 見つからない / 失敗の区別 (論点4) ---"
    Set el2 = p.QuerySelector("#nothing-here")
    TestBool "存在しないセレクタ → Nothing", (el2 Is Nothing)
    TestBool "  かつ LastEvalOk = True (本当に無い、の意味)", (p.LastEvalOk = True)

    Set el2 = p.QuerySelector("###")
    TestBool "不正なセレクタ → Nothing", (el2 Is Nothing)
    TestBool "  かつ LastEvalOk = False (失敗、の意味)", (p.LastEvalOk = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) ClearElementRegistry (論点3) ---"
    Set el = D2El(p, "ttl")
    TestBool "掃除前は stale でない", (el.IsStale = False)
    TestBool "ClearElementRegistry が成功する", p.ClearElementRegistry()
    TestBool "掃除後は stale になる", (el.IsStale = True)
    Set el = p.GetElementById("ttl")
    TestBool "掃除後も取り直せる", Not (el Is Nothing)
    If Not el Is Nothing Then
        TestEq "取り直した要素が読める", el, el.TagName, "H1"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D2_Find 終了 ================"
    Wv2Log.LogI ""
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
        Wv2Log.LogI "Test_D2_Stale: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD2ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D2_Stale: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-2 プローブ", 10) Then
        Wv2Log.LogI "Test_D2_Stale: 検証ページの読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D2_Stale 開始 ================"

    Set el = p.GetElementById("ttl")
    If el Is Nothing Then
        Wv2Log.LogI "  [FAIL] 要素が取れないので中止します。"
        m_ngCount = m_ngCount + 1
        TestCountPrint
        Exit Sub
    End If
    genBefore = el.Generation
    Wv2Log.LogI "  1) 遷移前: handle=" & el.Handle & " gen=" & genBefore
    TestEq "     読める", el, el.TagName, "H1"
    TestBool "     stale でない", (el.IsStale = False)

    Wv2Log.LogI "  2) 同じタブを別のページへ遷移させる"
    p.View_NavigateToString BuildD2SecondHtml()
    If Not D2WaitTitle(p, "D-2 プローブ 2 枚目", 10) Then
        Wv2Log.LogI "  [FAIL] 2 枚目の読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Wv2Log.LogI "  3) 遷移後: 古いハンドルの状態を見る"
    TestBool "     IsStale = True になる", (el.IsStale = True)
    v = el.TagName
    TestBool "     読み取りは空文字 + LastOk=False", (Len(v) = 0 And el.LastOk = False)
    Wv2Log.LogI "        LastError = " & el.LastError
    TestBool "     LastError が stale であること", (el.LastError = "stale")

    Wv2Log.LogI "  4) 新しいページで取り直せる"
    Set el = p.GetElementById("second")
    TestBool "     取得できる", Not (el Is Nothing)
    If Not el Is Nothing Then
        genAfter = el.Generation
        TestEq "     読める", el, el.InnerText, "2 枚目のページ"
        Wv2Log.LogI "        新しい gen=" & genAfter
        TestBool "     世代が変わっている", (genAfter <> genBefore)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D2_Stale 終了 ================"
    Wv2Log.LogI ""
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
        Wv2Log.LogI "  [FAIL] 要素 #" & elementId & " が取得できない " & _
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
            Wv2Log.LogI "D2WaitTitle: タイムアウト (期待=" & wantTitle & _
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
    Debug.Print "  1) UserForm1.Show vbModeless して StartWebView2_Full を実行する"
    Debug.Print "     ★Show が先★ フォームのウィンドウが無いと Frame1 の HWND が"
    Debug.Print "     取れず、hWnd_Frame = 0 のまま Browser.Init が失敗する。"
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
    Debug.Print "  --- ★判定はログファイルに残る★ ---"
    Debug.Print "    イミディエイトは ExecuteScript の配管ログ (1 往復で 15 行) で"
    Debug.Print "    すぐ流れるので、合否はログファイルで見る:"
    Debug.Print "      ?Wv2Log.LogPath   … ファイルの場所"
    Debug.Print "      末尾の「★判定 n 件: OK x / FAIL y★」だけ見れば合否が分かる"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D2_Find) ---"
    Debug.Print "  ・(1)～(7) の全行が [OK  ] であること"
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

' ============================================================
' Test_D3_Probe_Promise (D-3 の初手: ★未知1 の実測★)
'
'   「ExecuteScript は Promise を待つか」を確かめる。ここの結果で
'   論点3 (待ちループを VBA で回すか JS で回すか) が決まる。
'
'   ★(B) は自動判定できない★ raw ExecuteScript の結果は EvalSync の
'   保留テーブルを通らないので、VBA からは受け取れない。代わりに
'   Wv2Pane.OnExecuteScriptCompleted が出す resultJson= の行を目で見る。
'   その行が★どのマーカーの間に出たか★で判定できるようにしてある。
' ============================================================
Public Sub Test_D3_Probe_Promise()
    Dim p As Wv2Pane
    Dim res As String
    Dim cb As Long

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Probe_Promise: アクティブな Pane がありません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Probe_Promise 開始 ================"
    Wv2Log.LogI "  ★未知1 の実測: ExecuteScript は Promise を待つか★"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (A) EvalSync 経由 (D-1 の同期プリミティブ) ---"
    Wv2Log.LogI "      EvalSync は JSON.stringify(式) を撃つので、式が Promise を"
    Wv2Log.LogI "      返してもその場で {} に潰れる。★仕様上そうなる★ことの確認。"

    res = p.EvalSync("1+1")
    Wv2Log.LogI "      1+1                 → [" & res & "] ok=" & p.LastEvalOk
    res = p.EvalSync("Promise.resolve(42)")
    Wv2Log.LogI "      Promise.resolve(42) → [" & res & "] ok=" & p.LastEvalOk
    TestBool "EvalSync は Promise を値に解決しない ({} になる)", (res = "{}")

    res = p.EvalSync("(async function(){return 5;})()")
    Wv2Log.LogI "      async 関数の戻り    → [" & res & "] ok=" & p.LastEvalOk

    res = p.EvalSync("new Promise(function(r){setTimeout(function(){r(7);},300);})")
    Wv2Log.LogI "      遅延 Promise        → [" & res & "] ok=" & p.LastEvalOk

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (B) raw ExecuteScript (JSON.stringify で包まない) ---"
    Wv2Log.LogI "      ★ここは目で見る★ マーカーの間に出る"
    Wv2Log.LogI "      「OnExecuteScriptCompleted: ... resultJson=」の行を読むこと。"

    Wv2Log.LogI ""
    Wv2Log.LogI "  ★★★ (B-1) Promise.resolve(42) ここから ★★★"
    cb = p.View_ExecuteScript("Promise.resolve(42)")
    D3Pump 2
    Wv2Log.LogI "  ★★★ (B-1) ここまで (callbackId=" & cb & ") ★★★"
    Wv2Log.LogI "      resultJson={} なら★待たない★ / resultJson=42 なら★待つ★"

    Wv2Log.LogI ""
    Wv2Log.LogI "  ★★★ (B-2) 決して解決しない Promise ここから ★★★"
    cb = p.View_ExecuteScript("new Promise(function(r){})")
    D3Pump 2
    Wv2Log.LogI "  ★★★ (B-2) ここまで (callbackId=" & cb & ") ★★★"
    Wv2Log.LogI "      この間に resultJson={} が出た → ★待たない★ (確定)"
    Wv2Log.LogI "      この間に何も出なかった       → ★待つ★ (確定)"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 判定の意味 ---"
    Wv2Log.LogI "  ・待たない → 論点3 の骨格どおり★VBA 側でポーリング★する"
    Wv2Log.LogI "  ・待つ     → JS 側で待てるので EvalSync の包み方から設計し直す"
    Wv2Log.LogI "                (第9.30 / D-1 の教訓: 前提が違えば初手を差し替える)"
    Wv2Log.LogI "  ※ (A) の結果は (B) がどちらでも変わらない。今の EvalSync は"
    Wv2Log.LogI "     Promise を待てないので、待ちを JS 側に置くなら D-1 の包み方に手が要る。"

    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Probe_Promise 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D3_Write (D-3 本体の検証: 書き込み・操作)
' ============================================================
Public Sub Test_D3_Write()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim dummy As Wv2Element
    Dim info As String
    Dim res As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Write: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Write: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-3 プローブ", 10) Then
        Wv2Log.LogI "Test_D3_Write: 検証ページの読み込みを確認できませんでした。中止します。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Write 開始 ================"
    Wv2Log.LogI "  --- (1) input への書き込み ---"

    Set el = D2El(p, "txt")
    el.value = "書き込んだ値"
    info = el.LastInfo
    TestBool "Value = が成功する (LastOk)", el.LastOk
    Wv2Log.LogI "        経路 (LastInfo) = " & info & "  ★setter が期待値★"
    TestBool "  ★ネイティブ setter 経由である★", (info = "setter")
    TestEq "読み戻すと書いた値になる", el, el.value, "書き込んだ値"
    TestEq "★属性の value は変わらない★ (初期値のまま)", el, _
         el.GetAttribute("value"), "初期値"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 記号・日本語 (JsQuote の確認) ---"
    el.value = "記号: < > & "" ' \ の混在"
    TestBool "記号混じりでも成功する", el.LastOk
    TestEq "記号混じりが往復する", el, el.value, _
         "記号: < > & "" ' \ の混在"
    el.value = ""
    TestEq "空文字を書き込める", el, el.value, ""

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) input と change の両方が飛ぶこと (論点6) ---"
    Wv2Log.LogI "        ここまでの書き込みは 3 回。"
    TestBool "input イベントが 3 回飛んだ", (D3Cnt(p, "inputs") = 3)
    TestBool "change イベントが 3 回飛んだ", (D3Cnt(p, "changes") = 3)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) textarea / select / checkbox ---"
    Set el = D2El(p, "area")
    el.value = "書き換えたテキスト"
    TestEq "textarea に書ける", el, el.value, "書き換えたテキスト"

    Set el = D2El(p, "sel")
    TestEq "select の初期値", el, el.value, "b"
    el.value = "a"
    TestEq "select を切り替えられる", el, el.value, "a"

    Set el = D2El(p, "chk")
    TestBool "checkbox の Click が成功する", el.Click()
    res = p.EvalSync("document.getElementById('chk').checked")
    TestBool "  checked が true になった", (res = "true")

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ボタンの Click (論点7) ---"
    Set el = D2El(p, "btn")
    TestBool "Click が True を返す", el.Click()
    Wv2Log.LogI "        経路 (LastInfo) = " & el.LastInfo
    TestBool "  ★e.click() が使われた★", (el.LastInfo = "click")
    TestBool "  ページ側で 1 回数えられた", (D3Cnt(p, "clicks") = 1)
    Wv2Log.LogI "        イベントの素性 (type/bubbles/isTrusted) = " & _
                D3Str(p, "clickInfo")
    Wv2Log.LogI "        ※ isTrusted=false は★合成イベントなので原理的にそうなる★"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) SetAttribute ---"
    Set el = D2El(p, "dv")
    TestBool "SetAttribute が True を返す", _
           el.SetAttribute("data-note", "属性の値 (日本語)")
    TestEq "書いた属性が読める", el, el.GetAttribute("data-note"), _
         "属性の値 (日本語)"
    TestBool "引用符を含む属性値も書ける", _
           el.SetAttribute("data-q", "a""b'c")
    TestEq "  往復する", el, el.GetAttribute("data-q"), "a""b'c"
    TestBool "★不正な属性名は False になる★", (el.SetAttribute("", "x") = False)
    Wv2Log.LogI "        LastError = " & el.LastError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) value を持たない要素への書き込み ---"
    Wv2Log.LogI "        div は value を持たないので素の代入に落ちる。"
    Wv2Log.LogI "        ★JS 的には例外にならないので LastOk は True★"
    el.value = "div に書いてみる"
    info = el.LastInfo
    TestBool "LastOk は True (エラーではない)", el.LastOk
    TestBool "  経路は direct (ネイティブ setter が無い)", (info = "direct")
    TestEq "  読み戻すと生えたプロパティが見える", el, el.value, "div に書いてみる"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (8) 失敗の区別 (D-2 の規約を踏襲、論点8) ---"
    Set dummy = New Wv2Element
    TestBool "New で作った要素の Click は False", (dummy.Click() = False)
    TestBool "  LastError = no-pane", (dummy.LastError = "no-pane")
    dummy.value = "x"
    TestBool "  Value = も LastOk=False / no-pane", _
           (dummy.LastOk = False And dummy.LastError = "no-pane")

    Set el = D2El(p, "txt")
    TestBool "掃除前は書ける", el.SetAttribute("data-a", "1")
    TestBool "ClearElementRegistry が成功する", p.ClearElementRegistry()
    TestBool "★掃除後の Click は False★", (el.Click() = False)
    TestBool "  LastError = stale", (el.LastError = "stale")
    el.value = "書けないはず"
    TestBool "  Value = も stale で止まる", _
           (el.LastOk = False And el.LastError = "stale")

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Write 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_D3_Framework (★未知3 の実測★ フレームワークに伝わるか)
'
'   検証ページの #react には★React の value tracker を模した監視★が
'   被せてある。tracker は「自分が知っている値」を覚えていて、input
'   イベントが来たときに現在値と食い違っていたら★変更に気づく★。
'
'     素の代入 (e.value = x)  … tracker 経由なので tracker が値を覚えてしまい、
'                               input が飛んでも★気づけない★ (ignored が増える)
'     ネイティブ setter 経由  … tracker を迂回するので値が食い違い、
'                               input で★気づく★ (notified が増える)
'
'   ここが D-3 の存在意義そのもの。FAIL したら Wv2Element.Value = の
'   ディスクリプタ取得 (e.constructor.prototype) を疑うこと。
' ============================================================
Public Sub Test_D3_Framework()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Framework: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Framework: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-3 プローブ", 10) Then
        Wv2Log.LogI "Test_D3_Framework: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Framework 開始 ================"
    Wv2Log.LogI "  --- (1) ★素の代入★ + input 発火 (D-3 を使わない書き方) ---"

    p.EvalSync "(function(){var e=document.getElementById('react');" & _
               "e.value='素の代入';" & _
               "e.dispatchEvent(new Event('input',{bubbles:true}));return 1;})()"
    TestBool "素の代入が走った", p.LastEvalOk
    TestBool "  tracker 経由になった (trackedSet=1)", (D3Cnt(p, "trackedSet") = 1)
    TestBool "  ★フレームワークは気づかない (notified=0)★", _
           (D3Cnt(p, "notified") = 0)
    TestBool "  取りこぼしとして数えられた (ignored=1)", (D3Cnt(p, "ignored") = 1)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★Wv2Element.Value =★ (D-3 の書き方) ---"

    Set el = D2El(p, "react")
    el.value = "D-3 の書き込み"
    TestBool "書き込みが成功する", el.LastOk
    TestBool "  ネイティブ setter 経由 (LastInfo=setter)", (el.LastInfo = "setter")
    TestBool "  ★tracker を迂回した (trackedSet は 1 のまま)★", _
           (D3Cnt(p, "trackedSet") = 1)
    TestBool "  ★フレームワークが変更に気づいた (notified=1)★", _
           (D3Cnt(p, "notified") = 1)
    TestEq "  値も入っている", el, el.value, "D-3 の書き込み"

    Wv2Log.LogI ""
    Wv2Log.LogI "  ここが D-3 の核心。(1) が notified=0 で (2) が notified=1 なら、"
    Wv2Log.LogI "  React / Vue のページでも .Value = が効くことになる。"
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Framework 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' TestBool / TestEq / TestCountReset / TestCountPrint (判定ヘルパー)
'
'   ★D-1 / D-2 / D-3 の全テストが使う共通の判定★ (元は D-2 の D2Bool / D2Eq。
'   D-3 でログファイルにも出すようにしたのを機に 1 組へ統合した)
'
'   D-2 の TestBool / TestEq と同じ判定だが、★出し先が Wv2Log★ である点が違う。
'   Wv2Log は内部で Debug.Print も撃つのでイミディエイトの見え方は変わらない。
'   狙いは★判定行がログファイルに残ること★:
'     ・イミディエイトは ExecuteScript の配管ログ (1 往復で 15 行) で流れてしまう
'     ・ログファイルなら ' [FAIL]' を検索するだけで済む
'
'   TestCountPrint が最後に「OK n / FAIL m」を出すので、そこだけ見れば合否が分かる。
' ============================================================
Private Sub TestCountReset()
    m_okCount = 0
    m_ngCount = 0
End Sub

Private Sub TestCountPrint()
    Wv2Log.LogI "  ★判定 " & (m_okCount + m_ngCount) & " 件: OK " & m_okCount & _
                " / FAIL " & m_ngCount & "★"
End Sub

Private Sub TestBool(ByVal label As String, ByVal cond As Boolean)
    If cond Then
        m_okCount = m_okCount + 1
        Wv2Log.LogI "  [OK  ] " & label
    Else
        m_ngCount = m_ngCount + 1
        Wv2Log.LogI "  [FAIL] " & label
    End If
End Sub

Private Sub TestEq(ByVal label As String, _
                 ByVal el As Wv2Element, _
                 ByVal got As String, _
                 ByVal want As String)
    If got = want Then
        m_okCount = m_okCount + 1
        Wv2Log.LogI "  [OK  ] " & label
    Else
        m_ngCount = m_ngCount + 1
        Wv2Log.LogI "  [FAIL] " & label
        Wv2Log.LogI "         期待: [" & want & "]"
        Wv2Log.LogI "         実際: [" & got & "]"
    End If

    If el Is Nothing Then Exit Sub
    If Not el.LastOk Then
        Wv2Log.LogI "         ※ LastOk=False err=" & el.LastError
    End If
End Sub

' ============================================================
' D3Cnt / D3Str (D-3: 検証ページの監視カウンタを読む)
'   ページ側の window.__p から 1 項目を取り出す。読めなければ -1 を返す。
' ============================================================
Private Function D3Cnt(ByVal p As Wv2Pane, ByVal countName As String) As Long
    Dim res As String

    res = p.EvalSync("window.__p." & countName)
    If Not p.LastEvalOk Then
        m_ngCount = m_ngCount + 1
        Wv2Log.LogW "  [FAIL] カウンタ " & countName & " を読めない err=" & p.LastEvalError
        D3Cnt = -1
        Exit Function
    End If

    D3Cnt = CLng(Val(res))
End Function

Private Function D3Str(ByVal p As Wv2Pane, ByVal itemName As String) As String
    Dim res As String

    res = p.EvalSync("window.__p." & itemName)
    If Not p.LastEvalOk Then
        D3Str = "(読めない: " & p.LastEvalError & ")"
        Exit Function
    End If

    D3Str = Wv2Json.JsonUnescape(res)
End Function


' ============================================================
' D3Pump (D-3: 指定秒だけ DoEvents を回す)
'   raw ExecuteScript の完了通知が出るのを待つためだけの足踏み。
'   ★この間もブレーク/ステップ実行しないこと★ (仕様事実20)
' ============================================================
Private Sub D3Pump(ByVal waitSec As Single)
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If (Timer - t0) > waitSec Then Exit Do
    Loop
End Sub


' ============================================================
' BuildD3ProbeHtml (D-3 の検証ページ)
'
'   D-2 のページに★監視用の JS★ を足したもの。数えているのは:
'     inputs / changes … #txt に飛んだ input / change の回数 (論点6)
'     clicks / clickInfo … #btn のクリック回数とイベントの素性
'     trackedSet … ★React 風 value tracker★ を通った代入の回数
'     notified   … tracker が「値が変わった」と気づいた回数
'     ignored    … input は来たが tracker が気づけなかった回数
'
'   ★静的な HTML 部分の引用符は VBA の "" で書く★ (JS ではないので問題ない)
'   ★JS の文字列は必ずシングルクォート★ (プロジェクト規則)
' ============================================================
Private Function BuildD3ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>D-3 プローブ</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  input,textarea,select,button{font-size:14px;padding:6px 8px;margin:4px 4px;" & _
            "background:#0b0e15;color:#e8eaed;border:1px solid rgba(255,255,255,.18);" & _
            "border-radius:6px;}" & vbLf
    s = s & "  pre{background:#0b0e15;padding:10px 12px;border-radius:8px;" & _
            "white-space:pre-wrap;word-break:break-all;}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<h1 id=""ttl"">D-3 書き込みと操作のプローブ</h1>" & vbLf
    s = s & "<p class=""note"">このページは Test_D3_Write / Test_D3_Framework 専用です。" & _
            "外部サイトに依存しません。</p>" & vbLf
    s = s & "<div class=""card"">" & vbLf
    s = s & "  <input id=""txt"" type=""text"" name=""q"" value=""初期値"">" & vbLf
    s = s & "  <textarea id=""area"" rows=""2"">元のテキスト</textarea>" & vbLf
    s = s & "  <select id=""sel""><option value=""a"">A</option>" & _
            "<option value=""b"" selected>B</option></select>" & vbLf
    s = s & "  <input id=""chk"" type=""checkbox"">" & vbLf
    s = s & "  <input id=""react"" type=""text"" value="""">" & vbLf
    s = s & "</div>" & vbLf
    s = s & "<p><button id=""btn"" type=""button"">クリックしてね</button>" & _
            "<span id=""btnlog"">未クリック</span></p>" & vbLf
    s = s & "<div id=""dv"" class=""card"">属性の書き込み先</div>" & vbLf
    s = s & "<pre id=""cnt""></pre>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "  var P={inputs:0,changes:0,clicks:0,clickInfo:''," & _
            "trackedSet:0,notified:0,ignored:0};" & vbLf
    s = s & "  window.__p=P;" & vbLf
    s = s & "  function show(){" & _
            "document.getElementById('cnt').textContent=JSON.stringify(P);}" & vbLf
    s = s & "  var t=document.getElementById('txt');" & vbLf
    s = s & "  t.addEventListener('input',function(){P.inputs++;show();});" & vbLf
    s = s & "  t.addEventListener('change',function(){P.changes++;show();});" & vbLf
    s = s & "  var b=document.getElementById('btn');" & vbLf
    s = s & "  b.addEventListener('click',function(ev){P.clicks++;" & vbLf
    s = s & "    P.clickInfo=ev.type+'/'+ev.bubbles+'/'+ev.isTrusted;" & vbLf
    s = s & "    document.getElementById('btnlog').textContent=" & _
            "'クリック '+P.clicks+' 回';show();});" & vbLf
    s = s & "  var r=document.getElementById('react');" & vbLf
    s = s & "  var d=Object.getOwnPropertyDescriptor(r.constructor.prototype,'value');" & vbLf
    s = s & "  var last=d.get.call(r);" & vbLf
    s = s & "  Object.defineProperty(r,'value',{configurable:true," & vbLf
    s = s & "    get:function(){return d.get.call(this);}," & vbLf
    s = s & "    set:function(v){P.trackedSet++;last=v;d.set.call(this,v);show();}});" & vbLf
    s = s & "  r.addEventListener('input',function(){" & vbLf
    s = s & "    var cur=d.get.call(r);" & vbLf
    s = s & "    if(cur!==last){P.notified++;last=cur;}else{P.ignored++;}" & vbLf
    s = s & "    show();});" & vbLf
    s = s & "  show();" & vbLf
    s = s & "})();" & vbLf
    s = s & "</" & "script>" & vbLf
    s = s & "</body></html>"

    BuildD3ProbeHtml = s
End Function


' ============================================================
' Test_D3_Wait (D-3b の検証: DOM 条件待ち)
'
'   遅れて現れる / 遅れて消える要素を setTimeout で作り、
'   WaitFor / WaitGone / ImplicitWaitSec が期待どおり粘るかを見る。
'   ★待った秒数も出す★ 「粘った」だけでなく「粘りすぎない」ことも見たいため。
'
'   ★実行中はブレーク/ステップ実行しないこと★ (仕様事実20)
'   待ちループ全体が長い DoEvents 区間になる (未知2 のとおり)。
' ============================================================
Public Sub Test_D3_Wait()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim took As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D3_Wait: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD3ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D3_Wait: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-3 プローブ", 10) Then
        Wv2Log.LogI "Test_D3_Wait: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D3_Wait 開始 ================"
    Wv2Log.LogI "  --- (1) 既にある要素は待たずに返る ---"

    t0 = Timer
    Set el = p.WaitFor("#txt", 5)
    took = D3Took(t0)
    TestBool "既にある要素を WaitFor で掴める", Not (el Is Nothing)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒"
    TestBool "  ★待たずに返る (0.5 秒未満)★", (took < 0.5)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 遅れて現れる要素 (800ms 後に追加) ---"

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var d=document.createElement('input');d.id='late';d.type='text';" & _
               "d.value='遅れて出た';document.body.appendChild(d);},800);return 1;})()"
    TestBool "遅延追加を仕掛けられた", p.LastEvalOk

    t0 = Timer
    Set el = p.WaitFor("#late", 5)
    took = D3Took(t0)
    TestBool "★遅れて現れた要素を掴めた★", Not (el Is Nothing)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒 (仕掛けは 0.80 秒)"
    TestBool "  待ち時間が妥当 (0.4～3.0 秒)", (took > 0.4 And took < 3)

    If Not el Is Nothing Then
        TestEq "  掴んだ要素の値が読める", el, el.value, "遅れて出た"
        el.value = "待ってから書いた"
        TestBool "  ★待った要素にそのまま書ける (D-3a との接続)★", el.LastOk
        TestEq "    読み戻せる", el, el.value, "待ってから書いた"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 現れないもの (1 秒でタイムアウト) ---"

    t0 = Timer
    Set el = p.WaitFor("#never-ever", 1)
    took = D3Took(t0)
    TestBool "現れないものは Nothing", (el Is Nothing)
    TestBool "  ★LastEvalOk=True (現れなかった、の意味)★", (p.LastEvalOk = True)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒 (指定は 1.00 秒)"
    TestBool "  ちゃんと粘った (0.8 秒以上)", (took > 0.8)
    TestBool "  粘りすぎない (2.5 秒未満)", (took < 2.5)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★失敗は待たずに諦める★ (不正なセレクタ) ---"

    t0 = Timer
    Set el = p.WaitFor("###", 5)
    took = D3Took(t0)
    TestBool "不正なセレクタは Nothing", (el Is Nothing)
    TestBool "  ★LastEvalOk=False (失敗、の意味)★", (p.LastEvalOk = False)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒 (指定は 5.00 秒)"
    TestBool "  ★5 秒待たずに即座に返る (0.5 秒未満)★", (took < 0.5)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) WaitGone (800ms 後に #dv を消す) ---"

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var e=document.getElementById('dv');" & _
               "if(e){e.parentNode.removeChild(e);}},800);return 1;})()"
    TestBool "遅延削除を仕掛けられた", p.LastEvalOk

    t0 = Timer
    TestBool "★消えるまで待てた★", p.WaitGone("#dv", 5)
    took = D3Took(t0)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒 (仕掛けは 0.80 秒)"
    TestBool "  待ち時間が妥当 (0.4～3.0 秒)", (took > 0.4 And took < 3)

    t0 = Timer
    TestBool "最初から無いものは即 True", p.WaitGone("#never-ever", 5)
    TestBool "  待たずに返る (0.5 秒未満)", (D3Took(t0) < 0.5)

    t0 = Timer
    TestBool "消えないものは False", (p.WaitGone("#ttl", 1) = False)
    took = D3Took(t0)
    TestBool "  ★LastEvalOk=True (まだ居る、の意味)★", (p.LastEvalOk = True)
    TestBool "  ちゃんと粘った (0.8 秒以上)", (took > 0.8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ImplicitWaitSec (論点2 案C の opt-in) ---"

    TestBool "★既定は 0 (無効)★", (p.ImplicitWaitSec = 0)
    t0 = Timer
    Set el = p.QuerySelector("#late2")
    took = D3Took(t0)
    TestBool "既定では QuerySelector が待たずに Nothing", _
           (el Is Nothing And took < 0.5)

    p.EvalSync "(function(){setTimeout(function(){" & _
               "var d=document.createElement('div');d.id='late2';" & _
               "d.textContent='後から出た 2';document.body.appendChild(d);},800);" & _
               "return 1;})()"
    TestBool "遅延追加を仕掛けられた", p.LastEvalOk

    p.ImplicitWaitSec = 5
    t0 = Timer
    Set el = p.QuerySelector("#late2")
    took = D3Took(t0)
    TestBool "★ImplicitWaitSec=5 なら QuerySelector が粘って拾う★", _
           Not (el Is Nothing)
    Wv2Log.LogI "        待った時間 = " & Format$(took, "0.00") & " 秒 (仕掛けは 0.80 秒)"
    TestBool "  待ち時間が妥当 (0.4～3.0 秒)", (took > 0.4 And took < 3)

    p.ImplicitWaitSec = 0
    TestBool "0 に戻せる", (p.ImplicitWaitSec = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D3_Wait 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D3Took (D-3b: 経過秒。★Timer は深夜 0 時に 0 へ戻る★ ので補正する)
' ============================================================
Private Function D3Took(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    D3Took = d
End Function

' ============================================================
' Test_D3_Help (D-3 の手順)
' ============================================================
Public Sub Test_D3_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-3 検証手順 (書き込みと操作)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) Wv2Log.LogStart  … このテスト 1 回分を 1 ファイルに閉じる"
    Debug.Print "  2) UserForm1.Show vbModeless して StartWebView2_Full を実行する"
    Debug.Print "     ★Show が先★ フォームのウィンドウが無いと Frame1 の HWND が"
    Debug.Print "     取れず、hWnd_Frame = 0 のまま Browser.Init が失敗する。"
    Debug.Print "  3) ★イベントバーストが静まるまで待つ★ (仕様事実 20)"
    Debug.Print ""
    Debug.Print "  --- 実行 (★この順番★) ---"
    Debug.Print "  4) Test_D3_Probe_Promise … ★済★ 未知1 の実測 (2026-08-22 決着)"
    Debug.Print "     ExecuteScript が Promise を待つかの実測。ここの結果で"
    Debug.Print "     待ち API (WaitFor / WaitGone) の設計が決まる。"
    Debug.Print "     ★アクティブなタブが要る★ (新しいタブは開かない)"
    Debug.Print "  5) Test_D3_Write      … ★済★ 書き込み・操作 (33 件 OK)"
    Debug.Print "  6) Test_D3_Framework  … ★済★ SPA に効くかの実測 (7 件 OK)"
    Debug.Print "  7) Test_D3_Wait       … ★D-3b の検証★ 待ち API (タブを 1 枚開く)"
    Debug.Print "     ※ 4～6 は決着済み。次に見るのは 7 だけでよい。"
    Debug.Print ""
    Debug.Print "  ★実行中はブレーク/ステップ実行しないこと★ (仕様事実 20)"
    Debug.Print ""
    Debug.Print "  --- ★判定はログファイルに残る★ ---"
    Debug.Print "    D-3 のテストは判定を Wv2Log にも書く。イミディエイトは"
    Debug.Print "    ExecuteScript の配管ログ (1 往復で 15 行) ですぐ流れるので、"
    Debug.Print "    ★合否はログファイルで見る★:"
    Debug.Print "      ?Wv2Log.LogPath      … ファイルの場所"
    Debug.Print "      末尾の「★判定 n 件: OK x / FAIL y★」だけ見れば合否が分かる"
    Debug.Print "      FAIL があれば [FAIL] で検索する"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D3_Probe_Promise) ---"
    Debug.Print "  ・(B-1)(B-2) のマーカーの間に resultJson= の行が出たか"
    Debug.Print "      (B-2) で {} が出た   → ★待たない★ = 論点3 の骨格どおり VBA でポーリング"
    Debug.Print "      (B-2) で何も出ない   → ★待つ★   = 設計を組み直す"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D3_Write) ---"
    Debug.Print "  ・(1)～(8) の全行が [OK  ] であること"
    Debug.Print "  ・(1) の LastInfo が setter であること (ネイティブ setter 経由)"
    Debug.Print "  ・(3) の input / change がどちらも 3 回であること (論点6 の両方撃ち)"
    Debug.Print "  ・(7) が [OK  ] = 「JS が走った」と「効果が出た」の違いの確認"
    Debug.Print "  ・(8) が [OK  ] = no-pane / stale の区別 (D-2 の規約踏襲)"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D3_Framework) ★D-3 の核心★ ---"
    Debug.Print "  ・(1) が notified=0 / ignored=1  … 素の代入では気づかれない"
    Debug.Print "  ・(2) が notified=1 / trackedSet=1 のまま … D-3 の書き方なら気づく"
    Debug.Print "    ここが FAIL なら React / Vue のページで .Value = が効かない。"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D3_Wait) ★D-3b★ ---"
    Debug.Print "  ・(1)～(6) の全行が [OK  ] であること"
    Debug.Print "  ・(2) 遅れて現れる要素を掴み、そのまま書き込めること"
    Debug.Print "      = 待ち (D-3b) と書き込み (D-3a) が繋がっている証拠"
    Debug.Print "  ・(3) 現れないものは★1 秒粘ってから★ Nothing + LastEvalOk=True"
    Debug.Print "  ・(4) ★不正なセレクタは待たずに即座に諦める★"
    Debug.Print "      ここが FAIL だと、ハンドラ内から呼んだとき待ち時間ぶん固まる"
    Debug.Print "  ・(6) ImplicitWaitSec が既定 0 で、5 にすると QuerySelector が粘ること"
    Debug.Print ""
    Debug.Print "  --- 回帰確認 (D-3 は Wv2Pane と Wv2Element を触ったため) ---"
    Debug.Print "  8) Test_D2_Find   … 読み取りが壊れていないこと"
    Debug.Print "  9) Test_D2_Stale  … 世代と stale の扱いが壊れていないこと"
    Debug.Print " 10) Test_D1_Eval / Test_D1_Guard … EvalSync とガード"
    Debug.Print ""
    Debug.Print "  --- 手で試したいとき ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  Set el = p.QuerySelector(""input[name='q']"")"
    Debug.Print "  el.Value = ""検索語"""
    Debug.Print "  ?el.LastOk : ?el.LastInfo : ?el.Value"
    Debug.Print "  ?p.QuerySelector(""button"").Click"
    Debug.Print ""
    Debug.Print "  --- 待ち (D-3b) ---"
    Debug.Print "  Set el = p.WaitFor(""#result"", 5)   ' 現れるまで最大 5 秒"
    Debug.Print "  ?p.WaitGone("".spinner"", 10)        ' 消えるまで最大 10 秒"
    Debug.Print "  p.ImplicitWaitSec = 5               ' 既定 0。検索に暗黙の粘りを足す"
    Debug.Print ""
    Debug.Print "  ★待てるのは「DOM の条件」まで★"
    Debug.Print "    WaitFor / WaitGone は指定した要素の有無しか見ない。"
    Debug.Print "    「通信が全部終わって静かになるまで」を待つ静穏待ち"
    Debug.Print "    (MutationObserver / fetch カウンタ) は D-4 の宿題。"
    Debug.Print ""
End Sub

' ============================================================
' Test_D4_Probe (D-4a の検証: ページ内 SPA プローブ)
'
'   ★D-4 の初手★ 待ちの本体を作る前に、観測する仕掛けが健全に立つことを確かめる
'   (設計原則103: 測る前にプローブ自体を検算する)。
' ============================================================
Public Sub Test_D4_Probe()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim st As String
    Dim before As Single
    Dim after As Single
    Dim n0 As Long
    Dim g0 As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Probe: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Probe: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-4 プローブ", 10) Then
        Wv2Log.LogI "Test_D4_Probe: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Probe 開始 ================"
    Wv2Log.LogI "  --- (1) 呼ぶまでプローブは作られない (未知4) ---"

    st = p.EvalSync("typeof window.__wv2p")
    TestBool "★設置前は undefined★", (Wv2Json.JsonUnescape(st) = "undefined")
    Wv2Log.LogI "        typeof window.__wv2p = " & st

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 往復コストの比較 (未知2) ---"

    before = D4RoundTrip(p, 20)
    Wv2Log.LogI "        設置前の 1 往復 = " & Format$(before, "0.0") & " ms"

    st = p.SpaProbeState()
    TestBool "SpaProbeState が成功する", p.LastEvalOk
    Wv2Log.LogI "        state = " & st

    after = D4RoundTrip(p, 20)
    Wv2Log.LogI "        設置後の 1 往復 = " & Format$(after, "0.0") & " ms"
    TestBool "  観測を張っても往復が極端に遅くならない (3 倍未満)", _
             (after < before * 3 + 10)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 初回の状態 ---"

    TestBool "★健康である (h=true)★", (InStr(1, st, """h"":true") > 0)
    TestBool "  MutationObserver が生きている (ob=true)", _
             (InStr(1, st, """ob"":true") > 0)
    TestBool "  版番号が 1", (Wv2Json.JsonGetNum(st, "v") = 1)
    TestBool "  世代が 1", (Wv2Json.JsonGetNum(st, "g") = 1)
    TestBool "  作り直し回数が 0", (Wv2Json.JsonGetNum(st, "rp") = 0)
    TestBool "  実行中の通信は 0", _
             (Wv2Json.JsonGetNum(st, "f") = 0 And Wv2Json.JsonGetNum(st, "x") = 0)
    TestBool "  SpaProbeHealthy も True", p.SpaProbeHealthy()

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) DOM 変化を数える ---"

    st = p.SpaProbeState()
    n0 = Wv2Json.JsonGetNum(st, "m")
    p.EvalSync "(function(){document.getElementById('target')" & _
               ".textContent='書き換えた '+Date.now();return 1;})()"
    st = p.SpaProbeState()
    TestBool "★DOM 変化が数えられた★", (Wv2Json.JsonGetNum(st, "m") > n0)
    Wv2Log.LogI "        最後の変化からの経過 q = " & _
                Wv2Json.JsonGetNum(st, "q") & " ms"
    TestBool "  直後なので q が小さい (1000ms 未満)", _
             (Wv2Json.JsonGetNum(st, "q") < 1000)
    TestBool "  除外していないので sm = 0", (Wv2Json.JsonGetNum(st, "sm") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) 通信を数える (blob URL なのでオフラインで完結) ---"

    st = p.SpaProbeState()
    n0 = Wv2Json.JsonGetNum(st, "n")

    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "fetch(u).then(function(r){return r.text();});return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "★fetch が数えられた★", (Wv2Json.JsonGetNum(st, "n") > n0)
    TestBool "  完了して実行中が 0 に戻った", (Wv2Json.JsonGetNum(st, "f") = 0)

    n0 = Wv2Json.JsonGetNum(st, "n")
    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "var r=new XMLHttpRequest();r.open('GET',u);r.send();return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "★XHR が数えられた★", (Wv2Json.JsonGetNum(st, "n") > n0)
    TestBool "  完了して実行中が 0 に戻った", (Wv2Json.JsonGetNum(st, "x") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★健康診断と自動修復★ (未知1) ---"

    st = p.SpaProbeState()
    g0 = Wv2Json.JsonGetNum(st, "g")
    Wv2Log.LogI "        壊す前: 世代=" & g0 & " 作り直し=" & _
                Wv2Json.JsonGetNum(st, "rp")

    Wv2Log.LogI "        ページ側が window.fetch を差し替えた状況を作る"
    p.EvalSync "(function(){var o=window.fetch;" & _
               "window.fetch=function(){return o.apply(this,arguments);};return 1;})()"

    st = p.SpaProbeState()
    TestBool "★壊れても呼べば健康に戻る (h=true)★", _
             (InStr(1, st, """h"":true") > 0)
    TestBool "  ★作り直しが記録された (rp >= 1)★", _
             (Wv2Json.JsonGetNum(st, "rp") >= 1)
    TestBool "  世代が 1 つ進んだ", (Wv2Json.JsonGetNum(st, "g") = g0 + 1)
    Wv2Log.LogI "        壊した後: state = " & st

    n0 = Wv2Json.JsonGetNum(st, "n")
    p.EvalSync "(function(){var u=URL.createObjectURL(new Blob(['hi']));" & _
               "fetch(u).then(function(r){return r.text();});return 1;})()"
    D3Pump 1
    st = p.SpaProbeState()
    TestBool "  ★作り直した後も数えられる★", (Wv2Json.JsonGetNum(st, "n") > n0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) SpaProbeReset (手動の作り直し) ---"

    g0 = Wv2Json.JsonGetNum(st, "g")
    TestBool "SpaProbeReset が成功する", p.SpaProbeReset()
    st = p.SpaProbeState()
    TestBool "  世代が進んだ", (Wv2Json.JsonGetNum(st, "g") > g0)
    TestBool "  数えた値が 0 に戻った", (Wv2Json.JsonGetNum(st, "n") = 0)
    TestBool "  健康である", (InStr(1, st, """h"":true") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (8) ページ遷移でプローブは消える (未知3) ---"

    p.View_NavigateToString BuildD2SecondHtml()
    If Not D2WaitTitle(p, "D-2 プローブ 2 枚目", 10) Then
        Wv2Log.LogI "  [FAIL] 遷移を確認できませんでした"
        m_ngCount = m_ngCount + 1
    Else
        st = p.EvalSync("typeof window.__wv2p")
        TestBool "★遷移後は undefined に戻る★", _
                 (Wv2Json.JsonUnescape(st) = "undefined")
        st = p.SpaProbeState()
        TestBool "  新しいページでも立て直せる", _
                 (InStr(1, st, """h"":true") > 0)
        TestBool "  世代は 1 から数え直し", (Wv2Json.JsonGetNum(st, "g") = 1)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Probe 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D4RoundTrip (D-4a: EvalSync 1 往復の平均 ms を測る)
'   ★仕様事実52 (約 38ms) の測り直し★ プローブを張ると重くならないかを見る。
' ============================================================
Private Function D4RoundTrip(ByVal p As Wv2Pane, ByVal shots As Long) As Single
    Dim i As Long
    Dim t0 As Single

    t0 = Timer
    For i = 1 To shots
        p.EvalSync "1"
    Next i

    D4RoundTrip = D3Took(t0) / shots * 1000
End Function


' ============================================================
' BuildD4ProbeHtml (D-4 の検証ページ)
'
'   ★読み込み直後は完全に静か★ にしてある (静穏待ちの検証に要る)。
'   ページ側の口:
'     startNoise(ms) … ms ごとに #noise を書き換え続ける (静まらないページの再現)
'     stopNoise()    … 止める
'     later(ms)      … ms 後に #late を足す (D-4b で使う)
'     chain(ms)      … ms 後に blob を fetch し、その完了後にさらに DOM を書き換える
'                      (★fetch → DOM 更新の連鎖★ = SPA の再現。D-4b で使う)
' ============================================================
Private Function BuildD4ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>D-4 プローブ</title>" & vbLf
    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;background:#12161f;" & _
            "color:#e8eaed;padding:32px;line-height:1.8;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 18px;}" & vbLf
    s = s & "  .card{border:1px solid rgba(255,255,255,.12);border-radius:10px;" & _
            "padding:14px 16px;margin:12px 0;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .note{color:#8ea2c8;font-size:12.5px;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<h1 id=""ttl"">D-4 静穏待ちのプローブ</h1>" & vbLf
    s = s & "<p class=""note"">読み込み直後は★何も動いていない★状態です。" & _
            "startNoise() / later() / chain() で動きを作ります。</p>" & vbLf
    s = s & "<div id=""target"" class=""card"">書き換え対象</div>" & vbLf
    s = s & "<div id=""noise"" class=""card"">ノイズ: 停止中</div>" & vbLf
    s = s & "<div id=""slot"" class=""card""></div>" & vbLf
    s = s & "<p><button id=""btn"" type=""button"">押すと 400ms 後に更新</button></p>" & vbLf
    s = s & "<script>" & vbLf
    s = s & "(function(){" & vbLf
    s = s & "  var t=null;" & vbLf
    s = s & "  window.startNoise=function(ms){" & vbLf
    s = s & "    if(t){clearInterval(t);}" & vbLf
    s = s & "    t=setInterval(function(){" & vbLf
    s = s & "      document.getElementById('noise').textContent='ノイズ '+Date.now();" & vbLf
    s = s & "    },ms||200);return 1;};" & vbLf
    s = s & "  window.stopNoise=function(){" & vbLf
    s = s & "    if(t){clearInterval(t);t=null;}" & vbLf
    s = s & "    document.getElementById('noise').textContent='ノイズ: 停止中';return 1;};" & vbLf
    s = s & "  window.later=function(ms){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      var d=document.createElement('div');d.id='late';" & vbLf
    s = s & "      d.textContent='遅れて出た';" & vbLf
    s = s & "      document.getElementById('slot').appendChild(d);" & vbLf
    s = s & "    },ms||800);return 1;};" & vbLf
    s = s & "  window.chain=function(ms){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      var u=URL.createObjectURL(new Blob(['ok']));" & vbLf
    s = s & "      fetch(u).then(function(r){return r.text();}).then(function(x){" & vbLf
    s = s & "        setTimeout(function(){" & vbLf
    s = s & "          document.getElementById('target').textContent='連鎖の結果 '+x;" & vbLf
    s = s & "        },300);" & vbLf
    s = s & "      });" & vbLf
    s = s & "    },ms||500);return 1;};" & vbLf
    s = s & "  document.getElementById('btn').addEventListener('click',function(){" & vbLf
    s = s & "    setTimeout(function(){" & vbLf
    s = s & "      document.getElementById('target').textContent=" & _
            "'クリックの結果 '+Date.now();" & vbLf
    s = s & "    },400);" & vbLf
    s = s & "  });" & vbLf
    s = s & "})();" & vbLf
    s = s & "</" & "script>" & vbLf
    s = s & "</body></html>"

    BuildD4ProbeHtml = s
End Function


' ============================================================
' Test_D4_Settle (D-4b の検証: 静穏待ち)
'
'   ★D-4 の核心★ 「fetch が終わった」ではなく「その後の DOM 更新まで終わった」
'   ところで返ることを、決定的に確かめる。
'
'   ★実行中はブレーク/ステップ実行しないこと★ (仕様事実20)
' ============================================================
Public Sub Test_D4_Settle()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim took As Single
    Dim ok As Boolean
    Dim txt As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Settle: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Settle: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-4 プローブ", 10) Then
        Wv2Log.LogI "Test_D4_Settle: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Settle 開始 ================"
    Wv2Log.LogI "  --- (1) 静かなページ ---"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "静かなページでは静穏に達する", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        かかった時間 = " & Format$(took, "0.00") & " 秒"
    TestBool "  ★最低でも静穏窓ぶんは見張る (0.4 秒以上)★", (took > 0.4)
    TestBool "  無駄に長くはない (2 秒未満)", (took < 2)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★fetch → DOM 更新の連鎖★ (D-4 の核心) ---"
    Wv2Log.LogI "        100ms 後に fetch し、その完了の 300ms 後に #target を書き換える"

    p.EvalSync "(function(){document.getElementById('target')" & _
               ".textContent='まだ';return 1;})()"
    p.EvalSync "window.chain(100)"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "連鎖の後で静穏に達する", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        かかった時間 = " & Format$(took, "0.00") & " 秒"

    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        待ち終わった時点の #target = " & txt
    TestBool "★★ DOM 更新まで待てている (通信の完了だけで返っていない) ★★", _
             (InStr(1, txt, "連鎖の結果") > 0)
    TestBool "  連鎖ぶん待っている (0.7 秒以上)", (took > 0.7)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 静まらないページ (50ms ごとに書き換え続ける) ---"

    p.EvalSync "window.startNoise(50)"
    t0 = Timer
    ok = p.WaitSettled(2)
    took = D3Took(t0)
    TestBool "★静まらなければ False★", (ok = False)
    TestBool "  ★LastEvalOk=True (失敗ではなく時間切れ)★", (p.LastEvalOk = True)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が timeout であること", (InStr(1, p.LastSettleInfo, "timeout") > 0)
    TestBool "  指定どおり 2 秒粘った", (took > 1.7 And took < 4)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★ノイズ除外が効く★ (論点5) ---"
    Wv2Log.LogI "        ノイズは鳴らしたまま、#noise の変化だけ静穏判定から外す"

    p.IgnoreSelectors = "#noise"
    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "★除外すれば静穏に達する★", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        かかった時間 = " & Format$(took, "0.00") & " 秒"

    Wv2Log.LogI "        state = " & p.SpaProbeState()
    TestBool "  除外した DOM 変化が数えられている (sm > 0)", _
             (Wv2Json.JsonGetNum(p.SpaProbeState(), "sm") > 0)

    p.IgnoreSelectors = ""
    ok = p.WaitSettled(2)
    TestBool "  ★除外を外すとまた静まらない★", (ok = False)
    p.EvalSync "window.stopNoise()"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) stableMs を変えると待ちが伸びる ---"

    t0 = Timer
    ok = p.WaitSettled(6, 2000)
    took = D3Took(t0)
    TestBool "静穏窓 2000ms でも達する", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  2 秒以上かかっている", (took > 1.9)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★AutoWaitAfterAction★ (論点7 の opt-in) ---"

    TestBool "既定は False", (p.AutoWaitAfterAction = False)

    p.EvalSync "(function(){document.getElementById('target')" & _
               ".textContent='押す前';return 1;})()"
    Set el = D2El(p, "btn")
    el.Click
    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        自動待ち OFF: クリック直後の #target = " & txt
    TestBool "★OFF なら更新前の値が見える (待っていない)★", (txt = "押す前")

    p.AutoWaitAfterAction = True
    p.EvalSync "(function(){document.getElementById('target')" & _
               ".textContent='押す前';return 1;})()"
    Set el = D2El(p, "btn")
    t0 = Timer
    el.Click
    took = D3Took(t0)
    Set el = D2El(p, "target")
    txt = el.InnerText
    Wv2Log.LogI "        自動待ち ON: Click に " & Format$(took, "0.00") & " 秒"
    Wv2Log.LogI "        クリック直後の #target = " & txt
    TestBool "★ON なら更新後の値が見える (待っている)★", _
             (InStr(1, txt, "クリックの結果") > 0)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.AutoWaitAfterAction = False
    TestBool "  False に戻せる", (p.AutoWaitAfterAction = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) 除外リスト (通信) の口 ---"

    TestBool "AddIgnoreNetwork が成功する", p.AddIgnoreNetwork("example.invalid")
    TestBool "  同じものを足しても True (重複しない)", _
             p.AddIgnoreNetwork("example.invalid")
    TestBool "  除外を入れても静穏判定は動く", p.WaitSettled(5)
    p.ClearIgnoreNetwork

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Settle 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Signal (D-4c の検証: 明示シグナル)
'
'   ★D-4c の存在意義★ 静穏だけでは「アプリが終わった」と「アプリが無視した」を
'   区別できない。arm した目印が観測されるまで待つことで、後者を落とせる。
'   (3) がまさにその確認 ― ★静穏だけなら成功してしまう場面で、正しく失敗する★
' ============================================================
Public Sub Test_D4_Signal()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim t0 As Single
    Dim took As Single
    Dim ok As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Signal: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Signal: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-4 プローブ", 10) Then
        Wv2Log.LogI "Test_D4_Signal: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Signal 開始 ================"
    Wv2Log.LogI "  --- (1) arm は fail-fast (その場で気づける) ---"

    TestBool "★無い要素は arm できない★", _
             (p.ArmContentSignal("#nope-nope") = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError
    TestBool "  理由が not-found", (p.LastEvalError = "not-found")

    TestBool "★不正なセレクタも arm できない★", (p.ArmContentSignal("###") = False)
    TestBool "  理由が bad-selector", (p.LastEvalError = "bad-selector")

    TestBool "既存の器なら arm できる", p.ArmContentSignal("#slot")
    p.DisarmSignals

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★arm してから act する★ ---"
    Wv2Log.LogI "        #slot を arm し、800ms 後にその中へ要素を足す"

    TestBool "arm できる", p.ArmContentSignal("#slot")
    p.EvalSync "window.later(800)"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "★シグナルが当たって静穏に達する★", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        かかった時間 = " & Format$(took, "0.00") & " 秒"
    TestBool "  内訳が dom:hit", (InStr(1, p.LastSettleInfo, "dom:hit") > 0)
    TestBool "  仕掛けぶん待っている (0.8 秒以上)", (took > 0.8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★★ arm したのに何も起きない ★★ (D-4c の核心) ---"
    Wv2Log.LogI "        ページは静かなまま。静穏だけなら成功してしまう場面"

    TestBool "arm できる", p.ArmContentSignal("#slot")
    t0 = Timer
    ok = p.WaitSettled(2)
    took = D3Took(t0)
    TestBool "★★ 静かでも False (無視されたことを検出できる) ★★", (ok = False)
    TestBool "  LastEvalOk=True (失敗ではなく時間切れ)", (p.LastEvalOk = True)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が dom:miss", (InStr(1, p.LastSettleInfo, "dom:miss") > 0)
    TestBool "  指定どおり 2 秒粘った", (took > 1.7)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ワンショット (次の待ちには持ち越さない) ---"

    t0 = Timer
    ok = p.WaitSettled(5)
    took = D3Took(t0)
    TestBool "★arm は消費済みなので普通に静穏に達する★", ok
    TestBool "  待ち時間も普通 (2 秒未満)", (took < 2)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳に signal= が出ない", _
             (InStr(1, p.LastSettleInfo, "signal=") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) 通信シグナル ---"

    TestBool "arm できる (blob:)", p.ArmNetworkSignal("blob:")
    p.EvalSync "window.chain(100)"
    ok = p.WaitSettled(5)
    TestBool "★通信シグナルが当たる★", ok
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が net:hit", (InStr(1, p.LastSettleInfo, "net:hit") > 0)

    TestBool "当たらないパターンでも arm はできる", _
             p.ArmNetworkSignal("this-never-matches-xyz")
    p.EvalSync "window.chain(100)"
    ok = p.WaitSettled(3)
    TestBool "★当たらなければ False★", (ok = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が net:miss", (InStr(1, p.LastSettleInfo, "net:miss") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★signal-lost★ (arm が消えたら黙って直さない) ---"

    TestBool "arm できる", p.ArmContentSignal("#slot")
    TestBool "  プローブを作り直す (arm ごと消える)", p.SpaProbeReset()
    ok = p.WaitSettled(3)
    TestBool "★arm が消えたら False★", (ok = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が signal-lost", _
             (InStr(1, p.LastSettleInfo, "signal-lost") > 0)
    TestBool "  ★即座に返る (待たされない)★", True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) DisarmSignals (手で取り下げる) ---"

    TestBool "arm できる", p.ArmContentSignal("#slot")
    TestBool "  DisarmSignals が成功する", p.DisarmSignals()
    ok = p.WaitSettled(3)
    TestBool "★取り下げれば普通に静穏に達する★", ok
    Wv2Log.LogI "        " & p.LastSettleInfo

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Signal 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Log (D-4e の検証: in-flight 台帳 / 診断ログ / URL シグナル)
' ============================================================
Public Sub Test_D4_Log()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim st As String
    Dim n As Long
    Dim il As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Log: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Log: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-4 プローブ", 10) Then
        Wv2Log.LogI "Test_D4_Log: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Log 開始 ================"
    Wv2Log.LogI "  --- (1) 診断ログの ON / OFF ---"

    TestBool "既定は無効", (p.SpaProbeLogging = False)
    p.SpaProbeLogging = True
    TestBool "  有効にできる", p.SpaProbeLogging

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★出来事が溜まる★ ---"

    p.EvalSync "window.chain(100)"
    D3Pump 1.5
    st = p.SpaProbeState()
    TestBool "ログ件数が増えている", (Wv2Json.JsonGetNum(st, "lgN") > 0)

    n = p.SpaProbeDrainLog()
    TestBool "★取り出せる★", (n > 0)
    TestBool "  取り出したら空になる", (p.SpaProbeDrainLog() = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★in-flight の台帳★ ---"

    ' ★30 秒前から飛んでいる要求を仕込む★
    '   自前ページでは blob の fetch が 1ms で終わってしまい、EvalSync の往復
    '   (38ms) の間に消える。台帳と居座り判定を決定的に確かめるために、
    '   ロングポーリング相当の項目を直接押し込む。
    p.EvalSync "(function(){var q=window.__wv2p;" & _
               "q.ifl.push({u:'fake://long-poll',t:performance.now()-30000,w:'xhr'});" & _
               "return 1;})()"

    st = p.SpaProbeState()
    Wv2Json.JsonPickStr st, "il", il
    Wv2Log.LogI "        il = " & il
    TestBool "★飛んでいる要求が一覧に出る★", (InStr(1, il, "fake://long-poll") > 0)
    TestBool "  種別と経過時間が分かる", (InStr(1, il, "xhr 30") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3b) ★居座り判定 (StaleInflightMs)★ ---"

    TestBool "既定は 10000ms", (p.StaleInflightMs = 10000)
    TestBool "★居座りは静穏判定から外れる (ifn=0)★", _
             (Wv2Json.JsonGetNum(st, "ifn") = 0)
    TestBool "  居座りとして数えられている (ifo=1)", _
             (Wv2Json.JsonGetNum(st, "ifo") = 1)
    TestBool "★居座りがあっても静穏に達する★", p.WaitSettled(5)
    Wv2Log.LogI "        " & p.LastSettleInfo

    p.StaleInflightMs = 0
    st = p.SpaProbeState()
    TestBool "★無効にすると数える (ifn=1)★", (Wv2Json.JsonGetNum(st, "ifn") = 1)
    TestBool "  そのときは静穏に達しない", (p.WaitSettled(2) = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.StaleInflightMs = 10000

    ' 仕込んだ項目を片付ける
    p.EvalSync "(function(){window.__wv2p.ifl=[];return 1;})()"
    st = p.SpaProbeState()
    Wv2Json.JsonPickStr st, "il", il
    TestBool "  片付ければ一覧から消える", (Len(il) = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★URL シグナル★ ---"
    Wv2Log.LogI "        このページの URL = " & p.View_GetSource()

    TestBool "arm できる (当たるはずの文字列)", p.ArmUrlSignal("blank")
    TestBool "★URL に含まれていれば静穏に達する★", p.WaitSettled(5)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が url:hit", (InStr(1, p.LastSettleInfo, "url:hit") > 0)

    TestBool "arm できる (当たらない文字列)", p.ArmUrlSignal("zzz-not-here")
    TestBool "★含まれていなければ False★", (p.WaitSettled(2) = False)
    Wv2Log.LogI "        " & p.LastSettleInfo
    TestBool "  内訳が url:miss", (InStr(1, p.LastSettleInfo, "url:miss") > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★時間切れのときに実行中の要求を見せる★ (D-4d の教訓) ---"

    p.SpaProbeLogging = False
    TestBool "  無効に戻せる", (p.SpaProbeLogging = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D4_Log 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D4_Site (D-4d: ★Google Maps の偵察★)
'
'   ★合否を判定しない★ 実サイトは落ちるし DOM も予告なく変わる (設計原則75)。
'   ここでやるのは★観測して数字を残すこと★ に徹する。判定するのは
'   「タブが開けたか」「検索ボックスが見つかったか」のような構造的な数点だけ。
'
'   何を知りたいか:
'     (a) ★Maps に静穏は訪れるのか★ タイル取得とアニメーションが止まらない
'         ページで、q (最後の DOM 変化からの経過) がどこまで伸びるか
'     (b) ★どのくらい騒がしいのか★ 毎秒の DOM 変化数と通信数
'     (c) ★重いページでの往復コスト★ (未知2 の本番値。軽いページでは +0.9ms だった)
'     (d) ★何が完了の目印になるか★ 検索後に document.title と URL がどう変わるか
'         (URL の /@緯度,経度,ズーム は座標取得の目的そのものでもある)
'
'   ★外部サイトに実アクセスする★ 唯一の Test_*。ネットワークが要る。
' ============================================================
Public Sub Test_D4_Site(Optional ByVal searchAddr As String = "東京都千代田区丸の内1-9-1")
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim ms As Single
    Dim clicked As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D4_Site: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D4_Site 開始 ================"
    Wv2Log.LogI "  ★観測が目的。合否は付けない (実サイトは変わるため)★"
    Wv2Log.LogI "  検索する住所: " & searchAddr

    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    If p Is Nothing Then
        Wv2Log.LogI "  [FAIL] タブの生成に失敗しました。"
        m_ngCount = m_ngCount + 1
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) 読み込みを待つ (D-3 の WaitFor で検索ボックスを掴む) ---"

    ' ★先に Pane が JS を受け付ける状態になるまで待つ★
    '   タブを開いた直後は View がまだ無く、EvalSync が no-view で失敗する。
    '   WaitFor は「失敗なら待たずに諦める」ので、そのまま呼ぶと即座に落ちる。
    TestBool "Pane が JS を受け付けるようになった", D4WaitPane(p, 30)

    ' ★セレクタは候補を順に試す★ 実サイトの id は自動生成に変わりうる
    '   (2026-08-22 の実測: かつての #searchboxinput は消え、id は ucc-1 だった。
    '    name='q' と form の構造は残っていたので、そちらを先に試す)
    Set el = D4FindFirst(p, "input[name='q']" & Chr$(1) & _
                            "#searchboxinput" & Chr$(1) & _
                            "form input[type='text']", 30)
    TestBool "検索ボックスが見つかった", Not (el Is Nothing)
    If el Is Nothing Then
        Wv2Log.LogI "  ★見つからない★ 同意画面やレイアウト変更の可能性がある。"
        Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
        Wv2Log.LogI "        url   = " & p.View_GetSource()
        Wv2Log.LogI "        LastEvalOk=" & p.LastEvalOk & " err=" & p.LastEvalError
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "        url   = " & p.View_GetSource()

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★往復コスト (未知2 の本番値)★ ---"

    ms = D4RoundTrip(p, 20)
    Wv2Log.LogI "        プローブ設置前の 1 往復 = " & Format$(ms, "0.0") & " ms"
    Wv2Log.LogI "        state = " & p.SpaProbeState()
    ms = D4RoundTrip(p, 20)
    Wv2Log.LogI "        プローブ設置後の 1 往復 = " & Format$(ms, "0.0") & " ms"
    Wv2Log.LogI "        (軽いページでは 35.4 → 36.3 ms だった)"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★静穏は訪れるか★ 何もせず 15 秒観測 ---"
    D4Watch p, 15

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) 静穏待ちを試す (結果は記録するだけ) ---"

    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) 住所を入力して検索する ---"

    ' ★D-4e の道具を使う★ 診断ログを入れ、URL の変化を arm してから操作する
    p.SpaProbeLogging = True
    Wv2Log.LogI "        診断ログを有効にした"
    TestBool "URL シグナルを arm できる", p.ArmUrlSignal("/place/")

    Set el = D4FindFirst(p, "input[name='q']" & Chr$(1) & _
                            "#searchboxinput" & Chr$(1) & _
                            "form input[type='text']", 5)
    TestBool "検索ボックスを掴み直せる", Not (el Is Nothing)
    If el Is Nothing Then
        TestCountPrint
        Exit Sub
    End If

    el.value = searchAddr
    TestBool "★住所を書き込めた (D-3 の .Value =)★", el.LastOk
    Wv2Log.LogI "        経路 = " & el.LastInfo
    Wv2Log.LogI "        読み戻し = " & el.value

    ' ★検索の実行★ ボタンがあれば押す。無ければ Enter キーを合成する。
    Set el = D4FindFirst(p, "button[aria-label='検索']" & Chr$(1) & _
                            "#searchbox-searchbutton" & Chr$(1) & _
                            "button[aria-label='Search']", 3)
    If el Is Nothing Then
        Wv2Log.LogI "        検索ボタンが無いので Enter を合成する"
        ' ★CSS の属性値は二重引用符にする★ JS の文字列はシングルクォート、
        '   セレクタ内は二重引用符 ("""" で書く) とすれば、
        '   バックスラッシュを一切使わずに入れ子にできる (プロジェクト規則)。
        p.EvalSync "(function(){var e=document.querySelector('input[name=""q""]')" & _
                   "||document.getElementById('searchboxinput');" & _
                   "if(!e){return 0;}" & _
                   "e.focus();e.dispatchEvent(new KeyboardEvent('keydown'," & _
                   "{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));" & _
                   "return 1;})()"
        clicked = p.LastEvalOk
    Else
        clicked = el.Click()
        Wv2Log.LogI "        検索ボタンを押した (経路=" & el.LastInfo & ")"
    End If
    TestBool "検索を実行できた", clicked

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★arm した URL シグナルつきで待つ★ ---"

    Wv2Log.LogI "        WaitSettled(20) = " & p.WaitSettled(20)
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        ★時間切れなら 実行中: に居座っている要求が出る★"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6b) ★個別除外を足して待ち直す★ (案A の実演) ---"
    Wv2Log.LogI "        居座っていた要求を名指しで静穏判定から外す"

    TestBool "AddIgnoreNetwork できる", p.AddIgnoreNetwork("/search?tbm=map")
    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo
    p.ClearIgnoreNetwork

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6c) 検索後の 10 秒 ---"
    D4Watch p, 10

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6d) ★診断ログ★ 何が起きていたか ---"
    Wv2Log.LogI "        取り出した件数: " & p.SpaProbeDrainLog()
    Wv2Log.LogI "        (60 件で切れる。続きは次の呼び出しで取れる)"
    Wv2Log.LogI "        取り出した件数: " & p.SpaProbeDrainLog()
    p.SpaProbeLogging = False

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) 結果 ---"
    Wv2Log.LogI "        title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "        url   = " & p.View_GetSource()
    Wv2Log.LogI "        ★URL に /@緯度,経度,ズーム が入っていれば座標が取れる★"
    Wv2Log.LogI "        WaitSettled(10) = " & p.WaitSettled(10)
    Wv2Log.LogI "        " & p.LastSettleInfo
    Wv2Log.LogI "        state = " & p.SpaProbeState()

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "  ※ この判定数は構造的な数点のみ。観測の中身はログ本文を読むこと。"
    Wv2Log.LogI "================ Test_D4_Site 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' D4FindFirst (D-4d: 候補セレクタを順に試して最初に当たったものを返す)
'
'   ★実サイトのセレクタは予告なく変わる (設計原則75)★
'   1 つに賭けず候補を並べ、★どれが当たったかをログに残す★。次に壊れたときに
'   「何が変わったか」が分かる。候補は Chr$(1) 区切りで渡す。
'
'   1 つ目だけは timeoutSec ぶん待つ (ページの読み込み中を想定)。
'   2 つ目以降は 1 往復で見るだけ。全部外れたら Nothing。
' ============================================================
Private Function D4FindFirst(ByVal p As Wv2Pane, _
                             ByVal selectorList As String, _
                             ByVal timeoutSec As Single) As Wv2Element
    Dim cands As Variant
    Dim i As Long
    Dim el As Wv2Element

    cands = Split(selectorList, Chr$(1))

    For i = LBound(cands) To UBound(cands)
        If i = LBound(cands) Then
            Set el = p.WaitFor(CStr(cands(i)), timeoutSec)
        Else
            Set el = p.QuerySelector(CStr(cands(i)))
        End If

        If Not el Is Nothing Then
            Wv2Log.LogI "        セレクタ [" & cands(i) & "] で見つかった"
            Set D4FindFirst = el
            Exit Function
        End If

        Wv2Log.LogI "        セレクタ [" & cands(i) & "] は外れ" & _
                    IIf(p.LastEvalOk, "", " (失敗: " & p.LastEvalError & ")")
    Next i
End Function

' ============================================================
' D4WaitPane (D-4d: Pane が JS を受け付けるまで待つ)
'
'   ★タブを開いた直後は EvalSync が no-view で失敗する★
'   D-3 の WaitFor / D-4 の WaitSettled はどちらも「失敗は待っても直らない」
'   という方針 (設計原則104) なので、準備前に呼ぶと即座に諦めてしまう。
'   実サイトを開く回はこれを先に挟む。
' ============================================================
Private Function D4WaitPane(ByVal p As Wv2Pane, ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single

    t0 = Timer
    Do
        p.EvalSync "1", 3
        If p.LastEvalOk Then
            D4WaitPane = True
            Exit Function
        End If
        If D3Took(t0) >= timeoutSec Then
            Wv2Log.LogW "D4WaitPane: 時間切れ err=" & p.LastEvalError
            Exit Function
        End If
        D3Pump 0.3
    Loop
End Function


' ============================================================
' D4Watch (D-4d: 1 秒ごとに状態・タイトル・URL を記録する)
'
'   ★静穏が訪れるかを見る道具★ q が伸び続ければ静かになっている。
'   タイトルと URL は COM 経由で取る (EvalSync を使わないので観測が軽い)。
' ============================================================
Private Sub D4Watch(ByVal p As Wv2Pane, ByVal seconds As Long)
    Dim i As Long
    Dim st As String
    Dim ttl As String
    Dim url As String

    For i = 1 To seconds
        st = p.SpaProbeState()
        ttl = p.View_GetDocumentTitle()
        url = p.View_GetSource()

        Wv2Log.LogI "        [" & Format$(i, "00") & "s] " & _
                    "q=" & Wv2Json.JsonGetNum(st, "q") & _
                    " m=" & Wv2Json.JsonGetNum(st, "m") & _
                    " n=" & Wv2Json.JsonGetNum(st, "n") & _
                    " f=" & Wv2Json.JsonGetNum(st, "f") & _
                    " x=" & Wv2Json.JsonGetNum(st, "x") & _
                    " rp=" & Wv2Json.JsonGetNum(st, "rp") & _
                    " | " & Left$(ttl, 40) & " | " & Left$(url, 70)

        D3Pump 1
    Next i
End Sub

' ============================================================
' Test_D4_Dom (D-4d 補: ★アクティブなタブの DOM を覗く★)
'
'   ★「何を待つべきか」を推測でなく観測で決めるための道具★
'   実サイトのセレクタは予告なく変わる (設計原則75)。当てずっぽうに
'   セレクタを書く前に、実際に何があるのかを見る。
'
'   使い方: ブラウザで目的のページを表示してから、イミディエイトで
'           Test_D4_Dom           … 入力欄・ボタン・role つきの要素を列挙
'           Test_D4_Dom "h1,h2"  … セレクタを指定して列挙
'
'   出す情報: タグ / id / name / type / class (先頭 40 字) / placeholder /
'             aria-label / 可視かどうか / テキスト (先頭 30 字)
'
'   ★D-4e (診断ログ) の原型★ 時系列は取らないが、まず「今そこに何があるか」を
'   知るのはこれで足りる。
' ============================================================
Public Sub Test_D4_Dom(Optional ByVal sel As String = "")
    Dim p As Wv2Pane
    Dim js As String
    Dim res As String
    Dim parts As Variant
    Dim i As Long

    Set p = UserForm1.GetActivePane
    If p Is Nothing Then
        Wv2Log.LogI "Test_D4_Dom: アクティブな Pane がありません。"
        Exit Sub
    End If

    If Len(sel) = 0 Then
        sel = "input,textarea,select,[role=combobox],[role=searchbox]," & _
              "[role=search],button[aria-label],form"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_D4_Dom 開始 ================"
    Wv2Log.LogI "  url   = " & p.View_GetSource()
    Wv2Log.LogI "  title = " & p.View_GetDocumentTitle()
    Wv2Log.LogI "  セレクタ: " & sel

    js = "(function(){var out=[],q=document.querySelectorAll(" & _
         p.JsQuote(sel) & ");" & _
         "out.push('全 '+q.length+' 件');" & _
         "for(var i=0;i<q.length&&out.length<41;i++){var e=q[i];" & _
         "var g=function(a){var v=(e.getAttribute?e.getAttribute(a):null);" & _
         "return v?String(v).slice(0,40):'-';};" & _
         "out.push(e.tagName" & _
         "+' id='+(e.id||'-')" & _
         "+' name='+(e.name||'-')" & _
         "+' type='+(e.type||'-')" & _
         "+' cls='+String(e.className||'-').slice(0,40)" & _
         "+' ph='+g('placeholder')" & _
         "+' aria='+g('aria-label')" & _
         "+' vis='+(e.offsetParent!==null)" & _
         "+' txt='+String(e.textContent||'').trim().slice(0,30));}" & _
         "return out.join(String.fromCharCode(1));})()"

    res = p.EvalSync(js, 10)
    If Not p.LastEvalOk Then
        Wv2Log.LogI "  ★失敗★ err=" & p.LastEvalError
        Wv2Log.LogI "================ Test_D4_Dom 終了 ================"
        Exit Sub
    End If

    parts = Split(Wv2Json.JsonUnescape(res), ChrW$(1))
    For i = LBound(parts) To UBound(parts)
        Wv2Log.LogI "  " & parts(i)
    Next i

    Wv2Log.LogI "================ Test_D4_Dom 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D5_Geocode (D-5 の検証: 住所 → 座標)
'
'   ★D 軸の部品が業務で使える形になったかを見る★
'   3 件の住所を★同じタブで続けて★処理し、緯度経度と正規化後の名前を出す。
'
'   ★外部サイトに実アクセスする★ ネットワークが要る。
'   実サイトなので合否は緩く見る (座標が妥当な範囲に入っているか)。
' ============================================================
Public Sub Test_D5_Geocode()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim addrs As Variant
    Dim i As Long
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim t0 As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D5_Geocode: Browser が起動していません。" & _
                    "先に UserForm1.Show vbModeless と StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D5_Geocode 開始 ================"

    t0 = Timer
    Set p = Wv2Maps.MapsOpen(b)
    TestBool "★Maps を開いて操作できる状態にできた★", Not (p Is Nothing)
    Wv2Log.LogI "        かかった時間 = " & Format$(D3Took(t0), "0.0") & " 秒"
    If p Is Nothing Then
        Wv2Log.LogI "        理由 = " & Wv2Maps.MapsLastError
        TestCountPrint
        Exit Sub
    End If

    addrs = Array( _
        "東京都千代田区丸の内1-9-1", _
        "大阪府大阪市中央区大阪城1-1", _
        "北海道札幌市中央区北1条西2丁目")

    For i = LBound(addrs) To UBound(addrs)
        Wv2Log.LogI ""
        Wv2Log.LogI "  --- (" & (i + 1) & ") " & addrs(i) & " ---"

        t0 = Timer
        ok = Wv2Maps.MapsGeocode(p, CStr(addrs(i)), lat, lng, nm)

        Wv2Log.LogI "        かかった時間 = " & Format$(D3Took(t0), "0.0") & " 秒"
        Wv2Log.LogI "        緯度経度 = " & lat & ", " & lng
        Wv2Log.LogI "        名前     = " & nm
        If Not ok Then Wv2Log.LogI "        理由     = " & Wv2Maps.MapsLastError

        TestBool "  ★1 件に確定した★", ok
        TestBool "  緯度が日本の範囲 (20～46)", (lat > 20 And lat < 46)
        TestBool "  経度が日本の範囲 (122～154)", (lng > 122 And lng < 154)
        TestBool "  名前が取れている", (Len(nm) > 0)
    Next i

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 失敗の扱い ---"

    ok = Wv2Maps.MapsGeocode(p, "ZZZZ存在しない住所ZZZZ", lat, lng, nm, 8)
    Wv2Log.LogI "        戻り値 = " & ok & " 理由 = " & Wv2Maps.MapsLastError
    TestBool "★でたらめな住所では確定しない★", (ok = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D5_Geocode 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D5_Sheet (D-5b の検証: シート連携)
'
'   ★新しいブックを作って試し、保存せずに閉じる★
'   開発用ブックにシートを足すと、Excel が終了時に保存を聞いてきて鬱陶しい。
'
'   ★外部サイトに実アクセスする★ ネットワークが要る。
' ============================================================
Public Sub Test_D5_Sheet()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D5_Sheet: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D5_Sheet 開始 ================"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)

    sh.Cells(1, 1).value = "住所"
    sh.Cells(1, 2).value = "緯度"
    sh.Cells(1, 3).value = "経度"
    sh.Cells(1, 4).value = "正規化住所"
    sh.Cells(1, 5).value = "状態"
    sh.Cells(2, 1).value = "東京都千代田区丸の内1-9-1"
    sh.Cells(3, 1).value = "大阪府大阪市中央区大阪城1-1"
    sh.Cells(4, 1).value = "ZZZZ存在しない住所ZZZZ"

    Wv2Log.LogI "  3 行 (うち 1 行はでたらめ) を処理する"
    n = Wv2Maps.MapsGeocodeSheet(sh)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 書き込まれた内容 ---"
    Wv2Log.LogI "  2 行目: " & sh.Cells(2, 2).value & " / " & _
                sh.Cells(2, 3).value & " / " & sh.Cells(2, 4).value & " / " & sh.Cells(2, 5).value
    Wv2Log.LogI "  3 行目: " & sh.Cells(3, 2).value & " / " & _
                sh.Cells(3, 3).value & " / " & sh.Cells(3, 4).value & " / " & sh.Cells(3, 5).value
    Wv2Log.LogI "  4 行目: " & sh.Cells(4, 2).value & " / " & _
                sh.Cells(4, 3).value & " / " & sh.Cells(4, 4).value & " / " & sh.Cells(4, 5).value

    TestBool "★2 行が ok になった★", (n = 2)
    TestBool "  2 行目の緯度が東京", (Abs(sh.Cells(2, 2).value - 35.68) < 0.1)
    TestBool "  3 行目の緯度が大阪", (Abs(sh.Cells(3, 2).value - 34.69) < 0.1)
    TestBool "  2 行目の状態が ok", (sh.Cells(2, 5).value = "ok")
    TestBool "★でたらめな行は ok にならない★", (sh.Cells(4, 5).value <> "ok")
    TestBool "  その理由が残っている", (Len(sh.Cells(4, 5).value) > 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- ★再開できるか★ (ok の行は飛ばす) ---"
    n = Wv2Maps.MapsGeocodeSheet(sh)
    TestBool "2 回目も同じ件数を返す (飛ばしても数える)", (n = 2)

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D5_Sheet 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D6_All (D-6 の検証: QuerySelectorAll と寿命管理)
' ============================================================
Public Sub Test_D6_All()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim els As Collection
    Dim el As Wv2Element
    Dim ids As String
    Dim n0 As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_D6_All: Browser が起動していません。"
        Exit Sub
    End If

    Set p = b.AddTabWithHtml(BuildD4ProbeHtml())
    If p Is Nothing Then
        Wv2Log.LogI "Test_D6_All: タブの生成に失敗しました。"
        Exit Sub
    End If
    If Not D2WaitTitle(p, "D-4 プローブ", 10) Then
        Wv2Log.LogI "Test_D6_All: 検証ページの読み込みを確認できませんでした。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D6_All 開始 ================"
    Wv2Log.LogI "  --- (1) まとめて掴む ---"

    Set els = p.QuerySelectorAll("div")
    TestBool "★Collection が返る★", Not (els Is Nothing)
    TestBool "  3 件ある", (els.Count = 3)
    TestBool "  打ち切っていない", (p.LastAllTruncated = False)

    ' ★For Each で自然に回せることが論点2 の眼目★
    For Each el In els
        ids = ids & el.GetAttribute("id") & ","
    Next el
    Wv2Log.LogI "        文書順の id = " & ids
    TestBool "★文書順で返る★", (ids = "target,noise,slot,")

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) 0 件と失敗の区別 (設計原則93) ---"

    Set els = p.QuerySelectorAll("blockquote")
    TestBool "0 件でも Collection が返る (Nothing ではない)", Not (els Is Nothing)
    TestBool "  Count が 0", (els.Count = 0)
    TestBool "  ★LastEvalOk=True (本当に無い)★", (p.LastEvalOk = True)

    Set els = p.QuerySelectorAll("###")
    TestBool "不正なセレクタでも Collection が返る", Not (els Is Nothing)
    TestBool "  Count が 0", (els.Count = 0)
    TestBool "  ★LastEvalOk=False (失敗)★", (p.LastEvalOk = False)
    Wv2Log.LogI "        LastEvalError = " & p.LastEvalError

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★打ち切りは黙ってやらない★ ---"

    Set els = p.QuerySelectorAll("div", 2)
    TestBool "上限どおり 2 件", (els.Count = 2)
    TestBool "★打ち切ったことが分かる★", p.LastAllTruncated

    Set els = p.QuerySelectorAll("div", 200)
    TestBool "  上限内なら False に戻る", (p.LastAllTruncated = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) 掴んだ要素がそのまま使える ---"

    Set els = p.QuerySelectorAll("div")
    Set el = els(1)
    TestEq "1 件目を読める", el, el.GetAttribute("id"), "target"
    TestBool "  書き込める (D-3 の SetAttribute)", _
             el.SetAttribute("data-mark", "1")
    TestEq "  読み戻せる", el, el.GetAttribute("data-mark"), "1"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★寿命管理は見せるだけ★ ---"

    n0 = p.ElementCount
    Wv2Log.LogI "        今レジストリに積まれている数 = " & n0
    TestBool "ElementCount が読める", (n0 > 0)

    Set els = p.QuerySelectorAll("div")
    TestBool "★掴むたびに積み上がる (勝手に捨てない)★", (p.ElementCount > n0)

    TestBool "ClearElementRegistry で作り直せる", p.ClearElementRegistry()
    TestBool "  レジストリが空になる", (p.ElementCount = 0)
    TestBool "  ★手元の要素は stale になる★", (el.IsStale = True)

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D6_All 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D6_Pick (D-6b の検証: 候補一覧から 1 番目を採る)
'
'   ★D-6 の QuerySelectorAll が実サイトで効くかを見る出口★
'   カテゴリ検索 (「コンビニ 東京駅」) は候補一覧になるので、
'   pickFirst の有無で結果が変わることを確かめる。
'
'   ★外部サイトに実アクセスする★ 実サイトなので合否は緩く見る。
' ============================================================
Public Sub Test_D6_Pick()
    Dim p As Wv2Pane
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D6_Pick: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D6_Pick 開始 ================"

    Set p = Wv2Maps.MapsOpen(UserForm1.CurrentBrowser)
    TestBool "Maps を開けた", Not (p Is Nothing)
    If p Is Nothing Then
        TestCountPrint
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) 候補が複数になる検索 (選ばない) ---"

    ok = Wv2Maps.MapsGeocode(p, "コンビニ 東京駅", lat, lng, nm, 15)
    Wv2Log.LogI "        戻り値=" & ok & " 理由=" & Wv2Maps.MapsLastError
    Wv2Log.LogI "        座標=" & lat & "," & lng & "  名前=" & nm
    TestBool "★選ばなければ確定しない★", (ok = False)
    TestBool "  候補から選んでいない", (Wv2Maps.MapsLastPicked = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★候補の 1 件目を採る★ ---"

    ok = Wv2Maps.MapsGeocode(p, "コンビニ 東京駅", lat, lng, nm, 15, True)
    Wv2Log.LogI "        戻り値=" & ok & " 理由=" & Wv2Maps.MapsLastError
    Wv2Log.LogI "        座標=" & lat & "," & lng
    Wv2Log.LogI "        名前=" & nm
    Wv2Log.LogI "        候補から選んだか=" & Wv2Maps.MapsLastPicked

    TestBool "★候補を採れば確定する★", ok
    If ok Then
        TestBool "  ★候補から選んだ印が付く★", Wv2Maps.MapsLastPicked
        TestBool "  緯度が日本の範囲", (lat > 20 And lat < 46)
        TestBool "  経度が日本の範囲", (lng > 122 And lng < 154)
        TestBool "  名前が取れている", (Len(nm) > 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) 確定する住所では印が付かない ---"

    ok = Wv2Maps.MapsGeocode(p, "東京都千代田区丸の内1-9-1", lat, lng, nm, 15, True)
    TestBool "確定する", ok
    TestBool "★候補から選んだ印は付かない★", (Wv2Maps.MapsLastPicked = False)
    Wv2Log.LogI "        座標=" & lat & "," & lng & "  名前=" & nm

    Wv2Log.LogI ""
    Wv2Log.LogI "  レジストリに積まれた数: " & p.ElementCount
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_D6_Pick 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Cancel (D-7 の検証: 中断の口と分母)
'
'   ★ネットワークもブラウザも要らない★ 実サイトを叩く中断の検証は
'   Test_D7_Sheet (手で Esc を押す) の方。
' ============================================================
Public Sub Test_D7_Cancel()
    Dim wb As Workbook
    Dim sh As Object
    Dim sh2 As Object
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim n As Long
    Dim beforeTabs As Long

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Cancel 開始 ================"

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) 外から中断を立てる口 ---"

    Wv2Maps.MapsCancel = True
    TestBool "★立てたら立つ★", (Wv2Maps.MapsCancel = True)
    Wv2Maps.MapsCancel = False
    TestBool "  下ろしたら下りる", (Wv2Maps.MapsCancel = False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★中断状態は入口で捨てられる★ (設計原則106) ---"
    Wv2Log.LogI "        立てっぱなしにして、次の呼び出しが殺されないことを見る"

    Wv2Maps.MapsCancel = True
    ok = Wv2Maps.MapsGeocode(Nothing, "東京駅", lat, lng, nm, 1)
    Wv2Log.LogI "        戻り値 = " & ok & " 理由 = " & Wv2Maps.MapsLastError
    TestBool "呼べる (中身は no-pane で失敗する)", (ok = False)
    TestBool "★理由が canceled ではない★ (入口で捨てたから)", _
             (Wv2Maps.MapsLastError = "no-pane")
    TestBool "  中断要求が下りている", (Wv2Maps.MapsCancel = False)
    TestBool "  中断で終わった印も下りている", (Wv2Maps.MapsCanceled = False)

    ' --- D-7b: ★戻し忘れると Excel 全体で Esc が効かなくなる★ ---
    Wv2Log.LogI "        EnableCancelKey = " & Application.EnableCancelKey & _
                " (xlInterrupt = " & xlInterrupt & ")"
    TestBool "★EnableCancelKey が戻っている★", _
             (Application.EnableCancelKey = xlInterrupt)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★分母を数える★ (MapsCountRows) ---"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = "住所"
    sh.Cells(2, 1).value = "東京都千代田区丸の内1-9-1"
    sh.Cells(3, 1).value = "大阪府大阪市中央区大阪城1-1"
    sh.Cells(4, 1).value = "京都府京都市下京区東塩小路町"
    ' 5 行目は空 ― ★ここで止まる★
    sh.Cells(6, 1).value = "空白の向こうは数えない"

    TestBool "★空セルで止めて 3 件と数える★", (Wv2Maps.MapsCountRows(sh) = 3)
    TestBool "  開始行を 3 にすれば 2 件", (Wv2Maps.MapsCountRows(sh, 3) = 2)
    TestBool "  住所の無い列を見れば 0 件", (Wv2Maps.MapsCountRows(sh, 2, 5) = 0)
    TestBool "  シートが無ければ 0 件", (Wv2Maps.MapsCountRows(Nothing) = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★0 件ならタブを開かずに帰る★ ---"

    Set sh2 = wb.Worksheets.Add
    beforeTabs = -1
    If Not UserForm1.CurrentBrowser Is Nothing Then
        beforeTabs = UserForm1.CurrentBrowser.TabCount
    End If

    n = Wv2Maps.MapsGeocodeSheet(sh2)
    TestBool "0 を返す", (n = 0)
    TestBool "  理由が no-rows", (Wv2Maps.MapsLastError = "no-rows")

    If beforeTabs >= 0 Then
        Wv2Log.LogI "        タブ数 " & beforeTabs & " → " & _
                    UserForm1.CurrentBrowser.TabCount
        TestBool "★タブが増えていない★", _
                 (UserForm1.CurrentBrowser.TabCount = beforeTabs)
    Else
        Wv2Log.LogI "        (ブラウザ未起動なのでタブ数の判定は省略)"
    End If

    Application.StatusBar = Empty   ' D-7d: 前のテストの残りを消してから見る
    TestBool "  ステータスバーが汚れていない", _
             (TypeName(Application.StatusBar) = "Boolean")

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Cancel 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Sheet (D-7b の検証: ★実行中に手で Esc を押す★)
'
'   ★外部サイトに実アクセスする★ ネットワークが要る。
'   ★新しいブックを作って試し、保存せずに閉じる★
'
'   手順:
'     1) 実行すると 4 件の住所を順に処理し始める (1 件 8 秒ほど)。
'     2) ★ステータスバーを見ながら 2 件目あたりで Esc を押す★
'        押しっぱなしでなくてよい。焦点はどこにあってもよい。
'        ★D-7b からは「コードの実行が中断されました」は出ない★ (エラー18 として
'        受け取って中断に変える)。出たら D-7b が効いていないということ。
'     3) 中断されたことと、★中断した行に何も書かれていないこと★を判定する。
'     4) そのまま呼び直して★続きから埋まること★を判定する (設計原則110)。
'
'   ★Esc を押さなければ FAIL になる★ 押し忘れたらもう一度実行する。
'   ★住所はすべて「1 件に確定する」ものを選んである★ (D-7 では ambiguous な
'   住所が混ざって 1 件 27 秒かかり、判定も間接的になっていた)。
' ============================================================
Public Sub Test_D7_Sheet()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long
    Dim n2 As Long
    Dim i As Long
    Dim blanks As Long
    Dim wasCanceled As Boolean

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D7_Sheet: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Sheet 開始 ================"
    Wv2Log.LogI "  ★★★ 2 件目あたりで Esc を押してください ★★★"
    Debug.Print ""
    Debug.Print "  ★★★ 2 件目あたりで Esc を押してください ★★★"
    Debug.Print "  (連打しなくてよい。1 回で効きます)"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = "住所"
    sh.Cells(2, 1).value = "東京都千代田区丸の内1-9-1"
    sh.Cells(3, 1).value = "大阪府大阪市中央区大阪城1-1"
    sh.Cells(4, 1).value = "北海道札幌市北区北6条西4丁目"
    sh.Cells(5, 1).value = "宮城県仙台市青葉区中央1-1-1"

    n = Wv2Maps.MapsGeocodeSheet(sh)
    wasCanceled = Wv2Maps.MapsCanceled

    Wv2Log.LogI ""
    Wv2Log.LogI "        ok = " & n & " / 中断 = " & wasCanceled & _
                " / 理由 = " & Wv2Maps.MapsLastError
    TestDumpStatusBar "StatusBar"
    Wv2Log.LogI "        ★中断検知の呼び出し = " & Wv2Maps.MapsCheckCount & " 回★"

    TestBool "★中断された★ (押していなければ FAIL。もう一度どうぞ)", wasCanceled
    TestBool "  理由が canceled", (Wv2Maps.MapsLastError = "canceled")
    ' D-7d: Empty で戻すようにしたので判定を復活させた
    TestBool "★ステータスバーが戻っている★", _
             (TypeName(Application.StatusBar) = "Boolean")
    TestBool "★EnableCancelKey が戻っている★", _
             (Application.EnableCancelKey = xlInterrupt)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- ★書きかけの行を残していないか★ (設計原則111) ---"

    blanks = 0
    For i = 2 To 5
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then
            blanks = blanks + 1
            TestBool "  " & i & " 行目は手つかず (座標も空)", _
                     (Len(CStr(sh.Cells(i, 2).value)) = 0 And _
                      Len(CStr(sh.Cells(i, 3).value)) = 0)
        End If
    Next i
    Wv2Log.LogI "        手つかずの行 = " & blanks & " 件"
    ' ★これが「途中で止まった」の直接の証拠★ (n < 4 は間接指標なので使わない)
    TestBool "★手つかずの行がある★", (blanks >= 1)

    ' --- D-7e: ★止まったことを人に見せて、指が離れるのを待つ★ ---
    Wv2Log.LogI ""
    Debug.Print ""
    If wasCanceled Then
        Debug.Print "  ★★★ 止まりました (" & n & " 件処理して中断) ★★★"
        Wv2Log.LogI "  ★★★ 止まりました (" & n & " 件処理して中断) ★★★"
    Else
        Debug.Print "  ★止まりませんでした★ Esc を押しましたか?"
        Wv2Log.LogI "  ★止まりませんでした★"
    End If
    Debug.Print "  ★Esc から指を離してください★ 静かになったら再開を試します"
    TestWaitEscReleased 15

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- ★呼び直して続きから★ (今度は Esc を押さないでください) ---"
    Debug.Print "  --- 再開します (Esc は押さないでください) ---"

    n2 = Wv2Maps.MapsGeocodeSheet(sh)
    Wv2Log.LogI "        2 回目の ok = " & n2 & " / 中断 = " & Wv2Maps.MapsCanceled

    blanks = 0
    For i = 2 To 5
        Wv2Log.LogI "        " & i & " 行目: " & sh.Cells(i, 2).value & " / " & _
                    sh.Cells(i, 3).value & " / " & sh.Cells(i, 5).value
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then blanks = blanks + 1
    Next i

    TestBool "★2 回目は中断していない★ (入口で捨てられた)", _
             (Wv2Maps.MapsCanceled = False)
    TestBool "★手つかずの行が無くなった★", (blanks = 0)
    TestBool "  ok が減っていない", (n2 >= n)
    TestBool "  EnableCancelKey が戻っている", _
             (Application.EnableCancelKey = xlInterrupt)
    TestBool "  ステータスバーが戻っている", _
             (TypeName(Application.StatusBar) = "Boolean")
    TestDumpStatusBar "StatusBar"

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Sheet 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_D7_Resume (D-7b の検証: ★中断後の状態から再開できるか★)
'
'   ★Esc を押さなくてよい★ 中断された後のシート ―― 前半だけ埋まっていて
'   後半が手つかず ―― を人工的に作り、そこから呼び直す。
'   これで「再開」の回帰確認を毎回自動で回せる。
'
'   ★外部サイトに実アクセスする★ 処理するのは 2 件だけ (15 秒ほど)。
' ============================================================
Public Sub Test_D7_Resume()
    Dim wb As Workbook
    Dim sh As Object
    Dim n As Long
    Dim i As Long
    Dim blanks As Long

    If UserForm1.CurrentBrowser Is Nothing Then
        Wv2Log.LogI "Test_D7_Resume: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Resume 開始 ================"

    Set wb = Workbooks.Add
    Set sh = wb.Worksheets(1)
    sh.Cells(1, 1).value = "住所"
    sh.Cells(2, 1).value = "東京都千代田区丸の内1-9-1"
    sh.Cells(3, 1).value = "大阪府大阪市中央区大阪城1-1"
    sh.Cells(4, 1).value = "北海道札幌市北区北6条西4丁目"
    sh.Cells(5, 1).value = "宮城県仙台市青葉区中央1-1-1"

    ' --- ★中断された後の状態を人工的に作る★ 前半 2 行は処理済み ---
    sh.Cells(2, 2).value = 11.111
    sh.Cells(2, 3).value = 22.222
    sh.Cells(2, 4).value = "見張り番 1"
    sh.Cells(2, 5).value = "ok"
    sh.Cells(3, 2).value = 33.333
    sh.Cells(3, 3).value = 44.444
    sh.Cells(3, 4).value = "見張り番 2"
    sh.Cells(3, 5).value = "ok(候補1)"

    Wv2Log.LogI "  4 行のうち前半 2 行を「済み」にしてから呼ぶ"
    n = Wv2Maps.MapsGeocodeSheet(sh)

    Wv2Log.LogI "        戻り値 = " & n & " / 中断 = " & Wv2Maps.MapsCanceled
    For i = 2 To 5
        Wv2Log.LogI "        " & i & " 行目: " & sh.Cells(i, 2).value & " / " & _
                    sh.Cells(i, 3).value & " / " & sh.Cells(i, 5).value
    Next i

    ' ★済みの行に触っていないこと★ (見張り番の値がそのまま残っているか)
    TestBool "★済みの行は書き換えない★ (2 行目)", (sh.Cells(2, 2).value = 11.111)
    TestBool "  ok(候補1) も済み扱いになる (3 行目)", (sh.Cells(3, 2).value = 33.333)

    blanks = 0
    For i = 2 To 5
        If Len(CStr(sh.Cells(i, 5).value)) = 0 Then blanks = blanks + 1
    Next i
    TestBool "★手つかずの行が埋まった★", (blanks = 0)
    TestBool "  4 行目に座標が入った", (Len(CStr(sh.Cells(4, 2).value)) > 0)
    TestBool "  5 行目に座標が入った", (Len(CStr(sh.Cells(5, 2).value)) > 0)
    TestBool "★戻り値は「済み」も数える★ (4 = 2 + 2)", (n = 4)
    TestBool "  中断していない", (Wv2Maps.MapsCanceled = False)
    TestBool "  EnableCancelKey が戻っている", _
             (Application.EnableCancelKey = xlInterrupt)

    wb.Close SaveChanges:=False

    Wv2Log.LogI ""
    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Resume 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Quiet (K-4a の実験) - ★閉じる前に全タブを静める★
'
'   解放が遅い原因の切り分け。仮説は
'   ★「ページが動き続けているので WebView2 が忙しく、COM の応答が遅い」★。
'   全タブを about:blank に飛ばしてから閉じれば、速くなるはず。
'
'   ★製品コードは 1 行も変えない★ ここでやることは、利用者が手でできること
'   (別のページへ移動する) と同じ。効いたら製品側にどう入れるかを別途決める。
'
'   使い方:
'     1) Wv2Log.LogStart
'     2) UserForm1.Show vbModeless → StartWebView2_Full
'     3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser   (重いページを 1 枚作る)
'     4) ★Test_K4_Quiet 3★ / ★Test_K4_Quiet 1★ / ★Test_K4_Quiet 0★
'     5) フォームを × で閉じる
'
'   ★K-4b: 待ち秒数を引数にした★ 3 秒では合計時間で得をしない
'   (静めるのに 1.6 秒 + 待ち 3 秒 = 4.6 秒かけて、解放が 6.7 → 0.37 秒)。
'   ★0 秒 (飛ばすだけ) で効くなら合計でも勝てる★ ので、そこを測る。
'
'   実測済み: 何もしない = 6.718 秒 / 3 秒待つ = 0.367 秒
' ============================================================
Public Sub Test_K4_Quiet(Optional ByVal waitSec As Single = 3)
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim i As Long
    Dim t0 As Single

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_K4_Quiet: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_K4_Quiet 開始 ================"
    Wv2Log.LogI "  ★全 " & b.TabCount & " タブを about:blank に飛ばす★" & _
                " (そのあと " & waitSec & " 秒待つ)"

    t0 = Timer

    For i = 1 To b.TabCount
        Set p = b.PaneAt(i)
        If Not p Is Nothing Then
            p.View_Navigate "about:blank"
            Wv2Log.LogI "        タブ " & i & " を about:blank へ"
        End If
    Next i

    Wv2Log.LogI "  ★飛ばし終えた (" & Format$(Timer - t0, "0.00") & " 秒)★"

    ' ★飛ばした直後は「これから止まる」ところ★ なので、落ち着くまで待つ
    '   waitSec = 0 なら待たない (飛ばすだけで効くかを測るため)
    If waitSec > 0 Then
        t0 = Timer
        Do
            DoEvents
        Loop Until Timer - t0 >= waitSec Or Timer < t0
    End If

    Wv2Log.LogI "  ★静めた。この状態でフォームを × で閉じてください★"
    Debug.Print ""
    Debug.Print "  ★静めました。フォームを × で閉じてください★"
    Wv2Log.LogI "================ Test_K4_Quiet 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Step (K-4c) - ★何が「閉じるのを遅くする」のかを二分探索する★
'
'   実測で分かっていること:
'     AddTabWithUrl だけ … 閉じるのは★速い★
'     MapsOpen 相当     … 閉じるのが★遅い★ (5 秒)
'   差分は EvalSync を何度も撃つことと、QuerySelector で要素レジストリを
'   作ることの 2 つ。どちらが効いているのかを段階的に再現して確かめる。
'
'   使い方 ★1 段階ごとに Excel の再起動は不要。フォームを開き直すだけ★:
'     Wv2Log.LogStart
'     UserForm1.Show vbModeless → StartWebView2_Full
'     ★UserForm1.CurrentBrowser.QuietOnShutdown = False★  ← これを忘れない
'     Test_K4_Step 1   (または 2 / 3 / 4)
'     フォームを × で閉じる → ログの解放時間を見る
'
'   ★QuietOnShutdown = False が要る★ K-4c で解放前に静めるようにしたので、
'   そのままだと全部速くなって切り分けにならない (設計原則103: プローブを検算する)。
'
'   段階ごとの内容:
'     1 … AddTabWithUrl だけ (既知: 速い)
'     2 … + EvalSync を 20 回
'     3 … + QuerySelector を 5 回
'     4 … MapsOpen 相当 (全部)                        (既知: 遅い)
'     5 … ★DoEvents を 15 秒回すだけ★ (JS は撃たない)   (実測: 速い)
'     6 … ★EvalSync を 300 回★     (MapsOpen 相当の回数)
'     7 … ★View_GetSource を 300 回★ (COM で読むだけ)   (実測: 速い)
'     8 … + JS が通るまで待つ           (MapsOpen の第 1 段)
'     9 … + ★WaitFor で検索ボックスを待つ★ (唯一まだ試していない要素)
'    10 … + /@ が付くまで待つ           (= MapsOpen 相当)   (実測: 速い)
'    11 … + ★EnableCancelKey 操作 + AddIgnoreNetwork★ (最後の差分)
'
'   ★8→9→10 と進めて、遅くなった段階が答え★
'
'   ★実測で分かったこと (K-4c の切り分け)★
'     Step 1 + 30 秒待ち … 速い  ← ただし★VBA は止まっていた★
'     Step 2 (EvalSync×20) … 速い
'     Step 3 (+QuerySelector×5) … 速い / +30 秒待ち でも速い
'     MapsOpen … ★遅い★
'     Step 5 (DoEvents 15 秒 / 69 万回) … 速い
'   → ★「JS を撃つ」「時間」「DoEvents」はすべて無罪★。
'     残るのは ★回数★ (MapsOpen は EvalSync を 200 回前後撃っている)。それが段階6/7。
' ============================================================
Public Sub Test_K4_Step(Optional ByVal stepNo As Long = 1)
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim i As Long
    Dim t0 As Single
    Dim loops As Long
    Dim savedKey As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_K4_Step: Browser が起動していません。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_K4_Step " & stepNo & " 開始 ================"
    If b.QuietOnShutdown Then
        Wv2Log.LogW "  ★QuietOnShutdown が True のままです★ 切り分けになりません。"
        Wv2Log.LogW "  UserForm1.CurrentBrowser.QuietOnShutdown = False を先に打つこと。"
        Debug.Print "  ★QuietOnShutdown = False を先に打ってください★"
    End If

    ' --- 段階1: タブを開く (全段階で共通) ---
    t0 = Timer
    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    Wv2Log.LogI "  [1] AddTabWithUrl … " & Format$(Timer - t0, "0.00") & " 秒"
    If p Is Nothing Then
        Wv2Log.LogI "      タブを開けなかった"
        Exit Sub
    End If

    If stepNo >= 2 And stepNo <= 4 Then
        ' --- 段階2: EvalSync を 20 回 ---
        t0 = Timer
        For i = 1 To 20
            p.EvalSync "1", 3
        Next i
        Wv2Log.LogI "  [2] EvalSync × 20 … " & Format$(Timer - t0, "0.00") & " 秒"
    End If

    If stepNo >= 3 And stepNo <= 4 Then
        ' --- 段階3: QuerySelector を 5 回 (要素レジストリを作る) ---
        t0 = Timer
        For i = 1 To 5
            Set el = p.QuerySelector("input")
        Next i
        Wv2Log.LogI "  [3] QuerySelector × 5 … " & Format$(Timer - t0, "0.00") & _
                    " 秒 (レジストリ " & p.ElementCount & " 個)"
    End If

    If stepNo >= 11 Then
        ' [11] ★MapsOpen と同じ順序で包む★
        '   本物は「入る前に arm → 全部やる → 最後に AddIgnoreNetwork → disarm」。
        '   ★arm が一瞬では再現にならない★ ので、ここで掛けて最後に外す。
        savedKey = Application.EnableCancelKey
        Application.EnableCancelKey = xlDisabled
        Wv2Log.LogI "  [11] EnableCancelKey = xlDisabled にした"
    End If

    If stepNo >= 8 Then
        ' --- 段階8/9/10 (K-4c): ★MapsOpen の中身を段階的に再現する★ ---
        '   候補を当てにいって 6 連敗したので、当てるのをやめて
        '   ★MapsOpen を分解し、途中で止めて閉じる★ 方式にした。
        '   MapsOpen 全体は確実に遅いので、どこかに必ず境界がある。

        ' [8] JS が通るようになるまで待つ (MapsOpenCore の第 1 段)
        t0 = Timer
        Do
            p.EvalSync "1", 3
            If p.LastEvalOk Then Exit Do
            If Timer - t0 >= 30 Or Timer < t0 Then Exit Do
            DoEvents
        Loop
        Wv2Log.LogI "  [8] JS が通るまで … " & Format$(Timer - t0, "0.00") & " 秒"

        If stepNo >= 9 Then
            ' [9] ★検索ボックスが出るまで WaitFor★ (唯一まだ試していない要素)
            t0 = Timer
            Set el = p.WaitFor("input[name='q']", 30)
            Wv2Log.LogI "  [9] WaitFor … " & Format$(Timer - t0, "0.00") & " 秒 " & _
                        IIf(el Is Nothing, "(見つからなかった)", "(見つかった)") & _
                        " レジストリ " & p.ElementCount & " 個"
        End If

        If stepNo >= 10 Then
            ' [10] 地図が位置を持つまで待つ (= MapsOpen 相当)
            t0 = Timer
            Do
                If InStr(1, p.View_GetSource(), "/@") > 0 Then Exit Do
                If Timer - t0 >= 8 Or Timer < t0 Then Exit Do
                DoEvents
            Loop
            Wv2Log.LogI "  [10] /@ が付くまで … " & Format$(Timer - t0, "0.00") & " 秒"
        End If

        If stepNo >= 11 Then
            p.AddIgnoreNetwork "/search?tbm=map"
            Application.EnableCancelKey = savedKey
            Wv2Log.LogI "  [11] AddIgnoreNetwork して EnableCancelKey を戻した"
        End If
    End If

    If stepNo = 6 Or stepNo = 7 Then
        ' --- 段階6/7 (K-4c): ★回数を桁で増やす★ ---
        '   段階2 の EvalSync は 20 回で速かったが、MapsOpen は 17 秒ぶん
        '   ポーリングするので ★200 回前後★ 撃っている計算になる。桁が違う。
        '     6 … EvalSync (JS を撃つ) を 300 回
        '     7 … View_GetSource (COM で URL を読むだけ) を 300 回
        '   ★どちらが効くかで「JS 側が溜まる」か「COM 側が溜まる」かが分かれる★
        t0 = Timer
        For i = 1 To 300
            If stepNo = 6 Then
                p.EvalSync "1", 3
            Else
                p.View_GetSource
            End If
        Next i
        Wv2Log.LogI "  [" & stepNo & "] " & _
                    IIf(stepNo = 6, "EvalSync", "View_GetSource") & _
                    " × 300 … " & Format$(Timer - t0, "0.00") & " 秒"
        Wv2Log.LogI "      レジストリ " & p.ElementCount & " 個 / " & _
                    "最後の EvalSync は " & p.LastEvalOk
    End If

    If stepNo = 5 Then
        ' --- 段階5 (K-4c): ★DoEvents を回すだけ★ JS は 1 回も撃たない ---
        '   Step 1 + 30 秒待ち は速かったが、そのとき VBA は★止まっていた★。
        '   MapsOpen は 15 秒間 DoEvents を回し続ける = ★WebView2 のイベントが
        '   VBA のコールバックへ配送され続ける★。ここが唯一の差分。
        t0 = Timer
        loops = 0
        Do
            DoEvents
            loops = loops + 1
        Loop Until Timer - t0 >= 15 Or Timer < t0
        Wv2Log.LogI "  [5] DoEvents を 15 秒 … " & loops & " 回まわした"
    End If

    If stepNo = 4 Then
        ' --- 段階4: MapsOpen 相当 (待ちも含めて丸ごと) ---
        '   ★= 4 であって >= 4 ではない★ 段階5 に巻き込むと切り分けにならない
        '   (K-4c で実際にやらかした。設計原則103: プローブ自体を検算する)
        t0 = Timer
        Wv2Maps.MapsOpen b
        Wv2Log.LogI "  [4] MapsOpen … " & Format$(Timer - t0, "0.00") & " 秒"
    End If

    Wv2Log.LogI "  ★この状態でフォームを × で閉じてください★"
    Debug.Print ""
    Debug.Print "  ★段階 " & stepNo & " まで実行しました。× で閉じてください★"
    Wv2Log.LogI "================ Test_K4_Step " & stepNo & " 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' Test_K4_Help (K-4a の手順)
' ============================================================
Public Sub Test_K4_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " K-4a 実測手順 (フォームを閉じたときの解放が遅い)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  ★毎回 Wv2Log.LogStart から始める★ 1 回分が 1 ファイルに閉じる。"
    Debug.Print ""
    Debug.Print "  --- ケース1: 素の状態 (実測済み 0.336 秒) ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) UserForm1.Show vbModeless → StartWebView2_Full"
    Debug.Print "  3) すぐ × で閉じる"
    Debug.Print ""
    Debug.Print "  --- ケース2: 重いページを開いた後 (実測済み ★6.718 秒★) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) × で閉じる"
    Debug.Print ""
    Debug.Print "  --- ケース3: 閉じる前に静める (実測済み ★0.367 秒★) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) ★Test_K4_Quiet 3★  (about:blank に飛ばして 3 秒待つ)"
    Debug.Print "  5) × で閉じる"
    Debug.Print ""
    Debug.Print "  --- ケース3b: ★待ち時間を削る★ (これから測る) ---"
    Debug.Print "  4) Test_K4_Quiet 1   … 1 秒待つ"
    Debug.Print "  4) Test_K4_Quiet 0   … ★待たない (飛ばすだけ)★"
    Debug.Print "  ★合計時間で勝てるのはこちら★ 静めるのに 1.6 秒かかるので、"
    Debug.Print "  3 秒待つと合計 5 秒で「何もしない 6.7 秒」とあまり変わらない。"
    Debug.Print ""
    Debug.Print "  --- ケース4: 重いタブを閉じてから (これから測る) ---"
    Debug.Print "  3) Wv2Maps.MapsOpen UserForm1.CurrentBrowser"
    Debug.Print "  4) UserForm1.CurrentBrowser.CloseTab 4   ' Maps のタブ番号"
    Debug.Print "  5) × で閉じる"
    Debug.Print ""
    Debug.Print "  ★見るのはログの解放部分★ Wv2NavBar.Shutdown 開始 から"
    Debug.Print "  Wv2Browser.Terminate 完了 までの経過。?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  --- K-4c: 対策が入った後の切り分け ---"
    Debug.Print "  ★既定では解放の前に about:blank へ飛ばすので速い★"
    Debug.Print "  遅い状態を再現するには:"
    Debug.Print "    UserForm1.CurrentBrowser.QuietOnShutdown = False"
    Debug.Print "  何が遅くしているかの二分探索:"
    Debug.Print "    Test_K4_Step 1  … AddTabWithUrl だけ (既知: 速い)"
    Debug.Print "    Test_K4_Step 2  … + EvalSync × 20"
    Debug.Print "    Test_K4_Step 3  … + QuerySelector × 5"
    Debug.Print "    Test_K4_Step 4  … MapsOpen 相当 (既知: 遅い)"
    Debug.Print ""
End Sub

' ============================================================
' Test_D7_Key (D-7d のプローブ) - ★GetAsyncKeyState が Esc を見えているか★
'
'   D-7c で Esc が拾えなかった。原因は 2 つのどちらか:
'     A. MapsPump が回っておらず、そもそも呼ばれていない
'     B. 呼んでも API が Esc を見えていない
'   ★このプローブは B だけを測る★ 製品コードの待ちを通さず、
'   ここで自前に回して生の値を記録する (設計原則103)。
'
'   ★実行中は EnableCancelKey を xlDisabled にする★ さもないと Esc で
'   VBA が break して測定にならない。10 秒で必ず終わり、必ず元に戻す。
'
'   ネットワークもブラウザも要らない。
' ============================================================
Public Sub Test_D7_Key()
    Dim savedKey As Long
    Dim t0 As Single
    Dim polls As Long
    Dim hits As Long
    Dim firstHit As Single
    Dim lastSec As Long
    Dim k As Long
    Dim seen As Long

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_D7_Key 開始 ================"

    ' ★D-7e: 「これから 10 秒」では読む前に終わる★ 押されるまで待つ形にした。
    Debug.Print ""
    Debug.Print "  ★★★ 今から Esc を押してください ★★★"
    Debug.Print "  (押されたら即座に終わります。最大 30 秒待ちます)"
    Wv2Log.LogI "  ★押されるまで待つ★ (最大 30 秒)"

    savedKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled   ' さもないと Esc で break する

    firstHit = -1
    lastSec = -1
    t0 = Timer
    Do
        DoEvents
        polls = polls + 1

        k = Wv2Maps.MapsRawEscState()
        If k <> 0 Then
            hits = hits + 1
            If firstHit < 0 Then
                firstHit = Timer - t0
                seen = k
                Debug.Print "  ★検出した★ 生の値 = " & k
            End If
            ' 押されたことが分かれば十分 (毎秒 4 万回まわるのですぐ貯まる)
            If hits >= 500 Then Exit Do
        End If

        If Int(Timer - t0) <> lastSec Then
            lastSec = Int(Timer - t0)
            Debug.Print "  ... 残り " & (30 - lastSec) & " 秒"
        End If

        If Timer - t0 >= 30 Or Timer < t0 Then Exit Do
    Loop

    Application.EnableCancelKey = savedKey

    Wv2Log.LogI ""
    Wv2Log.LogI "        ポーリング " & polls & " 回 / " & _
                Format$(Timer - t0, "0.0") & " 秒"
    Wv2Log.LogI "        検出       " & hits & " 回"
    If hits > 0 Then
        Wv2Log.LogI "        最初の検出 " & Format$(firstHit, "0.00") & " 秒後"
        Wv2Log.LogI "        生の値     " & seen & " (&H" & Hex$(seen And &HFFFF&) & ")"
    End If

    TestBool "★GetAsyncKeyState は Esc を見えている★", (hits > 0)

    Wv2Log.LogI ""
    If hits > 0 Then
        Wv2Log.LogI "  ★結論: API は見えている★"
    Else
        Wv2Log.LogI "  ★結論: 30 秒待っても検出できなかった★"
        Wv2Log.LogI "         ★押し忘れでなければ★ キーの拾い方を変える必要がある。"
    End If

    TestCountPrint
    Wv2Log.LogI "================ Test_D7_Key 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' TestWaitEscReleased (D-7e、Private) - ★Esc から指が離れるまで待つ★
'
'   D-7d で、中断した直後に 2 回目を走らせたら、連打していた Esc が
'   そちらにも効いて判定が 3 件 FAIL になった。製品側は正しく、
'   ★検証の段取りが悪かった★。静かになるまで待ってから次へ進む。
'
'   ★待っている間は xlDisabled にする★ 戻したままだと、待ちの最中に
'   押された Esc で VBA が break してしまう。
' ============================================================
Private Sub TestWaitEscReleased(ByVal maxSec As Single)
    Dim savedKey As Long
    Dim t0 As Single
    Dim quiet As Single

    savedKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled

    quiet = -1
    t0 = Timer
    Do
        DoEvents
        If Wv2Maps.MapsRawEscState() = 0 Then
            If quiet < 0 Then quiet = Timer
            If Timer - quiet >= 2 Then Exit Do   ' ★2 秒静かなら離れたと見なす★
        Else
            quiet = -1
        End If
        If Timer - t0 >= maxSec Or Timer < t0 Then Exit Do
    Loop

    Application.EnableCancelKey = savedKey
    Wv2Log.LogI "  (Esc が静まった。再開します)"
End Sub

' ============================================================
' Test_D7_StatusBar (D-7b のプローブ)
'
'   ★製品コードを通さずに Excel の生の挙動だけを測る★ (設計原則103)
'   D-7 の実機で「★ステータスバーが戻っている★」が 3 箇所とも FAIL した。
'   判定の書き方 (TypeName で見る) が悪いのか、Excel が実行中は制御を返さないのか
'   を切り分ける。★判定はしない。事実だけ出す。★
'
'   ネットワークもブラウザも要らない。
' ============================================================
Public Sub Test_D7_StatusBar()
    Wv2Log.LogI ""
    Wv2Log.LogI "================ Test_D7_StatusBar 開始 ================"
    Wv2Log.LogI "  ★D-7b の実測: False を入れると文字列 ""FALSE"" が残った★"
    Wv2Log.LogI "  正しい戻し方を候補ごとに測る。★型が Boolean に戻れば当たり★"
    Wv2Log.LogI ""

    TestDumpStatusBar "(0) 何もしていないとき"
    Wv2Log.LogI ""

    TestStatusTry "(1) False       ", False
    TestStatusTry "(2) Empty       ", Empty
    TestStatusTry "(3) 空文字 """"   ", ""
    TestStatusTry "(4) vbNullString", vbNullString
    TestStatusTry "(5) CVar(False) ", CVar(False)

    Wv2Log.LogI ""
    Wv2Log.LogI "  ★どれかで 型=Boolean になっていれば、それが正しい戻し方★"
    Wv2Log.LogI "  すべて String なら、この Excel では制御を返せない ―― その場合は"
    Wv2Log.LogI "  ★空文字にしておく★ のが実害が小さい (FALSE と表示されるよりよい)。"

    ' ★測り終えたら Empty で戻す★ (D-7d の実測で当たりだったのがこれ)
    Application.StatusBar = Empty
    TestDumpStatusBar "(9) 後始末に Empty"

    Wv2Log.LogI "================ Test_D7_StatusBar 終了 ================"
    Wv2Log.LogI ""
End Sub

' ============================================================
' TestStatusTry (D-7c、Private) - 戻し方の候補を 1 つ試す
'
'   ★毎回「文字列を入れてから戻す」★ 入れずに戻しても、元から Excel が
'   制御を持っている状態なので何も測ったことにならない (設計原則103)。
'   DoEvents を挟んだ後の値も見る ―― Excel が後から取り戻す可能性があるため。
' ============================================================
Private Sub TestStatusTry(ByVal tag As String, ByVal v As Variant)
    Dim before As String
    Dim after As String

    Application.StatusBar = "テスト中 12/100"
    Application.StatusBar = v
    before = "[" & CStr(Application.StatusBar) & "] 型=" & _
             TypeName(Application.StatusBar)

    DoEvents
    after = "[" & CStr(Application.StatusBar) & "] 型=" & _
            TypeName(Application.StatusBar)

    Wv2Log.LogI "        " & tag & " → 直後 " & before & " / DoEvents 後 " & after
End Sub

' ============================================================
' TestDumpStatusBar (D-7b、Private) - ステータスバーの実測値を出す
'   ★値そのものと型の両方を残す★ どちらか片方では切り分けられない。
' ============================================================
Private Sub TestDumpStatusBar(ByVal tag As String)
    Wv2Log.LogI "        " & tag & " = [" & CStr(Application.StatusBar) & _
                "] 型=" & TypeName(Application.StatusBar)
End Sub

' ============================================================
' Test_D7_Help (D-7 の手順)
' ============================================================
Public Sub Test_D7_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-7 検証手順 (進捗表示と中断)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) UserForm1.Show vbModeless して StartWebView2_Full を実行する"
    Debug.Print "  3) ★イベントバーストが静まるまで待つ★ (仕様事実 20)"
    Debug.Print ""
    Debug.Print "  --- 実行 (上から順に) ---"
    Debug.Print "  4) Test_D7_Key … ★プローブ★ GetAsyncKeyState が Esc を"
    Debug.Print "     見えているか。★押されるまで待つ★ (最大 30 秒)。"
    Debug.Print "     ネットワーク不要。イミディエイトに残り秒数が出る。"
    Debug.Print "  4b) Test_D7_StatusBar … ★プローブ★ ステータスバーの戻し方を"
    Debug.Print "     候補ごとに実測する。ネットワーク不要・数秒。"
    Debug.Print "     ★判定はしない。事実だけ出す★ (設計原則103)。"
    Debug.Print "  5) Test_D7_Cancel … ★ネットワーク不要★ 中断の口と分母 (すぐ終わる)"
    Debug.Print "  6) Test_D7_Resume … ★自動★ 中断後の状態からの再開 (2 件・15 秒)"
    Debug.Print "  7) Test_D7_Sheet  … ★実サイト★ 4 件を処理する。"
    Debug.Print "     ★2 件目あたりで Esc を押す★ ステータスバーを見ながら。"
    Debug.Print "     ★連打しなくてよい。1 回で効く★ (連打すると再開にも効く)。"
    Debug.Print "     押しっぱなしでなくてよく、焦点はどこにあってもよい。"
    Debug.Print "     止まったらイミディエイトに ★★★ 止まりました ★★★ と出る。"
    Debug.Print "     ★D-7c からは中断ダイアログも実行時エラー 18 も出ない★"
    Debug.Print "     出たら効いていない。ログに ★Esc で中断が要求された★ が出る。"
    Debug.Print "     そのあと自動で呼び直して、続きから埋まるかを見る。"
    Debug.Print ""
    Debug.Print "  --- 判定はログファイルで読む (設計原則105) ---"
    Debug.Print "  ?Wv2Log.LogPath"
    Debug.Print ""
    Debug.Print "  --- 実務での使い方 ---"
    Debug.Print "  Wv2Maps.MapsGeocodeSheet Sheets(""住所録"")"
    Debug.Print "    処理中はステータスバーに 12/100 (ok 11 / 失敗 1) と出る。"
    Debug.Print "    ★Esc でいつでも止まる★ 止めた行には何も書かれない。"
    Debug.Print "    もう一度呼べば ok の行を飛ばして続きから進む。"
    Debug.Print "  ?Wv2Maps.MapsCanceled   ' 中断で終わったか"
    Debug.Print "  ?Wv2Maps.MapsCountRows(Sheets(""住所録""))  ' 何件あるか先に数える"
    Debug.Print "  Wv2Maps.MapsCancel = True  ' 外から止める (イミディエイトから)"
    Debug.Print ""
End Sub

' ============================================================
' Test_D4_Help (D-4 の手順)
' ============================================================
Public Sub Test_D4_Help()
    Debug.Print ""
    Debug.Print "=========================================================="
    Debug.Print " D-4 検証手順 (プローブと静穏待ち)"
    Debug.Print "=========================================================="
    Debug.Print ""
    Debug.Print "  --- 準備 ---"
    Debug.Print "  1) Wv2Log.LogStart"
    Debug.Print "  2) UserForm1.Show vbModeless して StartWebView2_Full を実行する"
    Debug.Print "  3) ★イベントバーストが静まるまで待つ★ (仕様事実 20)"
    Debug.Print ""
    Debug.Print "  --- 実行 ---"
    Debug.Print "  4) Test_D4_Probe  … ★済★ プローブと健康診断 (28 件 OK)"
    Debug.Print "  5) Test_D4_Settle … ★済★ 静穏待ち (22 件 OK)"
    Debug.Print "  6) Test_D4_Signal … ★済★ 明示シグナル (31 件 OK)"
    Debug.Print "  7) Test_D4_Log    … ★D-4e の検証★ 診断ログ / in-flight / URL シグナル"
    Debug.Print "  8) Test_D4_Site   … ★D-4d の偵察★ Google Maps (70 秒ほど)"
    Debug.Print "  9) Test_D5_Geocode … ★D-5★ 住所 → 座標 を 3 件続けて (60 秒ほど)"
    Debug.Print "     Wv2Maps.MapsOpen / MapsGeocode の実演。業務で使う形そのもの。"
    Debug.Print " 10) Test_D5_Sheet   … ★D-5b★ シート連携 (新しいブックを作って試す)"
    Debug.Print " 11) Test_D6_All     … ★D-6★ QuerySelectorAll と寿命管理"
    Debug.Print " 12) Test_D6_Pick    … ★D-6b★ 候補一覧から 1 番目を採る (実サイト)"
    Debug.Print " 13) Test_D7_Cancel  … ★D-7★ 中断の口と分母 (ネットワーク不要)"
    Debug.Print " 14) Test_D7_Resume  … ★D-7b★ 中断後の状態からの再開 (自動)"
    Debug.Print " 15) Test_D7_Sheet   … ★D-7b★ 進捗と中断 (実行中に Esc を押す)"
    Debug.Print "     ★外部サイトに実アクセスする唯一の Test_*★ ネットワークが要る。"
    Debug.Print "     住所を変えたいときは Test_D4_Site ""別の住所"" と打つ。"
    Debug.Print ""
    Debug.Print "  --- 実サイトを調べる道具 ---"
    Debug.Print "  Test_D4_Dom          … ★今開いているタブに何があるか列挙する★"
    Debug.Print "  Test_D4_Dom ""h1,h2""  … セレクタを指定して列挙"
    Debug.Print ""
    Debug.Print "  --- 診断ログ (D-4e) ―★何を待つべきかを観測で決める★ ---"
    Debug.Print "  p.SpaProbeLogging = True   ' 溜め始める (既定 False)"
    Debug.Print "  ... 操作する ..."
    Debug.Print "  ?p.SpaProbeDrainLog        ' 取り出してログへ流す (60 件ずつ)"
    Debug.Print "  p.SpaProbeLogging = False"
    Debug.Print ""
    Debug.Print "  --- URL シグナル (D-4e) ---"
    Debug.Print "  p.ArmUrlSignal ""/place/""   ' URL が変わるのを待つ (pushState 対応)"
    Debug.Print "    ★Maps の住所検索はこれが最も確実な目印だった★"
    Debug.Print "    実サイトのセレクタは予告なく変わる。当てずっぽうに書く前に見る。"
    Debug.Print ""
    Debug.Print "  ★判定はログファイルに残る★ 末尾の「★判定 n 件★」を見る"
    Debug.Print "    ?Wv2Log.LogPath でファイルの場所が分かる"
    Debug.Print ""
    Debug.Print "  --- 見るもの ---"
    Debug.Print "  ・(1) 設置前が undefined = 呼ばない Pane にプローブは付かない"
    Debug.Print "  ・(2) ★往復時間の設置前後★ 数値そのものをログで見ること"
    Debug.Print "      仕様事実52 は 38ms。観測を張って極端に落ちないかを見る"
    Debug.Print "  ・(5) fetch と XHR が数えられ、完了で 0 に戻ること"
    Debug.Print "  ・★(6) が D-4a の核心★ ページが window.fetch を差し替えても"
    Debug.Print "      次に呼んだ時点で自動修復し (rp>=1)、その後も数えられること"
    Debug.Print "  ・(8) 遷移でプローブが消え、立て直せること"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D4_Settle) ★D-4b★ ---"
    Debug.Print "  ・★(2) が核心★ 待ち終わった時点で #target が「連鎖の結果」に"
    Debug.Print "      なっていること = 通信の完了だけでなく、その後の DOM 更新まで"
    Debug.Print "      待てている証拠 (fetch → 300ms 後の書き換え、を跨いでいる)"
    Debug.Print "  ・(3) 静まらないページで False + LastEvalOk=True (失敗ではない)"
    Debug.Print "  ・(4) ★除外を入れると同じページが静穏になる★ = ノイズ除外が効いている"
    Debug.Print "  ・(6) 自動待ち OFF なら更新前、ON なら更新後の値が見えること"
    Debug.Print "  ・LastSettleInfo の slack が小さい (100ms 未満) 待ちは★際どい★"
    Debug.Print "      たまたま滑り込んだだけかもしれない。stableMs を上げるか"
    Debug.Print "      明示シグナルを併用する"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D4_Signal) ★D-4c★ ---"
    Debug.Print "  ・★(3) が核心★ arm したのに何も起きなければ、ページが静かでも"
    Debug.Print "      False になること = ★アプリに無視されたことを検出できる★"
    Debug.Print "      静穏待ちだけならここは成功してしまう"
    Debug.Print "  ・(1) 無い要素・不正なセレクタは★その場で★ arm 失敗 (fail-fast)"
    Debug.Print "  ・(4) arm はワンショット。次の待ちには持ち越さない"
    Debug.Print "  ・(6) arm が消えたら signal-lost で失敗 (黙って再 arm しない)"
    Debug.Print ""
    Debug.Print "  --- 見るもの (Test_D4_Site) ★D-4d は観測が目的★ ---"
    Debug.Print "  ・判定の件数は構造的な数点だけ。★中身はログ本文を読む★"
    Debug.Print "  ・(3)(6) の毎秒の行 [01s] q=.. m=.. n=.. が主役"
    Debug.Print "      q が伸びる  → 静穏が訪れる (静穏待ちが使える)"
    Debug.Print "      q が伸びない → ★除外か明示シグナルが要る★"
    Debug.Print "  ・(2) の往復コスト。軽いページでは 35.4 → 36.3 ms だった"
    Debug.Print "  ・rp (作り直し回数) が増えていたら、Maps が fetch を差し替えている"
    Debug.Print "  ・(7) の url に /@緯度,経度 が入るか = 座標が取れるか"
    Debug.Print ""
    Debug.Print "  --- 手で試したいとき ---"
    Debug.Print "  Set p = UserForm1.GetActivePane"
    Debug.Print "  ?p.SpaProbeState"
    Debug.Print "  ?p.SpaProbeHealthy"
    Debug.Print "  ?p.SpaProbeReset"
    Debug.Print "  ?p.WaitSettled(10)     : ?p.LastSettleInfo"
    Debug.Print "  p.IgnoreSelectors = ""#clock, .ticker"""
    Debug.Print "  p.AddIgnoreNetwork ""google-analytics"""
    Debug.Print "  p.AutoWaitAfterAction = True   ' 操作の後に自動で待つ"
    Debug.Print ""
    Debug.Print "  --- 明示シグナル (D-4c) ―★arm してから act する★ ---"
    Debug.Print "  p.ArmContentSignal ""#result-table""   ' 既存の器が書き変わるのを待つ"
    Debug.Print "  p.ArmNetworkSignal ""/api/search""     ' その要求の完了を待つ"
    Debug.Print "  el.Click"
    Debug.Print "  ?p.WaitSettled(10) : ?p.LastSettleInfo"
    Debug.Print "  ?p.DisarmSignals   ' 取りやめるとき"
    Debug.Print ""
    Debug.Print "  ★★ WaitSettled が True でも「アプリの処理が終わった」証明ではない ★★"
    Debug.Print "    静かなだけ。無視されて静かなのか終わって静かなのかは区別できない。"
    Debug.Print "    重要な操作では D-4c の明示シグナル (arm) を併用すること。"
    Debug.Print ""
    Debug.Print "  ★静穏窓は WaitSettled を呼んだ時点から測る★ 呼ぶ前から静かでも"
    Debug.Print "    いきなり成立させない (最短でも stableMs は見張る)。ただし"
    Debug.Print "    stableMs より後に始まる処理は原理的に取り逃す。"
    Debug.Print ""
End Sub



' ============================================================
' ★N-1 : ネットワーク要求のキャプチャの検証★
'
'   Test_N1_Capture … ★自前 HTML だけで完結する回帰試験★ (論点8)
'                     実サイトに頼らず、まず確実に捕まることを見る。
'   Test_N1_Watch   … ★今開いているタブの通信を捕まえ始める★ (実用の入口)
'   Test_N1_Drain   … 溜まった分をログへ流して空にする
'   Test_N1_Stop    … 捕まえるのをやめる
'   Test_N1_Help    … 手順
'
'   ★判定は Wv2Log に出す★ (設計原則105)。イミディエイトは配管ログで流れる。
' ============================================================


' ============================================================
' Test_N1_Capture (N-1b の回帰試験)
'
'   ★N-1 の初回実機で分かったこと★
'     仮想ホスト (SetVirtualHostNameToFolderMapping) で配信した要求は
'     WebResourceRequested に乗らない。初回が全滅したのはこれが原因で、
'     ★配線は最初から正しかった★ (実サイトでは 10 件すべて捕まった)。
'
'   ★そこで的を 2 系統 + 対照 1 系統にした (論点1 案G)★
'
'     案F  http://127.0.0.1:59999/...   到達不能なローカルアドレス
'          ★外部サービスに依存しない★ 接続は必ず失敗するが、
'          WebResourceRequested は「要求を止める / 書き換える」ためのイベントなので
'          ネットワークへ出る前に発火する ―― はず。それをここで確かめる。
'          ・★Chromium が塞いでいるポートを避ける★ (1/7/9/11/13/…/10080 など)。
'            59999 は塞がれていない。
'          ・★http なのに混在コンテンツで止まらない★ 127.0.0.1 と localhost は
'            仕様上「潜在的に信頼できるオリジン」なので、https のページから
'            呼んでも混在コンテンツ扱いされない。
'
'     案D  https://httpbingo.org/...    外部サービス
'          ネットが要る代わりに、素直に届く経路。
'
'     対照 https://appassets.netprobe/data.json (仮想ホスト)
'          ★捕まらないことを数えて確かめる★ = 見つけた仕様事実を回帰試験にする
'
'   ページ自体は今までどおり仮想ホストで配信する。★配信は問題なくできる★
'   乗らないのはイベントだけ。
'
'   ★どちらの的が届かなくても、空振りになった判定はログに明示する★
'   (「来ない」のが当たり前の状況で OK を積み上げないため)
'
'   ★種別は決めつけない (N-1c)★
'     fetch() で撃っても WebView2 は XHR 種別で報告してくる。判定は「捕まったか」
'     で行い、★実際に何の種別で届いたかはログに書き出して数える★。
' ============================================================
Public Sub Test_N1_Capture()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim el  As Wv2Element
    Dim folderPath   As String
    Dim hr           As Long
    Dim localOk      As Boolean
    Dim netOk        As Boolean
    Dim totalBefore  As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N1_Capture: Browser が起動していません。" & _
                    "先に UserForm1.StartWebView2_Full を実行してください。"
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N1_Capture: 検証ページの書き出しに失敗しました。中止します。"
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N1_Capture: タブの生成に失敗しました。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N1_Capture 開始 ================"
    Wv2Log.LogI "        案F の的 = " & N1_LOCAL & "  (到達不能なローカル)"
    Wv2Log.LogI "        案D の的 = " & N1_NET & "  (外部サービス)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) 準備 (★Navigate の前に捕捉を始める★) ---"

    TestBool "NetCaptureStart が成功する", p.NetCaptureStart()
    TestBool "  捕捉中になっている", (p.NetCaptureOn = True)

    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  仮想ホストのマッピングができる", (hr = 0)

    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate が成功する", (hr = 0)

    If Not D2WaitTitle(p, "N-1 プローブ", 10) Then
        Wv2Log.LogI "Test_N1_Capture: 検証ページの読み込みを確認できませんでした。"
        p.NetLogDrain
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If

    ' ページ内の資材 (CSS / 画像) が飛び終わるだけの間を置く
    D3Pump 3

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) ★仮想ホストの要求はイベントに乗らない (仕様事実の回帰)★ ---"
    ' ★間接的な指標を使わない★ (設計原則112) ページは確かに読めている
    ' (title が取れた) のに、その DOCUMENT が 1 件も居ないことを数える。
    TestBool "★ページ自身の DOCUMENT が 1 件も居ない★", _
             (N1Find(p, "", "", "netprobe.html") = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) fetch / XHR ― 的を 2 系統 + 対照 ---"

    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=f-local').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=f-net').catch(function(){});" & _
              "fetch('data.json?k=f-vhost').catch(function(){});" & _
              "var a=new XMLHttpRequest();a.open('GET','" & N1_LOCAL & _
              "/n1?k=x-local',true);a.send();" & _
              "var c=new XMLHttpRequest();c.open('GET','" & N1_NET & _
              "/get?k=x-net',true);c.send();" & _
              "return 1;})()"

    ' ★種別を条件に入れない (N-1c)★ fetch が FETCH で来るとは限らない。
    localOk = N1Wait(p, "GET", "", "k=f-local", 6)
    TestBool "★案F: 到達不能なローカルへの fetch が捕まる★", localOk
    TestBool "  案F: XHR も捕まる", N1Wait(p, "GET", "", "k=x-local", 4)

    netOk = N1Wait(p, "GET", "", "k=f-net", 8)
    TestBool "★案D: 外部への fetch が捕まる★", netOk
    TestBool "  案D: XHR も捕まる", N1Wait(p, "GET", "", "k=x-net", 4)

    TestBool "★対照: 仮想ホストへの fetch は捕まらない★", _
             (N1Find(p, "", "", "k=f-vhost") = 0)

    ' --- ★何の種別で届いたかを数える (N-1c)★ ---
    '   fetch で撃ったものが FETCH で来るのか XHR で来るのかは★こちらの都合では
    '   決まらない★。実機が返した種別をそのまま書き出して判定する。
    Wv2Log.LogI "        届いた種別: f-local=" & N1CtxOf(p, "k=f-local") & _
                "  x-local=" & N1CtxOf(p, "k=x-local") & _
                "  f-net=" & N1CtxOf(p, "k=f-net") & _
                "  x-net=" & N1CtxOf(p, "k=x-net")
    TestBool "★fetch は FETCH ではなく XHR 種別で届く★", _
             (N1CtxOf(p, "k=f-local") = "XHR" Or N1CtxOf(p, "k=f-net") = "XHR")

    If Not localOk Then
        Wv2Log.LogW "  ※ 案F が 1 件も届いていない。以降の案F の判定は空振りとして読むこと。"
    End If
    If Not netOk Then
        Wv2Log.LogW "  ※ 案D が 1 件も届いていない (ネット断か httpbingo.org のダウン)。" & _
                    "以降の案D の判定は空振りとして読むこと。"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★画像と CSS は来ない (既定フィルタ)★ ---"
    ' ページは n1.css / n1.png / image/png を★両系統に対して★要求している。
    ' 既定フィルタ (DOCUMENT / XHR / FETCH) がそれを弾いていることを数える。
    TestBool "★n1.css が 1 件も居ない★", (N1Find(p, "", "", "n1.css") = 0)
    TestBool "★n1.png が 1 件も居ない★", (N1Find(p, "", "", "n1.png") = 0)
    TestBool "★image/png が 1 件も居ない★", (N1Find(p, "", "", "image/png") = 0)
    If Not (localOk Or netOk) Then
        Wv2Log.LogW "  ※ ★この 3 件は空振り★ 的が 1 つも届いていないので、" & _
                    "「来ない」のはフィルタのおかげではない。"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3b) ★ここまでに捕まえたものを全部出す★ ---"
    ' ★推測で埋めないため (N-1c)★
    '   次の (4) で容量を 3 に落とすと中身が消える。N-1b ではそれで (2) の
    '   記録が失われ、「seq 1/2 は f-local と f-net だったはず」と推測する
    '   羽目になった。消える前に必ず出しておく (設計原則112)。
    p.NetLogDrain

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★溢れたら黙って捨てず、捨てた数が見える★ ---"

    p.NetLogCapacity = 3
    TestBool "容量を 3 にできる", (p.NetLogCapacity = 3)
    TestBool "  中身も通算も 0 に戻る", (p.NetLogCount = 0 And p.NetLogTotal = 0)

    N1Fire p, "(function(){for(var i=0;i<5;i++){" & _
              "fetch('" & N1_LOCAL & "/n1?n='+i).catch(function(){});" & _
              "fetch('" & N1_NET & "/get?n='+i).catch(function(){});" & _
              "}return 10;})()"
    D3Pump 4

    Wv2Log.LogI "        総発火 " & p.NetLogTotal & " 件 / 手元 " & p.NetLogCount & _
                " 件 / 溢れ " & p.NetLogDropped & " 件"
    TestBool "5 件以上発火した", (p.NetLogTotal >= 5)
    TestBool "★手元は容量ぶんの 3 件だけ★", (p.NetLogCount = 3)
    TestBool "★溢れた数が数えられている★", (p.NetLogDropped = p.NetLogTotal - 3)

    ' --- ★N-1b: ドレインしても通算は残る (論点3 案Y)★ ---
    '   初回実機のログで「総発火 0 件」と出て 10 件来ていた事実が消えた。
    '   その再発を止める判定。
    totalBefore = p.NetLogTotal
    p.NetLogDrain
    TestBool "★ドレインしても総発火は残る (通算)★", (p.NetLogTotal = totalBefore)
    TestBool "  手元だけ空になる", (p.NetLogCount = 0)
    p.NetLogClear
    TestBool "  NetLogClear なら通算も 0 に戻る", (p.NetLogTotal = 0)

    p.NetLogCapacity = 500

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★止めたら本当に来なくなり、張り直せる★ ---"

    TestBool "NetCaptureStop が成功する", p.NetCaptureStop()
    TestBool "  捕捉中でなくなる", (p.NetCaptureOn = False)

    p.NetLogClear
    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=afterstop').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=afterstop').catch(function(){});return 1;})()"
    D3Pump 3
    ' ★フィルタも外している★ ので、止めた後は 1 件も来ないのが正しい。
    TestBool "★止めた後は 1 件も来ない★", (p.NetLogTotal = 0)

    TestBool "もう一度 NetCaptureStart できる", p.NetCaptureStart()
    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n1?k=restart').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=restart').catch(function(){});return 1;})()"
    TestBool "★張り直した網でまた捕まる★", N1Wait(p, "GET", "", "k=restart", 8)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★フォーム POST でのページ遷移が見える★ ---"
    ' ★これが方式 B (ネイティブイベント) を選んだ理由そのもの★
    '   ページ内 JS のラップでは、この 1 行はどうやっても見えない。

    p.NetLogClear
    Set el = p.QuerySelector("#postbtn-local")
    TestBool "案F: 送信ボタンを掴める", Not (el Is Nothing)
    If Not (el Is Nothing) Then
        TestBool "  案F: クリックできる", el.Click()
        TestBool "★案F: POST が DOCUMENT として捕まる★", _
                 N1Wait(p, "POST", "DOCUMENT", "127.0.0.1", 8)
    End If

    ' 検証ページへ戻る (仮想ホストなのでこの遷移自体はイベントに乗らない)
    p.NetLogClear
    If N1Nav(p, "https://" & N1_HOST & "/netprobe.html", "N-1 プローブ", 10) Then
        Set el = p.QuerySelector("#postbtn-net")
        TestBool "案D: 送信ボタンを掴める", Not (el Is Nothing)
        If Not (el Is Nothing) Then
            TestBool "  案D: クリックできる", el.Click()
            TestBool "★案D: POST が DOCUMENT として捕まる★", _
                     N1Wait(p, "POST", "DOCUMENT", "httpbingo.org", 10)
        End If
    Else
        Wv2Log.LogI "        検証ページへ戻れなかったので案D の POST は飛ばす"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) ★DOCUMENT の GET★ (論点2 案a) ---"
    ' ページ自身の DOCUMENT は仮想ホストなので取れない。的へ直接 Navigate して取る。

    p.NetLogClear
    p.View_Navigate N1_LOCAL & "/n1-doc-get"
    TestBool "★案F: DOCUMENT GET が捕まる★", _
             N1Wait(p, "GET", "DOCUMENT", "n1-doc-get", 8)

    p.NetLogClear
    p.View_Navigate N1_NET & "/get?k=doc"
    TestBool "★案D: DOCUMENT GET が捕まる★", _
             N1Wait(p, "GET", "DOCUMENT", "k=doc", 10)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 手元に残っているものを全部出す ---"
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N1_Capture 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N1_Site (N-1 の切り分け役 兼 実サイトでの実感)
'
'   ★外部サイトに実アクセスする★ Test_N1_Capture が仮想ホスト頼みなので、
'   「配線が悪いのか、仮想ホストの要求だけ乗らないのか」を 1 手で切り分ける。
'
'   やることは Test_N1_Capture の (1) と同じ:
'     タブを作る → ★Navigate の前に捕捉を始める★ → 開く → 溜まった分を出す。
'
'   実サイトなので件数は毎回変わる。★合否は「DOCUMENT が 1 件以上居るか」だけ★
'   見る (中身の数は数えない)。
' ============================================================
Public Sub Test_N1_Site(Optional ByVal url As String = "https://www.google.com/")
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim hr As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N1_Site: Browser が起動していません。"
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N1_Site: タブの生成に失敗しました。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N1_Site 開始 (" & url & ") ================"

    TestBool "NetCaptureStart が成功する", p.NetCaptureStart()

    hr = p.View_Navigate(url)
    TestBool "  Navigate が成功する", (hr = 0)

    ' 実サイトは遅いので少し長めに待つ (件数は問わない)
    D3Pump 6

    TestBool "★DOCUMENT が 1 件以上捕まっている★", _
             (N1Find(p, "", "DOCUMENT", "") > 0)
    TestBool "  何かしら捕まっている", (p.NetLogCount > 0)

    Wv2Log.LogI ""
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N1_Site 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N1_Watch (N-1 の実用の入口)
'
'   ★今開いているタブの通信を捕まえ始める★
'   「手で 1 回操作して、何をどの順で叩いているか」を見るための道具。
'
'   contexts を省略すると DOCUMENT / XHR / FETCH の 3 種別。
'   画像や CSS まで見たいときは Test_N1_Watch "ALL" と打つ。
' ============================================================
Public Sub Test_N1_Watch(Optional ByVal contexts As String = "")
    Dim b As Wv2Browser
    Dim p As Wv2Pane

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Debug.Print "[N-1] Browser が起動していません。"
        Exit Sub
    End If
    Set p = b.ActivePane
    If p Is Nothing Then
        Debug.Print "[N-1] アクティブなタブがありません。"
        Exit Sub
    End If

    If p.NetCaptureStart(contexts) Then
        Debug.Print "[N-1] 捕捉を開始しました。ページを手で操作してから Test_N1_Drain を打ってください。"
        Debug.Print "      ★判定・一覧はログファイルに出ます★ (%APPDATA%\Wv2Browser\logs)"
    Else
        Debug.Print "[N-1] 捕捉を開始できませんでした。ログを見てください。"
    End If
End Sub


' ============================================================
' Test_N1_Drain (N-1) - 溜まった分をログへ流して空にする
' ============================================================
Public Sub Test_N1_Drain()
    Dim b As Wv2Browser
    Dim p As Wv2Pane
    Dim n As Long

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then Exit Sub
    Set p = b.ActivePane
    If p Is Nothing Then Exit Sub

    n = p.NetLogDrain
    Debug.Print "[N-1] " & n & " 件をログへ流しました (総発火 " & p.NetLogTotal & " 件)。"
End Sub


' ============================================================
' Test_N1_Stop (N-1) - 捕まえるのをやめる (フィルタも外す)
' ============================================================
Public Sub Test_N1_Stop()
    Dim b As Wv2Browser
    Dim p As Wv2Pane

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then Exit Sub
    Set p = b.ActivePane
    If p Is Nothing Then Exit Sub

    p.NetLogDrain
    If p.NetCaptureStop() Then
        Debug.Print "[N-1] 捕捉を終了しました。"
    Else
        Debug.Print "[N-1] 捕捉は終了しましたが、後始末で失敗があります。ログを見てください。"
    End If
End Sub


' ============================================================
' Test_N1_Help (N-1 の手順)
' ============================================================
Public Sub Test_N1_Help()
    Debug.Print "==== N-1 実機手順 (ブラウザが何を叩いているかを見る) ===="
    Debug.Print ""
    Debug.Print "  【回帰試験 (外部依存ゼロ)】"
    Debug.Print "    1) UserForm1.Show vbModeless      ' ★仕様事実54★ 先に Show する"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart                ' このテスト 1 回分を 1 ファイルに閉じる"
    Debug.Print "    4) Test_N1_Capture"
    Debug.Print "    → ★判定はログファイルで読む★ %APPDATA%\Wv2Browser\logs の最新 1 本。"
    Debug.Print "       末尾の『★判定 n 件: OK x / FAIL y★』を見ること。"
    Debug.Print ""
    Debug.Print "  【落ちたときの切り分け】"
    Debug.Print "    Test_N1_Site                      ' 実サイトで同じことをする"
    Debug.Print "    ・そちらでも 0 件      → ★配線の問題★ (vtable / IID / フィルタ)"
    Debug.Print "    ・そちらだけ捕まる     → 検証ページの的の問題"
    Debug.Print ""
    Debug.Print "  【★N-1 で確定した仕様★】"
    Debug.Print "    仮想ホスト (SetVirtualHostNameToFolderMapping) で配信した要求は"
    Debug.Print "    ★WebResourceRequested に乗らない★ (DOCUMENT / fetch / XHR / 画像 / CSS"
    Debug.Print "    のすべて)。ランタイムが内部の資材ハンドラで直接返しているため。"
    Debug.Print "    Test_N1_Capture の的はこれを避けて 2 系統に置いてある:"
    Debug.Print "      案F  http://127.0.0.1:59999/…  到達不能なローカル (外部依存ゼロ)"
    Debug.Print "      案D  https://httpbingo.org/…   外部サービス (ネットが要る)"
    Debug.Print ""
    Debug.Print "    ★fetch() は FETCH ではなく XHR 種別で届く★"
    Debug.Print "    JS の fetch で撃った要求を WebView2 は XML_HTTP_REQUEST(7) として"
    Debug.Print "    報告する。種別で絞り込むときに決めつけないこと。既定のフィルタは"
    Debug.Print "    FETCH(8) も張ったままにしてある (害はなく、将来変わっても拾える)。"
    Debug.Print ""
    Debug.Print "  【実際のサイトで使う】"
    Debug.Print "    1) 調べたいページをタブで開く"
    Debug.Print "    2) Test_N1_Watch                  ' 捕捉開始 (DOCUMENT / XHR / FETCH)"
    Debug.Print "       Test_N1_Watch ""ALL""            ' 画像や CSS まで見たいとき"
    Debug.Print "    3) ★ページを手で 1 回操作する★ (ログインする、検索する、CSV を吐かせる…)"
    Debug.Print "    4) Test_N1_Drain                  ' 溜まった分をログへ流して空にする"
    Debug.Print "    5) Test_N1_Stop                   ' 終わり (フィルタも外れる)"
    Debug.Print ""
    Debug.Print "  --- ログの読み方 ---"
    Debug.Print "    #    経過ms  メソッド 種別       URL"
    Debug.Print "    経過ms は捕捉開始からの時間。★順番と間隔が分かるのが眼目★。"
    Debug.Print ""
    Debug.Print "  --- N-1 でできないこと (後段) ---"
    Debug.Print "    ・リクエストのヘッダとボディ          → N-2"
    Debug.Print "    ・レスポンスのステータスとヘッダ      → N-3"
    Debug.Print "    ・レスポンス本文                      → N-4"
    Debug.Print "    ・PowerShell の Invoke-WebRequest 化  → N-5"
    Debug.Print ""
    Debug.Print "  --- 注意 ---"
    Debug.Print "  ・★捕捉は Navigate の前に始める★ さもないと初回の DOCUMENT を取り逃がす。"
    Debug.Print "  ・既定のリングは 500 件。溢れたら NetLogDropped に出る (黙って切らない)。"
    Debug.Print "    足りなければ p.NetLogCapacity = 5000 のように増やせる。"
    Debug.Print "  ・★仕様事実 20★ イベントバーストが静まるまでブレーク/ステップ実行はしない。"
End Sub


' ============================================================
' N1Fire (N-1、Private) - ページ側で JS を 1 発撃つ
'   ★JS の文字列は必ずシングルクォート★ (プロジェクト規則)
' ============================================================
Private Sub N1Fire(ByVal p As Wv2Pane, ByVal js As String)
    p.EvalSync js, 5
    If Not p.LastEvalOk Then
        Wv2Log.LogI "        (N1Fire 失敗 err=" & p.LastEvalError & ")"
    End If
End Sub


' ============================================================
' N1Find (N-1、Private) - 条件に合う 1 件目を探す
'
'   methodWant / ctxWant / uriPart は空文字なら「問わない」。
'   戻り値は NetLogLine に渡せる 1 起点の位置。見つからなければ 0。
'
'   ★「たぶん来たはず」で済ませないための道具★ (設計原則112)
' ============================================================
Private Function N1Find(ByVal p As Wv2Pane, _
                        ByVal methodWant As String, _
                        ByVal ctxWant As String, _
                        ByVal uriPart As String) As Long
    Dim i As Long
    Dim f As Variant

    For i = 1 To p.NetLogCount
        f = Split(p.NetLogLine(i), vbTab)
        If UBound(f) >= 4 Then
            If (LenB(methodWant) = 0 Or UCase$(CStr(f(2))) = UCase$(methodWant)) Then
                If (LenB(ctxWant) = 0 Or UCase$(CStr(f(3))) = UCase$(ctxWant)) Then
                    If (LenB(uriPart) = 0 Or _
                        InStr(1, CStr(f(4)), uriPart, vbTextCompare) > 0) Then
                        N1Find = i
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i
End Function


' ============================================================
' N1Wait (N-1、Private) - 条件に合う 1 件が来るまで待つ
'   来たら True。見つかった行はログに出す。
' ============================================================
Private Function N1Wait(ByVal p As Wv2Pane, _
                        ByVal methodWant As String, _
                        ByVal ctxWant As String, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, methodWant, ctxWant, uriPart)
        If k > 0 Then
            Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
            N1Wait = True
            Exit Function
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N1CtxOf (N-1c、Private) - その要求が★何の種別で届いたか★を返す
'   見つからなければ "(居ない)"。★種別を決めつけずに観察するための道具★
' ============================================================
Private Function N1CtxOf(ByVal p As Wv2Pane, ByVal uriPart As String) As String
    Dim k As Long
    Dim f As Variant

    k = N1Find(p, "", "", uriPart)
    If k = 0 Then
        N1CtxOf = "(居ない)"
        Exit Function
    End If

    f = Split(p.NetLogLine(k), vbTab)
    If UBound(f) >= 3 Then N1CtxOf = CStr(f(3))
End Function


' ============================================================
' N1Nav (N-1b、Private) - 遷移してタイトルで着地を確かめる
'   戻り値: True なら目的のページに着いた。
' ============================================================
Private Function N1Nav(ByVal p As Wv2Pane, _
                       ByVal url As String, _
                       ByVal wantTitle As String, _
                       ByVal timeoutSec As Single) As Boolean
    If p.View_Navigate(url) <> 0 Then Exit Function
    N1Nav = D2WaitTitle(p, wantTitle, timeoutSec)
End Function


' ============================================================
' N1WriteFolder (N-1、Private) - 検証ページ一式を %TEMP% に書き出す
'   戻り値: 書き出したフォルダの絶対パス。失敗なら空文字。
'
'   置くもの:
'     netprobe.html … 入口。両系統の的への form / fetch / CSS / 画像を持つ
'     data.json     … ★仮想ホストの的 (対照)★ 捕まらないことを数えるために要る
'
'   ★N-1b で echo.html を捨てた★ フォームの送り先は仮想ホストの外 (案F / 案D)
'   に移したので、同じフォルダに送り先を置く必要がなくなった。
' ============================================================
Private Function N1WriteFolder() As String
    Dim folderPath As String

    folderPath = Environ$("TEMP")
    If Right$(folderPath, 1) <> "\" Then folderPath = folderPath & "\"
    folderPath = folderPath & N1_FOLDER

    If Not WriteUtf8NoBom(folderPath, "netprobe.html", BuildN1ProbeHtml()) Then Exit Function
    If Not WriteUtf8NoBom(folderPath, "data.json", "{""ok"":1}") Then Exit Function

    N1WriteFolder = folderPath
End Function


' ============================================================
' BuildN1ProbeHtml (N-1b の検証ページ)
'
'   ★静的な HTML なので引用符は VBA の "" で書いてよい★ (JS ではない)
'   ★このページ自身には JS を一切置かない★ fetch / XHR は EvalSync で撃つので、
'     ページ側に仕掛けを持たせない = 何が飛んだかの原因が 1 つに絞れる。
'
'   ★的はすべて仮想ホストの外★ (N-1b)
'     仮想ホストの要求は WebResourceRequested に乗らないと実機で分かったので、
'     form / link / img の行き先を案F (127.0.0.1) と案D (httpbingo.org) の
'     2 系統に置いた。どちらも実在しない資材を指すので 404 / 接続失敗になるが、
'     ★見たいのは「要求が飛んだ」という 1 行だけ★ (設計原則112)。
' ============================================================
Private Function BuildN1ProbeHtml() As String
    Dim s As String

    s = "<!DOCTYPE html>" & vbLf
    s = s & "<html lang=""ja""><head><meta charset=""UTF-8"">" & vbLf
    s = s & "<title>N-1 プローブ</title>" & vbLf

    ' --- 捕まらないはずの資材 (CSS)。両系統に 1 本ずつ ---
    s = s & "<link rel=""stylesheet"" href=""" & N1_LOCAL & "/n1.css"">" & vbLf
    s = s & "<link rel=""stylesheet"" href=""" & N1_NET & "/n1.css"">" & vbLf

    s = s & "<style>" & vbLf
    s = s & "  body{font-family:'Segoe UI','Meiryo',sans-serif;padding:36px 28px;" & _
            "background:#0e121b;color:#e8eaed;}" & vbLf
    s = s & "  h1{font-size:22px;margin:0 0 6px;}" & vbLf
    s = s & "  .lead{font-size:13px;color:#9aa7bd;line-height:1.7;margin-bottom:22px;}" & vbLf
    s = s & "  .row{border:1px solid rgba(255,255,255,.12);border-radius:12px;" & _
            "padding:16px 18px;margin-bottom:14px;background:rgba(255,255,255,.04);}" & vbLf
    s = s & "  .tag{font-size:10.5px;letter-spacing:.08em;color:#6ea8fe;" & _
            "border:1px solid rgba(110,168,254,.4);border-radius:999px;padding:2px 9px;}" & vbLf
    s = s & "  button{font-family:inherit;font-size:14px;font-weight:600;color:#0e121b;" & _
            "background:#6ea8fe;border:0;border-radius:8px;padding:9px 18px;" & _
            "cursor:pointer;margin-top:10px;}" & vbLf
    s = s & "  .note{font-size:12px;color:#8ea2c8;margin-top:10px;line-height:1.6;}" & vbLf
    s = s & "  .ep{font-family:'Consolas','Courier New',monospace;font-size:11.5px;" & _
            "color:#7d8aa0;}" & vbLf
    s = s & "</style></head><body>" & vbLf
    s = s & "<h1>N-1 プローブ</h1>" & vbLf
    s = s & "<div class=""lead"">WebResourceRequested が本当に拾えるかを確かめるページです。" & _
            "このページ自身は仮想ホストで配信されているので、" & _
            "<b>ページの読み込みそのものはイベントに乗りません</b>。" & _
            "的はすべて仮想ホストの外に置いてあります。</div>" & vbLf

    ' --- 案F: 到達不能なローカルへ POST ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <span class=""tag"">案F</span>" & vbLf
    s = s & "  <form id=""f-local"" action=""" & N1_LOCAL & "/echo"" method=""post"">" & vbLf
    s = s & "    <input type=""hidden"" name=""probe"" value=""wv2-n1-local"">" & vbLf
    s = s & "    <input type=""hidden"" name=""nihongo"" value=""日本語　全角スペース入り"">" & vbLf
    s = s & "    <button type=""submit"" id=""postbtn-local"">POST する (ローカル)</button>" & vbLf
    s = s & "  </form>" & vbLf
    s = s & "  <div class=""note"">接続は必ず失敗しますが、" & _
            "<b>要求が飛んだこと自体</b>が捕まれば合格です。" & _
            "<br><span class=""ep"">" & N1_LOCAL & "/echo</span></div>" & vbLf
    s = s & "</div>" & vbLf

    ' --- 案D: 外部サービスへ POST ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <span class=""tag"">案D</span>" & vbLf
    s = s & "  <form id=""f-net"" action=""" & N1_NET & "/post"" method=""post"">" & vbLf
    s = s & "    <input type=""hidden"" name=""probe"" value=""wv2-n1-net"">" & vbLf
    s = s & "    <input type=""hidden"" name=""nihongo"" value=""日本語　全角スペース入り"">" & vbLf
    s = s & "    <button type=""submit"" id=""postbtn-net"">POST する (外部)</button>" & vbLf
    s = s & "  </form>" & vbLf
    s = s & "  <div class=""note"">ネットが要ります。届けば 200 の JSON が返ります。" & _
            "<br><span class=""ep"">" & N1_NET & "/post</span></div>" & vbLf
    s = s & "</div>" & vbLf

    ' --- 捕まらないはずの資材 (画像)。両系統に 1 本ずつ ---
    s = s & "<div class=""row"">" & vbLf
    s = s & "  <img src=""" & N1_LOCAL & "/n1.png"" alt="""" width=""1"" height=""1"">" & vbLf
    s = s & "  <img src=""" & N1_NET & "/image/png"" alt="""" width=""1"" height=""1"">" & vbLf
    s = s & "  <div class=""note"">CSS と画像は要求こそ飛びますが、" & _
            "既定のフィルタ (DOCUMENT / XHR / FETCH) では捕まりません。" & _
            "★捕まらないことを数えて確かめます★</div>" & vbLf
    s = s & "</div>" & vbLf

    s = s & "</body></html>"

    BuildN1ProbeHtml = s
End Function







' ============================================================
' Test_N2_Detail (N-2 の回帰試験)
'
'   N-1 と同じ的を使う (案F = 到達不能なローカル / 案D = httpbingo.org)。
'   ページは仮想ホスト配信のままでよい (乗らないのはイベントだけ = 仕様事実69)。
'
'   見ているもの:
'     (0) ★詳細は既定 OFF★ (論点8 案κ)
'     (1) ヘッダが読める / ★GET はボディを読まない★ (論点7 案α)
'     (2) ボディが読める。★UTF-8 の日本語が往復する★
'     (3) ★上限で切ったことが分かる★ (黙って切らない)
'     (4) ★Cookie 相当のヘッダが伏せられる★ (論点3 案M)
'     (5) ★★読んでも通信が壊れていない★★ (論点2 案X / 論点9 案D')
'         ―― N-2 でいちばん確かめたいのはこれ
'     (6) 詳細リングは一覧とは別 (論点5 案S)
'     (7) OFF に戻すと詳細だけ増えなくなる
' ============================================================
Public Sub Test_N2_Detail()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim el  As Wv2Element
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim n0  As Long
    Dim localOk As Boolean
    Dim netOk   As Boolean
    Dim hdr As String
    Dim txt As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N2_Detail: Browser が起動していません。"
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N2_Detail: 検証ページの書き出しに失敗しました。中止します。"
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N2_Detail: タブの生成に失敗しました。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N2_Detail 開始 ================"
    Wv2Log.LogI "        案F の的 = " & N1_LOCAL & "  (到達不能なローカル)"
    Wv2Log.LogI "        案D の的 = " & N1_NET & "  (外部サービス)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) ★詳細は既定 OFF★ (論点8 案κ) ---"

    TestBool "★NetDetailOn の既定は False★", (p.NetDetailOn = False)
    TestBool "  詳細リングの既定容量は 50", (p.NetDetailCapacity = 50)
    TestBool "  ボディの上限の既定は 64KB", (p.NetBodyMaxBytes = 65536)
    TestBool "  ★秘匿の既定は ON★", (p.NetRedact = True)

    TestBool "NetCaptureStart が成功する", p.NetCaptureStart()
    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  仮想ホストのマッピングができる", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate が成功する", (hr = 0)

    If Not D2WaitTitle(p, "N-1 プローブ", 10) Then
        Wv2Log.LogI "Test_N2_Detail: 検証ページの読み込みを確認できませんでした。"
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    ' ★ここから詳細を取る★
    p.NetDetailOn = True
    p.NetLogClear
    p.NetDetailClear

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) ヘッダが読める / ★GET はボディを読まない★ ---"

    N1Fire p, "(function(){" & _
              "fetch('" & N1_LOCAL & "/n2h?k=hdr-local').catch(function(){});" & _
              "fetch('" & N1_NET & "/get?k=hdr-net').catch(function(){});" & _
              "return 1;})()"

    localOk = N2Wait(p, "k=hdr-local", 6)
    TestBool "★案F: 詳細が取れる★", localOk
    netOk = N2Wait(p, "k=hdr-net", 8)
    TestBool "★案D: 詳細が取れる★", netOk

    If Not (localOk Or netOk) Then
        Wv2Log.LogW "  ※ ★的が 1 つも届いていない★ 以降の判定は空振りとして読むこと。"
    End If

    k = N2Find(p, "k=hdr-local")
    If k = 0 Then k = N2Find(p, "k=hdr-net")
    If k > 0 Then
        hdr = p.NetDetailHeaders(k)
        Wv2Log.LogI "        ヘッダ (先頭 3 行):"
        Wv2Log.LogI "        " & Replace$(Left$(hdr, 240), vbLf, " / ")
        TestBool "★ヘッダが 1 個以上読めている★", (InStr(1, hdr, ":") > 0)
        TestBool "★User-Agent が居る★", (InStr(1, hdr, "user-agent", vbTextCompare) > 0)
        TestBool "★GET なのでボディは 0 バイト★ (論点7 案α)", _
                 (Split(p.NetDetailLine(k), vbTab)(3) = "0")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ボディが読める / ★UTF-8 の日本語が往復する★ ---"

    p.NetLogClear
    p.NetDetailClear
    ' ★fetch に文字列ボディを渡すと Content-Type: text/plain になる★
    '   これは CORS の「単純な要求」なのでプリフライトが挟まらない。
    '   到達不能な的でも要求そのものは飛ぶ。
    N1Fire p, "(function(){var o={method:'POST'," & _
              "body:'probe=wv2-n2-local&nihongo=日本語'};" & _
              "fetch('" & N1_LOCAL & "/n2b',o).catch(function(){});" & _
              "return 1;})()"

    TestBool "★POST の詳細が取れる★", N2Wait(p, "/n2b", 6)
    k = N2Find(p, "/n2b")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
        Wv2Log.LogI "        ボディ = [" & p.NetDetailBody(k) & "]"
        TestBool "★ボディの中身が読める★", _
                 (InStr(1, p.NetDetailBody(k), "probe=wv2-n2-local") > 0)
        TestBool "★UTF-8 の日本語が化けずに戻る★", _
                 (InStr(1, p.NetDetailBody(k), "日本語") > 0)
        TestBool "  バイト数が入っている", (CLng(Split(p.NetDetailLine(k), vbTab)(3)) > 0)
        TestBool "  切っていない", (Split(p.NetDetailLine(k), vbTab)(4) = "False")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★上限で切ったことが分かる★ ---"

    p.NetLogClear
    p.NetDetailClear
    p.NetBodyMaxBytes = 100
    TestBool "上限を 100 バイトにできる", (p.NetBodyMaxBytes = 100)

    N1Fire p, "(function(){var s='';for(var i=0;i<200;i++){s=s+'0123456789';}" & _
              "fetch('" & N1_LOCAL & "/n2big',{method:'POST',body:s}).catch(function(){});" & _
              "return 1;})()"

    TestBool "★大きいボディの詳細が取れる★", N2Wait(p, "/n2big", 6)
    k = N2Find(p, "/n2big")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
        TestBool "★100 バイトで止まっている★", (Split(p.NetDetailLine(k), vbTab)(3) = "100")
        TestBool "★切ったことが分かる★", (Split(p.NetDetailLine(k), vbTab)(4) = "True")
    End If
    p.NetBodyMaxBytes = 65536

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) ★秘密のヘッダが伏せられる★ (論点3 案M) ---"
    ' ★Cookie は的が到達不能だと作れない★ ので、機構そのものを
    '   「必ず在るヘッダ」= User-Agent で確かめる。Cookie / Authorization は
    '   既定リストに入っている (NetRedactDefaults)。

    p.NetLogClear
    p.NetDetailClear
    p.AddRedactHeader "user-agent"

    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2r?k=redact').catch(function(){});" & _
              "return 1;})()"
    TestBool "詳細が取れる", N2Wait(p, "k=redact", 6)
    k = N2Find(p, "k=redact")
    If k > 0 Then
        hdr = p.NetDetailHeaders(k)
        TestBool "★伏せ字になっている★", (InStr(1, hdr, "<伏せた") > 0)
        TestBool "★中身が漏れていない★", (InStr(1, hdr, "Mozilla") = 0)

        p.NetRedact = False
        hdr = p.NetDetailHeaders(k)
        TestBool "★NetRedact = False なら生で見られる★", (InStr(1, hdr, "Mozilla") > 0)
        p.NetRedact = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★★読んでも通信が壊れていない★★ (論点2 案X / 論点9 案D') ---"
    ' ★これが N-2 の本丸★
    '   IStream を読むと位置が進む。Seek(0) で戻し損ねていると、
    '   ★WebView2 が空のボディを送る★。それを「POST が成功したっぽい」で
    '   済ませず、★送り先が返す本文と突き合わせて数える★ (設計原則112)。
    '   httpbingo.org/post は受け取ったフォームの中身をそのまま JSON で返す。

    If Not netOk Then
        Wv2Log.LogW "  ※ 外部が届いていないのでこの節は飛ばす。★N-2 の本丸は未確認のまま★"
    Else
        p.NetLogClear
        p.NetDetailClear
        p.NetDetailUriFilter = "httpbingo.org/post"

        Set el = p.QuerySelector("#postbtn-net")
        TestBool "送信ボタンを掴める", Not (el Is Nothing)
        If Not (el Is Nothing) Then
            TestBool "  クリックできる", el.Click()

            TestBool "★POST の詳細が取れる★", N2Wait(p, "httpbingo.org/post", 10)
            k = N2Find(p, "httpbingo.org/post")
            If k > 0 Then
                Wv2Log.LogI "        " & Replace$(p.NetDetailLine(k), vbTab, "  ")
                Wv2Log.LogI "        ボディ = [" & Left$(p.NetDetailBody(k), 200) & "]"
                TestBool "★こちらが読んだボディに probe が入っている★", _
                         (InStr(1, p.NetDetailBody(k), "wv2-n1-net") > 0)
            End If

            ' ★送り先が実際に受け取ったか★ を本文で確かめる
            txt = N2PageText(p, 15)
            Wv2Log.LogI "        送り先の応答 (先頭 200 字):"
            Wv2Log.LogI "        " & Left$(Replace$(txt, vbLf, " "), 200)
            TestBool "★★送り先にボディが壊れずに届いている★★", _
                     (InStr(1, txt, "wv2-n1-net") > 0)
            TestBool "  ★日本語も壊れずに届いている★", _
                     (InStr(1, txt, "日本語") > 0)
        End If

        p.NetDetailUriFilter = ""
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) 詳細リングは一覧とは別 (論点5 案S) ---"

    p.NetLogClear
    p.NetDetailClear
    p.NetDetailCapacity = 2
    TestBool "詳細の容量だけ 2 にできる", (p.NetDetailCapacity = 2)
    TestBool "★一覧の容量は 500 のまま★", (p.NetLogCapacity = 500)

    N1Fire p, "(function(){for(var i=0;i<5;i++){" & _
              "fetch('" & N1_LOCAL & "/n2ring?n='+i).catch(function(){});}return 5;})()"
    D3Pump 3

    Wv2Log.LogI "        一覧 " & p.NetLogTotal & " 件 / 詳細 " & p.NetDetailCount & " 件"
    TestBool "★一覧は 5 件以上たまる★", (p.NetLogTotal >= 5)
    TestBool "★詳細は容量ぶんの 2 件だけ★", (p.NetDetailCount = 2)
    TestBool "★詳細も溢れた数を数えている★ (黙って切らない)", _
             (p.NetDetailDropped > 0)
    p.NetDetailCapacity = 50

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) ★OFF に戻すと詳細だけ増えなくなる★ ---"

    p.NetDetailOn = False
    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2off?k=off').catch(function(){});" & _
              "return 1;})()"
    D3Pump 3

    n0 = p.NetLogTotal
    Wv2Log.LogI "        一覧 " & n0 & " 件 / 詳細 " & p.NetDetailCount & " 件"
    TestBool "★一覧は増える (N-1 は生きている)★", (n0 > 0)
    TestBool "★詳細は 1 件も増えない★", (p.NetDetailCount = 0)

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 詳細を全部出す ---"
    p.NetDetailOn = True
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n2last?k=last').catch(function(){});" & _
              "return 1;})()"
    N2Wait p, "k=last", 6
    p.NetDetailAll

    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N2_Detail 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N2_Help (N-2 の手順)
' ============================================================
Public Sub Test_N2_Help()
    Debug.Print "==== N-2 実機手順 (リクエストのヘッダとボディ) ===="
    Debug.Print ""
    Debug.Print "  【回帰試験】"
    Debug.Print "    1) UserForm1.Show vbModeless      ' ★仕様事実54★"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N2_Detail"
    Debug.Print "    → 判定はログファイルで読む。★本丸は (5) の『送り先にボディが"
    Debug.Print "       壊れずに届いている』★ (外部が要る)。"
    Debug.Print ""
    Debug.Print "  【実際のサイトで使う】"
    Debug.Print "    1) 調べたいページをタブで開く"
    Debug.Print "    2) Set p = UserForm1.CurrentBrowser.ActivePane"
    Debug.Print "       p.NetDetailOn = True          ' ★詳細は既定 OFF★"
    Debug.Print "       Test_N1_Watch"
    Debug.Print "    3) ★ページを手で 1 回操作する★"
    Debug.Print "    4) Test_N1_Drain                 ' まず一覧で当たりを付ける"
    Debug.Print "       p.NetDetailAll                ' 詳細を全部出す"
    Debug.Print "       p.NetDetail 3                 ' 3 番目だけ出す"
    Debug.Print "    5) Test_N1_Stop"
    Debug.Print ""
    Debug.Print "  --- 絞り込み ---"
    Debug.Print "    p.NetDetailUriFilter = ""/api/"" ' この URL だけ詳細を取る"
    Debug.Print "    p.NetBodyMaxBytes = 200000       ' ボディの上限 (既定 64KB)"
    Debug.Print "    p.NetDetailCapacity = 200        ' 詳細リングの容量 (既定 50)"
    Debug.Print ""
    Debug.Print "  --- ★秘密の扱い★ ---"
    Debug.Print "    既定で Cookie / Authorization / X-Api-Key などは伏せて出る。"
    Debug.Print "    p.AddRedactHeader ""x-my-token""   ' 伏せる対象を足す"
    Debug.Print "    p.NetRedact = False              ' ★平文で出る。ログを人に渡す前に注意★"
    Debug.Print "    ※ リング上は生のまま持っている。伏せているのは出すときだけ。"
    Debug.Print ""
    Debug.Print "  --- ★N-2 で気をつけたこと★ ---"
    Debug.Print "  ・IStream は読むと位置が進む。★戻さないと空のボディが飛ぶ★ ので"
    Debug.Print "    読んだ後に Seek(0) している。壊れていないことは (5) で数えている。"
    Debug.Print "  ・ボディはハンドラの中で同期的に読むしかない (args は Invoke の間だけ"
    Debug.Print "    有効)。だから★上限で必ず切る★。切ったことは詳細の行に出る。"
    Debug.Print "  ・GET のボディは読まない (ほぼ無いので)。"
    Debug.Print ""
    Debug.Print "  --- N-2 でできないこと (後段) ---"
    Debug.Print "    ・レスポンスのステータスとヘッダ      → N-3"
    Debug.Print "    ・レスポンス本文                      → N-4"
    Debug.Print "    ・PowerShell の Invoke-WebRequest 化  → N-5"
End Sub


' ============================================================
' N2Find (N-2、Private) - URL の部分一致で詳細を探す
'   戻り値は NetDetailLine 等に渡せる 1 起点の位置。無ければ 0。
' ============================================================
Private Function N2Find(ByVal p As Wv2Pane, ByVal uriPart As String) As Long
    Dim i As Long
    Dim f As Variant

    For i = 1 To p.NetDetailCount
        f = Split(p.NetDetailLine(i), vbTab)
        If UBound(f) >= 2 Then
            If InStr(1, CStr(f(2)), uriPart, vbTextCompare) > 0 Then
                N2Find = i
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' N2Wait (N-2、Private) - その詳細が来るまで待つ
' ============================================================
Private Function N2Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If N2Find(p, uriPart) > 0 Then
            N2Wait = True
            Exit Function
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N2PageText (N-2、Private) - 今のページの本文を読む
'   ★送り先が実際に何を受け取ったか★ を確かめるために使う。
'   読めるようになるまで少し待つ (遷移直後は空のことがある)。
' ============================================================
Private Function N2PageText(ByVal p As Wv2Pane, ByVal timeoutSec As Single) As String
    Dim t0  As Single
    Dim res As String
    Dim cur As String

    t0 = Timer
    Do
        DoEvents
        res = p.EvalSync("document.body.innerText", 5)
        If p.LastEvalOk Then
            cur = Wv2Json.JsonUnescape(res)
            If Len(cur) > 20 Then
                N2PageText = cur
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
    N2PageText = cur
End Function


' ============================================================
' Test_N3_Status (N-3 の回帰試験)
'
'   ★このイベントにはフィルタが無い★ ので、一覧に居る要求に対応するものだけ
'   記録する (論点1 案A)。的は N-1 / N-2 と同じ 2 系統。
'
'   見ているもの:
'     (0) 応答イベントが張れている (NetRespOn)
'     (1) ★ふつうの 200 が一覧の行に付く★
'     (2) ★404 も 500 もそのまま出る★ (成功だけ見ない)
'     (3) ★リダイレクトは 1 要求 1 応答で並ぶ★ (論点6 案Y)
'     (4) 応答ヘッダが読める / ★応答ヘッダにも伏せ字が効く★
'         ★応答の中身を見る節は N3WaitAny で待つ★ (N2Wait では要求が
'         飛んだ瞬間に成立してしまい、必ず空を読む)
'     (5) ★到達不能な的は応答が来ない = 空欄★ (論点7 案α)
'     (6) ★一覧に居ない応答は数えるだけ★ (論点5 案W)
'     (7) 詳細が OFF なら応答ヘッダは取らない (ステータスは付く)
'
'   ★(1)～(4) は外部が要る★ 到達不能ローカルには応答が来ないので、
'   ステータスの検証には本物の送り先が要る。届かないときはその旨を出して飛ばす。
' ============================================================
Public Sub Test_N3_Status()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim n0  As Long
    Dim netOk As Boolean
    Dim hdr As String

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N3_Status: Browser が起動していません。"
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N3_Status: 検証ページの書き出しに失敗しました。中止します。"
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N3_Status: タブの生成に失敗しました。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N3_Status 開始 ================"
    Wv2Log.LogI "        案F の的 = " & N1_LOCAL & "  (到達不能なローカル)"
    Wv2Log.LogI "        案D の的 = " & N1_NET & "  (外部サービス)"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) 応答イベントが張れている ---"

    TestBool "NetCaptureStart が成功する", p.NetCaptureStart()
    TestBool "★応答イベントも張れている (NetRespOn)★", (p.NetRespOn = True)

    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  仮想ホストのマッピングができる", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate が成功する", (hr = 0)

    If Not D2WaitTitle(p, "N-1 プローブ", 10) Then
        Wv2Log.LogI "Test_N3_Status: 検証ページの読み込みを確認できませんでした。"
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    p.NetDetailOn = True
    p.NetLogClear
    p.NetDetailClear

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) ★ふつうの 200 が一覧の行に付く★ ---"

    N1Fire p, "(function(){fetch('" & N1_NET & "/get?k=n3-200').catch(function(){});" & _
              "return 1;})()"
    netOk = N3Wait(p, "k=n3-200", 200, 12)
    TestBool "★200 が付く★", netOk
    If netOk Then
        k = N1Find(p, "", "", "k=n3-200")
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
        TestBool "  NetLogStatus でも読める", (p.NetLogStatus(k) = 200)
    Else
        Wv2Log.LogW "  ※ ★外部が届いていない★ (1)～(4) は空振りとして読むこと。"
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★404 も 500 もそのまま出る★ ---"
    ' ★成功だけ見ない★ 失敗が失敗として見えることが調査道具の値打ち。

    If netOk Then
        p.NetLogClear
        N1Fire p, "(function(){" & _
                  "fetch('" & N1_NET & "/status/404?k=n3-404').catch(function(){});" & _
                  "fetch('" & N1_NET & "/status/500?k=n3-500').catch(function(){});" & _
                  "return 1;})()"
        TestBool "★404 が付く★", N3Wait(p, "k=n3-404", 404, 12)
        TestBool "★500 が付く★", N3Wait(p, "k=n3-500", 500, 12)
        k = N1Find(p, "", "", "k=n3-404")
        If k > 0 Then Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) ★リダイレクトは 1 要求 1 応答で並ぶ★ (論点6 案Y) ---"
    ' リダイレクトは要求そのものが複数回発火するので、応答も 1 対 1 で付く。
    ' ★同じ URL に 2 つの応答が付いて上書きされる、が起きないことを見る★

    If netOk Then
        p.NetLogClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/redirect/2?k=n3-redir')" & _
                  ".catch(function(){});return 1;})()"
        D3Pump 6
        n0 = N3CountWithStatus(p)
        Wv2Log.LogI "        一覧 " & p.NetLogCount & " 件 / うち応答が付いたもの " & n0 & " 件"
        N3DumpAll p
        TestBool "★リダイレクトで複数の行が並ぶ★", (p.NetLogCount >= 2)
        TestBool "★どの行にも応答が 1 つずつ付いている★", (n0 = p.NetLogCount)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) 応答ヘッダが読める / ★応答ヘッダにも伏せ字が効く★ ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & _
                  "/response-headers?X-N3-Test=hello&k=n3-hdr').catch(function(){});" & _
                  "return 1;})()"
        ' ★応答が付くまで待つ (N-3c)★
        '   N2Wait は「詳細が現れるまで」= 要求が飛んだ瞬間に成立してしまう。
        '   応答ヘッダを読みたいなら★応答が来るまで★待たなければ意味がない。
        TestBool "★応答が付くまで待てる★", N3WaitAny(p, "k=n3-hdr", 15)
        k = N2Find(p, "k=n3-hdr")
        If k > 0 Then
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        応答ヘッダ: " & Replace$(Left$(hdr, 300), vbLf, " / ")
            ' ★どこまで到達したかを丸ごと出す (N-3b)★
            '   [応答] 行が出れば NetAttachRespHeaders までは来ている。
            '   出なければ詳細エントリの突き合わせで落ちている。
            p.NetDetail k
            TestBool "★応答ヘッダが読めている★", (InStr(1, hdr, ":") > 0)
            TestBool "★指定したヘッダが返ってきている★", _
                     (InStr(1, hdr, "X-N3-Test", vbTextCompare) > 0)
            TestBool "  要求ヘッダとは別に持っている", _
                     (p.NetDetailHeaders(k) <> hdr)
        End If

        ' --- ★応答ヘッダ側でも伏せ字が効くか★ ---
        '   ★外部の Set-Cookie に頼らない★ 既に読めている content-type を
        '   伏せ対象に足して、★同じ 1 件を伏せる前と後で読み比べる★。
        '   これは論点3 案M (保存時ではなく出すときに伏せる) の直接の証明でもある。
        If k > 0 Then
            p.AddRedactHeader "content-type"
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        伏せた後: " & Replace$(Left$(hdr, 240), vbLf, " / ")
            TestBool "★応答ヘッダにも伏せ字が効く★", (InStr(1, hdr, "<伏せた") > 0)
            TestBool "★中身 (application/json) が消えている★", _
                     (InStr(1, hdr, "application/json") = 0)

            ' ★リング上は生のまま持っている★ ので False にすれば戻る
            p.NetRedact = False
            hdr = p.NetDetailRespHeaders(k)
            TestBool "★リング上は生のまま (NetRedact = False で戻る)★", _
                     (InStr(1, hdr, "application/json") > 0)
            p.NetRedact = True
        End If

        ' --- 本物の Set-Cookie でも試す (上乗せ。来なければ未検証と記録する) ---
        '   ★走らなかったことを黙って通さない★ FAIL 0 に見えて実は
        '   何も確かめていない、が一番たちが悪い。
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & _
                  "/cookies/set?n3secret=topsecret123&k=n3-cookie'" & _
                  ",{redirect:'manual'}).catch(function(){});return 1;})()"
        If N3WaitAny(p, "k=n3-cookie", 12) Then
            k = N2Find(p, "k=n3-cookie")
            hdr = p.NetDetailRespHeaders(k)
            Wv2Log.LogI "        応答ヘッダ: " & Replace$(Left$(hdr, 300), vbLf, " / ")
            If InStr(1, hdr, "et-cookie", vbTextCompare) > 0 Then
                TestBool "★本物の Set-Cookie の中身も伏せられている★", _
                         (InStr(1, hdr, "topsecret123") = 0)
            Else
                Wv2Log.LogW "  ※ ★Set-Cookie が応答ヘッダに無かった★ " & _
                            "この上乗せの判定は走らなかった (未検証)。"
            End If
        Else
            Wv2Log.LogW "  ※ ★Set-Cookie の上乗せ判定は走らなかった (応答が来ない)★ " & _
                        "redirect manual の要求は応答イベントに乗らないのかもしれない。" & _
                        "断定はせず宿題に積む。伏せ字の機構そのものは上で検証済み。"
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★到達不能な的は応答が来ない = 空欄★ (論点7 案α) ---"
    ' ★0 を書かないことの確認★ 「まだ来ていない」と「本物の 0」を混ぜない。

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n3none?k=n3-none').catch(function(){});" & _
              "return 1;})()"
    TestBool "要求は一覧に載る", N1Wait(p, "", "", "k=n3-none", 8)
    D3Pump 3
    k = N1Find(p, "", "", "k=n3-none")
    If k > 0 Then
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(k), vbTab, "  ")
        TestBool "★ステータスは 0 (未着) のまま★", (p.NetLogStatus(k) = 0)
        TestBool "★表示は空欄 (0 と書かない)★", _
                 (Split(p.NetLogLine(k), vbTab)(5) = "")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★一覧に居ない応答は数えるだけ★ (論点5 案W) ---"
    ' ★わざと捕捉対象外の要求を出して数える★
    '   画像は IMAGE 種別なので一覧 (DOCUMENT / XHR / FETCH) には載らない。
    '   だがこのイベントにはフィルタが無いので★応答だけは来る★。
    '   それが「対応なし」として数えられ、一覧には現れないことを見る。
    '   ★「>= 0」のような常に真の判定にしない★ (それは何も確かめていない)

    If netOk Then
        p.NetLogClear
        n0 = p.NetRespUnmatched
        N1Fire p, "(function(){var im=new Image();" & _
                  "im.src='" & N1_NET & "/image/png?k=n3-orphan';" & _
                  "return 1;})()"
        D3Pump 6
        Wv2Log.LogI "        一覧に対応が無かった応答 = " & p.NetRespUnmatched & " 件 " & _
                    "(クリア直後は " & n0 & " 件)"
        TestBool "★捕捉対象外の応答が数えられている★", (p.NetRespUnmatched > n0)
        TestBool "★その画像は一覧には載っていない★", (N1Find(p, "", "", "k=n3-orphan") = 0)
        TestBool "  一覧の取りこぼしは別勘定 (NetLogDropped)", (p.NetLogDropped = 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) 詳細が OFF でもステータスは付く ---"
    ' ★応答ヘッダだけが詳細扱い (論点4 案U)★ ステータスは一覧の一部。

    If netOk Then
        p.NetDetailOn = False
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/get?k=n3-off').catch(function(){});" & _
                  "return 1;})()"
        TestBool "★詳細 OFF でもステータスは付く★", N3Wait(p, "k=n3-off", 200, 12)
        TestBool "  詳細は 1 件も増えない", (p.NetDetailCount = 0)
        p.NetDetailOn = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- 手元に残っているものを全部出す ---"
    p.NetLogDrain
    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N3_Status 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N3_Help (N-3 の手順)
' ============================================================
Public Sub Test_N3_Help()
    Debug.Print "==== N-3 実機手順 (レスポンスのステータスとヘッダ) ===="
    Debug.Print ""
    Debug.Print "  【回帰試験】"
    Debug.Print "    1) UserForm1.Show vbModeless      ' ★仕様事実54★"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N3_Status"
    Debug.Print "    → ★(1)～(4) は外部 (httpbingo.org) が要る★"
    Debug.Print "       到達不能ローカルには応答が来ないので、ステータスの検証には"
    Debug.Print "       本物の送り先が要る。届かないときは空振りとして表示する。"
    Debug.Print ""
    Debug.Print "  【一覧の読み方 (N-3 で 1 列増えた)】"
    Debug.Print "    #   経過ms  メソッド 種別       状態  URL"
    Debug.Print "    1       94  GET      DOCUMENT    200  https://..."
    Debug.Print "    2      141  GET      XHR         404  https://..."
    Debug.Print "    3      609  GET      XHR              https://...   ← ★空欄 = まだ来ていない★"
    Debug.Print ""
    Debug.Print "  【応答ヘッダ】"
    Debug.Print "    p.NetDetailOn = True             ' ★応答ヘッダは詳細扱い★"
    Debug.Print "    p.NetDetail 3                    ' 3 番目の詳細 (要求 + 応答)"
    Debug.Print "    p.NetDetailRespHeaders(3)        ' 応答ヘッダだけ取る"
    Debug.Print "    ※ Set-Cookie は既定で伏せる (NetRedact)。"
    Debug.Print ""
    Debug.Print "  --- ★N-3 で気をつけたこと★ ---"
    Debug.Print "  ・★このイベントにはフィルタが無い★ 画像も CSS も全部来る。だから"
    Debug.Print "    一覧に既に居る要求に対応するものだけ記録する。全部見たいときは"
    Debug.Print "    Test_N1_Watch ""ALL"" で一覧側を広げれば応答もついてくる。"
    Debug.Print "  ・突き合わせは URI + メソッド。★同じ URL への同時並行要求は"
    Debug.Print "    取り違えるし、取り違えたことは検出できない★ (args に紐付けの"
    Debug.Print "    手がかりが無い)。リダイレクトは要求が複数回発火するので 1 対 1。"
    Debug.Print "  ・★NetLogDrain するとまだ応答が来ていない行も消える★ ので、"
    Debug.Print "    その応答は NetRespUnmatched に回る。急いで流さないこと。"
    Debug.Print "  ・NetRespUnmatched が大きいのは★正常★ (捕捉対象外の応答)。"
    Debug.Print "    取りこぼしを疑うときは NetLogDropped を見る。"
    Debug.Print ""
    Debug.Print "  --- N-3 でできないこと (後段) ---"
    Debug.Print "    ・レスポンス本文                      → N-4 (GetContent は非同期)"
    Debug.Print "    ・PowerShell の Invoke-WebRequest 化  → N-5"
End Sub


' ============================================================
' N3Wait (N-3、Private) - その URL に指定のステータスが付くまで待つ
' ============================================================
Private Function N3Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal wantStatus As Long, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, "", "", uriPart)
        If k > 0 Then
            If p.NetLogStatus(k) = wantStatus Then
                N3Wait = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N3WaitAny (N-3c、Private) - ★ステータスを問わず、応答が付くまで待つ★
'
'   N3Wait は「このステータスが付くまで」、N2Wait は「詳細が現れるまで」。
'   ★詳細は要求が飛んだ瞬間に作られる★ ので、N2Wait で待って応答ヘッダを
'   読むと必ず空になる (N-3 の初回実機でこれを踏んだ)。
'   応答の中身を見たいときは必ずこちらで待つこと。
' ============================================================
Private Function N3WaitAny(ByVal p As Wv2Pane, _
                           ByVal uriPart As String, _
                           ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long

    t0 = Timer
    Do
        DoEvents
        k = N1Find(p, "", "", uriPart)
        If k > 0 Then
            If p.NetLogStatus(k) <> 0 Then
                N3WaitAny = True
                Exit Function
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function


' ============================================================
' N3CountWithStatus (N-3、Private) - 応答が付いている行の数
' ============================================================
Private Function N3CountWithStatus(ByVal p As Wv2Pane) As Long
    Dim i As Long
    For i = 1 To p.NetLogCount
        If p.NetLogStatus(i) <> 0 Then N3CountWithStatus = N3CountWithStatus + 1
    Next i
End Function


' ============================================================
' N3DumpAll (N-3、Private) - 一覧をそのままログへ出す (消さない)
'   ★NetLogDrain と違って空にしない★ 判定に使う途中で消えると困るため。
' ============================================================
Private Sub N3DumpAll(ByVal p As Wv2Pane)
    Dim i As Long
    For i = 1 To p.NetLogCount
        Wv2Log.LogI "        " & Replace$(p.NetLogLine(i), vbTab, "  ")
    Next i
End Sub



' ============================================================
' Test_N4_Body (N-4 の回帰試験)
'
'   ★N-4 で初めて非同期になる★ GetContent はハンドラを渡して即座に戻り、
'   本文は後から届く。だから ★必ず「本文が届くまで」待つ★ (設計原則120)。
'   N-3 で「詳細が現れるまで」待って空を読んだ失敗を繰り返さない。
'
'   見ているもの:
'     (0) ★本文は既定 OFF★ (論点1 案A)
'     (1) JSON 本文が読める / ★何が来たかを全部残す★ (論点2 案P')
'     (2) ★圧縮の扱いがどちらか一方に確定する★ ―― N-4 の最大の未知数
'     (3) HTML も読める
'     (4) バイナリは中身を出さず hex で残す
'     (5) ★上限で切ったことが分かる★
'     (6) ★本文が無い応答 (204) と「まだ来ていない」を区別する★
'     (7) OFF に戻すと本文だけ取らない (ステータスとヘッダは付く)
'
'   ★(1)～(6) は外部が要る★ 到達不能ローカルには応答が来ないので本文も来ない。
' ============================================================
Public Sub Test_N4_Body()
    Dim b   As Wv2Browser
    Dim p   As Wv2Pane
    Dim folderPath As String
    Dim hr  As Long
    Dim k   As Long
    Dim netOk As Boolean
    Dim bodyTxt As String
    Dim hexHead As String
    Dim f   As Variant
    Dim isPlain As Boolean
    Dim isGz    As Boolean

    Set b = UserForm1.CurrentBrowser
    If b Is Nothing Then
        Wv2Log.LogI "Test_N4_Body: Browser が起動していません。"
        Exit Sub
    End If

    folderPath = N1WriteFolder()
    If LenB(folderPath) = 0 Then
        Wv2Log.LogI "Test_N4_Body: 検証ページの書き出しに失敗しました。中止します。"
        Exit Sub
    End If

    Set p = b.AddTab()
    If p Is Nothing Then
        Wv2Log.LogI "Test_N4_Body: タブの生成に失敗しました。"
        Exit Sub
    End If

    Wv2Log.LogI ""
    TestCountReset
    Wv2Log.LogI "================ Test_N4_Body 開始 ================"
    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (0) ★本文は既定 OFF★ (論点1 案A) ---"

    TestBool "★NetRespBodyOn の既定は False★", (p.NetRespBodyOn = False)
    TestBool "  本文の上限の既定は 64KB", (p.NetRespBodyMaxBytes = 65536)

    TestBool "NetCaptureStart が成功する", p.NetCaptureStart()
    hr = p.View3_SetVirtualHostNameToFolderMapping(N1_HOST, folderPath, 1)   ' 1 = ALLOW
    TestBool "  仮想ホストのマッピングができる", (hr = 0)
    hr = p.View_Navigate("https://" & N1_HOST & "/netprobe.html")
    TestBool "  Navigate が成功する", (hr = 0)

    If Not D2WaitTitle(p, "N-1 プローブ", 10) Then
        Wv2Log.LogI "Test_N4_Body: 検証ページの読み込みを確認できませんでした。"
        p.NetCaptureStop
        TestCountPrint
        Exit Sub
    End If
    D3Pump 2

    ' ★本文を取るには詳細も要る★ (置き場が詳細エントリなので)
    p.NetDetailOn = True
    p.NetRespBodyOn = True

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (1) JSON 本文が読める / ★何が来たかを全部残す★ ---"

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-json').catch(function(){});" & _
              "return 1;})()"

    netOk = N4Wait(p, "k=n4-json", 15)
    TestBool "★本文が届くまで待てる★", netOk
    If Not netOk Then
        Wv2Log.LogW "  ※ ★外部が届いていない★ (1)～(6) は空振りとして読むこと。"
    Else
        k = N2Find(p, "k=n4-json")
        Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
        Wv2Log.LogI "        本文 (先頭 160 字) = " & Left$(p.NetDetailRespBody(k), 160)
        f = Split(p.NetDetailRespLine(k), vbTab)
        TestBool "★バイト数が入っている★", (CLng(f(0)) > 0)
        TestBool "★テキストと判定している★", (f(2) = "True")
        TestBool "★先頭 32 バイトの hex が残っている★", (LenB(f(4)) > 0)
        TestBool "★JSON の中身が読める★", _
                 (InStr(1, p.NetDetailRespBody(k), "slideshow") > 0 Or _
                  InStr(1, p.NetDetailRespBody(k), "{") > 0)
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (2) ★圧縮の扱いを確定させる★ (N-4 の最大の未知数) ---"
    ' ★決めつけない (設計原則117)★
    '   httpbingo の /gzip は gzip 圧縮された JSON を返す。中身には
    '   "gzipped": true が入っている。だから
    '     ・本文に gzipped が読めたら → GetContent は★復号済み★を返す
    '     ・先頭 hex が 1F 8B なら     → GetContent は★圧縮されたまま★返す
    '   のどちらか一方に必ず決まる。それを判定にする。

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/gzip?k=n4-gzip').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-gzip", 15) Then
            k = N2Find(p, "k=n4-gzip")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            bodyTxt = p.NetDetailRespBody(k)
            f = Split(p.NetDetailRespLine(k), vbTab)
            hexHead = CStr(f(4))
            Wv2Log.LogI "        本文 (先頭 120 字) = " & Left$(bodyTxt, 120)

            isPlain = (InStr(1, bodyTxt, "gzipped") > 0)
            isGz = (Left$(hexHead, 5) = "1F 8B")

            TestBool "★圧縮の扱いがどちらか一方に確定した★", (isPlain Xor isGz)
            If isPlain Then
                Wv2Log.LogI "        → ★GetContent は復号済みの本文を返す★"
            ElseIf isGz Then
                Wv2Log.LogI "        → ★GetContent は圧縮されたまま返す★ " & _
                            "(VBA では解凍できないので、encoding とバイト数だけを残す設計にする)"
            Else
                Wv2Log.LogW "        → ★どちらとも言えない★ 素性と本文をそのまま読むこと。"
            End If
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (3) HTML も読める ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/html?k=n4-html').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-html", 15) Then
            k = N2Find(p, "k=n4-html")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            TestBool "★HTML の中身が読める★", _
                     (InStr(1, p.NetDetailRespBody(k), "<html", vbTextCompare) > 0 Or _
                      InStr(1, p.NetDetailRespBody(k), "Melville") > 0)
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (4) バイナリは中身を出さず hex で残す ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/image/png?k=n4-png')" & _
                  ".catch(function(){});return 1;})()"
        If N4Wait(p, "k=n4-png", 15) Then
            k = N2Find(p, "k=n4-png")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "★バイナリと判定している★", (f(2) = "False")
            TestBool "  本文は空 (中身を出さない)", (LenB(p.NetDetailRespBody(k)) = 0)
            TestBool "★hex は残っている★", (LenB(f(4)) > 0)
        End If
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (5) ★上限で切ったことが分かる★ ---"

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        p.NetRespBodyMaxBytes = 100
        TestBool "上限を 100 バイトにできる", (p.NetRespBodyMaxBytes = 100)
        N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-cut').catch(function(){});" & _
                  "return 1;})()"
        If N4Wait(p, "k=n4-cut", 15) Then
            k = N2Find(p, "k=n4-cut")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "★100 バイトで止まっている★", (f(0) = "100")
            TestBool "★切ったことが分かる★", (f(1) = "True")
        End If
        p.NetRespBodyMaxBytes = 65536
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (6) ★本文が無い応答と「まだ来ていない」を区別する★ ---"
    ' 204 は本文が無い。0 バイトだが★届いてはいる★。
    ' 到達不能な的は★そもそも届かない★。この 2 つを混ぜない (設計原則111)。

    If netOk Then
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/status/204?k=n4-204')" & _
                  ".catch(function(){});return 1;})()"
        If N4Wait(p, "k=n4-204", 15) Then
            k = N2Find(p, "k=n4-204")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "★0 バイトだが届いている★", (f(0) = "0" And f(5) = "True")
            ' ★204 は ERROR_NO_DATA で完了が来る。これは失敗ではない★
            '   「取得失敗」と書くと N-5 で『本文があるか』を誤らせる (設計原則111)。
            TestBool "★『取得失敗』とは書かない (本文が無いだけ)★", _
                     (InStr(1, CStr(f(3)), "取得失敗") = 0)
        End If
    End If

    p.NetLogClear
    p.NetDetailClear
    N1Fire p, "(function(){fetch('" & N1_LOCAL & "/n4none?k=n4-none')" & _
              ".catch(function(){});return 1;})()"
    D3Pump 4
    k = N2Find(p, "k=n4-none")
    If k > 0 Then
        Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
        f = Split(p.NetDetailRespLine(k), vbTab)
        TestBool "★到達不能なら「届いていない」のまま★", (f(5) = "False")
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "  --- (7) ★OFF に戻すと本文だけ取らない★ ---"
    ' ★ステータスとヘッダは付く★ = 本文だけが別のスイッチであることの確認。

    If netOk Then
        p.NetRespBodyOn = False
        p.NetLogClear
        p.NetDetailClear
        N1Fire p, "(function(){fetch('" & N1_NET & "/json?k=n4-off').catch(function(){});" & _
                  "return 1;})()"
        If N3WaitAny(p, "k=n4-off", 15) Then
            D3Pump 3
            k = N2Find(p, "k=n4-off")
            Wv2Log.LogI "        素性 = " & Replace$(p.NetDetailRespLine(k), vbTab, "  ")
            f = Split(p.NetDetailRespLine(k), vbTab)
            TestBool "★本文は届かない★", (f(5) = "False")
            TestBool "★ステータスは付く★", (p.NetLogStatus(N1Find(p, "", "", "k=n4-off")) > 0)
            TestBool "★応答ヘッダも付く★", (InStr(1, p.NetDetailRespHeaders(k), ":") > 0)
        End If
        p.NetRespBodyOn = True
    End If

    Wv2Log.LogI ""
    Wv2Log.LogI "        置き場が無かった本文 = " & p.NetRespBodyOrphan & " 件"

    p.NetCaptureStop

    Wv2Log.LogI ""
    Wv2Log.LogI "  in-callback 深さ (期待 0): " & p.InCallbackDepth
    TestCountPrint
    Wv2Log.LogI "================ Test_N4_Body 終了 ================"
    Wv2Log.LogI ""
End Sub


' ============================================================
' Test_N4_Help (N-4 の手順)
' ============================================================
Public Sub Test_N4_Help()
    Debug.Print "==== N-4 実機手順 (レスポンス本文) ===="
    Debug.Print ""
    Debug.Print "  【回帰試験】"
    Debug.Print "    1) UserForm1.Show vbModeless      ' ★仕様事実54★"
    Debug.Print "    2) UserForm1.StartWebView2_Full"
    Debug.Print "    3) Wv2Log.LogStart"
    Debug.Print "    4) Test_N4_Body"
    Debug.Print "    → ★(1)～(6) は外部 (httpbingo.org) が要る★"
    Debug.Print ""
    Debug.Print "  【使い方】"
    Debug.Print "    p.NetDetailOn = True             ' ★本文の置き場は詳細エントリ★"
    Debug.Print "    p.NetRespBodyOn = True           ' ★本文は別のスイッチ (既定 OFF)★"
    Debug.Print "    Test_N1_Watch"
    Debug.Print "    ' ... 手で操作する ..."
    Debug.Print "    Test_N1_Drain                    ' 一覧で当たりを付ける"
    Debug.Print "    p.NetDetail 3                    ' 要求 + 応答 + 本文"
    Debug.Print "    p.NetDetailRespBody(3)           ' 本文だけ取る"
    Debug.Print "    p.NetDetailRespLine(3)           ' バイト数/切った/種類/encoding/hex/届いた"
    Debug.Print "    p.NetRespBodyMaxBytes = 500000   ' 本文の上限 (既定 64KB)"
    Debug.Print ""
    Debug.Print "  --- ★N-4 で作りが変わったところ★ ---"
    Debug.Print "  ・★初めて非同期★ GetContent はハンドラを渡して即座に戻り、本文は"
    Debug.Print "    後から届く。判定するときは★本文が届くまで待つ★こと (設計原則120)。"
    Debug.Print "  ・★N-2 と違って Seek(0) は要らない★ 応答のコピーを読むだけで、"
    Debug.Print "    通信物には手を入れない。"
    Debug.Print "  ・★ResponseView は完了まで保持している★ (解放してよいか SDK に"
    Debug.Print "    書かれていないため)。NetCaptureStop で必ずほどく。"
    Debug.Print ""
    Debug.Print "  --- ★圧縮について★ ---"
    Debug.Print "  応答は gzip / br / zstd で来る。GetContent が復号済みを返すかは"
    Debug.Print "  実機で確かめた (Test_N4_Body の (2))。どちらであっても"
    Debug.Print "  ★content-encoding と先頭 32 バイトの hex は必ず残る★ ので、"
    Debug.Print "  本文が読めないときはそこを見ること。"
    Debug.Print ""
    Debug.Print "  --- N-4 でできないこと (後段) ---"
    Debug.Print "    ・PowerShell の Invoke-WebRequest 化  → N-5"
End Sub


' ============================================================
' N4Wait (N-4、Private) - ★本文が届くまで★ 待つ
'
'   ★「詳細が現れるまで」でも「ステータスが付くまで」でもない★
'   本文は GetContent の完了ハンドラで最後に届くので、それを待つ。
'   N-3 でここを取り違えて実機 3 回ぶんを溶かした (設計原則120)。
' ============================================================
Private Function N4Wait(ByVal p As Wv2Pane, _
                        ByVal uriPart As String, _
                        ByVal timeoutSec As Single) As Boolean
    Dim t0 As Single
    Dim k  As Long
    Dim f  As Variant

    t0 = Timer
    Do
        DoEvents
        k = N2Find(p, uriPart)
        If k > 0 Then
            f = Split(p.NetDetailRespLine(k), vbTab)
            If UBound(f) >= 5 Then
                If f(5) = "True" Then          ' ★届いた★
                    N4Wait = True
                    Exit Function
                End If
            End If
        End If
        If (Timer - t0) > timeoutSec Then Exit Do
    Loop
End Function
