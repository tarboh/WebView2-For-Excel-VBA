VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Pane"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-2 段階 (リクエストのヘッダとボディ) ---
'
'   N-1 の一覧 (URL / メソッド / 種別) の上に、★1 本を再現するのに要るもの★を足す。
'   詳細一式は N-2 の節 (このファイルの末尾) にまとめてある。要点だけ:
'
'     ・★詳細はまるごと既定 OFF★ (NetDetailOn = True で ON)
'     ・★読んだ IStream は必ず Seek(0) で戻す★ さもないと空のボディが飛ぶ
'     ・★Cookie / Authorization は既定で伏せる★ (リング上は生のまま)
'     ・詳細は一覧とは別のリング (既定 50 件)。一覧は 500 件のまま
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-1b 段階 (実機で分かったことの反映) ---
'
'   ■ ★仕様事実: 仮想ホストの要求は WebResourceRequested に乗らない★
'     SetVirtualHostNameToFolderMapping で配信した要求 (DOCUMENT / fetch / XHR /
'     画像 / CSS のすべて) はイベントに乗らない。ランタイムが内部の資材ハンドラで
'     直接返しているため。★N-1 の初回検証が全滅した原因はこれ★ で、配線側は
'     最初から正しかった (実サイトでは 10 件すべて捕まっていた)。
'     検証ページの的は「到達不能なローカルアドレス」と「外部サービス」に置く。
'
'   ■ ★NetLogDrain は通算カウンタを 0 に戻さない (論点3 案Y)★
'     NetLogTotal / NetLogDropped は「★捕捉開始からの★ 通算」と説明しているのに、
'     Drain がリングごと作り直して 0 に戻していた。そのせいで実機ログの最後が
'       NetCaptureStop: 捕捉終了 (総発火 0 件 ...)
'     となり、★10 件来ていたのに「1 件も来なかった」と読める★ 行が残った。
'     ・NetLogDrain … リングの中身だけ空にする (NetRingEmpty)。通算は保つ
'     ・NetLogClear … 通算も含めて全部リセット (NetRingAlloc)
'     ・NetCaptureStart … NetLogClear を通るので毎回 0 から数え直す
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  N-1 段階 (WebResourceRequested の配線) ---
'
'   ★ブラウザが実際に何を叩いているかを Excel VBA だけで見る★
'   F12 の Network タブに相当するものの第一段。N-1 は URL / メソッド / 種別だけ
'   (ヘッダもボディもレスポンスも後段)。
'
'   ■ 追加した Public メンバ
'     NetCaptureStart(contexts, uriFilter) As Boolean  捕捉を始める
'     NetCaptureStop()  As Boolean                     捕捉をやめる
'     NetCaptureOn      As Boolean                     捕捉中か
'     NetLogDrain()     As Long     溜まった分を Wv2Log へ流して空にする
'     NetLogLine(i)     As String   ★流さずに 1 件だけ読む★ (古い順、1 起点)
'     NetLogCount / NetLogTotal / NetLogDropped As Long
'     NetLogCapacity    As Long     リングの容量 (Get/Let)
'     NetLogClear()                 リングを空にする
'     View_AddWebResourceRequested / View_RemoveWebResourceRequested
'     View_AddWebResourceRequestedFilter / View_RemoveWebResourceRequestedFilter
'     View_OnWebResourceRequested   (ComCallbackHandler からの受け口)
'
'   ■ ★記録だけしてすぐ返す (論点5)★
'     WebResourceRequested は「要求を止める / 書き換える」ためのイベントで、
'     ★ハンドラが遅いと通信そのものが詰まる★。だからハンドラの中では
'     ・Wv2Log も Debug.Print も撃たない
'     ・EvalSync は当然呼ばない (仕様事実14 / 20)
'     args から URL / メソッド / 種別を読んでリングに置くだけで戻る。
'
'   ■ ★記録はメモリ上のリングバッファ (論点6)★
'     要求は 1 ページで数十～数百件出る。ログに直書きすると量で溢れるので、
'     D-4 の SpaProbeDrainLog と同じ型にした (溜める ← まとめて取り出す)。
'     ★溢れた分は黙って捨てない★ NetLogDropped で見える (D-6 論点3 と同じ判断)。
'
'   ■ ★フィルタを 1 つも登録しないとイベントは 1 件も来ない★
'     AddWebResourceRequestedFilter(uri, ctx) を種別ごとに撃つ。既定は
'     DOCUMENT(1) / XML_HTTP_REQUEST(7) / FETCH(8) の 3 つ (論点4)。
'     ★登録したフィルタは覚えておいて NetCaptureStop で外す★ (設計原則106 の趣旨。
'     相手側に前の設定を残すと、次の Start が知らないうちに広い網になる)。
'
'   ■ ★vtable index は SDK のヘッダで確定させた (設計原則115)★
'     既存の配線からの逆算では add_WebResourceRequested を 54 と読み違えていた
'     (正解 55、OpenDevToolsWindow を数え落としていた)。道具は tools/vtable.py。
'       ICoreWebView2::add_WebResourceRequested              55
'       ICoreWebView2::remove_WebResourceRequested           56
'       ICoreWebView2::AddWebResourceRequestedFilter         57
'       ICoreWebView2::RemoveWebResourceRequestedFilter      58
'       ...WebResourceRequestedEventArgs::get_Request         3
'       ...WebResourceRequestedEventArgs::get_ResourceContext 7
'       ICoreWebView2WebResourceRequest::get_Uri              3
'       ICoreWebView2WebResourceRequest::get_Method           5
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  K-4a 段階 (解放経路のログを Wv2Log へ移す) ---
'
'   ★目的は計測★ フォームを閉じたときの解放が遅い (1 件ごとに 0.5 秒ほどの間)
'   という報告を受けたが、この経路は Debug.Print のままでログに何も残らず、
'   どこで時間を食っているのか分からなかった。
'
'   Wv2Log はミリ秒のタイムスタンプを持つので、★置き換えるだけで所要が読める★。
'   さらに Wv2Log.LogEcho = False / True で 2 回測れば、
'   ★「COM の解放が遅い」のか「イミディエイトへの出力が遅い」のか★ が切り分く。
'
'   ★この段階では直さない。測るだけ。★ (D-7 で推測を 3 回続けて外したため)
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-6 段階 (QuerySelectorAll と寿命管理) ---
'
'   ■ 追加した Public メンバ
'     QuerySelectorAll(selector, maxCount) As Collection
'     LastAllTruncated  As Boolean  直前の QuerySelectorAll が打ち切られたか
'     ElementCount      As Long     JS 側レジストリに積まれている要素の数
'
'   ■ ★複数は Collection で返す (論点2)★
'     For Each el In els と書けて VBA に最も自然。0 件は Count = 0 で素直。
'     ★配列にしなかった理由★ 空配列の UBound は実行時エラーになり、呼び出し側に
'     毎回の防御を強いる。★独自クラスにしなかった理由★ For Each を効かせるには
'     メンバー属性行 (Attribute Item.VB_UserMemId = 0) が要るが、import.ps1 は
'     属性行を落とす仕様 (仕様事実47) なので往復と相性が悪い。
'
'   ■ ★打ち切りは黙ってやらない (論点3)★
'     既定 200 件。超えたら LastAllTruncated = True にして警告ログも出す。
'     「返ってきた件数がページの実数」と思い込ませない。
'
'   ■ ★寿命管理は「見せる」だけ。勝手に捨てない (論点4)★
'     QuerySelectorAll は 1 回で 200 個積みうるので、レジストリは D-2 の想定より
'     速く伸びる。だが古いものを勝手に捨てると★手元の Wv2Element が黙って
'     stale になる★ (設計原則91 / D-2 論点3 の判断がそのまま効く)。
'     そこで:
'       ・ElementCount で今の数を見せる
'       ・一定数を超えたら LogW で知らせる
'       ・捨てるのは ClearElementRegistry を呼んだときだけ (D-2 から在る)
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-4b 段階 (静穏待ち) ---
'
'   D-4a で置いたプローブを使って★落ち着くまで待つ★。
'
'   ■ 追加した Public メンバ
'     WaitSettled(timeoutSec, stableMs) As Boolean  静穏になるまで待つ
'     AutoWaitAfterAction  As Boolean  ★既定 False★ 操作の後に自動で待つ (opt-in)
'     AddIgnoreNetwork(pattern)        静穏判定から外す URL の部分一致パターン
'     ClearIgnoreNetwork()             同上を空にする
'     IgnoreSelectors  As String       静穏判定から外す DOM 変化のセレクタ
'     LastSettleInfo   As String       直前の待ちの内訳 (診断用)
'
'   ■ 静穏の判定式 (論点4)
'     実行中の fetch + XHR が 0 かつ、最後の DOM 変化から stableMs 以上たったら静穏。
'     ★ただし静穏窓は「WaitSettled を呼んだ時点」から測る★ 呼ぶ前から静かでも
'     いきなり成立させない。したがって★最短でも stableMs は待つ★。
'     既定は stableMs = 500 / 上限 10 秒 / ポーリング 100ms (参考プロダクトに倣った)。
'
'   ★★★ WaitSettled が True でも「アプリの処理が終わった」証明にはならない ★★★
'     静かなだけである。**アプリが要求を無視して静かなのか、終わって静かなのかは
'     区別できない。** 重要な操作では D-4c の明示シグナル (arm) を併用すること。
'     ★名前を WaitIdle でなく WaitSettled にしたのはこのため★ (論点8)
'     「idle = 何も起きていない」と読めると保証と誤解される。
'
'   ■ 戻り値 (D-2 / D-3 の規約を踏襲、設計原則93)
'     True                     … 静穏に達した
'     False + LastEvalOk=True  … 時間内に静まらなかった (★正常な結果★)
'     False + LastEvalOk=False … 失敗 (in-callback / Close 済み等)
'     どちらでも LastSettleInfo に内訳が入る (slack / 実行中の通信 / 経過)。
'
'   ■ ノイズ除外 (論点5)
'     ★除外リストは VBA 側が持ち、毎回の往復でページへ送り込む★
'     ページ側に持たせると遷移で消えるため。送り込みは状態の読み取りと
'     同じ 1 往復に相乗りするので、往復は増えない。
'     ★広すぎる除外は「早すぎる静穏」を生む★ 参考プロダクトも警告している。
'     遮断 (要求そのものを止める) はやらない。ページの挙動を変えるため。
'
'   ■ ★プローブを置く前に始まった通信は数えられない★ (D-4a の仕様)
'     待ちの直前に操作する運用 (D-4c の arm-then-act) なら実害が無い。
'     ページを開いた直後の読み込み完了を待ちたいときは、既存の
'     NavigationCompleted / DOMContentLoaded を使うこと。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-4a 段階 (ページ内 SPA プローブと健康診断) ---
'
'   D-4 の目的は★静穏待ち★ (通信と DOM の変化が落ち着くまで待つ)。
'   D-3 の WaitFor / WaitGone は「何を待つか」を書ける場合の手段で、
'   **待つ対象が特定できないとき**には使えない。そこを埋めるのが D-4。
'
'   D-4a はその土台。★観測する仕掛け (プローブ) をページ内に置くところまで★で、
'   待ちの本体 (WaitSettled) と明示シグナル (Arm 系) は D-4b / D-4c。
'
'   ■ 追加した Public メンバ
'     SpaProbeState()  As String   プローブを整えてから状態 JSON を返す
'     SpaProbeReset()  As Boolean  プローブを捨てて作り直す (診断用)
'     SpaProbeHealthy() As Boolean 状態 JSON の h を読むだけの薄い口
'
'   ■ プローブが数えているもの (window.__wv2p)
'     fetchN / xhrN … ★実行中★ の fetch / XHR の数 (in-flight)
'     lastMut       … 最後に数えた DOM 変化の時刻。静穏窓はここから測る
'     netN / mutN   … 数えた総数 (診断用)
'     skipNet / skipMut … 除外した数 (ノイズ除外が効いているかの確認用)
'
'   ■ ★設計原則92 の拡張: 「無ければ作る」から「壊れていたら作り直す」へ★
'     fetch と XMLHttpRequest.prototype.open/send をラップして数えるが、
'     ★ページ側のフレームワークが後からこれらを差し替えることがある★
'     (参考プロダクト WebDriver BiDi for SeleniumVBA が健康診断を実装している
'      = 実際に壊れるということ)。そこで毎回
'       ・版番号が一致するか
'       ・window.fetch が今もこちらのラッパーか
'       ・XHR の open/send が今もこちらのラッパーか
'       ・MutationObserver が生きているか
'     を確かめ、1 つでも欠けたら★作り直す★。作り直した回数は repairs で読める。
'
'     ★作り直すときは「現在の window.fetch」を包む★
'     フレームワークが被せたラッパーごと包むので、向こうの機能を壊さない。
'     自分のラッパーを二重に包まない判定も入れてある (数え漏れ / 二重計上の防止)。
'
'   ■ ★いつ仕掛けるか (論点2 案A: 自己生成プロローグの拡張)★
'     AddScriptToExecuteOnDocumentCreated は使わない。D-2 と同じく
'     「使う式が自分で用意する」形にする。したがって
'     ★プローブを置いた時点より前に始まった通信は数えられない★
'     これは仕様。D-4c の「操作の直前に arm する」運用と組み合わせれば実害が無い
'     (arm した時点でプローブが立つので、必ず操作より前になる)。
'
'   ■ プローブが作られる場所
'     ★これらのメンバを呼んだ Pane にしか作られない★ タブバー / ナビバー /
'     設定タブの Pane は呼ばないので無傷 (D-2 で注入を避けた理由と同じ)。
'
''''''''''''''''''''''''''''''''''' --- Wv2Pane.cls  D-3 段階 (要素への書き込みと操作) ---
'
'   D-3 で Wv2Pane に増えたのは ★Elem_Exec (Friend) 1 本だけ★。
'   書き込み・操作の入口は Wv2Element 側 (Value = / Click / SetAttribute) にあり、
'   Wv2Pane は D-2 の Elem_Read と対になる「撃つ側」の経路を提供する。
'
'     Elem_Read (D-2) … 値を 1 個★読む★   変数 e を使う JS ★式★ を受け取る
'     Elem_Exec (D-3) … 操作を 1 回★撃つ★ 変数 e を使う JS ★文★ を受け取る
'
'   ■ Elem_Exec の JS 側の返し方 (Elem_Read と同じ形に揃えた)
'     {s:'gen', g:...}            世代違い = stale
'     {s:'nil', g:...}            添字に要素が無い = stale 扱い
'     {s:'ok',  g:..., v:'<rv>'}  成功。v は★操作が選んだ経路の名前★
'
'     ★JS 例外は try-catch しない★ View_ExecuteScriptSafe が既に包んでいるので、
'     例外は EvalSync の LastEvalError ("TypeError: ..." 等) として上がってくる。
'     二重に包むと理由が潰れる。
'
'   ■ 経路の名前 (v) を返す理由
'     .Value = の書き込みが★ネイティブ setter 経由 ('setter') か素の代入
'     ('direct') か★は、SPA に効いたかどうかの唯一の手掛かりになる。
'     Wv2Element.LastInfo で読める (D-3 の事後報告)。
'
'   ■ DOM 条件待ち (D-3b、論点2 案C: ★明示待ちを第一級に、暗黙待ちは opt-in★)
'
'     WaitFor(selector, timeoutSec)  As Wv2Element  現れるまで待つ
'     WaitGone(selector, timeoutSec) As Boolean     消えるまで待つ
'     ImplicitWaitSec                As Single      ★既定 0 = 無効★
'
'     ★未知1 は実測で決着した (2026-08-22)★
'       raw ExecuteScript に★決して解決しない Promise★ を渡しても、完了通知が
'       即座に resultJson={} で返ってきた。つまり ExecuteScript は Promise を待たない。
'       したがって待ちは JS 側では書けず、★VBA 側でポーリングする★ (論点3 の骨格)。
'       (そもそも EvalSync は JSON.stringify(式) を撃つので、待てたとしても
'        Promise はその場で {} に潰れる。JS 側で待つなら D-1 の包み方から作り直しになる)
'
'     ★ポーリング中はブレーク/ステップ実行しないこと★ (仕様事実20)
'       EvalSync の DoEvents 区間が待ち時間ぶん伸びる。10 秒待てば 10 秒間ずっと危険域。
'
'     ★Nothing / False の意味は LastEvalOk で分ける (設計原則93)★
'       WaitFor  → Nothing + LastEvalOk=True  … 時間内に現れなかった (正常)
'                  Nothing + LastEvalOk=False … 失敗 (セレクタ不正 / in-callback 等)
'       WaitGone → False   + LastEvalOk=True  … 時間内に消えなかった (正常)
'                  False   + LastEvalOk=False … 失敗
'       ★失敗と判った時点で待たずに抜ける★ in-callback で 10 秒固まるのを防ぐため。
'
'     ImplicitWaitSec は QuerySelector / GetElementById に暗黙のリトライを掛ける
'     (SeleniumVBA と同じ流儀)。★既定は 0 で、明示的に設定した人だけが影響を受ける★
'     理由: 既定で待つと「本当に無い」の確認が毎回タイムアウト待ちになり、
'     D-2 で決めた「Nothing + LastEvalOk=True = 本当に存在しない」の意味が濁る。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  K-1 段階 (デバッグログのファイル出力への移行) ---
'
'   K-1 の変更は ★ログの出し先だけ★ で、ロジックは 1 行も変えていない。
'     EvalSync (9 本) と ClearElementRegistry (2 本) の Debug.Print を
'     Wv2Log.LogE / LogW / LogD に置き換えた。
'
'   Wv2Log は内部で Debug.Print も撃つので、イミディエイトの見え方は変わらない
'   (Wv2Log.LogEcho = False で切れる)。ログは %APPDATA%\Wv2Browser\logs\ に
'   起動ごとの 1 ファイルとして残る。パスは ?Wv2Log.LogPath で分かる。
'
'   ★他のプロシージャの Debug.Print はまだ手つかず★ 段階的に移行する
'   (既存の全 Debug.Print を一度に書き換えない、が K-1 の合意)。
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-2 段階 (要素レジストリ + Wv2Element の読み取り) ---
'
'   D-2 のスコープ (JS 側の要素レジストリと、要素ハンドルの読み取り):
'     GetElementById / QuerySelector … 要素を 1 個探して Wv2Element を返す。
'     ClearElementRegistry           … JS 側レジストリを手動で作り直す。
'     CurrentDomGen                  … レジストリ世代のキャッシュ (高速パス用)。
'     Elem_Read / Elem_CheckStale    … Wv2Element 専用の Friend 経路。
'     ★やらないこと★ 書き込み (Value =) / Click / SetAttribute は D-3。
'                     QuerySelectorAll (複数返し) は D-3 以降 (論点7 案a)。
'
'   ■ 追加した Public メンバ
'     GetElementById(elementId) As Wv2Element
'     QuerySelector(selector)   As Wv2Element
'         見つかれば Wv2Element、見つからなければ ★Nothing★ を返す。
'         ★Nothing の意味は 2 通りあるので LastEvalOk で分ける (論点4)★
'           Nothing + LastEvalOk = True  … 本当に存在しない (正常)
'           Nothing + LastEvalOk = False … 失敗。理由は LastEvalError
'                                           (in-callback / timeout / セレクタの構文エラー等)
'     ClearElementRegistry() As Boolean
'         JS 側のレジストリを捨てて新しい世代で作り直す。既存の Wv2Element は
'         すべて stale になる。同一ページで大量に検索して配列が伸びたときの掃除用。
'     CurrentDomGen As String
'         最後に JS から受け取ったレジストリ世代。まだ一度も触っていなければ空文字。
'
'   ■ JS 側レジストリの形 (論点1 案c: ★注入しない★)
'     window.__wv2 = { gen: 'g<時刻>-<乱数>', els: [要素, 要素, ...] }
'
'     ★AddScriptToExecuteOnDocumentCreated も DOMContentLoaded での注入も使わない。★
'     代わりに★検索式そのものが「無ければ自分で作る」★形にした:
'
'       (function(){ if(typeof window.__wv2==='undefined'||!window.__wv2){
'                      window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]}; }
'                    var R=window.__wv2; ... })()
'
'     利点:
'       ・新しい COM 機構が要らない (前者は完了ハンドラ用の HandlerKind と vtable が要る)
'       ・往復が増えない (判定・生成・検索が 1 つの式に入る)
'       ・ページ遷移 / SPA の書き換え / about:blank のどれで消えても、
'         次のアクセスで自動的に作り直される (設計原則74 と同じ自己修復性)
'       ・タブバー / ナビバー / 設定タブの Pane に無駄な注入をしない
'     欠点:
'       ・ページ読み込み★前★にフックを仕込む用途には使えない。
'         それが要るときは AddScriptToExecuteOnDocumentCreated を独立段階で実装する。
'
'   ■ 世代番号 (論点2 案b: ★JS が発行する★)
'     世代は JS 側が 'g' + Date.now() + '-' + Math.random() で作る文字列。
'     ★VBA 側でナビゲーションを数える方式 (案a) は採らない。★
'       SPA の pushState では DOM が残るのにイベントは飛び、逆にページ内 JS が
'       要素を差し替えてもイベントは飛ばないので、実態とズレるため。
'     レジストリが消える = 世代も消える = 作り直しで必ず別値、という対応が保証され、
'     ★判定の権威は常に JS 側★になる。
'
'     m_domGen は「JS から最後に聞いた世代」のキャッシュで、高速パスにしか使わない:
'       手元の世代 <> m_domGen なら★往復せずに stale と確定できる★ (安全)
'       手元の世代 =  m_domGen でも JS が作り直しているかもしれないので必ず問い合わせる
'     キャッシュが古いときは必ず「問い合わせる」側に倒れるので誤判定は起きない。
'
'     ★Math.random() を混ぜる理由★ Date.now() はミリ秒精度なので、同一ミリ秒内に
'       レジストリが 2 回作り直されると衝突しうる。比較にしか使わない値なので
'       文字列にして衝突の心配ごと消す (論点2 の追加合意)。
'
'   ■ ハンドルの解放 (論点3 案a: ★解放しない★)
'     Wv2Element.Class_Terminate では JS 側スロットを消さない。Terminate が
'     いつ・どの文脈で走るかを制御できず、そこから EvalSync を撃つと D-1 の
'     in-callback ガードや Ctrl_Close 済みの Pane を踏みやすいため。
'     レジストリはページ遷移でまるごと消えるので実害は小さい。手動掃除は
'     ClearElementRegistry。本格的な寿命管理は D-4。
'
'   ■ JS 文字列の作り方 (★VBA 文字列内にバックスラッシュ + 二重引用符を書かない★)
'     セレクタや属性名はページ由来でなくユーザー由来なので、引用符や
'     バックスラッシュが入りうる。JsQuote が Wv2Json.JsonEscape を通して
'     Chr$(34) で括るので、VBA のソース上にその並びが現れない。
'     例: QuerySelector("input[name='q']") も QuerySelector("a[title=""x""]") も通る。
'
'   ■ 事後報告 (設計原則79)
'     (a) ★JsonPickStr を Wv2Json.bas へ移設した★ (Private → Public)
'         論点6 は「JsonUnescape を新設し JsonPickStr もそれを呼ぶ形に書き換える」
'         だったが、Wv2Element も同じ復号を要るため、Wv2Pane に Private で
'         残すと 2 実装になる。移設が論点6 の趣旨に最も忠実な形と判断した。
'         EvalSync 内の 3 箇所の呼び出しを Wv2Json.JsonPickStr に変更した以外、
'         ロジックは 1 文字も変えていない (走査部は Wv2Json.JsonUnescapeAt に同一移植)。
'     (b) ★Wv2Element に LastOk / LastError を足した (論点5 の 8 メンバ + 2)★
'         読み取り系は失敗時も空文字を返すため、「本当に空」と「stale / timeout」を
'         区別する手段が無いと使えない。Pane 側の LastEvalOk は JS 実行自体の成否
'         なので、stale のとき True になってしまい代用できない。
'     (c) ★Elem_Read / Elem_CheckStale の戻り値は Enum でなく Boolean + 文字列にした★
'         クラスモジュールの Public Enum が他モジュールから無修飾で見えるかは
'         この設計で未検証で、コンパイルエラーになると 1 往復無駄になる。
'         D-1 の LastEvalOk / LastEvalError と同じ流儀に揃える方が一貫もする。
'
''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  D-1 段階 (ExecuteScript の同期取得プリミティブ) ---
'
'   ★段階番号の体系が変わりました★
'     第9.32b までの「9.xx」は「WebView2 をブラウザとして仕立てる」という一本の軸の
'     里程標で、第9.32b で一区切りとした。D-1 からは軸ごとに記号を割り当てる:
'       (無印 9.xx) … ブラウザとしての機能整備 (完了・凍結)
'       D          … DOM / ページ自動制御 (MSHTML 相当)  ← 本段階はここ
'       C          … フォームコントロール (TreeView / ListView 相当)
'       K          … 基盤・借金返済 (Sentinel 等)
'
'   D-1 のスコープ (JS 式を投げて戻り値を VBA で同期的に受け取る):
'     EvalSync … 本段階の成果物。JS 式 1 本を撃ち、完了まで DoEvents で待って
'                戻り値の JSON 表現を String で返す。
'     ★やらないこと★ 要素レジストリ (__wv2els[])、Wv2Element、DOM API。D-2 以降。
'
'   ■ 追加した Public メンバ
'     EvalSync(jsExpr, [timeoutSec=5]) As String
'         成功: JS の戻り値を JSON.stringify した文字列を返す。
'               1+2       → 3
'               'abc'     → 引用符ごと abc (★文字列は引用符付きで返る★)
'               undefined → undefined   (JSON.stringify が undefined を返すため)
'               null      → null
'               ({a:1})   → {a:1} の JSON 表現
'         失敗: 空文字を返す。理由は LastEvalError を見る。
'     LastEvalOk    As Boolean … 直前の EvalSync が成功したか
'     LastEvalError As String  … 直前の失敗理由
'                                 closed / no-view / in-callback /
'                                 execute-failed(n) / timeout(n s) /
'                                 hresult(&Hxxxx) / ReferenceError: xxx /
'                                 syntax-error-or-null (...)
'     ResetCallbackGuard()     … in-callback ガードのカウンタを 0 に戻す (緊急用)
'     InCallbackDepth As Long  … ガードカウンタの現在値 (検証用)
'     Debug_SetInCallback(n)   … ガード発火を検証するための注入口 (検証専用)
'
'   ■ 論点と決定 (たーぼーさん合意済み)
'     論点1 メソッド名     : 接頭辞なしの EvalSync。View_ / Ctrl_ は「COM の
'                            インターフェースメソッドを 1 対 1 で叩く」意味なので
'                            待ちループ内蔵の高レベル API には付けない。
'     論点2 結果の受け口   : callbackId をキーにした保留テーブル (m_evalPending)。
'                            単一変数で受けると並行呼び出しで取り違える。
'                            UserForm1.OnExecuteScriptResult 直呼びは残す (ログのみ)。
'     論点3 戻り値の形     : String (= value の JSON 表現)。成否は LastEvalOk /
'                            LastEvalError で取る。型変換は D-2 以降で行う。
'     論点4 Safe 経由      : 常に View_ExecuteScriptSafe を通す。生の ExecuteScript は
'                            JS 例外が S_OK + null で返り、正常な null と区別できない。
'     論点5 タイムアウト   : 既定 5 秒。超過したら保留エントリを★破棄済みで残し★、
'                            遅れて来た完了通知はそれを見て捨てる (ログには残す)。
'     論点6 並行・ネスト   : 許容する。callbackId で分離されているので取り違えない。
'                            ネストは深さをログに出すだけ。
'     論点7 in-callback ガード:
'                            ★本段階で最も価値のある部分★ WebView2 のイベント
'                            ハンドラ内から EvalSync を呼ぶと、ランタイムはハンドラの
'                            復帰を待ち、こちらは JS の完了を待つのでタイムアウトまで
'                            固まる (仕様事実 14)。しかも仕様事実 20 でデバッガが
'                            使えない領域なので原因究明が苦しい。そこで
'                            m_inCallback > 0 なら JS を撃たずに即失敗させる。
'
'   ■ 実装上の事後報告 (設計原則 79)
'     (a) ★ハンドラを Core に分離した★ 合意時は「各 View_On* の入口で +1、出口で -1」
'         と説明したが、実際のコードを読むと View_OnNewWindowRequested は途中に
'         Exit Sub が 5 箇所あり (他も 1 箇所ずつ)、素直に 2 行足すとデクリメントを
'         飛ばす経路ができてしまう。そこで本体を Private ... _Core にリネームし、
'         同名の Public 薄いラッパーが +1 → Core 呼び出し → -1 を行う形にした。
'         対象は 7 個 (NavigationStarting / NavigationCompleted / DocumentTitleChanged /
'         WebMessageReceived / NewWindowRequested / HistoryChanged / DOMContentLoaded)。
'     (b) ★エラー時もカウンタを戻す。ただし握り潰さない★ ラッパーは
'         On Error GoTo GuardExit でカウンタを戻したあと、Err.Raise で★同じエラーを
'         投げ直す★。ComCallbackHandler 側にエラー処理は無く、従来はそのまま
'         伝播していたので、挙動を変えないためにこうした (握り潰すと不具合が
'         見えなくなる)。それでもカウンタが残る事故に備えて ResetCallbackGuard を用意。
'     (c) ★OnExecuteScriptCompleted にはガードを付けていない★ これは結果の受け口で
'         あり、中でユーザーコードが走らないため。付けると EvalSync の DoEvents 中に
'         カウンタが立ち、無害だが紛らわしい。
'     (d) ★JSON の取り出しは Wv2Pane 内の Private 関数にした★ Wv2Json.JsonGetStr は
'         設計原則 64 により★エスケープを解かない★実装で、次の閉じ引用符までを
'         そのまま返す。EvalSync の戻り値は JSON.stringify を二重に通るため
'         「バックスラッシュ + 引用符」や「バックスラッシュ + uXXXX」が必ず入る。
'         そこで JsonPickStr (エスケープを解く限定パーサ) を Private で用意した。
'         他所から要るようになったら Wv2Json へ昇格させる。
'     (e) ★JSON.stringify で包む形にした★ EvalSync は
'         View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")") を撃つ。こうすると
'         Safe が返す構造化 JSON の value が★必ず文字列★になり、(d) の限定パーサ
'         だけで型を問わず取り出せる。undefined のときは value キーごと消えるので
'         それを検出して undefined を返している。
'
'   ■ 利用上の注意 (D-2 以降で緩和予定)
'     ・jsExpr は「文」ではなく「式」でなければならない (View_ExecuteScriptSafe の
'       制約をそのまま継承)。複数文を書きたいときは即時関数で包む。
'     ・文字列を返す式の結果は引用符付きで返る。D-2 で Wv2Element のプロパティごとに
'       型を決めて剥がす。
'     ・★DoEvents ループ中にブレーク/ステップ実行しないこと★ (仕様事実 20)
''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''
' --- Wv2Pane.cls  第9.27 段階 (本丸: put_NewWindow で予備タブに遷移を載せる) ---
'
'   第9.27 のスコープ (POST リンクの新タブ展開を「自前のタブ」で正しく開く):
'     View_OnNewWindowRequested のモード分岐を 2 値 → 3 値に拡張した。
'       0 = NWM_LEGACY  : 現行動作 (put_Handled + RequestNewTabAsync)。Navigate は
'                         ★常に GET★ なので POST リンクは 405 になる。
'       1 = NWM_POPUP   : 9.26 のプローブ。put_Handled を立てず素通りし、ランタイムに
'                         ポップアップを開かせる。POST は保たれる (仕様事実 23)。
'       2 = NWM_PREWARM : ★9.27 の本命★ Browser が温めてある予備タブの CoreWebView2 を
'                         args.put_NewWindow (vtable index 4) に渡す。ランタイムが
'                         その WebView に元のリクエストを載せるので POST が保たれ、
'                         しかも自前のタブに収まる。
'
'   ★イベント内で何も待たない★ ここでやるのは
'       (1) m_browser.ClaimPrewarmViewPtr で予備の ViewPtr をもらう
'       (2) dcf(pArgs, 4, put_NewWindow, pNew)
'       (3) dcf(pArgs, 6, put_Handled, 1&)   ※ SetHandledWithNewWindow が True のとき
'       (4) m_browser.RequestAdoptPrewarmedAsync で採用を Timer に依頼
'     の 4 つだけ。Controller 生成も Deferral も無いので、仕様事実 14 (SafeTimer
'     デッドロック) にも 20 (EBMode の窓) にも触れない。★これがプリウォーム方式の要点★
'
'   ★フォールバック★ 予備が無い / 未 Ready / 既に Claim 済み / put_NewWindow 失敗 の
'     いずれでも、下の現行動作 (put_Handled + RequestNewTabAsync) にそのまま落ちる。
'     POST リンクなら 405 に戻るが壊れはしない。落ちた事実は必ず ★印つきでログに出す。
'
'   ★vtable index の根拠 (WebView2.h の ICoreWebView2NewWindowRequestedEventArgsVtbl)★
'       3 = get_Uri / 4 = put_NewWindow / 5 = get_NewWindow /
'       6 = put_Handled / 7 = get_Handled / 8 = get_IsUserInitiated / 9 = GetDeferral
'     put_Handled = 6 は 9.6b から使っている実績値で、この並びと一致する。
'
'   ★触らないもの★ get_Uri の取得とログ / 空 URL ガード / RequestNewTabAsync (案β
'     Timer、仕様事実 14 の回避) / 他のイベントハンドラ一切。
'
' --- Wv2Pane.cls  第9.26 段階 (NewWindowRequested 委譲プローブ) ---
'
'   第9.26 のスコープ (POST リンクが 405 になる件の切り分け。観測だけを行う回):
'     View_OnNewWindowRequested の冒頭に分岐を 1 個足しただけ。
'     m_browser.PopupDelegateMode が True のときは
'       ・args.put_Handled を立てない
'       ・m_browser.RequestNewTabAsync を呼ばない
'     で素通りし、WebView2 ランタイム自身にポップアップウィンドウを開かせる
'     (第9.6a で確認済みの挙動: Excel プロセス内に Chrome_WidgetWin_1 が生える)。
'     ランタイムが遷移そのものを担当するので、元のリクエストが POST ならボディごと
'     継続される。これで「ランタイムに委譲すれば POST は保たれるのか」を実測する。
'
'     ★なぜ『同一 View で View_Navigate する』案を採らないか★
'       9.6b にあったその形は ICoreWebView2::Navigate を呼ぶので★常に GET★。
'       新タブに流すか同一タブに流すかは関係なく、Navigate を通した時点で 405 のまま。
'       POST を保つ道は「ランタイムに遷移を渡す」しかない (今回の委譲か、次段の
'       put_NewWindow か) というのが切り分けの出発点。
'
'     ★m_browser が Nothing のとき★ 単体 Pane 運用 (SetBrowser 未実行) ではフラグの
'       置き場そのものが無いので、従来どおり同一 View の View_Navigate にフォールバック
'       する。挙動は 9.10e から変更なし。
'
'   ★触らないもの★ get_Uri の取得とログ / 空 URL ガード / put_Handled の失敗時の扱い /
'     RequestNewTabAsync (案β Timer、仕様事実 14 の回避) / 他のイベントハンドラ一切。
'     設計原則・仕様事実の追加は無し。
'
' --- Wv2Pane.cls 第9.19段階 ---
'
'   第9.19 の追加事項 (設定タブの重複防止フラグ):
'     ★m_isSettings フラグ (設定タブ識別) を新設★
'       この Pane が「設定タブ」かどうかを保持する真偽フラグ。Wv2Browser の
'       OpenSettingsTab が、設定タブ生成時にこの Pane に IsSettings = True を
'       立てる。[歯車] が再び押されたとき Browser は m_tabs を頭から走査し、
'       IsSettings = True の Pane があればそれを ActivateTab するだけで新規
'       生成しない ([歯車] 連打で設定タブが増えるのを防ぐ)。
'
'     ★なぜ Pane のフラグで識別するか (案A)★
'       設定タブは NavigateToString で data:text/html;base64,... を表示して
'       いるため URL/内容での判定は脆く高コスト。Pane に印を付けておけば、
'       毎回 m_tabs を走査して探すだけで「生存確認」も兼ねられる (閉じられた
'       設定タブは m_tabs から外れるので自然に見つからなくなる)。Browser 側に
'       Pane 参照を持たせて生存を自前で見張る案 (ダングリング検知が要る) より
'       綻びが少ない。
'
' --- Wv2Pane.cls 第9.18段階 ---
'
'   第9.18 の追加事項 (AddHostObjectToScript による JS への VBA オブジェクト公開):
'     ★View_AddHostObjectToScript(objName, instance) を新設★ (部品1)
'       ICoreWebView2.AddHostObjectToScript (View 上の vtable index 49) を叩き、
'       VBA のクラスインスタンスを JS 側に window.chrome.webview.hostObjects.<name>
'       として公開する。JS からそのメソッド/プロパティを直接呼べるようになる
'       (postMessage の文字列コマンドを介さない直接 RPC 経路)。
'
'     ★実装は旧 WebView2.cls (実機検証済み) の忠実な移植★
'       渡し方は既存メソッドと同じ dcf 作法:
'         Set vObj = instance      (オブジェクトを Variant に詰める)
'         dcf(m_pView, 49, ..., StrPtr(objName), VarPtr(vObj))
'       第1引数はホストオブジェクト名 (LPCWSTR)、第2引数は VARIANT* (詰めた
'       オブジェクトへのポインタ)。戻り値は HRESULT。Instancing は 1 - Private の
'       ままで機能する (旧実装で確認済み。VBA が生成する IDispatch を OS に
'       公開できれば足りる)。
'
'     ★用途 (第9.18 時点)★
'       Wv2Browser が設定タブを開くとき、この Pane に Wv2SettingsBridge を
'       attach し、設定画面の JS から hostObjects.sync.Settings.SetEngine("bing")
'       のように検索エンジンを直接操作させる。attach は NavigateToString の前に
'       行う (順序A: 登録してから Navigate すれば、読み込まれた JS が最初から
'       hostObjects を使える、という仮説を実機検証する。ダメなら ready
'       ハンドシェイク方式=順序B にフォールバックする)。
'
'     ★RemoveHostObjectFromScript (vtable 50) は足さない★ (YAGNI)
'       設定タブは閉じれば WebView2 ごと解放されるため、明示的な取り外しは
'       第9.18 では不要。必要になったステージで足す。
'
' --- Wv2Pane.cls 第9.12a段階 (旧称: Class9) ---
'
'   第9.12a の追加事項 (アドレスバー/ナビボタンの土台: URL 変化の通知):
'     ★構造は 9.11a の NotifyTitleChanged と全く同じ★ Pane が「自分の URL 状態が
'       変わった」ことを所属 Browser に伝え、Browser が index を逆引きして
'       RaiseEvent TabUrlChanged する。Pane はナビバーの存在を知らない (疎結合)。
'     - 新規 Private Sub: NotifyUrlState(pSender)
'         sender 経由で get_Source(4) / get_CanGoBack(38) / get_CanGoForward(39) を
'         取り、m_browser.NotifyUrlChanged Me, url, canBack, canFwd を呼ぶ。
'     - View_OnNavigationCompleted: 末尾で NotifyUrlState を呼ぶ (通常のページ遷移)。
'         ★pArgs=0 の早期 Exit Sub 経路でも呼ぶこと★ (両方の経路を通す)
'     - View_OnHistoryChanged: 既に url/canBack/canFwd を取得済みなので、その値で
'         m_browser.NotifyUrlChanged を呼ぶ (SPA の pushState、戻る/進むを捕捉)。
'     ★多重発火は無害★ NavigationCompleted と HistoryChanged が両方撃っても、
'       全体同期方式 (設計原則 60) なので同じ状態を 2 回送るだけで整合する。
'
'   第9.11c の追加事項 (VBA->JS タブ一覧描画の土台: タイトルキャッシュ + WebMessage シンク):
'     ★タイトルキャッシュ (選択肢3)★ タイトルを Pane オブジェクトに紐づけて
'       保持する。タブ index がずれてもタイトルが正しいペインに追従し、
'       全体同期の真実源を一元化できる。TabBar は Browser.PaneAt(i).DocumentTitle
'       で読む。
'       - 新規フィールド: m_title As String
'       - 新規 Public Property Get DocumentTitle() As String
'       - View_OnDocumentTitleChanged でタイトル取得後 m_title = title を保存
'         (9.11a の NotifyTitleChanged 通知ブロックの直前に 1 行足すだけ)。
'     ★汎用 WebMessage シンク (案P)★ JS から postMessage されたメッセージを
'       登録された「シンク」(TabBar 等) へ転送する仕組み。m_browser 後方参照
'       (NewWindowRequested 用) とは別系統の 1 本。
'       - 新規フィールド: m_webMsgSink As Object (受け側は OnPaneWebMessage を実装)
'       - 新規 Public Sub: SetWebMsgSink / ClearWebMsgSink
'       - View_OnWebMessageReceived: シンクがあれば OnPaneWebMessage(Me, source, msg)
'         を呼ぶブロックを追加。
'       - View_AddWebMessageReceived は既存 (9.4c) をそのまま使い、TabBar 側で
'         受信を有効化する。
'     ★解放時のシンク切断★ ClearBrowserRef で m_webMsgSink も Nothing にする
'       (設計原則 58。後方参照系の切断は Shutdown 経路に集約)。
'
'   第9.11a の追加事項 (疎結合タブバー用のタイトル通知):
'     - View_OnDocumentTitleChanged に 1 ブロック追加。タイトル取得後、
'       m_browser があれば m_browser.NotifyTitleChanged Me, title を呼ぶ。
'       後方参照 m_browser (9.10e で NewWindowRequested 用に導入) をそのまま
'       流用する。m_browser 未設定時 (後方参照なしの使い方) は Debug.Print の
'       みでフォールバックし、疎結合を壊さない。Pane 側の新規フィールド/
'       メソッドはなし (既存の m_browser を使うだけ)。
'
'   第9.10e の追加事項 (NewWindowRequested の新タブ化、案A + 案Y + 案β Timer):
'     ★デッドロック回避 (案β)★ NewWindowRequested ハンドラ内で AddTabWithUrl を
'       直接呼ぶと、AddTab の WaitPaneReady が ControllerCompleted を待つが、本
'       ハンドラが return するまで WebView2 はそれを配送しないためデッドロック
'       (タイムアウト) する (仕様事実 14)。そこで m_browser.RequestNewTabAsync に
'       依頼し、SafeTimer (ワンショット) 経由でハンドラ完了後に AddTabWithUrl を
'       実行させる。
'     - 新規フィールド: m_browser As Wv2Browser (所属 Browser への後方参照)
'     - 新規メソッド: SetBrowser / ClearBrowserRef (後方参照の設定・切断)
'     - View_OnNewWindowRequested を「同一 View Navigate」(9.6b) から
'       「m_browser.AddTabWithUrl で新タブを開く」に変更。m_browser 未設定時は
'       同一 View Navigate にフォールバック。window.open の戻り値は null になる
'       (put_NewWindow で新 View を渡さない簡易方式 = 案A。正攻法は将来の別段階)
'     ★循環参照注意 (設計原則 58)★ Browser.m_tabs ? Pane.m_browser の相互参照に
'       なる。Wv2Browser.Shutdown が各 Pane の ClearBrowserRef を呼んで断つ
'       (UserForm_Terminate 経由)。Class_Terminate では m_browser に触らない
'       (循環時は Class_Terminate 自体が呼ばれないため、Shutdown が唯一の経路)
'
'   第9.10d-2a の追加事項 (本セッション、SPA 用 https origin 基盤):
'     - 新規モジュール変数:
'         m_pView3 As LongPtr (ICoreWebView2_3 派生 IF のキャッシュ)
'     - Class_Initialize に m_pView3 = 0 を追加
'     - Class_Terminate / Ctrl_Close の Release 順序を拡張:
'         m_pView3 → m_pView2 → m_pView → m_pCtrl (最派生から先、設計原則 51)
'     - 新規 Private Function:
'         EnsureView3() As Boolean
'         (QueryInterface(IID_ICoreWebView2_3) で遅延取得、EnsureView2 と同形)
'     - 新規 Public Function:
'         View3_Ensure() As Boolean (検証用ラッパ、View2_Ensure と同形)
'         View3_SetVirtualHostNameToFolderMapping(hostName, folderPath, accessKind)
'           As Long
'           (ICoreWebView2_3 vtable index 71。ローカルフォルダを
'            https://<hostName>/ にマッピングし、about:blank origin の
'            pushState/replaceState 制約 (仕様事実 6) を回避する基盤。
'            単純引数系なので View_Navigate を見本にし、設計原則 45 は不要)
'     - IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'       (EnsureView3 のローカル変数で都度作る、9.9a 案 a を踏襲)
'
'   第9.10c-1 の追加事項 (可視制御):
'     - 新規 Public Function:
'         Ctrl_PutIsVisible(visible As Boolean) As Long
'         (ICoreWebView2Controller::put_IsVisible、vtable index 4)
'         Wv2Browser の ActivateTab がタブ可視制御 (アクティブタブだけ可視) に使う。
'         Win32 BOOL (True=1/False=0) なので dcf 前に 1/0 へ正規化する (仕様事実 11)。
'         単純引数系なので Ctrl_PutBounds を見本にする。
'
'   第9.9c の追加事項 (補完群一括処理):
'     - 新規 Public 関数: View_ExecuteScriptSafe(userCode) As Long
'         ユーザーコードを try-catch で包んで構造化 JSON を返す ExecuteScript ラッパ。
'         resultJson が必ず {ok:true,value:...} or {ok:false,errorName:...,...}
'         の形式になる (案 F: a + X + ★)。
'     - OnExecuteScriptCompleted から UserForm1.OnExecuteScriptResult を直接呼ぶ
'         (案 α 第二実装、案 E1-a-i = Wv2Pane は UserForm1 参照を保持しない。
'          VBA のグローバル UserForm 参照を活用、循環参照なし)。
'         Wv2Pane 内の Debug.Print も併存 (案 E2-X、リグレッション切り分け容易)。
'     - View_OnHistoryChanged のログ改善:
'         (α) about:blank なら「NavigateToString 環境では正常」と注記。
'         (β) pArgs=0 が仕様であることをコメント・ログで明示。
'     - 設計上ロジックの本質変更はなし (拡張のみ)。
'
'   第9.9b の追加事項 (継続、SPA 対応):
'     - 永続ハンドラ配列の上限を HK_NewWindowRequested(7) → HK_DOMContentLoaded(10)
'       に拡張
'     - 新規ヘルパー RemoveInterfaceFor(kind) As LongPtr
'     - HistoryChanged 系 (基底 ICoreWebView2、vtable 13/14):
'         View_AddHistoryChanged / View_RemoveHistoryChanged / View_OnHistoryChanged
'     - DOMContentLoaded 系 (ICoreWebView2_2、vtable 64/65):
'         View2_AddDOMContentLoaded / View2_RemoveDOMContentLoaded /
'         View2_OnDOMContentLoaded
'
'   ★ 第9.9b の設計原則 45 への警戒 ★
'     永続ハンドラを 2 個追加 = COM オブジェクトを WebView2 に渡す場面が 2 回。
'     9.8c で h.Slot を直接渡してクラッシュした教訓に基づき、
'     View_AddHistoryChanged / View2_AddDOMContentLoaded は両方とも
'     View_AddNavigationStarting を逐語的に見本にして実装し、
'     pVTableObj = h.Slot + VTABLE_OBJ_OFFSET を必ず通す。
'     View_PostWebMessageAsString のような単純引数系は見本にしない。
'
'   第9.9a の追加事項 (継続):
'     - 新規モジュール変数:
'         m_pView2 As LongPtr (ICoreWebView2_2 派生 IF のキャッシュ)
'     - Class_Initialize に m_pView2 = 0 を追加
'     - Class_Terminate の Release 順序を拡張:
'         (Settings → View2 → View → Ctrl、子から親へ厳格、設計原則 43)
'     - Ctrl_Close にも同じ Release 順序を適用 (m_pView2 を View より先に Release)
'     - 新規 Private Function:
'         EnsureView2() As Boolean
'         (m_pView2 が未取得なら QueryInterface(IID_ICoreWebView2_2) で取得、
'          EnsureSettings と同形パターン)
'     - 新規 Public Function:
'         View2_Ensure() As Boolean
'         (検証用ラッパ、Immediate ウィンドウから叩いて取得成功確認)
'
'   第9.9a の設計判断 (たーぼーさん合意済み):
'     ・案 1: m_pView2 フィールド + EnsureView2 遅延取得 (Settings 流儀を踏襲)
'     ・案 a: IID_ICoreWebView2_2 は EnsureView2 のローカル変数で都度作る
'             (Wv2Thunks に IID テーブル追加せず、9.9a を Wv2Thunks 最小変更に)
'     ・案 X: EnsureView2 / View2_Ensure とも As Boolean (EnsureSettings と同形)
'     ・案 i: 検証は Immediate で View2_Ensure 叩く + Test_AllEvents_Real で
'             リグレッションゼロ確認 (テスト関数は新設しない)
'     ・案 R1: Release 順序は Settings → View2 → View → Ctrl (子から親へ厳格)
'
'   ★ 第9.9a の意義 ★
'     ICoreWebView2_2 を取り出すこと自体が目的ではなく、第9.9b 以降の
'     SPA 対応 (DOMContentLoaded など、vtable 64/65 にあるメソッド) を
'     呼び出すための基盤を整える段階。HistoryChanged は基底 IF (vtable 13/14)
'     にあるので _2 が無くても 9.9b で実装可能だが、DOMContentLoaded は
'     _2 IF が必須。第 9.6a で確立した命名規約 (View2_xxx) を初めて
'     実地で適用する段階でもある。
'
'   第9.9a で踏襲する設計原則 (関連する主要なもの):
'     ・設計原則 31: 出自接頭辞による命名規約 (View2_xxx)
'     ・設計原則 42: 派生 IF の遅延取得・キャッシュ方式 (EnsureSettings の流儀)
'     ・設計原則 43: Release 順序は子から親へ
'     ・設計原則 45: COM オブジェクトとして渡すときは VTABLE_OBJ_OFFSET を足す
'                    (9.9a では「WebView2 から COM オブジェクトを受け取る」場面
'                     なので VTABLE_OBJ_OFFSET は登場しないが、9.9b で
'                     DOMContentLoaded ハンドラを登録するときに再警戒する)
'     ・設計原則 48: Close 後の安全性は「メンバクリア + フラグ」ハイブリッド
'
'   第9.8c の追加事項 (継続):
'     - 新規モジュール変数:
'         m_executeScriptCounter As Long (View_ExecuteScript の callback ID 採番)
'         m_closed As Boolean (Ctrl_Close 済みフラグ)
'     - Class_Initialize に上記 2 個の初期化追加
'     - 新規 Public 関数 (View_ 系):
'         View_ExecuteScript (jsCode As String) As Long  ' callback ID 返却
'     - 新規 Public Sub (動的バインディング受け取り口):
'         OnExecuteScriptCompleted (callbackId, errorCode, resultJson)
'     - Ctrl_Close の拡張 (案 E 1+3 ハイブリッド):
'         Controller::Close を撃つ + m_pSettings/m_pView/m_pCtrl の Release +
'         m_persistentTokens / m_persistentHandlers のクリア +
'         m_closed = True + m_state = Cs_Failed
'         これにより Close 後の Class_Terminate での二重処理が綺麗にスキップされる
'
'   ★ 第9.8c での初回バグ + 修正履歴 ★
'     View_ExecuteScript の初版では h.Slot を直接 dcf に渡してクラッシュ
'     (Excel プロセス即落ち)。実際は他の登録系 (View_AddNavigationStarting 等)
'     と同じく h.Slot + VTABLE_OBJ_OFFSET を渡す必要がある。
'     → 「サンクスロットを WebView2 に COM オブジェクトとして渡すとき、
'        必ず VTABLE_OBJ_OFFSET (= 176 byte) を足してから渡す」
'        という設計原則を改めて確認 (設計原則 45 として新設)。
'
'   第9.8c の設計判断 (たーぼーさん合意済み):
'     ・案 a: 同時並行 OK、スロット自身が完了管理
'     ・案 α: イベントコールバック方式、第一実装は OnExecuteScriptCompleted で
'             ログ出力のみ (将来 UserForm 通知に拡張可)
'     ・案 X: 連番 callback ID 返却
'     ・案 E 1+3: メンバクリア + m_closed フラグ
'
'   第9.8b の事項 (継続):
'     - m_pSettings + EnsureSettings + Settings_PutXxx 系 (3 個)
'     - View_Reload / View_Stop / View_GetCanGoBack / View_GetCanGoForward /
'       View_GetBrowserProcessId
'
'   第9.8a の事項 (継続、改名):
'     - クラス名 Class9 → Wv2Pane に正式固定
'     - Class7 → ComCallbackHandler / Class8 → Wv2Environment への参照置換
'     - 詳細は開発メモ.md の各段階セクションを参照

'   役割:
'     ICoreWebView2Controller と ICoreWebView2 の合体クラス (合体方針 = 案 a)。
'     SDK の階層構造との 1 対 1 対応は諦めるが、VBA ユーザ視点で
'     「1 つの WebView2 表示 = 1 クラス」という素直な対応が得られる。
'     UserForm の Frame コントロール領域に貼り付ける「1 枚の表示パネル」と
'     して捉えるのが直観に近い (= 「Pane」と命名した理由)。
'
'   実装済み永続ハンドラ:
'     HK_NavigationStarting      (View_AddNavigationStarting   / View_OnNavigationStarting)
'     HK_NavigationCompleted     (View_AddNavigationCompleted  / View_OnNavigationCompleted)
'     HK_WebMessageReceived      (View_AddWebMessageReceived   / View_OnWebMessageReceived)
'     HK_DocumentTitleChanged    (View_AddDocumentTitleChanged / View_OnDocumentTitleChanged)
'     HK_NewWindowRequested      (View_AddNewWindowRequested   / View_OnNewWindowRequested)
'     HK_HistoryChanged          (View_AddHistoryChanged       / View_OnHistoryChanged)        ' 第9.9b
'     HK_DOMContentLoaded        (View2_AddDOMContentLoaded    / View2_OnDOMContentLoaded)     ' 第9.9b
'
'   命名規約 (設計原則 31):
'     ・ICoreWebView2Controller のメソッド: Ctrl_xxx (基底) / Ctrl2_xxx (拡張、将来)
'     ・ICoreWebView2 のメソッド:           View_xxx (基底) / View2_xxx (拡張、第9.9a で初導入)
'     ・ICoreWebView2 のイベントレシーバー: View_OnXxx (基底) / View2_OnXxx (拡張、9.9b で初導入予定)
'     ・ICoreWebView2Environment のメソッド: Env_xxx 等 (将来必要時に追加)
'     ・1 ショット系コールバック (OnEnvironmentCompleted / OnControllerCompleted):
'       接頭辞なしで据え置き (派生インターフェース版が直接取れる関数がないため
'       バージョン識別の必要がない)
'
'   双方向通信フロー (第9.4c で確立、参照名のみ新名に更新):
'     [VBA → JS]
'       VBA: m_pane.View_PostWebMessageAsString("hello")
'         → dcf(m_pView, 33, ..., StrPtr("hello"))
'         → WebView2 ランタイムが JS に届ける
'         → JS: window.chrome.webview.addEventListener('message', ...) で受信
'     [JS → VBA]
'       JS: window.chrome.webview.postMessage("hello")
'         → WebView2 ランタイムが HK_WebMessageReceived 経由でハンドラ呼び出し
'         → ComCallbackHandler.Invoke[HK_WebMessageReceived]
'         → Wv2Pane.View_OnWebMessageReceived(pSender, pArgs)
'         → args.get_Source / args.TryGetWebMessageAsString でログ
''''''''''''''''''''''''''''''''''
Option Explicit

#If Win64 Then
    ' OK
#Else
    #Error "このクラスは 64 ビット VBA (x64) が必要です"
#End If

' --- N-1: 経過ミリ秒の時刻源 ---
'   ★Timer は使わない★ Timer は Single の「午前 0 時からの秒」なので、
'   午後の値 (5 万秒台) では Single の有効桁が足りず ms の分解能が消える。
'   さらに日付を跨ぐと巻き戻る。GetTickCount は起動からの ms で単調増加。
'   戻り値は Long なので約 49.7 日で一周する。差は Double に上げてから計算し、
'   負になったら 2^32 を足して戻す (NetElapsedMs)。
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long


Public Enum ControllerState
    Cs_Created = 0
    Cs_Waiting = 1
    Cs_Ready = 2
    Cs_Failed = 3
End Enum


' --- 内部状態 ---
' m_handler:     HK_ControllerCompleted ハンドラへの強参照 (1 ショット)
'                (OnControllerCompleted 後に Nothing にする)
' m_pCtrl:       WebView2 から受け取った ICoreWebView2Controller
'                (Class_Terminate で Release する責任を持つ)
' m_pView:       Ctrl_GetCoreWebView2 で取得した ICoreWebView2
'                (Class_Terminate で Release する責任を持つ)
Private m_handler As ComCallbackHandler
Private m_pCtrl   As LongPtr
Private m_pView   As LongPtr
'   第9.8b 段階で追加: ICoreWebView2Settings のキャッシュ
'     Wv2Pane が初めて Settings 系メソッドを叩かれたときに
'     View vtable 3 (get_Settings) で取得し、ここに保持する。
'     m_pView の寿命と一致するので Class_Terminate で Release する。
Private m_pSettings As LongPtr
'   第9.9a 段階で追加: ICoreWebView2_2 派生 IF のキャッシュ
'     Wv2Pane が初めて View2_xxx 系メソッドを叩かれたときに
'     QueryInterface(IID_ICoreWebView2_2) で取得し、ここに保持する。
'     EnsureView2 経由で遅延取得 (= EnsureSettings と同じ流儀、設計原則 42)。
'
'     m_pView と m_pView2 は同一 COM オブジェクトを別 IF で見たもの。
'     QueryInterface は別途 AddRef を撃つので、m_pView2 は m_pView とは
'     独立した参照カウントを持つ。Class_Terminate / Ctrl_Close では
'     m_pView より先に m_pView2 を Release する (子から親へ、設計原則 43)。
Private m_pView2 As LongPtr
'   第9.10d-2a 段階で追加: ICoreWebView2_3 派生 IF のキャッシュ
'     SetVirtualHostNameToFolderMapping (vtable 71) を叩くために必要。
'     EnsureView3 経由で遅延取得 (EnsureView2 と同形、設計原則 42/50)。
'     ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 の単一継承なので、
'     Release 順序では m_pView3 → m_pView2 → m_pView (最派生から先、設計原則 51)。
Private m_pView3 As LongPtr
'   第9.10e 段階で追加: 所属する Wv2Browser への後方参照
'     View_OnNewWindowRequested が新タブを開くために Wv2Browser.AddTabWithUrl を
'     呼べるようにする。Wv2Browser.AddTab が SetBrowser で自分をセットする。
'     ★循環参照 (Browser.m_tabs ? Pane.m_browser) になるので、解放は
'      Wv2Browser.Shutdown → ClearBrowserRef 経由で明示的に断つ (設計原則 58)。
Private m_browser As Wv2Browser
Private m_state   As ControllerState
Private m_lastError As Long

' --- 永続イベントハンドラの保持 (第9.4a 段階で追加、第9.6a/9.9b で上限拡張) ---
' m_persistentHandlers: HandlerKind の値域を index とした ComCallbackHandler 配列
' m_persistentTokens:   同 index の EventRegistrationToken (LongLong, 8 bytes)
'
' 第9.6a 段階で上限を HK_DocumentTitleChanged → HK_NewWindowRequested に拡張。
' 第9.9b 段階で上限を HK_NewWindowRequested(7) → HK_DOMContentLoaded(10) に拡張。
' N-1 段階で上限を HK_DOMContentLoaded(10) → HK_WebResourceRequested(11) に拡張。
'   これにより HK_HistoryChanged(9) / HK_DOMContentLoaded(10) も永続配列で
'   管理できる。
'
' 注意: 範囲 3?10 の中に HK_ExecuteScriptCompleted(8) が含まれるが、
'   ExecuteScript は 1 ショット系で同時並行のため永続配列には保持しない
'   (= index 8 は常に Nothing / 0 のまま未使用)。CleanupAllPersistentHandlers の
'   「If Not (m_persistentHandlers(k) Is Nothing)」判定で自然にスキップされる
'   ので実害はない。
'
' Class_Terminate ですべての非 Nothing なハンドラについて remove_ を呼ぶ
' (CleanupAllPersistentHandlers の LBound/UBound ループは変更不要、
'  これが「HandlerKind 連動配列で横展開を自動化」の本領発揮)。
Private m_persistentHandlers(HK_NavigationStarting To HK_WebResourceRequested) As ComCallbackHandler
Private m_persistentTokens(HK_NavigationStarting To HK_WebResourceRequested) As LongLong

' --- ページ読み込み完了フラグ (第9.4c 段階で追加、暫定実装) ---
'
'   View_OnNavigationCompleted で True にセットされる。
'   外部 (UserForm 等) が「ページ読み込み完了を待ってから PostMessage を撃つ」
'   のような同期に使う。
'
'   注意 (暫定実装):
'     これは「初回ナビゲーションの完了」しか意味しない。SPA のように
'     history.pushState や fetch でページ内容が更新される場合は捕捉できない。
'     SPA 対応 (DOMContentLoaded、HistoryChanged 等) は包括クラス導入時に
'     まとめて再設計予定。
Private m_pageLoaded As Boolean


' 第9.8c で追加:
'   m_executeScriptCounter: View_ExecuteScript の呼び出しごとに採番する
'     連番カウンタ (案 X 採用)。各 ComCallbackHandler.m_callbackId に格納される。
'     初期値 0、最初の呼び出しで 1 になる。
'   m_closed: Ctrl_Close が呼ばれたか、または Class_Terminate が走った後を示す。
'     True の間は View_xxx / Ctrl_xxx / Settings_xxx の冒頭で早期 return。
'     これにより Close 後に誤って各種メソッドを叩いても安全に弾く。
'     (案 E 1+3 ハイブリッド: メンバクリア + m_closed フラグ)
Private m_executeScriptCounter As Long
Private m_closed As Boolean

' --- 第9.11c 段階で追加 ---
'   m_title: このペインが表示しているドキュメントのタイトルキャッシュ
'     (選択肢3)。View_OnDocumentTitleChanged で更新される。タイトルを Pane
'     オブジェクトに紐づけることで、タブ index がずれてもタイトルが正しい
'     ペインに追従する (全体同期の真実源を一元化)。TabBar は
'     Browser.PaneAt(i).DocumentTitle で読む。
'   m_webMsgSink: JS から postMessage されたメッセージの転送先 (案P)。
'     TabBar が自分をシンクに登録する。受け側は OnPaneWebMessage を実装する
'     こと (遅延バインドなので As Object)。m_browser (NewWindowRequested 用の
'     後方参照) とは別系統。解放は ClearBrowserRef で断つ (設計原則 58)。
Private m_title As String
Private m_webMsgSink As Object
Private m_isSettings As Boolean   ' 第9.19 (この Pane が設定タブかどうか)

' --- 第9.27: NewWindowRequested の捌き方 (Wv2Browser.NewWindowMode と対応) ---
'   クラスモジュールでは Public Const が書けないので、Browser 側と同じ値を
'   Private Const で持つ。値の意味はヘッダーのコメントを参照。
Private Const NWM_LEGACY  As Long = 0
Private Const NWM_POPUP   As Long = 1
Private Const NWM_PREWARM As Long = 2


' --- D-1: EvalSync (同期取得プリミティブ) 用のモジュール変数と定数 ---
'   m_evalPending : callbackId をキーにした保留テーブル (論点2)。
'                   キーは cb & callbackId、値は Array(状態, errorCode, resultJson)。
'                   ★Collection の項目は書き換えられないので、更新は Remove → Add★
'   m_lastEvalOk / m_lastEvalError : 直前の EvalSync の成否と理由 (論点3)。
'   m_inCallback  : WebView2 のイベントハンドラ内にいる深さ (論点7 のガード)。
'                   Public ラッパーが +1 / -1 する。0 でなければ EvalSync は撃たない。
'   m_evalDepth   : EvalSync のネスト深さ。ログに出すだけで抑止はしない (論点6)。
Private m_evalPending   As Collection
Private m_lastEvalOk    As Boolean
Private m_lastEvalError As String
Private m_inCallback    As Long
Private m_evalDepth     As Long

' 保留テーブルの状態値
Private Const EVAL_WAITING   As Long = 0
Private Const EVAL_DONE      As Long = 1
Private Const EVAL_DISCARDED As Long = 2

' タイムアウトの既定値 (秒)。READY_TIMEOUT (Controller 生成待ち 10 秒) より短くする。
' JS 式 1 本の実行は本来ミリ秒オーダーなので、5 秒待つ時点で異常とみなす。
Private Const EVAL_TIMEOUT_DEFAULT As Single = 5

' --- D-3: DOM 条件待ちの設定 (論点2 案C / 論点4 案A) ---
'   m_implicitWaitSec : QuerySelector / GetElementById に掛ける暗黙待ち (秒)。
'                       ★既定 0 = 無効★ ImplicitWaitSec プロパティで設定する。
'   POLL_INTERVAL_SEC : ポーリング間隔 (論点4 案A の固定 100ms)。
'                       ★EvalSync 1 往復が実測 38ms なので、実際の判定間隔は
'                       約 140ms になる★ (2026-08-22 の実機ログから)。
'   WAIT_TIMEOUT_DEFAULT : WaitFor / WaitGone の既定の上限 (秒)。
Private m_implicitWaitSec As Single
Private Const POLL_INTERVAL_SEC    As Single = 0.1
Private Const WAIT_TIMEOUT_DEFAULT As Single = 10


' --- D-4a: ページ内 SPA プローブの版番号 ---
'   プローブの中身を変えたら★必ず上げる★。ページに古い版が残っていても、
'   版が違えば健康診断が落ちて作り直される (参考プロダクトと同じ契約)。
Private Const SPA_PROBE_VER As Long = 3


' --- D-6: QuerySelectorAll ---
'   ALL_MAX_DEFAULT    : 1 回に返す既定の上限。★黙って切らない★ (LastAllTruncated)
'   ELEMENT_WARN_COUNT : レジストリがこれを超えたら警告する (捨てはしない)
Private Const ALL_MAX_DEFAULT    As Long = 200
Private Const ELEMENT_WARN_COUNT As Long = 5000
Private m_allTruncated As Boolean


' --- D-4b: 静穏待ちの設定 ---
'   m_ignoreNet / m_ignoreSel : ★VBA 側が持つ★ 除外リスト (ページ側は遷移で消えるため)。
'   m_autoWait   : 操作の後に自動で WaitSettled するか。★既定 False★
'   m_lastSettle : 直前の待ちの内訳 (診断用)。
Private m_ignoreNet  As Collection
Private m_ignoreSel  As String
Private m_autoWait   As Boolean
Private m_lastSettle As String

' 静穏窓の既定 (ms)。参考プロダクト (WebDriver BiDi for SeleniumVBA) と同じ値。
Private Const SETTLE_STABLE_MS_DEFAULT As Long = 500

' 居座り判定の既定 (ms)。★10 秒を超えて飛んでいる要求は数えない★
Private Const STALE_INFLIGHT_MS_DEFAULT As Long = 10000


' --- D-4c: 明示シグナル (arm) ---
'   ★VBA 側も arm 済みかを覚える★ ページ側だけに持たせると、遷移でプローブごと
'   消えたときに「arm したのに消えた」ことに気づけない (黙って条件なしの待ちに
'   化けてしまう)。両方で持ち、食い違ったら signal-lost として失敗させる。
Private m_sigDom As Boolean
Private m_sigNet As Boolean
Private m_sigUrl As Boolean

' --- D-4f: 居座る要求を静穏判定から外す ---
'   ★ロングポーリングやストリーミングはどのサイトにもある★
'   Google Maps の検索は /search?tbm=map を開きっぱなしにして結果を流すので、
'   これを数えていると永久に静穏が来ない (D-4d で実測)。
'   一定時間を超えて飛んでいる要求は「流れっぱなしのもの」とみなして数えない。
'   0 にすると無効 (すべて数える)。
Private m_staleMs As Long


' --- N-2: リクエストの詳細 (ヘッダとボディ) ---
'   ★一覧とは別のリング★ (論点5 案S)。一覧は「順番と全体像」、
'   詳細は「これを再現したい 1 本」で用途が違う。
'
'   m_detOn        : 詳細を取るか。★既定 False★ (論点8 案κ)
'   m_detCap       : 詳細リングの容量 (既定 50)
'   m_detHead      : 次に書く位置 (0 起点)
'   m_detCount     : 今リングに入っている件数
'   m_detDropped   : ★溢れて捨てた件数★ 一覧と同じく黙って切らない
'   m_detBodyMax   : ボディを読む上限 (バイト。既定 65536)
'   m_detRedact    : 秘密のヘッダを伏せるか。★既定 True★ (論点3 案M)
'   m_detRedactList: 伏せる対象の名前を |小文字| で並べたもの
'   m_detUriFilter : 詳細を取る URL の部分一致 (空なら全部)
Private m_detOn         As Boolean
Private m_detCap        As Long
Private m_detHead       As Long
Private m_detCount      As Long
Private m_detDropped    As Long
Private m_detBodyMax    As Long
Private m_detRedact     As Boolean
Private m_detRedactList As String
Private m_detUriFilter  As String
Private m_detSeq()     As Long
Private m_detMethod()  As String
Private m_detUri()     As String
Private m_detHeaders() As String
Private m_detCType()   As String
Private m_detBytes()   As Long
Private m_detTrunc()   As Boolean
Private m_detIsText()  As Boolean
Private m_detBody()    As String
Private m_detHex()     As String

' 詳細リングの既定容量と上限。★一覧より桁で小さい★ 1 件あたりの重さが違う。
Private Const NET_DET_CAP_DEFAULT As Long = 50
Private Const NET_DET_CAP_MAX     As Long = 2000

' ボディの上限 (既定 64KB) と、その上限自体の上限 (8MB)。
'   ★ハンドラの中で同期的に読む★ ので青天井にしない。
Private Const NET_BODY_MAX_DEFAULT As Long = 65536
Private Const NET_BODY_MAX_LIMIT   As Long = 8388608
Private Const NET_BODY_CHUNK       As Long = 16384

' ヘッダを読む個数の上限。1 要求で COM 呼び出しが 3 回 x 個数ぶん走る。
Private Const NET_HDR_MAX As Long = 40

' バイナリのボディを見せるときの先頭バイト数。
Private Const NET_HEX_HEAD As Long = 32

' IStream::Seek の dwOrigin (Windows SDK objidl.h の STREAM_SEEK_SET)。
Private Const STREAM_SEEK_SET As Long = 0


' --- N-1: ネットワーク要求のキャプチャ (WebResourceRequested) ---
'   m_netOn      : 捕捉中か (add_WebResourceRequested 済みか)
'   m_netCap     : リングの容量 (件)。NetLogCapacity で変えられる。
'   m_netHead    : 次に書く位置 (0 起点)
'   m_netCount   : 今リングに入っている件数 (0 ～ m_netCap)
'   m_netTotal   : 捕捉開始からの★総発火数★ (溢れた分も数える)
'   m_netDropped : ★溢れて捨てた件数★ 黙って切らない
'   m_netT0      : 捕捉開始時の GetTickCount 値。各件の経過 ms の基準。
'   m_netFilters : 登録済みフィルタ (uri & vbTab & ctx)。Stop で外すために覚える。
Private m_netOn      As Boolean
Private m_netCap     As Long
Private m_netHead    As Long
Private m_netCount   As Long
Private m_netTotal   As Long
Private m_netDropped As Long
Private m_netT0      As Long
Private m_netFilters As Collection
Private m_netSeq()    As Long
Private m_netMs()     As Long
Private m_netCtx()    As Long
Private m_netMethod() As String
Private m_netUri()    As String

' リングの既定容量と上限。既定 500 件は「1 ページ分の操作」を丸ごと収める目安。
Private Const NET_CAP_DEFAULT As Long = 500
Private Const NET_CAP_MAX     As Long = 50000

' 既定で捕まえる種別 (論点4)。画像・CSS・フォント・スクリプトは除外する。
'   1 = DOCUMENT (フォーム POST / ページ遷移) / 7 = XML_HTTP_REQUEST / 8 = FETCH
Private Const NET_CTX_DEFAULT As String = "DOCUMENT,XHR,FETCH"


' --- D-2: 要素レジストリの世代キャッシュ ---
'   m_domGen : JS から最後に受け取った window.__wv2.gen。高速パス専用。
'              まだ一度も要素検索をしていなければ空文字。
'              ★判定の権威ではない★ (権威は常に JS 側。上のヘッダー参照)
Private m_domGen As String


' ============================================================
' Class_Initialize / Class_Terminate
' ============================================================
Private Sub Class_Initialize()
    m_staleMs = STALE_INFLIGHT_MS_DEFAULT
    m_state = Cs_Created
    m_pCtrl = 0
    m_pView = 0
    m_pView2 = 0    ' 第9.9a で追加 (派生 IF のキャッシュ、未取得状態)
    m_pView3 = 0    ' 第9.10d-2a で追加 (ICoreWebView2_3 派生 IF、未取得状態)
    Set m_browser = Nothing   ' 第9.10e (後方参照、AddTab 時に SetBrowser でセット)
    m_lastError = 0
    m_pageLoaded = False
    m_executeScriptCounter = 0
    m_closed = False
    m_title = ""                  ' 第9.11c (タイトルキャッシュ)
    Set m_webMsgSink = Nothing    ' 第9.11c (WebMessage シンク、未登録)
    m_isSettings = False          ' 第9.19 (既定は通常タブ)

    ' --- D-1: EvalSync 関連の初期化 ---
    Set m_evalPending = New Collection
    m_lastEvalOk = False
    m_lastEvalError = ""
    m_inCallback = 0
    m_evalDepth = 0

    ' --- D-2: 要素レジストリの世代キャッシュ ---
    m_domGen = ""

    ' --- N-1: ネットワーク要求のリングバッファ ---
    m_netOn = False
    m_netCap = NET_CAP_DEFAULT
    Set m_netFilters = New Collection
    NetRingAlloc

    ' --- N-2: 詳細 (ヘッダとボディ) ---
    m_detOn = False                       ' ★既定 OFF★ (論点8 案κ)
    m_detCap = NET_DET_CAP_DEFAULT
    m_detBodyMax = NET_BODY_MAX_DEFAULT
    m_detRedact = True                    ' ★既定は伏せる★ (論点3 案M)
    m_detRedactList = NetRedactDefaults()
    m_detUriFilter = ""
    NetDetailRingAlloc
End Sub

Private Sub Class_Terminate()
    ' COM 所有権規約: 受け取ったポインタは ComAddRef で VBA 側の所有権分を
    ' 確保しているので、ここで Release する。
    '
    ' 順序 (第9.8b で m_pSettings を追加、第9.9a で m_pView2 を追加):
    '   1. 永続イベントハンドラの解除 (m_pView がまだ生きている必要がある)
    '   2. m_pSettings の Release (Settings は View の子オブジェクト)
    '   3. m_pView2 の Release  ← 第9.9a で追加
    '                            (View2 は View から QI で取り出した派生 IF、
    '                             独立した参照カウントを持つので View より先)
    '   4. m_pView の Release
    '   5. m_pCtrl の Release
    '
    ' なぜこの順序か:
    '   - remove_*Event は m_pView (= ICoreWebView2) のメソッドなので
    '     m_pView が生きている間に撃つ必要がある
    '   - Settings は View が生成した子オブジェクト。View より先に Release する
    '   - View2 は View を QueryInterface で取り出した派生 IF。同一 COM
    '     オブジェクトを別 IF で見ているだけだが、QI は独立した AddRef を
    '     撃つため、View2 と View はそれぞれ独立した参照カウントを持つ。
    '     設計原則 43 (子から親へ) に従い View2 を先に Release する
    '   - View は Ctrl が生成した子オブジェクト。Controller 側を先に
    '     Release しても WebView2 内部の参照カウント管理は破綻しないが、
    '     寿命の依存関係 (Ctrl が親、View が子、Settings が孫) を
    '     VBA レベルでも明示するため、子から先に解放する。

    ' --- K-4a: ★所要時間を測るための入口★ ---
    '   ここから次の行 (Cleanup の 1 本目) までが remove_*Event の所要になる。
    Wv2Log.LogD "Wv2Pane.Terminate: ★開始★ (pCtrl=" & m_pCtrl & _
                " pView=" & m_pView & ")"

    ' --- 1. 永続イベントハンドラを一括解除 (第9.4a 段階で追加) ---
    CleanupAllPersistentHandlers
    Wv2Log.LogD "Wv2Pane.Terminate: 永続ハンドラの解除おわり"

    ' --- 2. m_pSettings の Release (第9.8b 段階で追加) ---
    If m_pSettings <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If

    ' --- 3-pre. m_pView3 の Release (第9.10d-2a 段階で追加) ---
    '   View3 は View を QI で取り出した最派生 IF (ICoreWebView2_3)。
    '   派生から順に解放するため、m_pView2 より先に Release (設計原則 51)。
    If m_pView3 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If

    ' --- 3. m_pView2 の Release (第9.9a 段階で追加) ---
    '   View2 は View を QueryInterface で取り出した派生 IF。View より先に Release。
    If m_pView2 <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If

    ' --- 4. m_pView の Release ---
    If m_pView <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If

    ' --- 5. m_pCtrl の Release ---
    If m_pCtrl <> 0 Then
        Wv2Log.LogD "Wv2Pane.Terminate: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' m_handler は通常 OnControllerCompleted で Nothing になっているはずだが、
    ' 万一 Init で例外が起きた等で残っている場合はここで切る。
    If Not (m_handler Is Nothing) Then
        Wv2Log.LogD "Wv2Pane.Terminate: m_handler が残っているので切る (異常終了パス)"
        Set m_handler = Nothing
    End If

    Wv2Log.LogD "Wv2Pane.Terminate: ★完了★"   ' K-4a
End Sub


' ============================================================
' CleanupAllPersistentHandlers (第9.4a 段階で追加、内部ヘルパー)
'
'   m_persistentHandlers に登録されているすべての永続ハンドラについて、
'   WebView2 側に解除依頼を出す。Class_Terminate から呼ばれる。
'
'   各 HandlerKind について、対応する remove_ vtable index を使い分ける。
'   解除に成功すれば WebView2 が ComCallbackHandler を Release → refcount=0 →
'   自動解放経路で ComCallbackHandler.ClearOwner → m_persistentHandlers(k) は
'   ClearOwner が走った時点で Nothing にしておくのが綺麗。
'
'   ただし「m_persistentHandlers(k) を先に Nothing にすると、
'   ClearOwner が触る m_owner が既に Nothing になっている」という
'   問題はない (ClearOwner は ComCallbackHandler.m_owner を Nothing にするだけで、
'   Wv2Pane 側の m_persistentHandlers には触らない)。
'
'   順序の整理:
'     1. View_RemoveEvent (remove_*Event) を WebView2 に依頼
'     2. WebView2 が ComCallbackHandler を Release → 基盤側の自動解放経路 →
'        ComCallbackHandler.ClearOwner で ComCallbackHandler.m_owner = Nothing
'     3. ここで Wv2Pane 側の m_persistentHandlers(k) = Nothing にする
'
'   第9.6a で配列上限が HK_NewWindowRequested に拡張されたが、
'   ループ自体は LBound/UBound を使っているのでコード変更不要
'   (= 設計原則 15 「HandlerKind 連動ループで横展開を自動化」の本領発揮)。
' ============================================================
Private Sub CleanupAllPersistentHandlers()
    If m_pView = 0 Then
        ' View が既に解放されている場合、remove_ を撃つ手段がない。
        ' 永続ハンドラはまだ生きているかもしれないが、何もできない (リーク容認)。
        Dim k1 As Long
        For k1 = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
            Set m_persistentHandlers(k1) = Nothing
            m_persistentTokens(k1) = 0^
        Next k1
        Exit Sub
    End If

    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        If Not (m_persistentHandlers(k) Is Nothing) Then
            Dim removeIndex As Long
            removeIndex = RemoveVtableIndexFor(CLng(k))
            If removeIndex > 0 Then
                ' 第9.9b: remove を撃つべきインターフェースを kind から決める。
                '   DOMContentLoaded(_2 IF) は m_pView2、それ以外は m_pView。
                '   設計原則 45 に忠実 (派生 IF のメソッドは派生 IF のポインタで叩く)。
                Dim pIfRemove As LongPtr
                pIfRemove = RemoveInterfaceFor(CLng(k))
                If pIfRemove <> 0 Then
                    Wv2Log.LogD "Wv2Pane.Cleanup: removing persistent handler kind=" & k & _
                                " vtable=" & removeIndex & " token=" & m_persistentTokens(k) & _
                                " iface=" & pIfRemove
                    dcf pIfRemove, removeIndex, "View.remove_Event", m_persistentTokens(k)
                Else
                    ' 解除に必要なインターフェース (例: m_pView2) が未取得 / 解放済み。
                    ' remove を撃つ手段がないので参照だけ切る (リーク容認)。
                    Wv2Log.LogD "Wv2Pane.Cleanup: kind=" & k & " の remove 用 IF が 0、" & _
                                "remove スキップ (参照のみ解放)"
                End If
            End If
            Set m_persistentHandlers(k) = Nothing
            m_persistentTokens(k) = 0^
        End If
    Next k
End Sub


' ============================================================
' RemoveInterfaceFor (第9.9b 段階で追加、内部ヘルパー)
'
'   HandlerKind に対応する remove_*Event を撃つべきインターフェースの
'   ポインタを返す。RemoveVtableIndexFor が「vtable index」を返すのに対し、
'   こちらは「どのインターフェース (m_pView / m_pView2 / ...) で叩くか」を返す。
'
'   設計原則 45 (派生 IF のメソッドは派生 IF のポインタで叩く):
'     DOMContentLoaded は ICoreWebView2_2 のイベント (vtable 65) なので
'     m_pView2 で remove する。それ以外は基底 ICoreWebView2 のイベントなので
'     m_pView で remove する。
'
'     9.9a の観察で「m_pView2 == m_pView (同じ値)」と判明しているが、
'     それに依拠して m_pView で vtable 65 を叩くのは「派生 IF のメソッドを
'     基底ポインタで叩く」危険な近道 (= 9.8c でクラッシュを招いた思考と同根)。
'     意味的に正しいポインタを使うことで、将来の SDK 変更にも耐える。
'
'   戻り値:
'     対応するインターフェースのポインタ。
'     0 を返した場合は「remove を撃つ手段がない」(例: DOMContentLoaded だが
'     m_pView2 が未取得 / 既に解放済み) ので、呼び出し側は remove をスキップする。
'
'   将来の拡張:
'     ICoreWebView2_3 以降のイベントが増えたら、ここに Case を 1 行足すだけで
'     横展開できる (RemoveVtableIndexFor との対称性を保つ)。
' ============================================================
Private Function RemoveInterfaceFor(ByVal kind As Long) As LongPtr
    Select Case kind
        Case HK_DOMContentLoaded
            ' _2 IF のイベント。m_pView2 で叩く (未取得なら 0 を返す)。
            RemoveInterfaceFor = m_pView2
        Case Else
            ' 基底 ICoreWebView2 のイベント。m_pView で叩く。
            RemoveInterfaceFor = m_pView
    End Select
End Function


' ============================================================
' RemoveVtableIndexFor (第9.4a 段階で追加、内部ヘルパー)
'
'   HandlerKind に対応する remove_*Event の vtable index を返す。
'   add_*Event とペアになっており、SDK で対応関係が決められている。
'
'   ICoreWebView2 vtable layout (基底):
'     0  QueryInterface
'     1  AddRef
'     2  Release
'     3  get_Settings
'     ...
'     5  Navigate
'     ...
'     7  add_NavigationStarting       8  remove_NavigationStarting
'     ...
'    13  add_HistoryChanged          14  remove_HistoryChanged       ← 第9.9b で追加
'    15  add_NavigationCompleted     16  remove_NavigationCompleted
'     ...
'    34  add_WebMessageReceived      35  remove_WebMessageReceived
'     ...
'    44  add_NewWindowRequested      45  remove_NewWindowRequested  ← 第9.6a で追加
'    46  add_DocumentTitleChanged    47  remove_DocumentTitleChanged
'    55  add_WebResourceRequested    56  remove_WebResourceRequested  ← N-1 で追加
'
'   ICoreWebView2_2 vtable layout (基底の 61 個に続く派生 IF):
'    64  add_DOMContentLoaded        65  remove_DOMContentLoaded     ← 第9.9b で追加
'       ※ ただし add/remove の vtable index は m_pView2 (ICoreWebView2_2) 上の
'         オフセット。RemoveInterfaceFor が DOMContentLoaded には m_pView2 を返す。
'
'   第9.6a で HK_NewWindowRequested = 45 を追加。
'   第9.9b で HK_HistoryChanged = 14、HK_DOMContentLoaded = 65 を追加。
' ============================================================
Private Function RemoveVtableIndexFor(ByVal kind As Long) As Long
    Select Case kind
        Case HK_NavigationStarting:   RemoveVtableIndexFor = 8
        Case HK_NavigationCompleted:  RemoveVtableIndexFor = 16
        Case HK_WebMessageReceived:   RemoveVtableIndexFor = 35
        Case HK_DocumentTitleChanged: RemoveVtableIndexFor = 47
        Case HK_NewWindowRequested:   RemoveVtableIndexFor = 45   ' 第9.6a で追加
        Case HK_HistoryChanged:       RemoveVtableIndexFor = 14   ' 第9.9b で追加 (基底 IF)
        Case HK_DOMContentLoaded:     RemoveVtableIndexFor = 65   ' 第9.9b で追加 (_2 IF)
        Case HK_WebResourceRequested: RemoveVtableIndexFor = 56   ' N-1 で追加 (基底 IF)
        Case Else:                    RemoveVtableIndexFor = 0   ' 未知 → 解除不可
    End Select
End Function


' ============================================================
' Init
'
'   ICoreWebView2Controller の作成を開始する。
'   このメソッドはコールバックを待たずに即座にリターンする (非同期)。
'   完了通知は OnControllerCompleted で受ける。
'
'   呼び出し側は ctrl.IsReady が True になるまで DoEvents で待つ。
'   ただし WebView2Loader は同期コールバックする可能性があり、その場合は
'   Init から戻った時点で既に Ready / Failed になっている。
' ============================================================
Public Sub Init(ByVal env As Wv2Environment, ByVal hwnd As LongPtr)

    If m_state <> Cs_Created Then
        MsgBox "Wv2Pane.Init は二重呼び出し不可", vbCritical
        Exit Sub
    End If

    If env Is Nothing Then
        MsgBox "Wv2Pane.Init: env が Nothing", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    If Not env.IsReady Then
        MsgBox "Wv2Pane.Init: env がまだ Ready でない (state=" & env.state & ")", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    Dim pEnv As LongPtr
    pEnv = env.EnvironmentPtr
    If pEnv = 0 Then
        MsgBox "Wv2Pane.Init: env.EnvironmentPtr が 0", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 1. ハンドラ (ComCallbackHandler) を 1 個 Acquire ---
    Set m_handler = AcquireHandlerFor(HK_ControllerCompleted, Me)
    If m_handler Is Nothing Then
        MsgBox "Wv2Pane.Init: ハンドラの Acquire に失敗", vbCritical
        m_state = Cs_Failed
        Exit Sub
    End If

    ' --- 2. WebView2 が握る obj アドレスを計算 ---
    Dim pVTableObj As LongPtr
    pVTableObj = m_handler.Slot + VTABLE_OBJ_OFFSET

    ' --- 3. m_state = Cs_Waiting を API 呼び出しの「前」に立てる ---
    '   (第9.3a で確立されたパターン: 同期コールバック対応)
    m_state = Cs_Waiting

    ' --- 4. Environment->CreateCoreWebView2Controller(hwnd, handler) を呼ぶ ---
    '
    '   ICoreWebView2Environment vtable レイアウト:
    '     0 : QueryInterface
    '     1 : AddRef
    '     2 : Release
    '     3 : CreateCoreWebView2Controller(parentWindow HWND, handler ICoreWebView2CreateCoreWebView2ControllerCompletedHandler*)
    '     4 : ... (以降は SDK バージョンによる)
    '
    '   引数:
    '     hwnd       : Frame コントロールの HWND
    '     pVTableObj : ComCallbackHandler ハンドラの vtableObj アドレス
    Debug.Print "Wv2Pane.Init: dcf(pEnv=" & pEnv & ", 3, ""Env.CreateController"", hwnd=" & _
                hwnd & ", handler=" & pVTableObj & ")"

    Dim hr As Long
    hr = dcf(pEnv, 3, "Env.CreateCoreWebView2Controller", hwnd, pVTableObj)

    If hr <> 0 Then
        MsgBox "CreateCoreWebView2Controller failed. HRESULT: &H" & Hex(hr), vbCritical
        m_state = Cs_Failed
        m_lastError = hr
        ' 失敗時のリーク対策は第9.3a / 9.3b では未実装 (Wv2Environment と同じ方針)。
        ' Thunks_Shutdown で全スロット強制クリーンアップに頼る。
        Exit Sub
    End If
End Sub


' ============================================================
' OnControllerCompleted
'
'   ComCallbackHandler.Handler_Invoke の HK_ControllerCompleted 分岐から
'   m_owner.OnControllerCompleted が動的バインディングで呼ばれる。
'
'   Wv2Environment.OnEnvironmentCompleted と完全に同形。Public Sub であることが必須
'   (Friend だと IDispatch に乗らない)。
'
'   注 (第9.6a 命名規約): 1 ショット系コールバックは出自接頭辞なしの命名を維持。
'   派生 IF (ICoreWebView2Environment2 等) を直接受け取る関数が存在しない
'   ためバージョン識別の必要がない (たーぼーさん確認済み、第9.6a 段階)。
'
'   引数:
'     errorCode   : HRESULT (S_OK = 0 で成功)
'     pController : ICoreWebView2Controller ポインタ
'                   (WebView2 が一時的に AddRef した状態で渡される。
'                    保持する場合は VBA 側で AddRef 必須。)
' ============================================================
Public Sub OnControllerCompleted( _
    ByVal errorCode As Long, _
    ByVal pController As LongPtr)

    Debug.Print "Wv2Pane.OnControllerCompleted: errorCode=" & errorCode & _
                ", pController=" & pController

    m_lastError = errorCode

    If errorCode = 0 And pController <> 0 Then
        ' VBA 側で保持する分の参照を確保
        m_pCtrl = pController
        ComAddRef m_pCtrl
        m_state = Cs_Ready
    Else
        m_pCtrl = 0
        m_state = Cs_Failed
    End If

    ' 1 ショットハンドラパターン (案 D1-α): ハンドラへの参照を切る
    Set m_handler = Nothing
End Sub


' ============================================================
' プロパティ群
' ============================================================
Public Property Get state() As ControllerState
    state = m_state
End Property

Public Property Get IsReady() As Boolean
    IsReady = (m_state = Cs_Ready)
End Property

Public Property Get IsFailed() As Boolean
    IsFailed = (m_state = Cs_Failed)
End Property

Public Property Get LastError() As Long
    LastError = m_lastError
End Property

' --- 生ポインタの公開 (内部観察 / デバッグ用) ---
Public Property Get CtrlPtr() As LongPtr
    CtrlPtr = m_pCtrl
End Property

Public Property Get ViewPtr() As LongPtr
    ViewPtr = m_pView
End Property

' --- IsPageLoaded (第9.4c で追加) ---
'   View_OnNavigationCompleted で True にセットされる。
'   外部 (UserForm 等) からのページ読み込み完了待ちの同期に使う。
'   将来 SPA 対応する際は包括クラス側に移管予定。
Public Property Get IsPageLoaded() As Boolean
    IsPageLoaded = m_pageLoaded
End Property

' --- DocumentTitle (第9.11c で追加、タイトルキャッシュの読み取り口) ---
'   View_OnDocumentTitleChanged が保存した最新タイトルを返す。
'   TabBar の PushSyncToJs が Browser.PaneAt(i).DocumentTitle で読む。
'   まだタイトル未取得のペインは空文字を返す。
Public Property Get DocumentTitle() As String
    DocumentTitle = m_title
End Property


' --- IsSettings (第9.19 で追加、設定タブ識別フラグ) ---
'   この Pane が設定タブとして開かれたものかどうか。Wv2Browser.OpenSettingsTab
'   が生成時に True を立て、[歯車] 再クリック時の重複チェックで参照する。
Public Property Get IsSettings() As Boolean
    IsSettings = m_isSettings
End Property

Public Property Let IsSettings(ByVal value As Boolean)
    m_isSettings = value
End Property


' ============================================================
' Ctrl_PutBounds (ICoreWebView2Controller::put_Bounds、vtable index 6)
'
'   Frame コントロール等の HWND クライアント領域に対応する RECT を渡して、
'   WebView2 の表示矩形を設定する。
'
'   引数の bounds は GetClientRect で取った RECT。座標系は親 HWND の
'   クライアント領域内で (0, 0) 始まり。
'
'   x64 ABI 上は 16 bytes 構造体は値渡しだが、旧版から継承して
'   VarPtr(rc) でポインタ渡しすると WebView2 が正しく解釈する
'   (実機検証済み)。dcf には LongPtr として渡せばよい。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function Ctrl_PutBounds(ByRef bounds As RECT) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutBounds = &H80004003   ' E_POINTER
        Exit Function
    End If
    Ctrl_PutBounds = dcf(m_pCtrl, 6, "Ctrl.put_Bounds", VarPtr(bounds))
End Function


' ============================================================
' Ctrl_PutIsVisible (ICoreWebView2Controller::put_IsVisible、vtable index 4)
'   第9.10c-1 段階で追加。
'
'   Controller の可視/不可視を切り替える。Wv2Browser のタブ切替
'   (ActivateTab) で「アクティブタブだけ可視、他は不可視」を実現するために使う。
'
'   引数: visible (True = 可視 / False = 不可視)
'   戻り値: HRESULT (S_OK = 0 で成功)
'
'   実装メモ:
'     put_IsVisible は BOOL (整数) を 1 個受ける単純引数メソッド。
'     COM オブジェクトを渡すハンドラ系ではないので、見本は Ctrl_PutBounds の
'     ような単純引数系 (dcf で直接呼ぶ) で正しい。
'     VARIANT_BOOL / Win32 BOOL の区別: ICoreWebView2Controller::put_IsVisible の
'     引数は Win32 BOOL (= 4 byte int、True=1/False=0)。VBA の True は -1 なので、
'     dcf に渡す前に 1 / 0 へ正規化する (仕様事実 11)。
' ============================================================
Public Function Ctrl_PutIsVisible(ByVal visible As Boolean) As Long
    If m_pCtrl = 0 Then
        Ctrl_PutIsVisible = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim b As Long
    If visible Then b = 1 Else b = 0
    Ctrl_PutIsVisible = dcf(m_pCtrl, 4, "Ctrl.put_IsVisible", b)
End Function


' ============================================================
' Ctrl_GetCoreWebView2 (ICoreWebView2Controller::get_CoreWebView2、vtable index 25)
'
'   ICoreWebView2 ポインタを out 引数で取得し、m_pView に格納する。
'
'   COM 規約: get_xxx 系メソッドは取得したポインタを既に AddRef した
'   状態で返す (= 受け取る側で再 AddRef は不要、Release のみ必要)。
'   m_pView は Class_Terminate で Release される。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
'
'   注意: このメソッドは Init チェーン内で 1 回だけ呼ばれる前提。
'         m_pView が既に非ゼロの場合、二重取得を防ぐため
'         先に既存ポインタを Release する (本来 Init 時にしか呼ばれない
'         はずだが、安全のため)。
' ============================================================
Public Function Ctrl_GetCoreWebView2() As Long
    If m_pCtrl = 0 Then
        Ctrl_GetCoreWebView2 = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' 既存の m_pView があれば先に Release (二重取得防止)
    If m_pView <> 0 Then
        ComRelease m_pView
        m_pView = 0
    End If

    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pCtrl, 25, "Ctrl.get_CoreWebView2", VarPtr(pv))
    If hr = 0 Then
        m_pView = pv
    End If
    Ctrl_GetCoreWebView2 = hr
End Function


' ============================================================
' View_Navigate (ICoreWebView2::Navigate、vtable index 5)
'
'   指定 URL へ遷移する。Navigate 自体は同期的に true を返すだけで、
'   実際のページ読み込み完了は HK_NavigationCompleted で受ける。
'
'   引数の URL は VBA String (= UTF-16 LE) で、StrPtr で LPCWSTR を取り出す。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View_Navigate(ByVal url As String) As Long
    If m_pView = 0 Then
        View_Navigate = &H80004003   ' E_POINTER
        Exit Function
    End If
    ' 新しいナビゲーション開始でフラグをリセット
    m_pageLoaded = False
    View_Navigate = dcf(m_pView, 5, "View.Navigate", StrPtr(url))
End Function


' ============================================================
' View_AddNavigationStarting (第9.4a 段階で新規、第9.6a で View_ 接頭辞付きに改名)
'
'   ICoreWebView2::add_NavigationStarting (vtable index 7) を呼んで
'   永続イベントハンドラを登録する。
'
'   流れ:
'     1. ComCallbackHandler (HK_NavigationStarting) を Acquire (owner = Me)
'     2. dcf(m_pView, 7, handler, &token) で WebView2 に登録依頼
'     3. WebView2 が AddRef を撃つ (ComCallbackHandler の refcount が上がる)
'     4. 取得した token を m_persistentTokens(HK_NavigationStarting) に保存
'     5. ComCallbackHandler を m_persistentHandlers(HK_NavigationStarting) に保存
'
'   呼び出し後は Bing 上のリンククリック等で
'   ComCallbackHandler.Invoke[HK_NavigationStarting] → Wv2Pane.View_OnNavigationStarting
'   の経路でイベントを受け取れる。
'
'   既に登録済みなら何もしない (二重登録防止)。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View_AddNavigationStarting() As Long
    If m_pView = 0 Then
        View_AddNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' 既に登録済みなら何もしない
    If Not (m_persistentHandlers(HK_NavigationStarting) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: 既に登録済みなのでスキップ"
        View_AddNavigationStarting = 0
        Exit Function
    End If

    ' ComCallbackHandler を Acquire (owner = Me で動的バインディング経由でイベント受信)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationStarting, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: ハンドラの Acquire に失敗"
        View_AddNavigationStarting = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_NavigationStarting を撃つ (vtable index 7)
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 7, "View.add_NavigationStarting", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationStarting: add_NavigationStarting 失敗 hr=&H" & Hex(hr)
        ' ComCallbackHandler をリリースする (refcount=0 で自動解放経路)
        ' h はローカル変数として参照しているので Set h = Nothing するだけでは
        ' 基盤側の m_handlers(idx) からの参照が残る。
        ' WebView2 が AddRef を撃てていない可能性があるため、ComCallbackHandler のスロットは
        ' 第9.5 段階のセンチネル機構が対応するためここでは個別対処不要。
        Set h = Nothing
        View_AddNavigationStarting = hr
        Exit Function
    End If

    ' 成功: token と ComCallbackHandler を保存
    m_persistentTokens(HK_NavigationStarting) = token
    Set m_persistentHandlers(HK_NavigationStarting) = h
    Debug.Print "Wv2Pane.View_AddNavigationStarting OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationStarting = 0
End Function


' ============================================================
' View_RemoveNavigationStarting (第9.4a 段階で新規、第9.6a で改名)
'
'   登録済みの NavigationStarting ハンドラを解除する。
'   ユーザが明示的に解除したい場合に呼ぶ。Class_Terminate からは
'   CleanupAllPersistentHandlers が一括で行うので、これを直接
'   呼ぶ必要はない。
'
'   流れ:
'     1. View_RemoveEvent (vtable index 8) で WebView2 側から解除
'     2. WebView2 が ComCallbackHandler を Release → 基盤の自動解放経路
'     3. Wv2Pane 側の m_persistentHandlers(k) = Nothing, token = 0
'
'   戻り値: HRESULT (S_OK = 0 で成功、未登録なら S_OK で何もしない)
' ============================================================
Public Function View_RemoveNavigationStarting() As Long
    If m_pView = 0 Then
        View_RemoveNavigationStarting = &H80004003   ' E_POINTER
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationStarting) Is Nothing Then
        View_RemoveNavigationStarting = 0   ' 未登録は no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 8, "View.remove_NavigationStarting", _
             m_persistentTokens(HK_NavigationStarting))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationStarting OK"
    End If

    ' 成功でも失敗でも Wv2Pane 側の参照は切る
    Set m_persistentHandlers(HK_NavigationStarting) = Nothing
    m_persistentTokens(HK_NavigationStarting) = 0^

    View_RemoveNavigationStarting = hr
End Function


' ============================================================
' View_OnNavigationStarting (第9.4a 段階で新規、第9.6a で改名)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   ComCallbackHandler.Handler_Invoke の HK_NavigationStarting 分岐から
'   m_owner.View_OnNavigationStarting が動的バインディングで呼ばれる。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
'
'   引数:
'     pSender : WebView2 視点の Invoke 第 1 引数 = ICoreWebView2 ポインタ
'               (= m_pView と同じはずだが、念のため別変数で受ける)
'     pArgs   : WebView2 視点の Invoke 第 2 引数 =
'               ICoreWebView2NavigationStartingEventArgs ポインタ
'
'   動作:
'     args.get_Uri (vtable index 3) を呼んで URL をログに出力する。
'     get_Uri は LPWSTR* を out 引数で返すので、GetStringProperty が
'     PtrToString + CoTaskMemFree までを一括で処理する。
'
'   args の参照カウント:
'     WebView2 が args を一時的に AddRef した状態で Invoke を呼んでいる。
'     コールバックから戻った後 WebView2 が Release するため、
'     View_OnNavigationStarting 内で args を保持する必要がなければ AddRef 不要。
'     URL を取得してログに出すだけなので、保持しない。
' ============================================================
Public Sub View_OnNavigationStarting(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnNavigationStarting_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationStarting_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationStarting: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationStarting_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationStarting: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs が 0 なので URL 取得をスキップ"
        Exit Sub
    End If

    ' args.get_Uri を呼んで URL を取得
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NavStartingArgs.get_Uri")
    Debug.Print "  → 遷移先 URL: " & url
End Sub


' ============================================================
' View_AddNavigationCompleted (第9.4b 段階で新規、第9.6a で改名)
'
'   ICoreWebView2::add_NavigationCompleted (vtable index 15) を呼んで
'   永続イベントハンドラを登録する。
'
'   View_AddNavigationStarting と同形パターン。違いは:
'     - 種別: HK_NavigationCompleted
'     - vtable index: 15 (add) / 16 (remove)
'     - args 型: ICoreWebView2NavigationCompletedEventArgs
'                (get_IsSuccess, get_WebErrorStatus, get_NavigationId を持つ)
' ============================================================
Public Function View_AddNavigationCompleted() As Long
    If m_pView = 0 Then
        View_AddNavigationCompleted = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NavigationCompleted) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: 既に登録済みなのでスキップ"
        View_AddNavigationCompleted = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NavigationCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: ハンドラの Acquire に失敗"
        View_AddNavigationCompleted = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 15, "View.add_NavigationCompleted", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNavigationCompleted: add_NavigationCompleted 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddNavigationCompleted = hr
        Exit Function
    End If

    m_persistentTokens(HK_NavigationCompleted) = token
    Set m_persistentHandlers(HK_NavigationCompleted) = h
    Debug.Print "Wv2Pane.View_AddNavigationCompleted OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNavigationCompleted = 0
End Function


' ============================================================
' View_RemoveNavigationCompleted (第9.4b 段階で新規、第9.6a で改名)
'   View_RemoveEvent (vtable index 16) で解除。
'   未登録なら no-op (S_OK で返す)。
' ============================================================
Public Function View_RemoveNavigationCompleted() As Long
    If m_pView = 0 Then
        View_RemoveNavigationCompleted = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NavigationCompleted) Is Nothing Then
        View_RemoveNavigationCompleted = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 16, "View.remove_NavigationCompleted", _
             m_persistentTokens(HK_NavigationCompleted))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNavigationCompleted OK"
    End If

    Set m_persistentHandlers(HK_NavigationCompleted) = Nothing
    m_persistentTokens(HK_NavigationCompleted) = 0^

    View_RemoveNavigationCompleted = hr
End Function


' ============================================================
' View_OnNavigationCompleted (第9.4b 段階で新規、第9.6a で改名)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   args = ICoreWebView2NavigationCompletedEventArgs の get_IsSuccess
'   (vtable index 3) を呼んで成功/失敗をログ出力。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
' ============================================================
Public Sub View_OnNavigationCompleted(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnNavigationCompleted_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNavigationCompleted_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNavigationCompleted: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNavigationCompleted_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNavigationCompleted: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs が 0 なので IsSuccess 取得をスキップ"
        ' pArgs が無くてもページロード完了の通知としては有効
        m_pageLoaded = True
        ' --- 第9.12a: この経路でも URL 状態を通知する (早期 return に埋もれさせない) ---
        NotifyUrlState pSender
        Exit Sub
    End If

    ' args.get_IsSuccess を呼んで成否を取得
    Dim isOk As Boolean
    isOk = GetBoolProperty(pArgs, 3, "NavCompletedArgs.get_IsSuccess")
    Debug.Print "  → isSuccess: " & isOk

    ' ページ読み込み完了フラグを立てる (第9.4c で追加)
    ' 暫定実装: 成功/失敗にかかわらず NavigationCompleted が来たら True にする。
    ' これは初回ロード待ちの同期用。SPA 等での追加遷移を区別したい場合は
    ' 包括クラス側で別途設計する。
    m_pageLoaded = True

    ' --- 第9.12a: 所属 Browser に URL 状態を通知 (アドレスバー更新用) ---
    NotifyUrlState pSender
End Sub


' ============================================================
' NotifyUrlState (第9.12a、URL 状態を Browser へ通知)
'
'   sender (ICoreWebView2) から現在の URL / 戻れるか / 進めるか を読み取り、
'   所属 Browser の NotifyUrlChanged に渡す。Browser がオブジェクト参照の一致で
'   index を逆引きして RaiseEvent TabUrlChanged する (9.11a の NotifyTitleChanged
'   と全く同じ構造)。
'
'   ★Pane はナビバーの存在を知らない★ ただ「自分の状態が変わった」と
'   Browser に言うだけ。聞き手 (ナビバー) がいるかどうかは関知しない (疎結合)。
'
'   呼び出し元:
'     - View_OnNavigationCompleted : 通常のページ遷移の完了
'     - View_OnHistoryChanged      : SPA の pushState / 戻る / 進む
'   両方から撃つが、全体同期方式 (設計原則 60) なので多重発火は無害。
' ============================================================
Private Sub NotifyUrlState(ByVal pSender As LongPtr)
    If m_browser Is Nothing Then Exit Sub
    If pSender = 0 Then Exit Sub

    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    m_browser.NotifyUrlChanged Me, url, canBack, canFwd
End Sub


' ============================================================
' View_AddDocumentTitleChanged (第9.4b 段階で新規、第9.6a で改名)
'
'   ICoreWebView2::add_DocumentTitleChanged (vtable index 46) を呼んで
'   永続イベントハンドラを登録する。
'
'   args は IUnknown* (情報なし)。タイトルは sender.get_DocumentTitle
'   (vtable 48) を撃って取得する。
' ============================================================
Public Function View_AddDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_AddDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: 既に登録済みなのでスキップ"
        View_AddDocumentTitleChanged = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DocumentTitleChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: ハンドラの Acquire に失敗"
        View_AddDocumentTitleChanged = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 46, "View.add_DocumentTitleChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddDocumentTitleChanged: add_DocumentTitleChanged 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddDocumentTitleChanged = hr
        Exit Function
    End If

    m_persistentTokens(HK_DocumentTitleChanged) = token
    Set m_persistentHandlers(HK_DocumentTitleChanged) = h
    Debug.Print "Wv2Pane.View_AddDocumentTitleChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddDocumentTitleChanged = 0
End Function


' ============================================================
' View_RemoveDocumentTitleChanged (第9.4b 段階で新規、第9.6a で改名)
'   View_RemoveEvent (vtable index 47) で解除。
' ============================================================
Public Function View_RemoveDocumentTitleChanged() As Long
    If m_pView = 0 Then
        View_RemoveDocumentTitleChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DocumentTitleChanged) Is Nothing Then
        View_RemoveDocumentTitleChanged = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 47, "View.remove_DocumentTitleChanged", _
             m_persistentTokens(HK_DocumentTitleChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveDocumentTitleChanged OK"
    End If

    Set m_persistentHandlers(HK_DocumentTitleChanged) = Nothing
    m_persistentTokens(HK_DocumentTitleChanged) = 0^

    View_RemoveDocumentTitleChanged = hr
End Function


' ============================================================
' View_OnDocumentTitleChanged (第9.4b 段階で新規、第9.6a で改名)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   sender = ICoreWebView2 の get_DocumentTitle (vtable index 48) を呼んで
'   現在のタイトルをログ出力する。
'
'   args は IUnknown* で情報がないので、sender 経由で取りに行く必要がある。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
' ============================================================
Public Sub View_OnDocumentTitleChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnDocumentTitleChanged_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnDocumentTitleChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnDocumentTitleChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnDocumentTitleChanged: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' sender (ICoreWebView2).get_DocumentTitle (vtable index 48) を呼んでタイトル取得
    Dim title As String
    title = GetStringProperty(pSender, 48, "View.get_DocumentTitle")
    Debug.Print "  → タイトル: " & title

    ' --- 第9.11c: タイトルをこのペインにキャッシュ (選択肢3、真実源の一元化) ---
    m_title = title

    ' --- 第9.11a: 所属 Browser にタイトル変更を通知 (後方参照 m_browser 経由) ---
    '   NewWindowRequested で実績のある後方参照をそのまま流用する。
    '   m_browser 未設定時 (後方参照なしの使い方) は Debug.Print のみで
    '   フォールバックし、疎結合を壊さない。
    If Not m_browser Is Nothing Then
        m_browser.NotifyTitleChanged Me, title
    End If
End Sub


' ============================================================
' View_AddWebMessageReceived (第9.4c 段階で新規、第9.6a で改名)
'
'   ICoreWebView2::add_WebMessageReceived (vtable index 34) を呼んで
'   永続イベントハンドラを登録する。
'
'   JS 側から window.chrome.webview.postMessage(text) で送信されたメッセージを
'   受信する。受信時に Wv2Pane.View_OnWebMessageReceived が呼ばれる。
' ============================================================
Public Function View_AddWebMessageReceived() As Long
    If m_pView = 0 Then
        View_AddWebMessageReceived = &H80004003
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_WebMessageReceived) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: 既に登録済みなのでスキップ"
        View_AddWebMessageReceived = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebMessageReceived, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: ハンドラの Acquire に失敗"
        View_AddWebMessageReceived = &H8000FFFF
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 34, "View.add_WebMessageReceived", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebMessageReceived: add_WebMessageReceived 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddWebMessageReceived = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebMessageReceived) = token
    Set m_persistentHandlers(HK_WebMessageReceived) = h
    Debug.Print "Wv2Pane.View_AddWebMessageReceived OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddWebMessageReceived = 0
End Function


' ============================================================
' View_RemoveWebMessageReceived (第9.4c 段階で新規、第9.6a で改名)
'   View_RemoveEvent (vtable index 35) で解除。
' ============================================================
Public Function View_RemoveWebMessageReceived() As Long
    If m_pView = 0 Then
        View_RemoveWebMessageReceived = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_WebMessageReceived) Is Nothing Then
        View_RemoveWebMessageReceived = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 35, "View.remove_WebMessageReceived", _
             m_persistentTokens(HK_WebMessageReceived))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveWebMessageReceived OK"
    End If

    Set m_persistentHandlers(HK_WebMessageReceived) = Nothing
    m_persistentTokens(HK_WebMessageReceived) = 0^

    View_RemoveWebMessageReceived = hr
End Function


' ============================================================
' View_OnWebMessageReceived (第9.4c 段階で新規、第9.6a で改名)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   JS が window.chrome.webview.postMessage(text) で送信したメッセージを
'   受信する。Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
'
'   args = ICoreWebView2WebMessageReceivedEventArgs の vtable index:
'     3 : get_Source                 (送信元 URL、本実装で使用)
'     4 : get_WebMessageAsJson       (未使用、必要なら追加)
'     5 : TryGetWebMessageAsString   (JS が string で送ったときに成功、本実装で使用)
'
'   TryGetWebMessageAsString は HRESULT を返すが、GetStringProperty は
'   HRESULT に関係なく成功時のみ文字列を返す (失敗時は空文字)。
'   JS が postMessage(jsonObj) でオブジェクト送信した場合はここが空文字に
'   なる (その場合は get_WebMessageAsJson で取る)。
' ============================================================
Public Sub View_OnWebMessageReceived(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnWebMessageReceived_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebMessageReceived_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnWebMessageReceived: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnWebMessageReceived_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnWebMessageReceived: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs が 0 なのでメッセージ取得をスキップ"
        Exit Sub
    End If

    ' args.get_Source を呼んで送信元 URL を取得
    Dim source As String
    source = GetStringProperty(pArgs, 3, "WebMsgArgs.get_Source")
    Debug.Print "  → source: " & source

    ' args.TryGetWebMessageAsString を呼んでメッセージ本文を取得
    ' JS が文字列で送ったとき成功、オブジェクトで送ったとき空文字
    Dim msg As String
    msg = GetStringProperty(pArgs, 5, "WebMsgArgs.TryGetWebMessageAsString")
    Debug.Print "  → message: " & msg

    ' --- 第9.11c: 登録されたシンク (TabBar 等) へメッセージを転送 (案P) ---
    '   シンクは OnPaneWebMessage(pane, source, msg) を実装していること。
    '   遅延バインド呼び出しなので、実装がなければ実行時エラーになるが、
    '   シンク登録側 (TabBar) の責務として正しいメソッドを備える。
    If Not m_webMsgSink Is Nothing Then
        m_webMsgSink.OnPaneWebMessage Me, source, msg
    End If
End Sub


' ============================================================
' View_AddNewWindowRequested (第9.6a 段階で新規)
'
'   ICoreWebView2::add_NewWindowRequested (vtable index 44) を呼んで
'   永続イベントハンドラを登録する。
'
'   ページ内の「新しいウィンドウを開く」要求 (target="_blank" リンク、
'   window.open、Ctrl+クリック等) で発火する。
'
'   第9.6a 段階の動作:
'     URL をログに出すだけ。put_Handled を立てないため、ブラウザは
'     デフォルト挙動 (= OS の別ウィンドウ、または OS のデフォルトブラウザで
'     開く) で振る舞う。Bing のリンクは現状この経路で「VBA からは開けない」
'     状態が続く (たーぼーさん検証済みの従来挙動)。
'
'   第9.6b 段階で put_Handled = TRUE + 同一 View で View_Navigate(uri) する
'   挙動を追加し、Bing のリンクが「同じ画面で」開けるようになる予定。
' ============================================================
Public Function View_AddNewWindowRequested() As Long
    If m_pView = 0 Then
        View_AddNewWindowRequested = &H80004003   ' E_POINTER
        Exit Function
    End If

    If Not (m_persistentHandlers(HK_NewWindowRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: 既に登録済みなのでスキップ"
        View_AddNewWindowRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_NewWindowRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: ハンドラの Acquire に失敗"
        View_AddNewWindowRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 44, "View.add_NewWindowRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddNewWindowRequested: add_NewWindowRequested 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddNewWindowRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_NewWindowRequested) = token
    Set m_persistentHandlers(HK_NewWindowRequested) = h
    Debug.Print "Wv2Pane.View_AddNewWindowRequested OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddNewWindowRequested = 0
End Function


' ============================================================
' View_RemoveNewWindowRequested (第9.6a 段階で新規)
'   View_RemoveEvent (vtable index 45) で解除。
'   未登録なら no-op (S_OK で返す)。
' ============================================================
Public Function View_RemoveNewWindowRequested() As Long
    If m_pView = 0 Then
        View_RemoveNewWindowRequested = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_NewWindowRequested) Is Nothing Then
        View_RemoveNewWindowRequested = 0
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 45, "View.remove_NewWindowRequested", _
             m_persistentTokens(HK_NewWindowRequested))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveNewWindowRequested OK"
    End If

    Set m_persistentHandlers(HK_NewWindowRequested) = Nothing
    m_persistentTokens(HK_NewWindowRequested) = 0^

    View_RemoveNewWindowRequested = hr
End Function


' ============================================================
' View_OnNewWindowRequested (第9.6a で新規、第9.6b で捌きを追加)
'
'   args = ICoreWebView2NewWindowRequestedEventArgs の vtable レイアウト
'   (旧版 WebView2.cls の IDL コメントより):
'      3 : get_Uri(LPWSTR* uri)            ← 9.6a で取得
'      4 : put_NewWindow(ICoreWebView2*)
'      5 : get_NewWindow(ICoreWebView2**)
'      6 : put_Handled(BOOL)               ← 9.6b で利用
'      7 : get_Handled(BOOL*)
'      8 : get_IsUserInitiated(BOOL*)
'      9 : GetDeferral(ICoreWebView2Deferral**)
'     10 : get_WindowFeatures(ICoreWebView2WindowFeatures**)
'
'   第9.6b 段階の動作:
'     1. URL を get_Uri で取得してログ出力 (9.6a から継承)
'     2. args.put_Handled(1) で「VBA 側が捌く」と宣言
'     3. 同じ View で View_Navigate(url) を撃って「同じ画面で」遷移
'
'   put_Handled を立てない場合の挙動 (9.6a で確認済み):
'     WebView2 ランタイムが Excel プロセス内に新規 WebView2 ウィンドウを生成する。
'     UIA Inspect で確認した結果:
'       - クラス名: Chrome_WidgetWin_1
'       - Name: "Bing 画像" (= タイトルのみ、ブラウザ名なし)
'       - タスクマネージャー上では「ユーティリティ」グループ内の WebView2 プロセス
'     つまり Edge にバトンタッチするわけではなく、WebView2 が自分で別ウィンドウを
'     開いている。put_Handled = TRUE でこれを抑制できる。
'
'   失敗時の挙動 (案 a + 案 c、たーぼーさん合意):
'     put_Handled 失敗時: WebView2 が勝手に新規ウィンドウを開く (9.6a の挙動に戻る)
'     View_Navigate 失敗時: 現在のページのまま、リンクは開かない (ハングはしない)
'     どちらも致命的でないため Debug.Print だけして処理続行する。
'
'   将来 (タブブラウザ化、第9.6c 相当):
'     View_Navigate の代わりに args.put_NewWindow で別の Wv2Pane インスタンスを
'     渡す方針。put_Handled は据え置き。TabWebView2 包括クラスで対応予定。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
' ============================================================
Public Sub View_OnNewWindowRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnNewWindowRequested_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnNewWindowRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnNewWindowRequested: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnNewWindowRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View_OnNewWindowRequested: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs が 0 なので URL 取得をスキップ"
        Exit Sub
    End If

    ' --- 1. args.get_Uri (vtable index 3) で遷移先 URL を取得 (9.6a から継承) ---
    Dim url As String
    url = GetStringProperty(pArgs, 3, "NewWindowReqArgs.get_Uri")
    Debug.Print "  → 新規ウィンドウ要求 URL: " & url

    ' --- 第9.26/9.27: NewWindowRequested の捌き方をモードで切り替える ---
    '   m_browser が Nothing (単体 Pane 運用 = SetBrowser 未実行) のときはモードの
    '   置き場そのものが無いので、常に LEGACY 相当 (下の現行動作) になる。
    Dim nwMode As Long
    nwMode = NWM_LEGACY
    If Not (m_browser Is Nothing) Then nwMode = m_browser.NewWindowMode

    If nwMode = NWM_POPUP Then
        ' ★モード 1 (9.26 のプローブ)★ put_Handled を立てず、新タブ要求も出さずに
        '   Exit する。以降の捌きはランタイム任せで、自前のポップアップウィンドウに
        '   元のリクエスト (POST ならボディごと) が継続される = 仕様事実 23。
        '   ★検証用★ 開いたウィンドウはこちらの管理外 (手動で閉じる)。
        Debug.Print "  [9.26 委譲モード] put_Handled を立てずにランタイムへ委譲する"
        Debug.Print "    (ポップアップが開けば元のリクエストが継続されている = POST 生存)"
        Debug.Print "    ※自前の新タブは作らない。ポップアップは手動で閉じること。"
        Exit Sub

    ElseIf nwMode = NWM_PREWARM Then
        ' ★モード 2 (9.27 の本命)★ 温めてある予備タブの CoreWebView2 を
        '   args.put_NewWindow に渡す。ランタイムはその WebView に元のリクエストを
        '   載せるので、POST のメソッド/ボディ/Referer が保たれたまま自前のタブで開く。
        Dim pNew As LongPtr
        pNew = m_browser.ClaimPrewarmViewPtr()

        If pNew = 0 Then
            ' 予備が無い / 未 Ready / 既に Claim 済み。壊さずに現行動作へ落ちる。
            Debug.Print "  ★[9.27 プリウォーム] 予備タブが使えないので現行動作にフォールバックする★"
            Debug.Print "    → この遷移は GET に化ける (POST リンクなら 405 になる)。補充を予約する。"
            m_browser.RequestPrewarmAsync
            ' ↓ 下の現行動作 (put_Handled + RequestNewTabAsync) へ落ちる
        Else
            Dim hrNw As Long
            hrNw = dcf(pArgs, 4, "NewWindowReqArgs.put_NewWindow", pNew)

            If hrNw <> 0 Then
                ' put_NewWindow が失敗。予約を戻してから現行動作へ落ちる。
                Debug.Print "  ★[9.27 プリウォーム] put_NewWindow 失敗 hr=&H" & Hex(hrNw) & _
                            " → 現行動作にフォールバックする★"
                m_browser.ReleasePrewarmClaim
                ' ↓ 下の現行動作へ落ちる
            Else
                Debug.Print "  [9.27 プリウォーム] put_NewWindow OK (予備タブの View=" & pNew & ")"

                ' put_Handled の併用は Browser 側のスイッチで切り替える (論点5 の保険)。
                If m_browser.SetHandledWithNewWindow Then
                    hrNw = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
                    If hrNw <> 0 Then
                        Debug.Print "    put_Handled 失敗 hr=&H" & Hex(hrNw)
                    Else
                        Debug.Print "    put_Handled OK (put_NewWindow と併用)"
                    End If
                Else
                    Debug.Print "    put_Handled は立てない (SetHandledWithNewWindow = False)"
                End If

                ' 採用 (m_tabs へ登録 → TabAdded → ActivateTab) と補充は Timer に任せる。
                ' ★ここで待たない★ のがプリウォーム方式の要点 (仕様事実 14 を回避)。
                m_browser.RequestAdoptPrewarmedAsync url
                Exit Sub
            End If
        End If
    End If

    ' --- 第9.6b で追加: 同一 View で開く捌き ---
    If url = "" Then
        Debug.Print "  url が空なので put_Handled / View_Navigate をスキップ"
        Exit Sub
    End If

    ' --- 2. put_Handled = TRUE ---
    '   WebView2 ランタイムに「VBA 側で捌く」と宣言する。
    '   引数 1& は BOOL の TRUE (Win32 では BOOL = Long、& サフィックスで
    '   明示的に Long であることを保証、設計原則 29 に準ずる)。
    Dim hr As Long
    hr = dcf(pArgs, 6, "NewWindowReqArgs.put_Handled", 1&)
    If hr <> 0 Then
        Debug.Print "  put_Handled 失敗 hr=&H" & Hex(hr) & _
                    " (WebView2 が新規ウィンドウを開く挙動になる)"
    Else
        Debug.Print "  put_Handled OK (新規ウィンドウは抑制される)"
    End If

    ' --- 3. 新タブで開く (第9.10e、案A + 案Y) ---
    '   9.6b の「同一 View で View_Navigate」から「Wv2Browser の新タブで開く」に
    '   変更。m_browser 後方参照経由で AddTabWithUrl を呼ぶ (案Y)。
    '   ★window.open の戻り値は null になる★ put_NewWindow で新 View を渡さない
    '     簡易方式 (案A) のため。正攻法 (put_NewWindow + Deferral) は将来の別段階。
    '   m_browser 未設定 (SetBrowser 未実行) の場合は、壊さないよう同一 View の
    '     View_Navigate にフォールバックする (単体 Pane 運用でも動くように)。
    If m_browser Is Nothing Then
        Debug.Print "  [警告] m_browser 未設定 (SetBrowser 未実行)。同一 View Navigate に" & _
                    "フォールバック。"
        hr = View_Navigate(url)
        If hr <> 0 Then Debug.Print "  View_Navigate 失敗 hr=&H" & Hex(hr)
        Exit Sub
    End If

    ' ★デッドロック回避 (案β Timer)★ ここで m_browser.AddTabWithUrl を直接呼ぶと、
    '   AddTab の WaitPaneReady が新 Controller の ControllerCompleted を待つが、
    '   本ハンドラ (NewWindowRequested) が return するまで WebView2 はそれを配送
    '   しないためデッドロック (タイムアウト) する (仕様事実 14)。そこで
    '   RequestNewTabAsync に依頼し、ワンショット Timer 経由でハンドラ完了後に
    '   AddTabWithUrl を実行させる。
    m_browser.RequestNewTabAsync url
    Debug.Print "  新タブ生成を Timer に依頼 (RequestNewTabAsync): " & url
End Sub


' ============================================================
' SetBrowser / ClearBrowserRef (第9.10e、Wv2Browser 後方参照の設定・切断)
'
'   SetBrowser: Wv2Browser.AddTab がタブ生成時に自分自身 (Me) を渡す。これにより
'     Wv2Pane は View_OnNewWindowRequested から m_browser.AddTabWithUrl を呼べる。
'   ClearBrowserRef: Wv2Browser.Shutdown が循環参照 (Browser.m_tabs ? Pane.m_browser)
'     を断つために呼ぶ。これを呼ばずに UserForm が m_browser=Nothing すると、
'     循環参照で Browser も Pane も解放されずリークする (設計原則 58)。
'
'   ★設計原則 58★ VBA の参照カウント GC では相互参照が自動解放されない。
'     後方参照を持つ場合は「明示的に断ち切る Shutdown」を上位に用意する。
' ============================================================
Public Sub SetBrowser(ByVal browser As Wv2Browser)
    Set m_browser = browser
End Sub

Public Sub ClearBrowserRef()
    Set m_browser = Nothing
    Set m_webMsgSink = Nothing   ' 第9.11c: シンク後方参照も一緒に断つ (設計原則 58)
End Sub


' ============================================================
' SetWebMsgSink / ClearWebMsgSink (第9.11c、WebMessage シンクの設定・切断、案P)
'
'   SetWebMsgSink: JS からの postMessage をこのペインが受けたとき、指定した
'     シンクオブジェクトの OnPaneWebMessage(pane, source, msg) へ転送する。
'     TabBar が Init 時に自分自身をシンクに登録する。受信を実際に有効化する
'     には別途 View_AddWebMessageReceived を呼ぶこと (登録は 9.4c から既存)。
'   ClearWebMsgSink: シンクを切る。ClearBrowserRef からも呼ばれる (解放時)。
'
'   ★m_browser (NewWindowRequested 用の後方参照) とは別系統★ こちらは
'     WebMessage 転送専用。両者を混ぜないことで責務を分離する。
' ============================================================
Public Sub SetWebMsgSink(ByVal sink As Object)
    Set m_webMsgSink = sink
End Sub

Public Sub ClearWebMsgSink()
    Set m_webMsgSink = Nothing
End Sub


' ============================================================
' View_AddHistoryChanged (第9.9b 段階で新規、SPA 対応)
'
'   ICoreWebView2::add_HistoryChanged (vtable index 13) を呼んで
'   永続イベントハンドラを登録する。基底 ICoreWebView2 のイベントなので
'   m_pView を使う (m_pView2 は不要)。
'
'   発火タイミング:
'     SPA で history.pushState / replaceState / back / forward が
'     呼ばれて履歴状態が変化したとき。通常のページ遷移 (Navigate) でも
'     履歴に積まれれば発火する。SPA のルート変更を検知する主要な手段。
'
'   ★ 実装は View_AddNavigationStarting を逐語的な見本にしている ★
'     (設計原則 45: COM オブジェクトを WebView2 に渡すときは
'      pVTableObj = h.Slot + VTABLE_OBJ_OFFSET を必ず通す。
'      View_PostWebMessageAsString のような単純引数系は見本にしない)
'
'   違い:
'     - 種別: HK_HistoryChanged
'     - vtable index: 13 (add) / 14 (remove)
'     - args 型: IUnknown* (情報なし。sender 経由で get_Source 等を取る)
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View_AddHistoryChanged() As Long
    If m_pView = 0 Then
        View_AddHistoryChanged = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' 既に登録済みなら何もしない
    If Not (m_persistentHandlers(HK_HistoryChanged) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: 既に登録済みなのでスキップ"
        View_AddHistoryChanged = 0
        Exit Function
    End If

    ' ComCallbackHandler を Acquire (owner = Me で動的バインディング経由でイベント受信)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_HistoryChanged, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: ハンドラの Acquire に失敗"
        View_AddHistoryChanged = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_HistoryChanged を撃つ (vtable index 13)
    '   重要: WebView2 に渡すハンドラ COM オブジェクトのアドレスは
    '   h.Slot そのものではなく h.Slot + VTABLE_OBJ_OFFSET (設計原則 45)。
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 13, "View.add_HistoryChanged", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddHistoryChanged: add_HistoryChanged 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddHistoryChanged = hr
        Exit Function
    End If

    ' 成功: token と ComCallbackHandler を保存
    m_persistentTokens(HK_HistoryChanged) = token
    Set m_persistentHandlers(HK_HistoryChanged) = h
    Debug.Print "Wv2Pane.View_AddHistoryChanged OK: token=" & token & ", handler.slot=" & h.Slot

    View_AddHistoryChanged = 0
End Function


' ============================================================
' View_RemoveHistoryChanged (第9.9b 段階で新規)
'   View_RemoveEvent (vtable index 14) で解除。基底 IF なので m_pView。
'   未登録なら no-op (S_OK で返す)。
' ============================================================
Public Function View_RemoveHistoryChanged() As Long
    If m_pView = 0 Then
        View_RemoveHistoryChanged = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_HistoryChanged) Is Nothing Then
        View_RemoveHistoryChanged = 0   ' 未登録は no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView, 14, "View.remove_HistoryChanged", _
             m_persistentTokens(HK_HistoryChanged))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View_RemoveHistoryChanged OK"
    End If

    Set m_persistentHandlers(HK_HistoryChanged) = Nothing
    m_persistentTokens(HK_HistoryChanged) = 0^

    View_RemoveHistoryChanged = hr
End Function


' ============================================================
' View_OnHistoryChanged (第9.9b 段階で新規)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   args は IUnknown* で情報がないので、sender (= ICoreWebView2) 経由で
'   現在の状態を取りに行く (DocumentTitleChanged と同じパターン)。
'
'   取得する情報 (案 N2-a):
'     - get_Source        (vtable 4)  : 現在の URL
'     - get_CanGoBack     (vtable 38) : 「戻る」が可能か
'     - get_CanGoForward  (vtable 39) : 「進む」が可能か
'   SPA で pushState すると URL が変わり CanGoBack が True になる、という
'   SPA の本質的な挙動を 1 イベントで可視化できる。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
' ============================================================
Public Sub View_OnHistoryChanged(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View_OnHistoryChanged_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnHistoryChanged_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View_OnHistoryChanged: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View_OnHistoryChanged_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' 第9.9c (β): pArgs=0 は HistoryChanged の仕様 (args は IUnknown* で情報なし)。
    ' sender 経由で get_Source / get_CanGoBack / get_CanGoForward を取りに行く。
    Debug.Print "Wv2Pane.View_OnHistoryChanged: pSender=" & pSender & _
                ", pArgs=" & pArgs & " (args=0 は仕様)"
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    ' sender (ICoreWebView2) 経由で現在の状態を取得
    Dim url As String
    url = GetStringProperty(pSender, 4, "View.get_Source")
    Dim canBack As Boolean
    canBack = GetBoolProperty(pSender, 38, "View.get_CanGoBack")
    Dim canFwd As Boolean
    canFwd = GetBoolProperty(pSender, 39, "View.get_CanGoForward")

    ' 第9.9c (α): NavigateToString 環境では about:blank になるのが正常。
    ' 利用者が「URL が空っぽに見える」と困惑しないよう注記を入れる。
    If url = "about:blank" Then
        Debug.Print "  → URL: about:blank (NavigateToString 環境では正常、第9.9b 仕様事実 6)"
    Else
        Debug.Print "  → URL: " & url
    End If
    Debug.Print "  → CanGoBack: " & canBack & ", CanGoForward: " & canFwd

    ' --- 第9.12a: 所属 Browser に URL 状態を通知 (SPA の pushState / 戻る / 進む) ---
    '   ここでは既に url / canBack / canFwd を取得済みなので、NotifyUrlState を
    '   経由せず直接渡す (get_* の二度読みを避ける)。
    If Not m_browser Is Nothing Then
        m_browser.NotifyUrlChanged Me, url, canBack, canFwd
    End If
End Sub


' ============================================================
' View2_AddDOMContentLoaded (第9.9b 段階で新規、SPA 対応、ICoreWebView2_2)
'
'   ICoreWebView2_2::add_DOMContentLoaded (vtable index 64) を呼んで
'   永続イベントハンドラを登録する。これは派生 IF (ICoreWebView2_2) の
'   イベントなので、m_pView ではなく m_pView2 を使う。
'   (= 第9.9a で実装した EnsureView2 / m_pView2 の初の実戦投入)
'
'   発火タイミング:
'     ページの DOM が構築完了したとき (DOMContentLoaded、初回ロード時に 1 回)。
'     NavigationCompleted より前、DOM が触れるようになった時点で発火する。
'     SPA でも初回ロード時に発火する (pushState では発火しない = それは
'     HistoryChanged の役目)。
'
'   ★ 実装は View_AddHistoryChanged / View_AddNavigationStarting を
'      逐語的な見本にしている (設計原則 45) ★
'     違いは「m_pView の代わりに EnsureView2() + m_pView2 を使う」点のみ。
'     pVTableObj = h.Slot + VTABLE_OBJ_OFFSET は同じく必ず通す。
'
'   命名: 派生 IF 由来なので View2_ 接頭辞 (第9.6a/9.9b 命名規約)。
'
'   違い (基底版との比較):
'     - 種別: HK_DOMContentLoaded
'     - インターフェース: m_pView2 (EnsureView2 で遅延取得)
'     - vtable index: 64 (add) / 65 (remove)  ※ m_pView2 上のオフセット
'     - args 型: ICoreWebView2DOMContentLoadedEventArgs
'                (get_NavigationId = UINT64、vtable 3)
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View2_AddDOMContentLoaded() As Long
    ' _2 IF が必要。EnsureView2 で m_pView2 を取得 (未取得なら遅延取得)。
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: EnsureView2 失敗 (m_pView2 取得不可)"
        View2_AddDOMContentLoaded = &H80004003   ' E_POINTER
        Exit Function
    End If

    ' 既に登録済みなら何もしない
    If Not (m_persistentHandlers(HK_DOMContentLoaded) Is Nothing) Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: 既に登録済みなのでスキップ"
        View2_AddDOMContentLoaded = 0
        Exit Function
    End If

    ' ComCallbackHandler を Acquire (owner = Me で動的バインディング経由でイベント受信)
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_DOMContentLoaded, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: ハンドラの Acquire に失敗"
        View2_AddDOMContentLoaded = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    ' add_DOMContentLoaded を撃つ (vtable index 64、m_pView2 上)
    '   重要: WebView2 に渡すハンドラ COM オブジェクトのアドレスは
    '   h.Slot そのものではなく h.Slot + VTABLE_OBJ_OFFSET (設計原則 45)。
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView2, 64, "View2.add_DOMContentLoaded", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_AddDOMContentLoaded: add_DOMContentLoaded 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View2_AddDOMContentLoaded = hr
        Exit Function
    End If

    ' 成功: token と ComCallbackHandler を保存
    m_persistentTokens(HK_DOMContentLoaded) = token
    Set m_persistentHandlers(HK_DOMContentLoaded) = h
    Debug.Print "Wv2Pane.View2_AddDOMContentLoaded OK: token=" & token & ", handler.slot=" & h.Slot

    View2_AddDOMContentLoaded = 0
End Function


' ============================================================
' View2_RemoveDOMContentLoaded (第9.9b 段階で新規、ICoreWebView2_2)
'   remove_DOMContentLoaded (vtable index 65) で解除。_2 IF なので m_pView2。
'   未登録なら no-op (S_OK で返す)。
' ============================================================
Public Function View2_RemoveDOMContentLoaded() As Long
    ' _2 IF が必要。ただし解除なので、既に m_pView2 を持っているはず。
    ' 念のため EnsureView2 を通す (未取得なら 0 を返して弾く)。
    If Not EnsureView2() Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: m_pView2 が未取得"
        View2_RemoveDOMContentLoaded = &H80004003
        Exit Function
    End If

    If m_persistentHandlers(HK_DOMContentLoaded) Is Nothing Then
        View2_RemoveDOMContentLoaded = 0   ' 未登録は no-op
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView2, 65, "View2.remove_DOMContentLoaded", _
             m_persistentTokens(HK_DOMContentLoaded))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded: 解除失敗 hr=&H" & Hex(hr)
    Else
        Debug.Print "Wv2Pane.View2_RemoveDOMContentLoaded OK"
    End If

    Set m_persistentHandlers(HK_DOMContentLoaded) = Nothing
    m_persistentTokens(HK_DOMContentLoaded) = 0^

    View2_RemoveDOMContentLoaded = hr
End Function


' ============================================================
' View2_OnDOMContentLoaded (第9.9b 段階で新規、ICoreWebView2_2)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。
'
'   args = ICoreWebView2DOMContentLoadedEventArgs の get_NavigationId
'   (vtable index 3) を呼んで NavigationId をログ出力する。
'   NavigationId は UINT64 なので GetLongLongProperty で取得 (64bit 専用)。
'
'   命名: 派生 IF 由来なので View2_On 接頭辞。
'
'   Public Sub 必須 (Friend だと IDispatch::Invoke に乗らない)。
'
'   ★ 32bit 化ポイント ★
'     GetLongLongProperty は #If Win64 で囲まれた 64bit 専用ヘルパー。
'     32bit 対応段階では NavigationId の取得方法を別途設計する。
' ============================================================
Public Sub View2_OnDOMContentLoaded(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    ' ============================================================
    ' D-1 論点7: in-callback ガードの薄いラッパー
    '   本体は View2_OnDOMContentLoaded_Core。ここでは m_inCallback を +1 してから Core を呼び、
    '   ★どの経路で抜けても★ -1 する。Core には途中 Exit Sub があるため、
    '   Core 内に +1 / -1 を直書きする形では戻し漏れが出る (事後報告 (a))。
    '   Core がエラーで抜けた場合はカウンタを戻したうえで同じエラーを投げ直す
    '   (握り潰さない = 従来の挙動を変えない。事後報告 (b))。
    ' ============================================================
    Dim errNum As Long
    Dim errSrc As String
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View2_OnDOMContentLoaded_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errSrc = Err.source
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: Core でエラー " & errNum & " " & errDsc & _
                    " (in-callback ガードは解除済み、エラーは再送出)"
        Err.Raise errNum, errSrc, errDsc
    End If
End Sub


Private Sub View2_OnDOMContentLoaded_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Debug.Print "Wv2Pane.View2_OnDOMContentLoaded: pSender=" & pSender & ", pArgs=" & pArgs
    If pSender <> m_pView Then
        Debug.Print "  [警告] pSender が m_pView と一致しない (pSender=" & pSender & _
                    ", m_pView=" & m_pView & ")"
    End If

    If pArgs = 0 Then
        Debug.Print "  pArgs が 0 なので NavigationId 取得をスキップ"
        Exit Sub
    End If

    ' args.get_NavigationId (vtable index 3、UINT64) を取得
    Dim navId As LongLong
    navId = GetLongLongProperty(pArgs, 3, "DOMContentLoadedArgs.get_NavigationId")
    Debug.Print "  → navigationId: " & navId
End Sub


' ============================================================
' View_PostWebMessageAsString (第9.4c 段階で新規)
'
'   ICoreWebView2::PostWebMessageAsString (vtable index 33) を呼んで
'   JS にメッセージを送信する。
'
'   JS 側で受け取るには、以下のリスナーを HTML 内に登録しておく:
'     window.chrome.webview.addEventListener('message', event => {
'         console.log(event.data);   // 受信内容
'     });
'
'   注意: WebView2 はメッセージを「ページ単位」で隔離する。
'         ページが NavigationCompleted する前 (= まだリスナーが登録されていない)
'         に PostMessage を撃つと、メッセージはどこにも届かない。
'         必ず NavigationCompleted を待ってから撃つこと。
' ============================================================
Public Function View_PostWebMessageAsString(ByVal webMessage As String) As Long
    If m_pView = 0 Then
        View_PostWebMessageAsString = &H80004003
        Exit Function
    End If
    View_PostWebMessageAsString = dcf(m_pView, 33, _
                                     "View.PostWebMessageAsString", _
                                     StrPtr(webMessage))
End Function


' ============================================================
' ナビゲーション制御メソッド (第9.4b 段階で新規)
'
'   ICoreWebView2 のナビゲーション関連メソッドを Public で公開する。
'   ユーザは UserForm 等から m_pane.View_xxx() で呼べる。
'
'   各メソッドの vtable index は旧版で実機検証済み:
'     4  : get_Source
'     6  : NavigateToString
'     40 : GoBack
'     41 : GoForward
'     48 : get_DocumentTitle
' ============================================================

' --- HTML 文字列を直接 WebView2 に表示させる ---
Public Function View_NavigateToString(ByVal htmlContent As String) As Long
    If m_pView = 0 Then
        View_NavigateToString = &H80004003
        Exit Function
    End If
    ' 新しいナビゲーション開始でフラグをリセット
    m_pageLoaded = False
    View_NavigateToString = dcf(m_pView, 6, "View.NavigateToString", StrPtr(htmlContent))
End Function

' --- VBA オブジェクトを JS に公開する (第9.18、AddHostObjectToScript, vtable 49) ---
'   objName  : JS 側で window.chrome.webview.hostObjects.<objName> となる名前
'   instance : 公開する VBA クラスインスタンス (Instancing 1-Private でよい)
'
'   ★旧 WebView2.cls の実機検証済み実装を現行作法へ移植★
'     Variant にオブジェクトを詰め、その VarPtr を第2引数 (VARIANT*) に渡す。
'     第1引数はホストオブジェクト名の LPCWSTR。既存の View_Navigate (StrPtr) と
'     get_Settings (VarPtr) を組み合わせた形。
'
'   ★呼ぶタイミング (順序A)★ 設定タブでは NavigateToString の「前」に呼ぶ。
'     登録が Navigate より先なら、読み込まれた JS が最初から hostObjects を
'     使える、という仮説。実機で undefined になるようなら ready ハンドシェイク
'     方式 (順序B) にフォールバックする。
Public Function View_AddHostObjectToScript(ByVal objName As String, ByVal instance As Object) As Long
    If m_pView = 0 Then
        View_AddHostObjectToScript = &H80004003   ' E_POINTER
        Exit Function
    End If
    Dim vObj As Variant
    Set vObj = instance
    View_AddHostObjectToScript = dcf(m_pView, 49, "View.AddHostObjectToScript", _
                                     StrPtr(objName), VarPtr(vObj))
End Function

' --- 履歴を 1 つ戻る ---
Public Function View_GoBack() As Long
    If m_pView = 0 Then
        View_GoBack = &H80004003
        Exit Function
    End If
    View_GoBack = dcf(m_pView, 40, "View.GoBack")
End Function

' --- 履歴を 1 つ進む ---
Public Function View_GoForward() As Long
    If m_pView = 0 Then
        View_GoForward = &H80004003
        Exit Function
    End If
    View_GoForward = dcf(m_pView, 41, "View.GoForward")
End Function

' --- 現在表示中のページの URL を取得 ---
Public Function View_GetSource() As String
    If m_pView = 0 Then Exit Function
    View_GetSource = GetStringProperty(m_pView, 4, "View.get_Source")
End Function

' --- 現在表示中のページの DocumentTitle を取得 ---
Public Function View_GetDocumentTitle() As String
    If m_pView = 0 Then Exit Function
    View_GetDocumentTitle = GetStringProperty(m_pView, 48, "View.get_DocumentTitle")
End Function


' ============================================================
' 第9.8b 段階で追加: ICoreWebView2 基底 IF のメソッド網羅 (View_ 系)
'                  + ICoreWebView2Controller の Ctrl_Close
'                  + ICoreWebView2Settings のヘルパー方式実装 (Settings_ 系)
'
'   設計原則:
'     ・View_ 系は既存パターン (m_pView 経由 + GetXxxProperty / dcf) を踏襲
'     ・Settings 系は EnsureSettings で m_pSettings をキャッシュし、
'       Wv2Pane 内で完結 (Wv2Settings クラスを増やさない方針、9.8b-A 採用)
'     ・BOOL 取得系 (View_GetCanGoBack 等) は As Boolean を返すスタイル
'       (HRESULT は内部で確認、IsPageLoaded と整合、9.8b-C 採用)
'     ・void 系 (View_Reload, View_Stop, Ctrl_Close) は As Long で HRESULT 返却
'       (View_Navigate と整合)
' ============================================================

' --- ページを再読み込み (F5 相当、View vtable 31) ---
Public Function View_Reload() As Long
    If m_pView = 0 Then
        View_Reload = &H80004003
        Exit Function
    End If
    ' 再読み込みなのでページ読み込み完了フラグもリセット
    m_pageLoaded = False
    View_Reload = dcf(m_pView, 31, "View.Reload")
End Function

' --- ナビゲーション・読み込みを停止 (Esc 相当、View vtable 43) ---
Public Function View_Stop() As Long
    If m_pView = 0 Then
        View_Stop = &H80004003
        Exit Function
    End If
    View_Stop = dcf(m_pView, 43, "View.Stop")
End Function

' --- 「戻る」が可能か (View vtable 38、UI の Back ボタン活性制御用) ---
'   戻り値: True なら View_GoBack を呼べる状態
'   HRESULT は内部処理、失敗時は False を返す
Public Function View_GetCanGoBack() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoBack = GetBoolProperty(m_pView, 38, "View.get_CanGoBack")
End Function

' --- 「進む」が可能か (View vtable 39、UI の Forward ボタン活性制御用) ---
'   戻り値: True なら View_GoForward を呼べる状態
'   HRESULT は内部処理、失敗時は False を返す
Public Function View_GetCanGoForward() As Boolean
    If m_pView = 0 Then Exit Function
    View_GetCanGoForward = GetBoolProperty(m_pView, 39, "View.get_CanGoForward")
End Function

' --- WebView2 ブラウザプロセスの PID を取得 (View vtable 37、診断用) ---
'   戻り値: ブラウザプロセスの Process ID (Long)
'   タスクマネージャと突き合わせて WebView2 プロセスを特定するときに有用
'   HRESULT は内部処理、失敗時は 0 を返す
Public Function View_GetBrowserProcessId() As Long
    If m_pView = 0 Then Exit Function
    View_GetBrowserProcessId = GetLongProperty(m_pView, 37, "View.get_BrowserProcessId")
End Function


' ============================================================
' EnsureSettings (Private、内部ヘルパー)
'
'   m_pSettings が未取得 (= 0) なら View vtable 3 (get_Settings) を叩いて
'   ICoreWebView2Settings* を取得し、m_pSettings にキャッシュする。
'   既に取得済みならそのまま使う。
'
'   COM 所有権:
'     get_Settings は [out, retval] でポインタを返す。返ってきた時点で
'     refcount は既に立っているので、VBA 側で別途 AddRef は不要。
'     m_pSettings の解放責任は Class_Terminate にある。
'
'   戻り値: True なら m_pSettings が有効、False なら取得失敗
' ============================================================
Private Function EnsureSettings() As Boolean
    If m_pSettings <> 0 Then
        EnsureSettings = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: m_pView が未取得"
        Exit Function
    End If

    Dim pSettings As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 3, "View.get_Settings", VarPtr(pSettings))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: get_Settings 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    If pSettings = 0 Then
        Debug.Print "Wv2Pane.EnsureSettings: pSettings が 0"
        Exit Function
    End If

    m_pSettings = pSettings
    Debug.Print "Wv2Pane.EnsureSettings: m_pSettings 取得 = " & m_pSettings
    EnsureSettings = True
End Function

' --- JS の実行可否を設定 (Settings vtable 4 = put_IsScriptEnabled) ---
'   既定値 True。False にすると JS が一切動かなくなる。
'   IsScriptEnabled は WebView2 起動後・ナビゲーション前に設定するのが本来
'   だが、ナビゲーション後でも反映される (次回ロードから有効)。
Public Function Settings_PutIsScriptEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutIsScriptEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutIsScriptEnabled = dcf(m_pSettings, 4, "Settings.put_IsScriptEnabled", bv)
End Function

' --- 右クリックメニューの表示可否を設定 (Settings vtable 14 = put_AreDefaultContextMenusEnabled) ---
'   既定値 True。False にすると Ctrl+クリック等を含む右クリック系メニューが
'   一切出なくなる。業務アプリで右クリックを抑制したいとき有効。
Public Function Settings_PutAreDefaultContextMenusEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDefaultContextMenusEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDefaultContextMenusEnabled = dcf(m_pSettings, 14, _
                                              "Settings.put_AreDefaultContextMenusEnabled", bv)
End Function

' --- DevTools (F12) の利用可否を設定 (Settings vtable 12 = put_AreDevToolsEnabled) ---
'   既定値 True。False にすると F12 を押しても DevTools が開かない。
'   配布アプリで DevTools を封じたいときに利用。
Public Function Settings_PutAreDevToolsEnabled(ByVal v As Boolean) As Long
    If Not EnsureSettings() Then
        Settings_PutAreDevToolsEnabled = &H80004003
        Exit Function
    End If
    Dim bv As Long
    If v Then bv = 1 Else bv = 0
    Settings_PutAreDevToolsEnabled = dcf(m_pSettings, 12, "Settings.put_AreDevToolsEnabled", bv)
End Function


' ============================================================
' 第9.9a 段階で追加: ICoreWebView2_2 派生 IF の遅延取得 (View2_ 系の基盤)
'
'   設計判断 (たーぼーさん合意済み):
'     ・案 1: m_pView2 フィールド + EnsureView2 遅延取得 (Settings 流儀を踏襲)
'     ・案 a: IID_ICoreWebView2_2 は EnsureView2 のローカル変数で都度作る
'             (Wv2Thunks に IID テーブル追加せず、9.9a を Wv2Thunks 最小変更に)
'     ・案 X: 戻り値は As Boolean (EnsureSettings と同形)
'
'   IID_ICoreWebView2_2 の出典:
'     WebView2.h L5078:
'       EXTERN_C __declspec(selectany) const IID IID_ICoreWebView2_2 =
'           {0x9E8F0CF8,0xE670,0x4B5E,{0xB2,0xBC,0x73,0xE0,0x61,0xE3,0x18,0x4C}};
'     MIDL_INTERFACE 文字列: "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"
'
'   ICoreWebView2_2 で追加されるメソッド (基底 ICoreWebView2 の 61?67):
'     61 : add_WebResourceResponseReceived
'     62 : remove_WebResourceResponseReceived
'     63 : NavigateWithWebResourceRequest
'     64 : add_DOMContentLoaded         ← 第9.9b で利用予定
'     65 : remove_DOMContentLoaded      ← 第9.9b で利用予定
'     66 : get_CookieManager
'     67 : get_Environment
' ============================================================


' ============================================================
' EnsureView2 (Private、内部ヘルパー)
'
'   m_pView2 が未取得 (= 0) なら QueryInterface(IID_ICoreWebView2_2) を
'   叩いて ICoreWebView2_2* を取得し、m_pView2 にキャッシュする。
'   既に取得済みならそのまま使う。
'
'   COM 所有権:
'     QueryInterface は成功時にポインタを AddRef した状態で返す。
'     m_pView2 は m_pView とは独立した参照カウントを持つ
'     (= 同一 COM オブジェクトを別 IF で見ているだけだが、参照カウントは別)。
'     m_pView2 の解放責任は Class_Terminate / Ctrl_Close にある (子から先)。
'
'   QueryInterface 呼び出し方:
'     IUnknown::QueryInterface (vtable 0):
'       HRESULT QueryInterface(REFIID riid, void** ppvObject);
'     第 1 引数 riid は GUID 構造体のポインタ。
'     第 2 引数 ppvObject は LongPtr の out 引数。
'     dcf 経由なら VarPtr(localIID) と VarPtr(pv) で渡す。
'
'   IID は EnsureView2 のローカル変数で都度作る (案 a 採用):
'     Wv2Thunks の m_iidTable に追加せず、必要なときに必要なところで生成。
'     初回 1 回しか呼ばれない (キャッシュ済みなら即 True で抜ける) ため
'     オーバーヘッドは極小。Wv2Thunks の FillGUID を Public で呼ぶ。
'
'   戻り値: True なら m_pView2 が有効、False なら取得失敗 (案 X、EnsureSettings と同形)
' ============================================================
Private Function EnsureView2() As Boolean
    If m_pView2 <> 0 Then
        EnsureView2 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: m_pView が未取得"
        Exit Function
    End If

    ' --- IID_ICoreWebView2_2 をローカル GUID 構造体に組み立てる ---
    '   {9E8F0CF8-E670-4B5E-B2BC-73E061E3184C}
    '   FillGUID は Wv2Thunks の Public Sub (第9.9a で Private → Public 昇格)
    Dim iidView2 As GUID
    FillGUID iidView2, "9E8F0CF8-E670-4B5E-B2BC-73E061E3184C"

    ' --- QueryInterface(riid, &ppv) を叩く ---
    '   vtable index 0 が QueryInterface (IUnknown 基底)。
    '   引数:
    '     #1: REFIID riid     → VarPtr(iidView2)  (GUID 構造体への参照)
    '     #2: void** ppvObject → VarPtr(pv)        (出力先 LongPtr の参照)
    '   戻り値: HRESULT (S_OK = 0 で成功)
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_2)", _
             VarPtr(iidView2), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView2: QueryInterface 成功だが pv が 0"
        Exit Function
    End If

    m_pView2 = pv
    Debug.Print "Wv2Pane.EnsureView2: m_pView2 取得 = " & m_pView2
    EnsureView2 = True
End Function


' ============================================================
' View2_Ensure (第9.9a で新規、Public 検証用ラッパ)
'
'   Immediate ウィンドウから ?UserForm1.GetActivePane.View2_Ensure で
'   叩いて m_pView2 取得を確認するための薄いラッパ。
'
'   9.9b 以降は View2_AddDOMContentLoaded 等の各メソッドが内部で
'   EnsureView2 を呼ぶ形になるので、この関数の直接利用は基本的に検証時のみ。
'
'   案 i 採用: 検証は本関数 + Test_AllEvents_Real のリグレッションゼロ確認で実施。
'             テスト関数 (Test_View2_Ensure_Real) は新設しない。
'
'   戻り値: True なら m_pView2 が有効 (取得成功 or 既に保持済み)
' ============================================================
Public Function View2_Ensure() As Boolean
    View2_Ensure = EnsureView2()
End Function


' ============================================================
' EnsureView3 (第9.10d-2a で新規、ICoreWebView2_3 派生 IF の遅延取得)
'
'   ICoreWebView2_3 を QueryInterface で取り出して m_pView3 にキャッシュする。
'   EnsureView2 (第9.9a) と完全に同形。基底 m_pView から QI する
'   (ICoreWebView2_3 : ICoreWebView2_2 : ICoreWebView2 の単一継承なので
'    基底から一段で取り出せる)。
'
'   SetVirtualHostNameToFolderMapping (vtable 71) を叩く前提の基盤。
'
'   IID_ICoreWebView2_3 = {A0D6DF20-3B92-416D-AA0C-437A9C727857}
'     (WebView2.h: 0xA0D6DF20,0x3B92,0x416D,{0xAA,0x0C,0x43,0x7A,0x9C,0x72,0x78,0x57})
'
'   戻り値: True なら m_pView3 が有効 (取得成功 or 既に保持済み)
' ============================================================
Private Function EnsureView3() As Boolean
    If m_pView3 <> 0 Then
        EnsureView3 = True
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: m_pView が未取得"
        Exit Function
    End If

    ' --- IID_ICoreWebView2_3 をローカル GUID 構造体に組み立てる ---
    '   FillGUID は Wv2Thunks の Public Sub (第9.9a で Public 昇格)
    Dim iidView3 As GUID
    FillGUID iidView3, "A0D6DF20-3B92-416D-AA0C-437A9C727857"

    ' --- QueryInterface(riid, &ppv) を叩く (vtable index 0) ---
    Dim pv As LongPtr
    Dim hr As Long
    hr = dcf(m_pView, 0, "View.QueryInterface(ICoreWebView2_3)", _
             VarPtr(iidView3), VarPtr(pv))
    If hr <> 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If
    If pv = 0 Then
        Debug.Print "Wv2Pane.EnsureView3: QueryInterface 成功だが pv が 0"
        Exit Function
    End If

    m_pView3 = pv
    Debug.Print "Wv2Pane.EnsureView3: m_pView3 取得 = " & m_pView3
    EnsureView3 = True
End Function


' ============================================================
' View3_Ensure (第9.10d-2a で新規、Public 検証用ラッパ)
'   Immediate から ?UserForm1.GetActivePane.View3_Ensure で m_pView3 取得を確認。
'   View2_Ensure (第9.9a) と同形。
' ============================================================
Public Function View3_Ensure() As Boolean
    View3_Ensure = EnsureView3()
End Function


' ============================================================
' View3_SetVirtualHostNameToFolderMapping (第9.10d-2a で新規、ICoreWebView2_3)
'
'   ICoreWebView2_3::SetVirtualHostNameToFolderMapping (vtable index 71) を呼び、
'   ローカルフォルダを仮想ホスト名にマッピングする。これにより
'   https://<hostName>/ という secure origin でローカルファイルを配信でき、
'   about:blank origin (仕様事実 6) の pushState/replaceState 制約を回避できる。
'
'   ★これは単純引数系メソッド★ COM オブジェクト (ハンドラ) を WebView2 に
'   渡さないので、設計原則 45 (VTABLE_OBJ_OFFSET) は不要。View_Navigate の
'   ような文字列引数系を見本にする (View_AddNavigationStarting 等のハンドラ系は
'   見本にしない)。
'
'   シグネチャ (WebView2.h):
'     HRESULT SetVirtualHostNameToFolderMapping(
'       LPCWSTR hostName,     ' 仮想ホスト名 (例 "appassets.example")
'       LPCWSTR folderPath,   ' マッピング先ローカルフォルダの絶対パス
'       COREWEBVIEW2_HOST_RESOURCE_ACCESS_KIND accessKind)  ' DENY=0/ALLOW=1/DENY_CORS=2
'
'   ★呼び出しタイミング★ Navigate の前に設定する必要がある。かつ View3 が
'   必要なので、先に EnsureView3 を通す。
'
'   accessKind は Win32 の enum (4 byte int)。利用者が選べるよう引数で受ける
'   (通常は 1 = ALLOW を渡す)。dcf には Long としてそのまま渡す
'   (put_Handled 1& と同じ、単純な整数引数)。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View3_SetVirtualHostNameToFolderMapping( _
    ByVal hostName As String, _
    ByVal folderPath As String, _
    ByVal accessKind As Long) As Long

    If Not EnsureView3() Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: EnsureView3 失敗"
        View3_SetVirtualHostNameToFolderMapping = &H80004003   ' E_POINTER
        Exit Function
    End If

    Dim hr As Long
    hr = dcf(m_pView3, 71, "View3.SetVirtualHostNameToFolderMapping", _
             StrPtr(hostName), StrPtr(folderPath), accessKind)

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping: 失敗 hr=&H" & _
                    Hex(hr) & " (host=" & hostName & ", folder=" & folderPath & ")"
    Else
        Debug.Print "Wv2Pane.View3_SetVirtualHostNameToFolderMapping OK: https://" & _
                    hostName & "/ -> " & folderPath & " (accessKind=" & accessKind & ")"
    End If

    View3_SetVirtualHostNameToFolderMapping = hr
End Function


' ============================================================
' Ctrl_Close (Controller vtable 24)
'
'   WebView2 controller を明示的に閉じる。Controller の寿命を VBA 側から
'   即座に終わらせたい場合に使う。
'
'   第9.8c での拡張 (案 E 1+3 ハイブリッド):
'     Close 後の安全性を高めるため、以下を実施:
'       (1) 通常通り Controller::Close を撃つ (vtable 24)
'       (2) m_pSettings / m_pView2 / m_pView / m_pCtrl をローカルにキャッシュして
'           VBA 側でも Release を撃ち、メンバを 0 にクリア
'           (第9.9a で m_pView2 を追加、Release 順序は子から親へ厳格 = 設計原則 43)
'       (3) m_persistentTokens を 0 にクリア (Cleanup の二重 remove 防止)
'           ※ 永続ハンドラは WebView2 側が Close で勝手に Release してくれるので、
'              VBA 側で remove_xxx を撃ってはならない (= 既に refCount=0 の
'              スロットに remove_xxx を撃つと AV のリスク)
'       (4) m_closed = True を立てる (View_xxx / Settings_xxx の早期 return)
'       (5) m_state = Cs_Failed (Init 直後とは異なる状態として示す)
'
'   これにより:
'     - Close 後に Class_Terminate が走っても、CleanupAllPersistentHandlers は
'       m_persistentTokens がすべて 0 なのでスキップ (= 安全)
'     - Release も m_pSettings/m_pView2/m_pView/m_pCtrl がすべて 0 なのでスキップ
'     - Close 後に利用者が誤って View_xxx を叩いても m_closed で早期 return
'
'   注意:
'     - Ctrl_Close の戻り値は Controller::Close 自体の HRESULT
'       (= 後続の VBA 側クリア処理で失敗しても、戻り値には影響しない)
'     - 同じ Wv2Pane インスタンスで再度 Init を呼ぶ用途は想定していない
'       (= 再利用するなら新規 New Wv2Pane で別インスタンスを作る)
' ============================================================
Public Function Ctrl_Close() As Long
    If m_closed Then
        Debug.Print "Wv2Pane.Ctrl_Close: 既に Close 済み、無視"
        Ctrl_Close = 0  ' 二重 Close は無害なので S_OK 返却
        Exit Function
    End If
    If m_pCtrl = 0 Then
        Ctrl_Close = &H80004003
        Exit Function
    End If

    Debug.Print "Wv2Pane.Ctrl_Close: m_pCtrl = " & m_pCtrl
    Dim hr As Long
    hr = dcf(m_pCtrl, 24, "Ctrl.Close")

    ' --- (2) (3) (4) (5): Close 後の状態に遷移 ---
    '
    ' Close を撃った時点で WebView2 ランタイム側は登録済みハンドラを Release 済み
    ' のはず (実機検証で観察、第9.8b の知見「Ctrl_Close の副作用」)。
    ' その状態で VBA 側も m_pSettings/m_pView2/m_pView/m_pCtrl の Release を撃って、
    ' VBA 側の AddRef 分を清算する。
    '
    ' 順序は Class_Terminate と同じ Settings → View2 → View → Ctrl (子から親へ厳格、
    ' 第9.9a で m_pView2 を追加、設計原則 43)。
    If m_pSettings <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pSettings = " & m_pSettings
        ComRelease m_pSettings
        m_pSettings = 0
    End If
    If m_pView3 <> 0 Then
        ' 第9.10d-2a で追加: View3 は最派生 IF、m_pView2 より先に Release (設計原則 51)。
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView3 = " & m_pView3
        ComRelease m_pView3
        m_pView3 = 0
    End If
    If m_pView2 <> 0 Then
        ' 第9.9a で追加: View2 は View から QI で取り出した派生 IF、View より先に Release。
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView2 = " & m_pView2
        ComRelease m_pView2
        m_pView2 = 0
    End If
    If m_pView <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pView = " & m_pView
        ComRelease m_pView
        m_pView = 0
    End If
    If m_pCtrl <> 0 Then
        Debug.Print "Wv2Pane.Ctrl_Close: Releasing m_pCtrl = " & m_pCtrl
        ComRelease m_pCtrl
        m_pCtrl = 0
    End If

    ' (3) 永続ハンドラの token を 0 にクリア:
    '   WebView2 が Close で勝手に Release を撃っているので、
    '   ここで remove_xxx を撃つと「既に解放済みのハンドラに remove を撃つ」
    '   ことになり危険。token を 0 にすることで CleanupAllPersistentHandlers が
    '   スキップする (token=0 の要素は処理しない実装になっている、第9.4a)。
    '
    '   第9.9b: ループ範囲を固定値から LBound/UBound に変更。
    '   配列上限が HK_DOMContentLoaded に拡張されても自動追従する
    '   (= CleanupAllPersistentHandlers と同じ「横展開自動化」の流儀)。
    Dim k As Long
    For k = LBound(m_persistentHandlers) To UBound(m_persistentHandlers)
        m_persistentTokens(k) = 0^
        Set m_persistentHandlers(k) = Nothing
    Next k

    ' (4) m_closed フラグを立てる
    m_closed = True
    m_state = Cs_Failed   ' Close 後は無効状態とみなす

    Ctrl_Close = hr
End Function


' ============================================================
' 第9.8c 段階で追加: View_ExecuteScript + OnExecuteScriptCompleted
'                  (案 a = 同時並行 OK、案 α = イベントコールバック方式、
'                   案 X = 連番 callback ID 返却)
' ============================================================


' --- 任意の JavaScript を WebView2 内で実行 (View vtable 29) ---
'
'   引数:
'     jsCode  : 実行する JS コード (String)
'
'   戻り値 (As Long):
'     >= 1    : 採番された callback ID (正常な呼び出し開始)
'     <= 0    : エラー (m_pView = 0、Close 済み、dcf 失敗等)
'                 -1 → m_closed (Close 済み)
'                 -2 → m_pView = 0 (View 未取得)
'                 -3 → Acquire 失敗
'                 -4 → dcf 失敗 (戻り値の絶対値は HRESULT の下位、
'                       詳細は Debug.Print に出力)
'
'   流れ:
'     1. callback ID を採番 (m_executeScriptCounter を 1 増やす)
'     2. AcquireHandlerFor(HK_ExecuteScriptCompleted, Me) で新規ハンドラ取得
'     3. ハンドラに callbackId を設定 (CallbackId プロパティ Let)
'     4. View::ExecuteScript(jsCode, handler.Slot) を撃つ
'     5. ローカルの h を Nothing に倒す (= Wv2Pane は ComCallbackHandler を保持しない、
'        WebView2 ランタイム側だけが AddRef 経由で握る状態にする)
'     6. callback ID を返却
'
'   後続:
'     WebView2 ランタイムが JS 実行 → 完了時に handler.Invoke を撃つ
'     → ComCallbackHandler.Invoke[HK_ExecuteScriptCompleted] が走る
'     → 動的バインディングで Wv2Pane.OnExecuteScriptCompleted が呼ばれる
'     → WebView2 が handler を Release → refCount=0 → auto-released
'
'   同時並行性:
'     利用者が View_ExecuteScript を 10 回連続で叩いても、それぞれが独立した
'     スロット + 独立した callback ID を持つので干渉しない (案 a の核心)。
Public Function View_ExecuteScript(ByVal jsCode As String) As Long
    If m_closed Then
        Debug.Print "Wv2Pane.View_ExecuteScript: Close 済みのため拒否"
        View_ExecuteScript = -1
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: m_pView が未取得"
        View_ExecuteScript = -2
        Exit Function
    End If

    ' --- 1. callback ID を採番 ---
    m_executeScriptCounter = m_executeScriptCounter + 1
    Dim cbid As Long
    cbid = m_executeScriptCounter

    ' --- 2. ハンドラを Acquire ---
    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_ExecuteScriptCompleted, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_ExecuteScript: ハンドラの Acquire 失敗"
        View_ExecuteScript = -3
        Exit Function
    End If

    ' --- 3. callback ID を設定 ---
    h.callbackId = cbid

    ' --- 4. ExecuteScript を撃つ (vtable 29) ---
    '   シグネチャ: HRESULT ExecuteScript(LPCWSTR javaScript, IExecuteScriptCompletedHandler*)
    '
    '   重要: WebView2 に渡す「ハンドラ COM オブジェクト」のポインタは
    '   h.Slot そのものではなく、h.Slot + VTABLE_OBJ_OFFSET。
    '   サンクヘッダの先頭ではなく、サンク内の VTable オブジェクト部分を
    '   指すアドレスを渡す必要がある (= 既存の add_NavigationStarting 等と
    '   同じパターン、設計原則 8 系)。
    '
    '   これを忘れると WebView2 がデタラメなアドレスを vtable として辿って
    '   即クラッシュする。
    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim hr As Long
    hr = dcf(m_pView, 29, "View.ExecuteScript", StrPtr(jsCode), pVTableObj)

    ' --- 5. ローカル参照を破棄 (Wv2Pane は保持しない) ---
    '   WebView2 が AddRef を撃って自分で保持しているので、こちらは捨てて OK。
    '   完了通知で m_owner (= Me Wv2Pane) を経由して呼ばれる。
    Set h = Nothing

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_ExecuteScript: dcf 失敗 hr=&H" & Hex(hr) & _
                    " (callbackId=" & cbid & " は破棄、ハンドラは即解放見込み)"
        View_ExecuteScript = -4
        Exit Function
    End If

    Debug.Print "Wv2Pane.View_ExecuteScript OK: callbackId=" & cbid & _
                ", jsLen=" & Len(jsCode)
    View_ExecuteScript = cbid
End Function


' ============================================================
' View_ExecuteScriptSafe (第9.9c 段階で新規、JS 例外の構造化結果ヘルパー)
'
'   ユーザーが指定した JS 式を try-catch で包んでから ExecuteScript に渡す。
'   完了通知の resultJson が必ず以下のいずれかの構造化 JSON になる:
'     成功時: {"ok":true, "value":<JS の戻り値>}
'     失敗時: {"ok":false, "errorName":"<例外名>", "errorMessage":"<メッセージ>",
'              "stack":"<スタックトレース>"}
'
'   これにより、9.8c で判明した「JS 例外時に HRESULT=S_OK + result=null」で
'   区別不能だった問題を VBA 側で構造的に判別できるようになる。
'
'   設計判断 (たーぼーさん合意済み):
'     ・F1: 名前は View_ExecuteScriptSafe (簡潔、Safe = try-catch で包む)
'     ・F2: VBA 側では JSON 文字列のまま OnExecuteScriptCompleted に流す
'           (VBA で JSON パースは 64bit で標準機能なし。利用者が InStr 等で簡易解析)
'     ・F3: シンプル結合 (ユーザーコードの " エスケープは利用者の責任)
'
'   ★ 制約 (利用者への注意) ★
'     ユーザーコードは「文」ではなく「式」でなければならない。
'     (= var __r = (USER_CODE) で代入できる単一の式が前提)
'
'     具体的に NG / OK の例:
'       NG: "throw new Error('x')"   ← throw は文 (statement)。式の括弧内に
'                                       置けないため、wrappedJs がパース段階で
'                                       構文エラーになり、ExecuteScript は
'                                       resultJson = "null" を返す
'                                       (= try-catch には届かない)
'       OK: "(()=>{ throw new Error('x') })()"
'                                     ← 即時関数で包めば「式」として扱える
'                                       (戻り値は throw が走るので決まらないが、
'                                        catch 節が確実に発火する)
'       OK: "(function(){ var x=1; return x+2; })()"
'                                     ← 複数文 + return を即時関数で包む
'       OK: "nonexistentFunction()"   ← 式 (関数呼び出し)。実行時の
'                                       ReferenceError は catch で受けられる
'
'     ★ 実機検証 (第9.9c) で判明した事実 ★
'       VBA 標準のドキュメントには書きにくいが、JS の構文エラーは
'       ExecuteScript の HRESULT=S_OK + result="null" で返ってくる。
'       実行時例外 (ReferenceError 等) と区別できない。本ヘルパーは
'       実行時例外には対応するが、構文エラーには対応できない。
'       (構文エラーを catch する正攻法は eval() を使うことだが過剰)
'
'     文字列リテラルに " を含めたいときは利用者側で \" にする
'       (VBA の "" エスケープではなく JS の \" エスケープ)。
'
'   使用例:
'     cbid = m_pane.View_ExecuteScriptSafe("1 + 2")
'       → resultJson = "{\"ok\":true,\"value\":3}"
'     cbid = m_pane.View_ExecuteScriptSafe("(()=>{throw new Error('boom')})()")
'       → resultJson = "{\"ok\":false,\"errorName\":\"Error\",
'                         \"errorMessage\":\"boom\",\"stack\":\"...\"}"
'       ※ "throw new Error('boom')" を直接渡すと resultJson="null" になる
'         (上の制約参照)
'     cbid = m_pane.View_ExecuteScriptSafe("nonexistentFunction()")
'       → resultJson = "{\"ok\":false,\"errorName\":\"ReferenceError\",...}"
'
'   戻り値: View_ExecuteScript と同じ (callback ID か負のエラーコード)
' ============================================================
Public Function View_ExecuteScriptSafe(ByVal userCode As String) As Long
    ' ユーザーコードを try-catch で包んで構造化 JSON を返す JS 式に組み立てる。
    ' 改行は入れず 1 行で組む (ExecuteScript はソースマップが意味を持たないので可読性 < 単純性)。
    Dim wrappedJs As String
    wrappedJs = "(function(){try{var __r=(" & userCode & ");" & _
                "return {ok:true,value:__r};}" & _
                "catch(e){return {ok:false,errorName:e.name," & _
                "errorMessage:e.message,stack:e.stack};}})()"

    Debug.Print "Wv2Pane.View_ExecuteScriptSafe: wrappedJsLen=" & Len(wrappedJs)
    View_ExecuteScriptSafe = View_ExecuteScript(wrappedJs)
End Function


' --- ExecuteScript の完了通知を受け取る (動的バインディング受け取り口) ---
'
'   ComCallbackHandler.Handler_Invoke[HK_ExecuteScriptCompleted] から
'   m_owner.OnExecuteScriptCompleted で呼ばれる。Public でなければ動的バインディングで
'   呼べないため Public 必須 (= 設計原則 14 系)。
'
'   引数:
'     callbackId  : View_ExecuteScript が返した連番 ID
'     errorCode   : HRESULT (0 = S_OK)
'     resultJson  : JS の戻り値を JSON エンコードした文字列
'                   例:
'                     ExecuteScript("1 + 2")     → "3"
'                     ExecuteScript("'foo'")     → "\"foo\""
'                     ExecuteScript("document.title")
'                                                → "\"検索 - Microsoft Bing\""
'                     ExecuteScript("undefined") → "null" (※ undefined は JSON 化で null)
'                     ExecuteScript("throw 1")   → "" (errorCode が非 0)
'
'   案 α 第二実装 (第9.9c で拡張):
'     1. Wv2Pane 内のログ出力 (= 9.8c 第一実装、リグレッション切り分け用に残す)
'     2. UserForm1.OnExecuteScriptResult への通知 (= 9.9c で追加)
'
'   ★ UserForm1 への通知の設計判断 (案 E1-a-i) ★
'     Wv2Pane は UserForm1 への参照を **保持しない**。VBA の UserForm は
'     モジュールレベルでグローバルにアクセスできる (= UserForm1 と書けば
'     自動生成される) ので、ここで直接 UserForm1.OnExecuteScriptResult を
'     呼べば循環参照なしで通知できる。
'
'     これは「Wv2Pane が UserForm1 の存在を知っている」という意味で
'     階層的にはやや乱暴 (= 通常は下位が上位を知らない方が綺麗) だが、
'     9.10 で TabWebView2 が入る際に「結果通知を上位に流す経路」自体が
'     再設計されるので、9.9c では暫定として直接呼びの形にする。
'
'     UserForm1 側に On Error Resume Next 等の防御は入れないので、
'     UserForm1.OnExecuteScriptResult はエラーを投げない実装にする。
Public Sub OnExecuteScriptCompleted( _
    ByVal callbackId As Long, _
    ByVal errorCode As Long, _
    ByVal resultJson As String)

    ' --- (1) Wv2Pane 内のログ (9.8c で実装、案 E2-X = 残す) ---
    Debug.Print "Wv2Pane.OnExecuteScriptCompleted: callbackId=" & callbackId & _
                ", errorCode=" & errorCode & _
                ", resultJson=" & resultJson

    ' --- (1b) D-1: 保留テーブルの更新 (論点2・論点5) ---
    '   EvalSync が待っている callbackId なら結果を書き込む。
    '   既に破棄済み (タイムアウト後の遅延到着) ならログを残して捨てる。
    '   EvalSync を経由しない View_ExecuteScript 直呼びの結果はテーブルに
    '   居ないので、何もせず素通りする。
    Dim k As String
    Dim ent As Variant
    Dim found As Boolean

    k = "cb" & callbackId
    If Not m_evalPending Is Nothing Then
        found = False
        On Error Resume Next
        ent = m_evalPending(k)
        found = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0

        If found Then
            If ent(0) = EVAL_DISCARDED Then
                Debug.Print "Wv2Pane.OnExecuteScriptCompleted: ★破棄済み★ の " & k & _
                            " に遅延到着したので捨てる (タイムアウト後の完了)"
                m_evalPending.Remove k
            Else
                m_evalPending.Remove k
                m_evalPending.Add Array(EVAL_DONE, errorCode, resultJson), k
            End If
        End If
    End If

    ' --- (2) UserForm1 への通知 (第9.9c で追加、案 α 第二実装) ---
    '   VBA のグローバル UserForm 参照を活用 (循環参照なし、案 E1-a-i)。
    '   UserForm1 が存在しないケース (起動していない、または既に Terminate 済み)
    '   への配慮: VBA は UserForm1 を参照しただけで自動生成するので、
    '   起動していなくても呼び出し自体は通る。Terminate 済みの場合は
    '   再び自動生成 + Initialize される。
    '   ただし、本来 ExecuteScript の完了通知は Wv2Pane が生きている
    '   = UserForm1 が表示中であるべき状況でしか飛んでこないので、
    '   この自動生成パスは現実的には踏まれない。
    UserForm1.OnExecuteScriptResult callbackId, errorCode, resultJson
End Sub


' ============================================================
' ★D-1 の本体: EvalSync とその周辺★
' ============================================================


' ============================================================
' EvalSync (D-1) - JS 式を撃って戻り値を同期的に受け取る
'
'   引数:
'     jsExpr     : 評価する JS ★式★ (文は不可。複数文は即時関数で包む)
'     timeoutSec : 待ち時間の上限 (既定 5 秒)。0 以下なら既定値に丸める。
'
'   戻り値 (As String):
'     成功 : JS の戻り値を JSON.stringify した文字列。LastEvalOk = True
'     失敗 : 空文字。理由は LastEvalError に入る。LastEvalOk = False
'
'   流れ:
'     1. 拒否条件を確認 (Close 済み / View 未取得 / ★in-callback★)
'     2. View_ExecuteScriptSafe で JSON.stringify(式) を撃って callbackId を得る
'     3. 保留テーブルに待機中エントリを登録
'     4. DoEvents を回して完了を待つ (WaitPaneReady と同じ形)
'     5. 完了 → JSON をほどいて返す / 超過 → 破棄済みにして空文字を返す
'
'   ★なぜ JSON.stringify で包むのか (事後報告 (e))★
'     Safe ラッパーは ok と value を返すが、value が数値・オブジェクト・文字列と
'     型が揺れると VBA 側の取り出しが型ごとに分岐してしまう。JSON.stringify を
'     通しておけば value は★常に文字列★になり、限定パーサ 1 本で足りる。
'
'   ★DoEvents ループ中はブレーク/ステップ実行しないこと (仕様事実 20)★
' ============================================================
Public Function EvalSync(ByVal jsExpr As String, _
                         Optional ByVal timeoutSec As Single = 5) As String
    Dim res As String
    Dim cbid As Long
    Dim k As String
    Dim t As Single
    Dim ent As Variant
    Dim timedOut As Boolean
    Dim errCode As Long
    Dim resultJson As String
    Dim v As String
    Dim eName As String
    Dim eMsg As String

    res = ""
    m_lastEvalOk = False
    m_lastEvalError = ""

    ' --- 1. 拒否条件 ---
    If m_closed Then
        m_lastEvalError = "closed"
        Wv2Log.LogW "Wv2Pane.EvalSync: Close 済みのため拒否"
        EvalSync = ""
        Exit Function
    End If

    If m_pView = 0 Then
        m_lastEvalError = "no-view"
        Wv2Log.LogW "Wv2Pane.EvalSync: m_pView が未取得"
        EvalSync = ""
        Exit Function
    End If

    If m_inCallback > 0 Then
        ' ★論点7 のガード★ ここで撃つと仕様事実 14 でタイムアウトまで固まる。
        m_lastEvalError = "in-callback"
        Wv2Log.LogW "Wv2Pane.EvalSync: ★イベントハンドラ内から呼ばれたので拒否★ " & _
                    "(深さ=" & m_inCallback & ")。ハンドラ内で JS の結果が要るときは " & _
                    "SafeTimer で一拍置いてから呼ぶこと (仕様事実 14)。" & _
                    "ハンドラ外なのにこれが出るならカウンタ残留を疑い " & _
                    "ResetCallbackGuard を実行すること。"
        EvalSync = ""
        Exit Function
    End If

    If timeoutSec <= 0 Then timeoutSec = EVAL_TIMEOUT_DEFAULT
    If m_evalPending Is Nothing Then Set m_evalPending = New Collection

    ' --- 2. Safe ラッパー経由で撃つ (論点4) ---
    cbid = View_ExecuteScriptSafe("JSON.stringify(" & jsExpr & ")")
    If cbid <= 0 Then
        m_lastEvalError = "execute-failed(" & cbid & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: ExecuteScript の発行に失敗 rc=" & cbid
        EvalSync = ""
        Exit Function
    End If

    ' --- 3. 保留テーブルに登録 ---
    k = "cb" & cbid
    m_evalPending.Add Array(EVAL_WAITING, 0&, ""), k

    m_evalDepth = m_evalDepth + 1
    If m_evalDepth > 1 Then
        Wv2Log.LogD "Wv2Pane.EvalSync: ネスト実行 深さ=" & m_evalDepth & " (" & k & ")"
    End If

    ' --- 4. 完了待ち (WaitPaneReady と同じ形) ---
    timedOut = False
    t = Timer
    Do
        DoEvents
        ent = m_evalPending(k)
        If ent(0) = EVAL_DONE Then Exit Do
        If (Timer - t) > timeoutSec Then
            timedOut = True
            ' ★論点5★ 消さずに破棄済みで残す。遅延到着を捨てるため。
            m_evalPending.Remove k
            m_evalPending.Add Array(EVAL_DISCARDED, 0&, ""), k
            Exit Do
        End If
    Loop

    m_evalDepth = m_evalDepth - 1

    If timedOut Then
        m_lastEvalError = "timeout(" & timeoutSec & "s)"
        Wv2Log.LogE "Wv2Pane.EvalSync: ★タイムアウト★ " & k & " を破棄済みにした " & _
                    "(遅れて完了通知が来ても捨てる)"
        EvalSync = ""
        Exit Function
    End If

    ' --- 5. 結果をほどく ---
    errCode = ent(1)
    resultJson = ent(2)
    m_evalPending.Remove k

    If errCode <> 0 Then
        m_lastEvalError = "hresult(&H" & Hex$(errCode) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: ExecuteScript が失敗 " & m_lastEvalError
        EvalSync = ""
        Exit Function
    End If

    If InStr(1, resultJson, """ok"":true") > 0 Then
        If Wv2Json.JsonPickStr(resultJson, "value", v) Then
            m_lastEvalOk = True
            res = v
        Else
            ' JSON.stringify(undefined) は undefined を返し、外側の stringify で
            ' value キーごと消える。これは戻り値が undefined だった合図。
            m_lastEvalOk = True
            res = "undefined"
        End If

    ElseIf InStr(1, resultJson, """ok"":false") > 0 Then
        Wv2Json.JsonPickStr resultJson, "errorName", eName
        Wv2Json.JsonPickStr resultJson, "errorMessage", eMsg
        m_lastEvalError = eName & ": " & eMsg
        Wv2Log.LogW "Wv2Pane.EvalSync: JS 例外 " & m_lastEvalError

    Else
        ' Safe ラッパーは必ず ok を返すので、ここに来るのは JS の構文エラー
        ' (ExecuteScript が S_OK + null を返すケース。第9.9c の実測)。
        m_lastEvalError = "syntax-error-or-null (resultJson=" & Left$(resultJson, 60) & ")"
        Wv2Log.LogE "Wv2Pane.EvalSync: " & m_lastEvalError
    End If

    EvalSync = res
End Function


' ============================================================
' LastEvalOk / LastEvalError (D-1) - 直前の EvalSync の結果
'   EvalSync が空文字を返したとき、それが「JS が空文字を返した」のか
'   「失敗した」のかを区別するために使う。
' ============================================================
Public Property Get LastEvalOk() As Boolean
    LastEvalOk = m_lastEvalOk
End Property

Public Property Get LastEvalError() As String
    LastEvalError = m_lastEvalError
End Property


' ============================================================
' InCallbackDepth / ResetCallbackGuard / Debug_SetInCallback (D-1 論点7)
'
'   ResetCallbackGuard:
'     イベントハンドラが想定外の抜け方をしてカウンタが残った場合の緊急脱出口。
'     ラッパーが必ず戻すので通常は不要だが、EvalSync が in-callback で
'     拒否され続けるときはこれを叩く。
'
'   Debug_SetInCallback:
'     ★検証専用★ ガードの発火を Test_D1_Guard から再現するための注入口。
'     本番コードからは呼ばない (Debug_ 接頭辞は Wv2Browser の Debug_* と同じ規約)。
' ============================================================
Public Property Get InCallbackDepth() As Long
    InCallbackDepth = m_inCallback
End Property

Public Sub ResetCallbackGuard()
    Debug.Print "Wv2Pane.ResetCallbackGuard: m_inCallback = " & m_inCallback & " → 0"
    m_inCallback = 0
End Sub

Public Sub Debug_SetInCallback(ByVal n As Long)
    Debug.Print "Wv2Pane.Debug_SetInCallback: " & m_inCallback & " → " & n & " (検証専用)"
    m_inCallback = n
End Sub


' ============================================================
' JsQuote (D-2、Friend) - VBA 文字列を JS の文字列リテラルにする
'
'   ★VBA のソース上にバックスラッシュ + 二重引用符を書かないための唯一の口★
'     Wv2Json.JsonEscape が引用符・バックスラッシュ・制御文字を JSON 仕様で
'     エスケープし、それを Chr$(34) で括る。エスケープ列は★実行時に組み立てられる★
'     ので、VBA のソースには現れない (プロジェクトの JS-in-VBA 規則に適合)。
'
'   例:
'     JsQuote("input[name='q']")   → "input[name='q']"          (そのまま)
'     JsQuote("a[title=""x""]")    → 引用符が JSON 仕様でエスケープされた形になる
'                                    (エスケープ列は★実行時に組み立てられる★)
'     JsQuote("C:\temp")           → "C:\\temp"
'
'   Friend にしているのは Wv2Element.GetAttribute が属性名を包むのに使うため。
' ============================================================
Friend Function JsQuote(ByVal s As String) As String
    JsQuote = Chr$(34) & Wv2Json.JsonEscape(s) & Chr$(34)
End Function


' ============================================================
' JsRegistryPrologue (D-2、Private) - レジストリの自己生成コード
'
'   D-2 の全 JS 式の先頭に置く共通部分 (論点1 案c)。
'   無ければ作り、変数 R にレジストリを束縛する。
'   ★ここで作られた場合、世代は新しい値になる★ = 既存ハンドルは全て stale。
'
'   ★JS の文字列リテラルはシングルクォート★ (プロジェクト規則)。
' ============================================================
Private Function JsRegistryPrologue() As String
    JsRegistryPrologue = _
        "if(typeof window.__wv2==='undefined'||!window.__wv2){" & _
        "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};}" & _
        "var R=window.__wv2;"
End Function


' ============================================================
' GetElementById / QuerySelector (D-2) - 要素を 1 個探す
'
'   見つかれば Wv2Element を、見つからなければ Nothing を返す (論点4 案a)。
'
'   ★Nothing の意味は 2 通り★
'     Nothing + LastEvalOk = True  … 本当に存在しない (正常)
'     Nothing + LastEvalOk = False … 失敗。LastEvalError に理由
'
'   使用例:
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     If el Is Nothing Then
'         Debug.Print "無い (ok=" & p.LastEvalOk & " err=" & p.LastEvalError & ")"
'     Else
'         Debug.Print el.TagName & " / " & el.Value
'     End If
' ============================================================
'   ★D-3 から ImplicitWaitSec > 0 なら見つかるまでリトライする★
'     既定は 0 なので、設定しない限り D-2 と同じ「1 往復で答える」ままである。
Public Function GetElementById(ByVal elementId As String) As Wv2Element
    Set GetElementById = FindWithWait("document.getElementById(" & JsQuote(elementId) & ")", _
                                      "GetElementById", elementId, m_implicitWaitSec)
End Function

Public Function QuerySelector(ByVal selector As String) As Wv2Element
    Set QuerySelector = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                                     "QuerySelector", selector, m_implicitWaitSec)
End Function


' ============================================================
' FindOne (D-2、Private) - 検索の共通実装
'
'   finderExpr : 要素を 1 個返す JS 式 (見つからなければ null を返すこと)
'   apiName / arg : ログ用
'
'   JS 側は次を返す:
'     {g:'<世代>', found:true,  h:<添字>}   見つかった
'     {g:'<世代>', found:false, h:0}        見つからなかった
'
'   ★h に負値を使わない理由★ Wv2Json.JsonGetNum は負数を扱わない (設計原則64)。
'     found フラグで分けるので h は常に 0 以上でよい。
' ============================================================
'   quiet (D-3): True なら「見つからない」のログを出さない。ポーリング中に
'                1 往復ごとに出ると、イミディエイトが埋まってしまうため。
Private Function FindOne(ByVal finderExpr As String, _
                         ByVal apiName As String, _
                         ByVal arg As String, _
                         Optional ByVal quiet As Boolean = False) As Wv2Element
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim h As Long
    Dim el As Wv2Element

    Set FindOne = Nothing

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=" & finderExpr & ";" & _
         "if(!e){return {g:R.gen,found:false,h:0};}" & _
         "R.els.push(e);" & _
         "return {g:R.gen,found:true,h:R.els.length-1};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane." & apiName & ": 失敗 (" & arg & ") err=" & m_lastEvalError
        Exit Function
    End If

    ' 世代キャッシュを更新する (成功したときだけ)
    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """found"":true") = 0 Then
        ' 見つからなかった。★これは正常★ なので m_lastEvalOk は True のまま。
        If Not quiet Then
            Debug.Print "Wv2Pane." & apiName & ": 見つからない (" & arg & ")"
        End If
        Exit Function
    End If

    h = Wv2Json.JsonGetNum(res, "h")

    Set el = New Wv2Element
    el.InitElement Me, h, gen
    Set FindOne = el

    Debug.Print "Wv2Pane." & apiName & ": 取得 (" & arg & ") handle=" & h & " gen=" & gen
End Function


' ============================================================
' CurrentDomGen (D-2) - 世代キャッシュの読み取り口
'
'   ★これは権威ではない★ 「JS から最後に聞いた値」でしかない。
'   stale 判定の高速パスと、検証・デバッグ表示に使う。
' ============================================================
Public Property Get CurrentDomGen() As String
    CurrentDomGen = m_domGen
End Property


' ============================================================
' ClearElementRegistry (D-2、論点3) - レジストリを手動で作り直す
'
'   JS 側のレジストリを新しい世代で作り直す。
'   ★既存の Wv2Element はすべて stale になる。★
'
'   同一ページ内で大量に検索して els が伸びたときの掃除用。通常は不要
'   (ページ遷移でまるごと消えるため)。
'
'   戻り値: True = 成功。False = 失敗 (理由は LastEvalError)
' ============================================================
Public Function ClearElementRegistry() As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    js = "(function(){" & _
         "window.__wv2={gen:'g'+Date.now()+'-'+Math.random(),els:[]};" & _
         "return window.__wv2.gen;})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogE "Wv2Pane.ClearElementRegistry: 失敗 err=" & m_lastEvalError
        ClearElementRegistry = False
        Exit Function
    End If

    gen = Wv2Json.JsonUnescape(res)
    m_domGen = gen
    Wv2Log.LogD "Wv2Pane.ClearElementRegistry: レジストリを作り直した (新しい世代=" & gen & ")"
    ClearElementRegistry = True
End Function


' ============================================================
' Elem_Read (D-2、Friend) - Wv2Element のプロパティ 1 個を読む
'
'   引数:
'     elemHandle : JS 側 R.els[] の添字
'     elemGen    : 要素を取得したときの世代
'     valueExpr  : 変数 e (対象要素) を使う JS 式。例 "e.innerText"
'     outVal     : 読み取れた値 (エスケープ復号済み)
'     outStatus  : 失敗理由。成功時は空文字
'   戻り値:
'     True  = 読み取り成功 (値が null / undefined だった場合も True。outVal は空文字)
'     False = stale または失敗。理由は outStatus
'
'   JS 側は次を返す:
'     {s:'gen', g:'<世代>'}                 世代が違う = stale
'     {s:'nil', g:'<世代>'}                 添字に要素が無い = stale 扱い
'     {s:'ok',  g:'<世代>', has:false, v:''} 値が null / undefined
'     {s:'ok',  g:'<世代>', has:true,  v:'<値>'}
'
'   ★値は必ず String(x) で文字列に正規化してから返す★
'     JSON の型のまま返すと VBA 側で型ごとの取り出しが要る。読み取り API の
'     戻り値は全て String なので、JS 側で寄せた方が単純 (仕様事実31 の undefined
'     消失問題も踏まない)。
' ============================================================
Friend Function Elem_Read(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal valueExpr As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Read = False

    ' --- 高速パス: 手元の世代がキャッシュと違えば、往復せずに stale と確定できる ---
    '     (キャッシュが古い場合は一致側に倒れて必ず問い合わせるので、誤判定はしない)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var x=" & valueExpr & ";" & _
         "if(x===null||x===undefined){return {s:'ok',g:R.gen,has:false,v:''};}" & _
         "return {s:'ok',g:R.gen,has:true,v:String(x)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    If InStr(1, res, """has"":true") = 0 Then
        ' 値が null / undefined。★エラーではない★ (属性が無い / value を持たない要素)
        outVal = ""
        Elem_Read = True
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Read = True
End Function


' ============================================================
' Elem_CheckStale (D-2、Friend) - ハンドルが無効になっていないか
'
'   戻り値: True = stale (もう使えない)
'   outStatus: 判定に失敗した場合の理由。成功時は空文字
'
'   ★判定できなかった場合も True (安全側) を返す。★
'     区別が必要な呼び出し側は outStatus を見ること。
' ============================================================
Friend Function Elem_CheckStale(ByVal elemHandle As Long, _
                                ByVal elemGen As String, _
                                ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outStatus = ""
    Elem_CheckStale = True

    ' 高速パス (Elem_Read と同じ理屈)
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then Exit Function
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "var ok=(R.gen===" & JsQuote(elemGen) & ")&&(!!R.els[" & CStr(elemHandle) & "]);" & _
         "return {s:ok?'ok':'stale',g:R.gen};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    Elem_CheckStale = (InStr(1, res, """s"":""ok""") = 0)
End Function


' ============================================================
' Elem_Exec (D-3、Friend) - 要素への書き込み・操作を 1 回撃つ
'
'   引数:
'     elemHandle  : JS 側 R.els[] の添字
'     elemGen     : 要素を取得したときの世代
'     actionStmts : 変数 e (対象要素) を使う JS ★文★ (末尾のセミコロン込み)。
'                   結果の手掛かりを返したいときは変数 rv に文字列を入れる。
'     outVal      : rv の内容 (経路の名前)。使わない操作では空文字
'     outStatus   : 失敗理由。成功時は空文字
'   戻り値:
'     True  = 操作が JS 側で最後まで走った
'     False = stale または失敗。理由は outStatus
'
'   ★成功 = 「JS が例外なく走った」であって「意図した効果が出た」ではない★
'     例えば disabled なボタンへの Click は例外にならないが何も起きない。
'     効果を確かめたいときは読み取り (Elem_Read) で確認すること。
'
'   高速パスと世代キャッシュの扱いは Elem_Read と同じ (安全側にしか倒れない)。
' ============================================================
Friend Function Elem_Exec(ByVal elemHandle As Long, _
                          ByVal elemGen As String, _
                          ByVal actionStmts As String, _
                          ByRef outVal As String, _
                          ByRef outStatus As String) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    outVal = ""
    outStatus = ""
    Elem_Exec = False

    ' --- 高速パス: 手元の世代がキャッシュと違えば、往復せずに stale と確定できる ---
    If Len(m_domGen) > 0 Then
        If elemGen <> m_domGen Then
            outStatus = "stale"
            Exit Function
        End If
    End If

    js = "(function(){" & JsRegistryPrologue() & _
         "if(R.gen!==" & JsQuote(elemGen) & "){return {s:'gen',g:R.gen};}" & _
         "var e=R.els[" & CStr(elemHandle) & "];" & _
         "if(!e){return {s:'nil',g:R.gen};}" & _
         "var rv='';" & actionStmts & _
         "return {s:'ok',g:R.gen,v:String(rv)};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        outStatus = m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    If InStr(1, res, """s"":""ok""") = 0 Then
        outStatus = "stale"
        Exit Function
    End If

    Wv2Json.JsonPickStr res, "v", outVal
    Elem_Exec = True
End Function

' ============================================================
' ★D-3b の本体: DOM 条件待ち★
'
'   未知1 の実測 (2026-08-22) で「ExecuteScript は Promise を待たない」が確定した
'   ため、待ちは論点3 の骨格どおり★VBA 側のポーリング★で実装する。
'   JS 側に仕掛けを置く案 (postMessage で完了を通知する) は採らない。保留テーブルが
'   増えるうえ、D-1 で作った EvalSync の型から外れるため。
' ============================================================


' ============================================================
' ImplicitWaitSec (D-3、論点2 案C の opt-in 部分)
'
'   0 より大きくすると、QuerySelector / GetElementById が★見つかるまで★
'   リトライするようになる (SeleniumVBA の ImplicitMaxWait と同じ流儀)。
'
'   ★既定は 0 (無効)★ 既定で待たせない理由はヘッダー参照。
'   WaitFor / WaitGone は明示的な待ちなので★この設定の影響を受けない★。
'
'   使用例:
'     p.ImplicitWaitSec = 5      ' 以後の検索は最大 5 秒粘る
'     p.ImplicitWaitSec = 0      ' 元に戻す
' ============================================================
Public Property Get ImplicitWaitSec() As Single
    ImplicitWaitSec = m_implicitWaitSec
End Property

Public Property Let ImplicitWaitSec(ByVal waitSec As Single)
    If waitSec < 0 Then waitSec = 0
    m_implicitWaitSec = waitSec
    Wv2Log.LogD "Wv2Pane.ImplicitWaitSec = " & waitSec & " 秒"
End Property


' ============================================================
' WaitFor (D-3、論点2 案C の第一級 API)
'   セレクタに一致する要素が★現れるまで★待つ。
'
'   引数:
'     selector   : CSS セレクタ
'     timeoutSec : 上限 (秒)。既定 10。0 以下なら既定値に丸める
'   戻り値:
'     見つかった            → Wv2Element
'     時間内に現れなかった  → ★Nothing + LastEvalOk = True★ (正常)
'     失敗 (セレクタ不正等) → ★Nothing + LastEvalOk = False★ (LastEvalError に理由)
'
'   ★存在だけを見る (論点5 案A)★
'     display:none でも「現れた」と判定する。SPA では DOM にあるが不可視、が
'     普通にあるため。可視まで待ちたくなったら WaitVisible を別に用意する。
'
'   ★失敗したら待たずに即座に返す★
'     in-callback (仕様事実14) やセレクタの構文エラーは、待っても直らない。
'     ここで粘ると、イベントハンドラ内から呼んだときに timeoutSec ぶん固まる。
'
'   ★待っている間はブレーク/ステップ実行しないこと★ (仕様事実20)
'
'   ★タブを開いた直後に呼ばないこと★
'     View がまだ無いうちは EvalSync が no-view で失敗し、上の方針により
'     ★待たずに即座に Nothing を返す★。ページを開いた直後は、まず
'     EvalSync が通るようになるまで待つこと (検証では D4WaitPane が実例)。
'
'   使用例:
'     Dim el As Wv2Element
'     Set el = p.WaitFor("#result", 5)
'     If el Is Nothing Then
'         Debug.Print IIf(p.LastEvalOk, "現れなかった", "失敗: " & p.LastEvalError)
'     End If
' ============================================================
Public Function WaitFor(ByVal selector As String, _
                        Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Wv2Element
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    Set WaitFor = FindWithWait("document.querySelector(" & JsQuote(selector) & ")", _
                               "WaitFor", selector, timeoutSec)
End Function


' ============================================================
' WaitGone (D-3)
'   セレクタに一致する要素が★無くなるまで★待つ。
'
'   戻り値:
'     True                  … 消えた (または最初から 1 個も無かった)
'     False + LastEvalOk=True  … 時間内に消えなかった (正常な「まだ居る」)
'     False + LastEvalOk=False … 失敗 (理由は LastEvalError)
'
'   ★存在確認にレジストリを使わない★ (D-3 の事後報告)
'     FindOne を回すと、まだ消えていない要素をポーリングのたびに R.els へ
'     push してしまう (10 秒待てば 70 個ほど積む)。WaitGone は掴む必要が
'     無いので、found の真偽だけを返す専用の式 (ExistsNow) を使う。
'
'   使用例:
'     If Not p.WaitGone(".spinner", 10) Then Debug.Print "まだ回っている"
' ============================================================
Public Function WaitGone(ByVal selector As String, _
                         Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT) As Boolean
    Dim t0 As Single
    Dim found As Boolean
    Dim ok As Boolean

    WaitGone = False
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT

    t0 = Timer
    Do
        found = ExistsNow(selector, ok)

        If Not ok Then
            ' 失敗は待っても直らない。LastEvalOk = False のまま返す。
            Debug.Print "Wv2Pane.WaitGone: 失敗 (" & selector & ") err=" & m_lastEvalError
            Exit Function
        End If

        If Not found Then
            WaitGone = True
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            Debug.Print "Wv2Pane.WaitGone: 時間切れ (" & selector & ") " & _
                        timeoutSec & " 秒たっても消えなかった"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' FindWithWait (D-3、Private) - 見つかるまで FindOne を繰り返す
'
'   waitSec <= 0 なら 1 回だけ試す (= D-2 までの挙動そのまま)。
'   ★見つからない (Nothing + LastEvalOk=True) ときだけ粘る★
'   失敗 (LastEvalOk=False) は即座に諦める。
' ============================================================
Private Function FindWithWait(ByVal finderExpr As String, _
                              ByVal apiName As String, _
                              ByVal arg As String, _
                              ByVal waitSec As Single) As Wv2Element
    Dim el As Wv2Element
    Dim t0 As Single

    t0 = Timer
    Do
        Set el = FindOne(finderExpr, apiName, arg, (waitSec > 0))

        If Not el Is Nothing Then
            Set FindWithWait = el
            Exit Function
        End If

        If Not m_lastEvalOk Then Exit Function
        If waitSec <= 0 Then Exit Function
        If ElapsedSec(t0) >= waitSec Then
            Debug.Print "Wv2Pane." & apiName & ": 時間切れ (" & arg & ") " & _
                        waitSec & " 秒たっても現れなかった"
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' ExistsNow (D-3、Private) - 要素があるかだけを 1 往復で調べる
'
'   ★レジストリに登録しない★ WaitGone のポーリング用 (上のコメント参照)。
'   outOk : JS の実行自体が成功したか。False なら戻り値に意味は無い。
' ============================================================
Private Function ExistsNow(ByVal selector As String, ByRef outOk As Boolean) As Boolean
    Dim js As String
    Dim res As String
    Dim gen As String

    ExistsNow = False
    outOk = False

    js = "(function(){" & JsRegistryPrologue() & _
         "var e=document.querySelector(" & JsQuote(selector) & ");" & _
         "return {g:R.gen,found:!!e};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then Exit Function

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    outOk = True
    ExistsNow = (InStr(1, res, """found"":true") > 0)
End Function


' ============================================================
' PollPause / ElapsedSec (D-3、Private) - ポーリングの足踏みと経過時間
'
'   PollPause : POLL_INTERVAL_SEC (100ms) だけ DoEvents を回す。
'               ★Sleep を使わない★ WebView2 の完了通知を回すため、
'               この間もメッセージポンプを止めてはいけない。
'
'   ElapsedSec: ★Timer は深夜 0 時に 0 へ戻る★ ので、素の引き算だと負になり
'               待ちループが永久に抜けなくなる。またいだ分を足して補正する
'               (D-3 の事後報告。実機検証が 1:53 に走ったので他人事ではない)。
' ============================================================
Private Sub PollPause()
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        If ElapsedSec(t0) >= POLL_INTERVAL_SEC Then Exit Do
    Loop
End Sub

Private Function ElapsedSec(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    ElapsedSec = d
End Function

' ============================================================
' ★D-4a の本体: ページ内 SPA プローブ★
' ============================================================


' ============================================================
' SpaProbeState (D-4a) - プローブを整えてから状態を読む
'
'   戻り値: 状態 JSON (成功時)。失敗時は空文字 + LastEvalOk = False
'
'   JSON の中身:
'     v  版番号        h  健康か (True/False)   g  設置世代    rp 作り直した回数
'     f  実行中 fetch  x  実行中 XHR
'     q  最後の DOM 変化からの経過 ms   qn 最後の通信完了からの経過 ms
'     m  数えた DOM 変化  sm 除外した DOM 変化
'     n  数えた通信        sn 除外した通信
'     up プローブ設置からの経過 ms      ob MutationObserver が生きているか
'     sda DOM シグナルを arm 済みか     sdh その当たり回数
'     sna 通信シグナルを arm 済みか     snh その当たり回数
'     sua URL シグナルを arm 済みか      suh 今それに一致しているか
'     lgN 診断ログの件数                 lgOn 診断ログが有効か
'     il  ★今飛んでいる要求★ (最大 5 件。"xhr 1234ms https://..." の形)
'     ifn 飛んでいる要求のうち★新しいもの★の数 (静穏判定はこれを見る)
'     ifo 同じく★居座っているもの★の数 (StaleInflightMs を超えたもの)
'
'   ★呼ぶたびに健康診断が走り、壊れていれば作り直す★
'   したがって「まず Ensure、次に Read」の 2 往復は要らない。
'
'   使用例:
'     Debug.Print p.SpaProbeState
'     → {""v"":1,""h"":true,""g"":1,...}
' ============================================================
Public Function SpaProbeState() As String
    Dim js As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {v:P.v,h:HP(P),g:P.gen,rp:P.repairs," & _
         "f:P.fetchN,x:P.xhrN," & _
         "q:Math.round(N()-P.lastMut),qn:Math.round(N()-P.lastNet)," & _
         "m:P.mutN,sm:P.skipMut,n:P.netN,sn:P.skipNet," & _
         "up:Math.round(N()-P.since),ob:!!P.ob," & _
         "sda:!!P.sdArmed,sdh:P.sdHits,sna:!!P.snArmed,snh:P.snHits," & _
         "sua:!!P.suArmed," & _
         "suh:(P.suArmed&&String(location.href).indexOf(P.suPat)>=0)?1:0," & _
         "lgN:P.log.length,lgOn:!!P.logOn," & _
         "il:P.ifl.map(function(o){" & _
         "return o.w+' '+Math.round(N()-o.t)+'ms '+o.u;}).slice(0,5).join(' # ')," & _
         "ifn:P.ifl.filter(function(o){return (N()-o.t)<" & StaleMsJs() & ";}).length," & _
         "ifo:P.ifl.filter(function(o){return (N()-o.t)>=" & StaleMsJs() & ";}).length};})()"

    SpaProbeState = EvalSync(js)

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeState: 失敗 err=" & m_lastEvalError
        SpaProbeState = ""
    End If
End Function


' ============================================================
' SpaProbeHealthy (D-4a) - 健康かどうかだけを見る薄い口
'   ★判定できなかった場合は False を返す★ (安全側)。区別が要るなら LastEvalOk。
' ============================================================
Public Function SpaProbeHealthy() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Function

    SpaProbeHealthy = (InStr(1, st, """h"":true") > 0)
End Function


' ============================================================
' SpaProbeReset (D-4a) - プローブを捨てて作り直す
'
'   ★通常は不要★ 健康診断が壊れを検出して自動で作り直すため。
'   数えた値をゼロに戻したいときと、健康診断そのものを検証したいときに使う。
'
'   ★元の fetch / XHR には戻さない★ ラップを外すと、フレームワークが
'   こちらのラッパーを掴んでいた場合に壊れる。直後に現在の window.fetch を包み直す。
'
'   ★実装は「版番号を壊してから作り直させる」形★ (window.__wv2p を null に
'   しない)。こうすると世代と作り直し回数が増え続け、除外リストも引き継がれる。
'   null にすると世代が 1 に戻ってしまい、いつ作り直したのかが追えなくなる。
' ============================================================
Public Function SpaProbeReset() As Boolean
    Dim js As String

    js = "(function(){var o=window.__wv2p;if(o){o.v=-1;}" & SpaProbeJs() & _
         "return {g:P.gen,h:HP(P)};})()"

    EvalSync js
    SpaProbeReset = m_lastEvalOk

    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeReset: 失敗 err=" & m_lastEvalError
    Else
        Wv2Log.LogD "Wv2Pane.SpaProbeReset: プローブを作り直した"
    End If
End Function


' ============================================================
' SpaProbeJs (D-4a、Private) - プローブを整える JS (自己生成プロローグの拡張)
'
'   実行すると、変数 P にプローブが束縛された状態になる。
'   無ければ作り、★壊れていれば作り直す★。健康診断は HP(P) で呼べる。
'
'   ★JS の文字列リテラルはシングルクォート★ (プロジェクト規則)。
'   ★数える対象★
'     fetch          … 呼び出しで +1、解決/失敗で -1
'     XHR            … send で +1、loadend で -1 (二重減算を __wv2d で防ぐ)
'     MutationObserver … childList / characterData / subtree / attributes
'   除外リスト (ignoreNet / ignoreSel) は★参照を毎回読む★ので、
'   作り直さずに差し替えられる (D-4b でその口を足す)。
' ============================================================
Private Function SpaProbeJs() As String
    Dim js As String

    js = "var V=" & CStr(SPA_PROBE_VER) & ",N=function(){return performance.now();};"
    js = js & "var XP=(window.XMLHttpRequest?XMLHttpRequest.prototype:null);"
    js = js & "var HP=function(p){return !!p&&p.v===V" & _
                      "&&(typeof window.fetch!=='function'||window.fetch===p.fw)" & _
                      "&&(!XP||(XP.open===p.xo&&XP.send===p.xs))&&!!p.ob;};"
    js = js & "var P=window.__wv2p;"
    js = js & "if(!HP(P)){var prev=P;"

    ' --- 新しいプローブの器 (除外リストは引き継ぐ) ---
    js = js & "var q={v:V,gen:(prev?((prev.gen||0)+1):1)," & _
                      "repairs:(prev?((prev.repairs||0)+1):0)," & _
                      "since:N(),lastMut:N(),lastNet:N()," & _
                      "fetchN:0,xhrN:0,mutN:0,netN:0,skipMut:0,skipNet:0," & _
                      "sdArmed:false,sdSel:'',sdHits:0,sdObs:null," & _
                      "snArmed:false,snPat:'',snHits:0," & _
                      "suArmed:false,suPat:''," & _
                      "ifl:[],log:[],logOn:(prev&&prev.logOn?true:false)," & _
                      "ignoreNet:(prev&&prev.ignoreNet?prev.ignoreNet:[])," & _
                      "ignoreSel:(prev&&prev.ignoreSel?prev.ignoreSel:'')};"

    ' --- 除外の判定 (毎回 q から読むので後から差し替えられる) ---
    js = js & "var ig=function(u){var a=q.ignoreNet,i;" & _
                      "for(i=0;i<a.length;i++){if(u.indexOf(a[i])>=0){return true;}}return false;};"
    js = js & "var sk=function(t){if(!q.ignoreSel){return false;}" & _
                      "try{var e=(t&&t.nodeType===1)?t:(t?t.parentElement:null);" & _
                      "return !!(e&&e.closest(q.ignoreSel));}catch(x){return false;}};"

    ' --- D-4e: in-flight の台帳と診断ログ ---
    '   ifl は「今飛んでいる要求」の一覧。★居座る要求の正体を見るため★
    '   log は最近の出来事のリングバッファ (既定は無効。logOn で入れる)。
    js = js & "var lg=function(k,d){if(!q.logOn){return;}" & _
                      "q.log.push(Math.round(N())+' '+k+' '+String(d).slice(0,120));" & _
                      "if(q.log.length>200){q.log.shift();}};q.lg=lg;"
    js = js & "var ifA=function(u,w){var o={u:String(u).slice(0,150),t:N(),w:w};" & _
                      "q.ifl.push(o);lg('NET+',w+' '+o.u);return o;};"
    js = js & "var ifD=function(o){var i=q.ifl.indexOf(o);" & _
                      "if(i>=0){q.ifl.splice(i,1);}" & _
                      "lg('NET-',Math.round(N()-o.t)+'ms '+o.u);};"

    ' --- D-4c: 明示シグナルの当たり判定と arm 系 ---
    '   ★当たったら lastMut を今にする★ = 証拠の後にもう一度まるごと静穏窓を取る
    '   (参考プロダクトと同じ考え方。1 ポーリング分で満足しない)。
    js = js & "var sn=function(u){if(q.snArmed&&q.snPat&&" & _
                      "String(u).indexOf(q.snPat)>=0){q.snHits++;q.lastMut=N();" & _
                      "lg('SIG','net hit '+String(u).slice(0,80));}};q.sn=sn;"
    js = js & "q.armUrl=function(pat){q.suPat=pat;q.suArmed=true;return 0;};"
    js = js & "q.armDom=function(sel){var n;" & _
                      "try{n=document.querySelector(sel);}catch(e){return -2;}" & _
                      "if(!n){return -1;}" & _
                      "if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdSel=sel;q.sdHits=0;q.sdArmed=true;" & _
                      "q.sdObs=new MutationObserver(function(ms){" & _
                      "q.sdHits+=ms.length;q.lastMut=N();" & _
                      "lg('SIG','dom hit x'+ms.length);});" & _
                      "q.sdObs.observe(n," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "return 0;};"
    js = js & "q.armNet=function(pat){q.snPat=pat;q.snHits=0;q.snArmed=true;return 0;};"
    js = js & "q.disarm=function(){if(q.sdObs){try{q.sdObs.disconnect();}catch(e){}}" & _
                      "q.sdObs=null;q.sdArmed=false;q.snArmed=false;q.suArmed=false;" & _
                      "return 0;};"

    ' --- fetch のラップ (★現在の window.fetch を包む★) ---
    js = js & "if(typeof window.fetch==='function'){"
    js = js & "var of=(prev&&prev.of&&prev.fw===window.fetch)?prev.of:window.fetch;q.of=of;"
    '   ★除外した要求でも明示シグナルの判定はする★
    '   除外は「静穏判定から外す」だけの意味で、arm した目印は別扱いにする。
    js = js & "q.fw=function(u,o){" & _
                      "var r=(typeof u==='string')?u:((u&&u.url)?u.url:'');" & _
                      "var sp=ig(r);if(sp){q.skipNet++;}else{q.netN++;q.fetchN++;}" & _
                      "var en=sp?null:ifA(r,'fetch');" & _
                      "var d=function(){if(!sp){q.fetchN--;q.lastNet=N();ifD(en);}sn(r);};" & _
                      "return of.apply(this,arguments).then(" & _
                      "function(v){d();return v;},function(e){d();throw e;});};"
    js = js & "window.fetch=q.fw;}"

    ' --- XHR のラップ ---
    js = js & "if(XP){"
    js = js & "var oo=(prev&&prev.oo&&prev.xo===XP.open)?prev.oo:XP.open;"
    js = js & "var os=(prev&&prev.os&&prev.xs===XP.send)?prev.os:XP.send;q.oo=oo;q.os=os;"
    js = js & "q.xo=function(m,u){try{this.__wv2u=String(u);}catch(e){}" & _
                      "return oo.apply(this,arguments);};"
    js = js & "q.xs=function(){var s=this;try{var u=s.__wv2u||'';" & _
                      "var sp=ig(u);if(sp){q.skipNet++;}else{q.netN++;q.xhrN++;}" & _
                      "var en=sp?null:ifA(u,'xhr');" & _
                      "s.addEventListener('loadend',function(){" & _
                      "if(!s.__wv2d){s.__wv2d=true;" & _
                      "if(!sp){q.xhrN--;q.lastNet=N();ifD(en);}sn(u);}});}catch(e){}" & _
                      "return os.apply(this,arguments);};"
    js = js & "XP.open=q.xo;XP.send=q.xs;}"

    ' --- DOM 変化の観測 ---
    js = js & "try{q.ob=new MutationObserver(function(ms){" & _
                      "var i,c=0,t;for(i=0;i<ms.length;i++){t=ms[i].target;" & _
                      "if(ms[i].type==='characterData'){t=(t&&t.parentElement)?t.parentElement:t;}" & _
                      "if(sk(t)){q.skipMut++;}else{c++;}}" & _
                      "if(c>0){q.mutN+=c;q.lastMut=N();lg('DOM','x'+c);}});"
    js = js & "q.ob.observe(document," & _
                      "{childList:true,characterData:true,subtree:true,attributes:true});" & _
                      "}catch(e){q.ob=null;}"

    js = js & "window.__wv2p=q;P=q;}"

    SpaProbeJs = js
End Function

' ============================================================
' ★D-4b の本体: 静穏待ち★
' ============================================================


' ============================================================
' WaitSettled (D-4b) - 通信と DOM の変化が落ち着くまで待つ
'
'   引数:
'     timeoutSec : 上限 (秒)。既定 10。0 以下なら既定値に丸める
'     stableMs   : 静穏窓 (ms)。既定 500。この時間 DOM が動かなければ静穏とみなす
'   戻り値:
'     True                     … 静穏に達した
'     False + LastEvalOk=True  … 時間内に静まらなかった (正常)
'     False + LastEvalOk=False … 失敗 (理由は LastEvalError)
'
'   ★★★ True は「アプリの処理が終わった」証明ではない ★★★
'     静かなだけ。無視されて静かなのか終わって静かなのかは区別できない。
'     重要な操作では D-4c の明示シグナルを併用すること。
'
'   ★待っている間はブレーク/ステップ実行しないこと★ (仕様事実20)
'     10 秒待てば 10 秒間ずっと DoEvents 区間にいる。
'
'   使用例:
'     el.Click
'     If p.WaitSettled(10) Then
'         Set r = p.QuerySelector("#result")
'     Else
'         Debug.Print "落ち着かなかった: " & p.LastSettleInfo
'     End If
' ============================================================
Public Function WaitSettled(Optional ByVal timeoutSec As Single = WAIT_TIMEOUT_DEFAULT, _
                            Optional ByVal stableMs As Long = SETTLE_STABLE_MS_DEFAULT) As Boolean
    Dim t0 As Single
    Dim st As String
    Dim inflight As Long
    Dim quietMs As Long
    Dim quietEff As Long
    Dim elapsedMs As Long
    Dim polls As Long
    Dim sigOk As Boolean
    Dim inflightList As String
    Dim stale As Long

    WaitSettled = False
    m_lastSettle = ""
    If timeoutSec <= 0 Then timeoutSec = WAIT_TIMEOUT_DEFAULT
    If stableMs < 0 Then stableMs = SETTLE_STABLE_MS_DEFAULT

    t0 = Timer
    Do
        polls = polls + 1
        st = SpaProbeState()

        If Not m_lastEvalOk Then
            ' 失敗は待っても直らない (設計原則104)。
            m_lastSettle = "error " & m_lastEvalError
            Wv2Log.LogW "Wv2Pane.WaitSettled: 失敗 err=" & m_lastEvalError
            m_sigDom = False
            m_sigNet = False
            m_sigUrl = False
            Exit Function
        End If

        ' ★居座っている要求は数えない★ (D-4f)
        '   ifn = 飛んでいる要求のうち StaleInflightMs 以内に始まったものの数。
        '   ロングポーリングやストリーミングを待ち続けないための現実的な妥協で、
        '   ★「静かに見える」を判定しているのであって「全部終わった」ではない★
        '   という WaitSettled の意味とは矛盾しない。正確さが要るなら
        '   AddIgnoreNetwork か明示シグナルを使う。
        inflight = Wv2Json.JsonGetNum(st, "ifn")
        stale = Wv2Json.JsonGetNum(st, "ifo")
        quietMs = Wv2Json.JsonGetNum(st, "q")

        ' --- D-4c: arm したシグナルの状態を見る ---
        '   ★arm したはずなのにページ側が arm していない = 遷移などで消えた★
        '   黙って再 arm も無視もせず、失敗として返す (fail-visible)。
        If m_sigDom Or m_sigNet Or m_sigUrl Then
            If (m_sigDom And InStr(1, st, """sda"":true") = 0) Or _
               (m_sigNet And InStr(1, st, """sna"":true") = 0) Or _
               (m_sigUrl And InStr(1, st, """sua"":true") = 0) Then
                m_lastSettle = "signal-lost (遷移などで arm が消えた)"
                Wv2Log.LogW "Wv2Pane.WaitSettled: ★signal-lost★ " & _
                            "arm したシグナルがページ側から消えている。" & _
                            "遷移をまたいだ可能性がある (arm し直すこと)"
                DisarmSignals
                Exit Function
            End If
        End If

        sigOk = True
        If m_sigDom Then
            If Wv2Json.JsonGetNum(st, "sdh") = 0 Then sigOk = False
        End If
        If m_sigNet Then
            If Wv2Json.JsonGetNum(st, "snh") = 0 Then sigOk = False
        End If
        If m_sigUrl Then
            If Wv2Json.JsonGetNum(st, "suh") = 0 Then sigOk = False
        End If

        ' --- ★静穏窓は「呼び出した時点」から測る★ (D-4b の事後報告) ---
        '   ページ側の q は「最後の DOM 変化からの経過」なので、呼び出す前から
        '   静かだった場合は最初の 1 往復でいきなり条件を満たしてしまう。それだと
        '   ★「まだ始まっていない」と「もう終わった」を区別できない★
        '   (操作の効果が数十 ms 遅れて始まるのは普通にある)。
        '   そこで呼び出しからの経過で頭打ちにし、★最低でも stableMs は見張る★。
        '   この間に何か動けば q が戻るので、始まりが遅れても捕まえられる。
        '   ただし stableMs より後に始まる処理は原理的に取り逃す。
        '   そこは D-4c の明示シグナル (arm) の担当。
        elapsedMs = CLng(ElapsedSec(t0) * 1000)
        quietEff = quietMs
        If quietEff > elapsedMs Then quietEff = elapsedMs

        If inflight = 0 And quietEff >= stableMs And sigOk Then
            m_lastSettle = "stable q=" & quietMs & "ms(eff=" & quietEff & ") slack=" & _
                           (quietEff - stableMs) & "ms polls=" & polls & _
                           " elapsed=" & Format$(ElapsedSec(t0), "0.00") & "s" & _
                           IIf(stale > 0, " (居座り " & stale & " 件は除外)", "") & _
                           SigInfoOf(st)
            WaitSettled = True
            DisarmSignals
            Exit Function
        End If

        If ElapsedSec(t0) >= timeoutSec Then
            ' ★時間切れの原因が通信なら、何が飛んでいるかを見せる★ (D-4d の教訓)
            '   Maps では終わらない XHR が 1 本居座って静穏が来なかった。
            '   URL が分かれば AddIgnoreNetwork で外せる。
            Wv2Json.JsonPickStr st, "il", inflightList

            m_lastSettle = "timeout inflight=" & inflight & _
                           IIf(stale > 0, "(居座り " & stale & " 件は除外)", "") & _
                           " q=" & quietMs & _
                           "ms (stableMs=" & stableMs & ") polls=" & polls & _
                           SigInfoOf(st)
            If Len(inflightList) > 0 Then
                m_lastSettle = m_lastSettle & " | 実行中: " & inflightList
            End If
            Wv2Log.LogD "Wv2Pane.WaitSettled: 時間切れ " & m_lastSettle
            DisarmSignals
            Exit Function
        End If

        PollPause
    Loop
End Function


' ============================================================
' LastSettleInfo (D-4b) - 直前の待ちの内訳
'
'   成功時: stable q=<静穏 ms> slack=<余裕 ms> polls=<往復回数> elapsed=<秒>
'   失敗時: timeout inflight=<実行中の通信> q=<静穏 ms> ... / error <理由>
'
'   ★slack が小さい (100ms 未満) 待ちは「際どい」★ 参考プロダクトの知見。
'   たまたま静穏窓に滑り込んだだけかもしれないので、stableMs を上げるか
'   明示シグナル (D-4c) を併用することを検討する。
' ============================================================
Public Property Get LastSettleInfo() As String
    LastSettleInfo = m_lastSettle
End Property


' ============================================================
' AutoWaitAfterAction (D-4b、論点7 案C の opt-in 部分)
'
'   True にすると、Wv2Element の ★.Value = / .Click / .SetAttribute が成功した後★
'   自動で WaitSettled を呼ぶ。既定 False。
'
'   ★D-3 の契約は変えない★ 操作の戻り値は「JS が例外なく走った」のままで、
'   待ちの成否は p.LastSettleInfo / p.WaitSettled を自分で見る。
'   ImplicitWaitSec (D-3) と同じ「明示が第一級、自動は opt-in」の流儀。
' ============================================================
Public Property Get AutoWaitAfterAction() As Boolean
    AutoWaitAfterAction = m_autoWait
End Property

Public Property Let AutoWaitAfterAction(ByVal onOff As Boolean)
    m_autoWait = onOff
    Wv2Log.LogD "Wv2Pane.AutoWaitAfterAction = " & onOff
End Property


' ============================================================
' AddIgnoreNetwork / ClearIgnoreNetwork / IgnoreSelectors (D-4b、論点5)
'
'   静穏の判定から外すものを指定する。★遮断ではない★ 要求は普通に飛ぶ。
'
'   AddIgnoreNetwork : URL の★部分一致★パターン。常時鳴っている計測系に使う
'                      例: p.AddIgnoreNetwork "google-analytics"
'   IgnoreSelectors  : DOM 変化を無視する要素の CSS セレクタ (カンマ区切り)
'                      例: p.IgnoreSelectors = ".clock, .ticker, #ads"
'                      ★closest で判定する★ ので、指定要素の子孫の変化も外れる
'
'   ★リストは VBA 側が持つ★ ページ側に置くと遷移で消えるため。
'   毎回の状態読み取りに相乗りして送り込むので往復は増えない。
'
'   ★★★ 広すぎる除外は「早すぎる静穏」を生む ★★★
'     意味のある通信まで外すと、まだ処理中なのに静穏と判定してしまう。
'     除外は「ログを見て、無関係だと確かめたものだけ」に留めること。
' ============================================================
Public Function AddIgnoreNetwork(ByVal urlPattern As String) As Boolean
    Dim v As Variant

    If Len(urlPattern) = 0 Then Exit Function
    If m_ignoreNet Is Nothing Then Set m_ignoreNet = New Collection

    For Each v In m_ignoreNet
        If CStr(v) = urlPattern Then
            AddIgnoreNetwork = True
            Exit Function
        End If
    Next v

    m_ignoreNet.Add urlPattern
    Wv2Log.LogD "Wv2Pane.AddIgnoreNetwork: " & urlPattern & _
                " (計 " & m_ignoreNet.Count & " 件)"
    AddIgnoreNetwork = True
End Function

Public Sub ClearIgnoreNetwork()
    Set m_ignoreNet = New Collection
    Wv2Log.LogD "Wv2Pane.ClearIgnoreNetwork: 除外リストを空にした"
End Sub

Public Property Get IgnoreSelectors() As String
    IgnoreSelectors = m_ignoreSel
End Property

Public Property Let IgnoreSelectors(ByVal cssSelectors As String)
    m_ignoreSel = cssSelectors
    Wv2Log.LogD "Wv2Pane.IgnoreSelectors = " & cssSelectors
End Property


' ============================================================
' SpaProbeApplyJs (D-4b、Private) - 除外リストをページへ送り込む JS
'
'   状態の読み取りと同じ往復に相乗りさせる。変数 P (プローブ) が束縛済みである
'   ことが前提。★空であっても必ず送る★ (送らないとページ側に前の設定が残る)。
' ============================================================
Private Function SpaProbeApplyJs() As String
    Dim js As String
    Dim arr As String
    Dim v As Variant

    If Not m_ignoreNet Is Nothing Then
        For Each v In m_ignoreNet
            If Len(arr) > 0 Then arr = arr & ","
            arr = arr & JsQuote(CStr(v))
        Next v
    End If

    ' ★空でも必ず送る★ (D-4b の実機検証で踏んだ)
    '   「空だから送らない」にすると、★ページ側に前の設定が残り続ける★。
    '   IgnoreSelectors = "" や ClearIgnoreNetwork が効かなくなり、
    '   一度除外したものが二度と戻らない。毎回上書きすれば VBA 側が唯一の権威になる。
    js = "P.ignoreNet=[" & arr & "];P.ignoreSel=" & _
         JsQuote(m_ignoreSel) & ";"

    SpaProbeApplyJs = js
End Function

' ============================================================
' ★D-4c: 明示シグナル (arm してから操作する)★
'
'   ★静穏だけでは「アプリが終わった」と「アプリが無視した」を区別できない★
'   どちらも静かだからである。重要な操作では、完了の目印を★操作の直前に arm★し、
'   次の WaitSettled でそれが観測されるまで待つ。
'
'   ■ ワンショット
'     arm したシグナルは★次の WaitSettled 1 回で消費される★。同じ操作を 2 回
'     するなら 2 回 arm し直す。消費は成功でも時間切れでも起きる。
'
'   ■ 当たったら静穏窓を測り直す
'     証拠が来た直後の 1 ポーリング分で満足せず、★その後もう一度まるごと
'     静穏窓を取る★ (参考プロダクトと同じ考え方)。
'
'   ■ ★arm が消えたら黙って直さない (fail-visible)★
'     ページ遷移でプローブごと消えると arm も消える。そのまま待つと
'     「条件なしの待ち」に化けて、静かになった瞬間に成功を返してしまう。
'     VBA 側でも arm 済みを覚えておき、食い違ったら signal-lost で失敗させる。
'
'   ■ 除外との関係
'     ★除外した要求でも明示シグナルの判定はする★ 除外は「静穏判定から外す」
'     だけの意味であって、「見なかったことにする」ではない。
'
'   使用例 (arm してから act する):
'     p.ArmContentSignal "#result-table"     ' 既存の器が書き変わるのを待つ
'     el.Click
'     If Not p.WaitSettled(10) Then Debug.Print p.LastSettleInfo
' ============================================================


' ============================================================
' ArmContentSignal (D-4c) - 既存の部分木が書き変わるのを待つ
'
'   ★arm した時点で対象が存在しなければ False★ (fail-fast)
'   セレクタの綴り間違いや、まだ作られていない要素を指定した事故を、
'   待ってから気づくのではなくその場で気づけるようにする。
'   ★操作の後にできる要素は指定できない★ 既にある親を指定すること。
'
'   戻り値: True = arm できた / False = 対象が無い・不正なセレクタ・通信失敗
'           理由は LastEvalError (not-found / bad-selector / その他)
' ============================================================
Public Function ArmContentSignal(ByVal selector As String) As Boolean
    Dim js As String
    Dim res As String

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armDom(" & JsQuote(selector) & ")};})()"

    res = EvalSync(js)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: 失敗 err=" & m_lastEvalError
        Exit Function
    End If

    ' ★JsonGetNum は負数を扱わない (設計原則64)★ ので、戻り値は文字列として読む。
    If InStr(1, res, """r"":0") = 0 Then
        If InStr(1, res, "-2") > 0 Then
            m_lastEvalError = "bad-selector"
        Else
            m_lastEvalError = "not-found"
        End If
        m_lastEvalOk = False
        Wv2Log.LogW "Wv2Pane.ArmContentSignal: arm できない (" & selector & _
                    ") 理由=" & m_lastEvalError
        Exit Function
    End If

    m_sigDom = True
    Wv2Log.LogD "Wv2Pane.ArmContentSignal: arm した (" & selector & ")"
    ArmContentSignal = True
End Function


' ============================================================
' ArmNetworkSignal (D-4c) - 特定の要求の完了を待つ
'
'   urlPattern は URL の★部分一致★。応答が返った時点で当たりとする。
'   ★存在確認はできない★ (これから飛ぶ要求なので)。綴りを間違えると
'   当たらないまま時間切れになる。LastSettleInfo に net:miss と出るので、
'   そのときは D-4d の診断ログで実際の URL を見て直すこと。
' ============================================================
Public Function ArmNetworkSignal(ByVal urlPattern As String) As Boolean
    Dim js As String

    If Len(urlPattern) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armNet(" & JsQuote(urlPattern) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmNetworkSignal: 失敗 err=" & m_lastEvalError
        Exit Function
    End If

    m_sigNet = True
    Wv2Log.LogD "Wv2Pane.ArmNetworkSignal: arm した (" & urlPattern & ")"
    ArmNetworkSignal = True
End Function


' ============================================================
' DisarmSignals (D-4c) - arm を解除する
'   WaitSettled が終わるときに自動で呼ばれる (ワンショット)。
'   arm したあと操作を取りやめたときは手で呼ぶ。
' ============================================================
Public Function DisarmSignals() As Boolean
    m_sigDom = False
    m_sigNet = False
    m_sigUrl = False

    EvalSync "(function(){var p=window.__wv2p;" & _
             "if(p&&p.disarm){p.disarm();}return 1;})()"

    DisarmSignals = m_lastEvalOk
End Function


' ============================================================
' SigInfoOf (D-4c、Private) - 待ちの内訳に載せるシグナルの当たり外れ
' ============================================================
Private Function SigInfoOf(ByVal st As String) As String
    Dim r As String

    If m_sigDom Then
        r = r & " dom:"
        If Wv2Json.JsonGetNum(st, "sdh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigNet Then
        r = r & " net:"
        If Wv2Json.JsonGetNum(st, "snh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If m_sigUrl Then
        r = r & " url:"
        If Wv2Json.JsonGetNum(st, "suh") > 0 Then r = r & "hit" Else r = r & "miss"
    End If

    If Len(r) > 0 Then r = " signal=" & Trim$(r)
    SigInfoOf = r
End Function

' ============================================================
' ArmUrlSignal (D-4e) - URL が指定の文字列を含むまで待つ
'
'   ★D-4d の実測で必要と分かった★ SPA の完了が DOM でも通信でもなく
'   ★URL に現れる★ ことがある。Google Maps の住所検索がまさにそれで、
'   検索が通ると URL が /place/... と /@緯度,経度,ズーム に変わる。
'   これは「完了の目印」であると同時に「欲しかった座標そのもの」でもある。
'
'   pushState による変化は DOM 変化として観測できないので、
'   ★location.href を毎回の状態読み取りで見る★ 形にしてある。
'
'   ワンショット。次の WaitSettled 1 回で消費される。
'   ★arm した時点で既に一致していても当たり扱いになる★ ので、
'   遷移前の URL に含まれない文字列を指定すること。
'
'   使用例:
'     p.ArmUrlSignal "/place/"
'     el.Click
'     If p.WaitSettled(15) Then Debug.Print p.View_GetSource()
' ============================================================
Public Function ArmUrlSignal(ByVal urlFragment As String) As Boolean
    Dim js As String

    If Len(urlFragment) = 0 Then Exit Function

    js = "(function(){" & SpaProbeJs() & SpaProbeApplyJs() & _
         "return {r:P.armUrl(" & JsQuote(urlFragment) & ")};})()"

    EvalSync js
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.ArmUrlSignal: 失敗 err=" & m_lastEvalError
        Exit Function
    End If

    m_sigUrl = True
    Wv2Log.LogD "Wv2Pane.ArmUrlSignal: arm した (" & urlFragment & ")"
    ArmUrlSignal = True
End Function


' ============================================================
' SpaProbeLogging (D-4e) - 診断ログの ON / OFF
'
'   ★既定は False★ 有効にすると、ページ側が最近 200 件の出来事を溜める:
'     NET+ <fetch|xhr> <URL>     要求が始まった
'     NET- <経過ms> <URL>        要求が終わった
'     DOM  x<件数>               DOM 変化のバースト
'     SIG  <種別> hit ...        arm したシグナルが当たった
'
'   ★何を待つべきかを推測でなく観測で決めるための道具★
'   参考プロダクト (WebDriver BiDi for SeleniumVBA) の Discovery Log に当たる。
'   arm なしで 1 回流し、ログに出た URL や DOM の動きを見て、次に何を arm するか /
'   何を除外するかを決める。
'
'   ★常時 ON にしない★ 溜め込みの分だけページ側が重くなる。
' ============================================================
Public Property Get SpaProbeLogging() As Boolean
    Dim st As String

    st = SpaProbeState()
    If Not m_lastEvalOk Then Exit Property

    SpaProbeLogging = (InStr(1, st, """lgOn"":true") > 0)
End Property

Public Property Let SpaProbeLogging(ByVal onOff As Boolean)
    Dim js As String

    js = "(function(){" & SpaProbeJs() & _
         "P.logOn=" & IIf(onOff, "true", "false") & ";" & _
         "if(!P.logOn){P.log=[];}return {r:1};})()"

    EvalSync js
    Wv2Log.LogD "Wv2Pane.SpaProbeLogging = " & onOff & _
                IIf(m_lastEvalOk, "", " (失敗: " & m_lastEvalError & ")")
End Property


' ============================================================
' SpaProbeDrainLog (D-4e) - 溜まった診断ログを取り出して Wv2Log へ流す
'
'   ★取り出したらページ側は空になる★ (溜め続けない)。
'   戻り値: 流した件数。0 なら何も溜まっていなかった。
'
'   1 回の EvalSync で運べる量には限りがあるので、★1 回あたり 60 件まで★。
'   残りは次の呼び出しで取れる (戻り値が 60 なら、まだ残っている可能性がある)。
' ============================================================
Public Function SpaProbeDrainLog() As Long
    Dim js As String
    Dim res As String
    Dim txt As String
    Dim lines As Variant
    Dim i As Long

    js = "(function(){" & SpaProbeJs() & _
         "var a=P.log.splice(0,60);" & _
         "return {n:a.length,s:a.join(String.fromCharCode(1))};})()"

    res = EvalSync(js, 10)
    If Not m_lastEvalOk Then
        Wv2Log.LogW "Wv2Pane.SpaProbeDrainLog: 失敗 err=" & m_lastEvalError
        Exit Function
    End If

    SpaProbeDrainLog = Wv2Json.JsonGetNum(res, "n")
    If SpaProbeDrainLog = 0 Then Exit Function

    Wv2Json.JsonPickStr res, "s", txt
    lines = Split(txt, ChrW$(1))

    Wv2Log.LogI "  --- 診断ログ (" & SpaProbeDrainLog & " 件) ---"
    For i = LBound(lines) To UBound(lines)
        Wv2Log.LogI "    " & lines(i)
    Next i
End Function

' ============================================================
' StaleInflightMs (D-4f) - 居座る要求を静穏判定から外す境目 (ms)
'
'   ★既定 10000 (10 秒)★ 0 にすると無効 (すべての要求を数える)。
'
'   ■ なぜ要るか (D-4d の実測)
'     Google Maps の住所検索は /search?tbm=map を★開きっぱなし★にして結果を
'     流し続ける。検索自体は 2 秒で終わって DOM も静まり URL も変わっているのに、
'     この 1 本のせいで永久に inflight=0 にならず、静穏待ちが時間切れになった。
'     ロングポーリングやストリーミングはどのサイトにもあるので、個別の除外
'     (AddIgnoreNetwork) だけに頼ると、知らないサイトでは必ず一度は踏む。
'
'   ■ 何を捨てているか (正直に)
'     ★本当に遅いだけの要求も、この時間を超えれば数えなくなる★
'     ただしそのとき DOM も静穏窓のあいだ動いていないので、
'     「見た目は落ち着いている」という WaitSettled の意味とは矛盾しない
'     (True は「アプリの処理が終わった」証明ではない、という前提のまま)。
'     正確さが要る場面では AddIgnoreNetwork で個別に外すか、
'     D-4c の明示シグナルで「何を待つか」を書くこと。
' ============================================================
Public Property Get StaleInflightMs() As Long
    StaleInflightMs = m_staleMs
End Property

Public Property Let StaleInflightMs(ByVal ms As Long)
    If ms < 0 Then ms = 0
    m_staleMs = ms
    Wv2Log.LogD "Wv2Pane.StaleInflightMs = " & ms & _
                IIf(ms = 0, " (無効: すべて数える)", "")
End Property


' ============================================================
' StaleMsJs (D-4f、Private) - 居座り境目を JS の数値リテラルにする
'   0 (無効) のときは巨大な値にして「すべて新しい」扱いにする。
' ============================================================
Private Function StaleMsJs() As String
    If m_staleMs <= 0 Then
        StaleMsJs = "1e15"
    Else
        StaleMsJs = CStr(m_staleMs)
    End If
End Function

' ============================================================
' ★D-6 の本体: QuerySelectorAll と寿命管理★
' ============================================================


' ============================================================
' QuerySelectorAll (D-6) - 一致する要素をまとめて掴む
'
'   引数:
'     selector : CSS セレクタ
'     maxCount : 返す上限 (既定 200)。0 以下なら既定値に丸める
'   戻り値:
'     Wv2Element の Collection。★見つからなければ空の Collection★ (Nothing ではない)
'     失敗したときも空の Collection を返す。区別は LastEvalOk で行う (設計原則93)。
'
'   ★件数が上限に達したら LastAllTruncated = True★
'     「返ってきた件数 = ページにある数」と思い込ませないため。
'
'   ★文書順で返る★ querySelectorAll の順序そのまま。
'
'   使用例:
'     Dim els As Collection, el As Wv2Element
'     Set els = p.QuerySelectorAll("a[href]")
'     Debug.Print els.Count & " 件"
'     For Each el In els
'         Debug.Print el.GetAttribute("href")
'     Next el
'
'   ★1 プロパティにつき 1 往復する★ (D-2 の割り切り) ので、200 件から属性を
'   全部読むと 200 往復 = 8 秒ほどかかる。まとめて取りたいときは EvalSync で
'   JS 式を 1 本書く方が速い。
' ============================================================
Public Function QuerySelectorAll(ByVal selector As String, _
                                 Optional ByVal maxCount As Long = ALL_MAX_DEFAULT) As Collection
    Dim js As String
    Dim res As String
    Dim gen As String
    Dim handles As String
    Dim parts As Variant
    Dim i As Long
    Dim el As Wv2Element
    Dim outCol As Collection

    Set outCol = New Collection
    Set QuerySelectorAll = outCol
    m_allTruncated = False

    If maxCount <= 0 Then maxCount = ALL_MAX_DEFAULT

    ' ★D-2 の FindOne を push からループに変えただけ★
    '   ハンドルはカンマ区切りの文字列で返す (Wv2Json は配列を扱えないため。
    '   パーサを拡張するより Split の方が小さい ― 論点5)。
    js = "(function(){" & JsRegistryPrologue() & _
         "var q=document.querySelectorAll(" & JsQuote(selector) & ");" & _
         "var mx=" & CStr(maxCount) & ",h=[],i;" & _
         "for(i=0;i<q.length&&i<mx;i++){R.els.push(q[i]);h.push(R.els.length-1);}" & _
         "return {g:R.gen,n:q.length,k:h.length,cut:(q.length>mx)," & _
         "h:h.join(','),tot:R.els.length};})()"

    res = EvalSync(js)

    If Not m_lastEvalOk Then
        Debug.Print "Wv2Pane.QuerySelectorAll: 失敗 (" & selector & ") err=" & m_lastEvalError
        Exit Function
    End If

    If Wv2Json.JsonPickStr(res, "g", gen) Then
        m_domGen = gen
    End If

    m_allTruncated = (InStr(1, res, """cut"":true") > 0)
    If m_allTruncated Then
        Wv2Log.LogW "Wv2Pane.QuerySelectorAll: ★打ち切った★ (" & selector & _
                    ") ページには " & Wv2Json.JsonGetNum(res, "n") & _
                    " 件あるが " & maxCount & " 件で止めた"
    End If

    Wv2Json.JsonPickStr res, "h", handles
    If Len(handles) = 0 Then
        Debug.Print "Wv2Pane.QuerySelectorAll: 0 件 (" & selector & ")"
        WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")
        Exit Function
    End If

    parts = Split(handles, ",")
    For i = LBound(parts) To UBound(parts)
        Set el = New Wv2Element
        el.InitElement Me, CLng(parts(i)), gen
        outCol.Add el
    Next i

    WarnIfElementsPiled Wv2Json.JsonGetNum(res, "tot")

    Debug.Print "Wv2Pane.QuerySelectorAll: " & outCol.Count & " 件 (" & _
                selector & ") gen=" & gen
End Function


' ============================================================
' LastAllTruncated (D-6) - 直前の QuerySelectorAll が上限で打ち切られたか
'   ★True なら「返ってきた件数 = ページにある数」ではない★
'   全部要るなら maxCount を上げるか、セレクタを絞ること。
' ============================================================
Public Property Get LastAllTruncated() As Boolean
    LastAllTruncated = m_allTruncated
End Property


' ============================================================
' ElementCount (D-6) - JS 側レジストリに積まれている要素の数
'
'   ★掴んだ要素は解放しない★ (D-2 論点3)。ページ遷移でレジストリごと消えるので
'   実害は小さいが、QuerySelectorAll は 1 回で 200 個積みうるので伸びは速い。
'   気になったらこれで数を見て、ClearElementRegistry で作り直す。
'   ★作り直すと手元の Wv2Element はすべて stale になる★
'
'   戻り値: 積まれている数。失敗したら -1 (LastEvalError に理由)
' ============================================================
Public Property Get ElementCount() As Long
    Dim res As String

    res = EvalSync("(function(){" & JsRegistryPrologue() & _
                  "return {n:R.els.length};})()")

    If Not m_lastEvalOk Then
        ElementCount = -1
        Exit Property
    End If

    ElementCount = Wv2Json.JsonGetNum(res, "n")
End Property


' ============================================================
' WarnIfElementsPiled (D-6、Private) - 積み上がりを知らせる
'   ★捨てはしない★ 勝手に捨てると手元の Wv2Element が黙って stale になる。
'   知らせるだけにして、捨てるかどうかは呼び出し側が決める。
' ============================================================
Private Sub WarnIfElementsPiled(ByVal total As Long)
    If total < ELEMENT_WARN_COUNT Then Exit Sub

    Wv2Log.LogW "Wv2Pane: ★要素レジストリに " & total & " 個積まれている★ " & _
                "(掴んだ要素は解放しない設計)。要らなければ ClearElementRegistry で" & _
                "作り直せる。ただし手元の Wv2Element はすべて stale になる。"
End Sub



' ============================================================
' ★N-1 : ネットワーク要求のキャプチャ (WebResourceRequested)★
'
'   ここから下が N-1 段階で足した一式。並びは
'     (1) 使う側の口   NetCaptureStart / NetCaptureStop / NetLogDrain ほか
'     (2) COM の配線   View_AddWebResourceRequested ほか 4 本
'     (3) 受け口       View_OnWebResourceRequested (+ in-callback ガード)
'     (4) リング       NetRingAlloc / NetPush / NetElapsedMs / NetRingIndex
'     (5) 種別の名前   NetCtxName / NetCtxFromName
'
'   使い方 (最小):
'     p.NetCaptureStart          ' ★Navigate の前に呼べば初回の DOCUMENT から捕まる★
'     ... 手で 1 回操作する ...
'     p.NetLogDrain              ' 溜まった分をログファイルへ流して空にする
'     p.NetCaptureStop
' ============================================================


' ============================================================
' NetCaptureStart (N-1) - 捕捉を始める
'
'   contexts  : 捕まえる種別をカンマ区切りで指定する。空なら DOCUMENT,XHR,FETCH。
'               名前でも数字でもよい。"ALL" にすると全種別 (画像も CSS も来る)。
'                 DOCUMENT / STYLESHEET / IMAGE / MEDIA / FONT / SCRIPT /
'                 XHR (= XMLHTTPREQUEST) / FETCH / TEXTTRACK / EVENTSOURCE /
'                 WEBSOCKET / MANIFEST / SIGNEDEXCHANGE / PING /
'                 CSPVIOLATIONREPORT / OTHER / ALL
'               ★数字も受ける★ SDK が種別を増やしてもここを直さずに使えるため。
'   uriFilter : URL のフィルタ。既定 "*" (すべて)。* と ? が使える。
'
'   ★既に捕捉中なら一度止めてから張り直す★
'     種別を変えたつもりが前の網も残っている、という事故を防ぐ (設計原則106 の趣旨)。
'
'   ★フィルタが 1 つも張れなかったらハンドラごと外して False を返す★
'     フィルタゼロではイベントが 1 件も来ないので、「捕捉中なのに何も出ない」
'     という一番たちの悪い状態を作らない。
'
'   戻り値: True なら捕捉中。
' ============================================================
Public Function NetCaptureStart(Optional ByVal contexts As String = "", _
                                Optional ByVal uriFilter As String = "*") As Boolean
    Dim parts As Variant
    Dim i     As Long
    Dim ctx   As Long
    Dim added As Long
    Dim hr    As Long
    Dim spec  As String

    If m_closed Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: このペインは既に閉じている"
        Exit Function
    End If
    If m_pView = 0 Then
        Wv2Log.LogW "Wv2Pane.NetCaptureStart: m_pView が未取得 (先にタブを作ること)"
        Exit Function
    End If

    If m_netOn Then NetCaptureStop

    spec = Trim$(contexts)
    If LenB(spec) = 0 Then spec = NET_CTX_DEFAULT
    If LenB(Trim$(uriFilter)) = 0 Then uriFilter = "*"

    ' --- (1) ハンドラを登録する ---
    hr = View_AddWebResourceRequested()
    If hr <> 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: add_WebResourceRequested 失敗 hr=&H" & Hex(hr)
        Exit Function
    End If

    ' --- (2) フィルタを張る (★1 つも張らないと 1 件も来ない★) ---
    parts = Split(spec, ",")
    For i = LBound(parts) To UBound(parts)
        ctx = NetCtxFromName(CStr(parts(i)))
        If ctx < 0 Then
            Wv2Log.LogW "Wv2Pane.NetCaptureStart: 種別 [" & Trim$(CStr(parts(i))) & _
                        "] が分からないので飛ばす"
        Else
            hr = View_AddWebResourceRequestedFilter(uriFilter, ctx)
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStart: フィルタ登録に失敗 " & _
                            NetCtxName(ctx) & " hr=&H" & Hex(hr)
            Else
                m_netFilters.Add uriFilter & vbTab & ctx
                added = added + 1
            End If
        End If
    Next i

    If added = 0 Then
        Wv2Log.LogE "Wv2Pane.NetCaptureStart: ★フィルタが 1 つも張れなかった★ " & _
                    "この状態ではイベントが 1 件も来ないので、ハンドラも外して中止する。"
        View_RemoveWebResourceRequested
        Exit Function
    End If

    NetLogClear
    m_netOn = True
    NetCaptureStart = True

    Wv2Log.LogI "Wv2Pane.NetCaptureStart: 捕捉開始 uri=" & uriFilter & _
                " 種別=" & spec & " (フィルタ " & added & " 本)"
End Function


' ============================================================
' NetCaptureStop (N-1) - 捕捉をやめる
'
'   ★張ったフィルタは必ず外す★ (設計原則106 の趣旨)
'   ハンドラを外すだけでもイベントは来なくなるが、相手側に網が残ったままだと
'   次の NetCaptureStart が知らないうちに広い網で動くことになる。
'
'   戻り値: True なら後始末まで含めて成功。
' ============================================================
Public Function NetCaptureStop() As Boolean
    Dim v     As Variant
    Dim parts As Variant
    Dim hr    As Long
    Dim allOk As Boolean

    If Not m_netOn Then
        NetCaptureStop = True
        Exit Function
    End If

    allOk = True

    ' --- (1) フィルタを外す ---
    '   ★既に閉じている / View が無いときは COM を叩けない★
    '   相手ごと消えているので網も残らない。覚えている分を捨てるだけにする。
    If m_closed Or m_pView = 0 Then
        Wv2Log.LogD "Wv2Pane.NetCaptureStop: View が無いのでフィルタ解除は省略"
    Else
        For Each v In m_netFilters
            parts = Split(CStr(v), vbTab)
            hr = View_RemoveWebResourceRequestedFilter(CStr(parts(0)), CLng(parts(1)))
            If hr <> 0 Then
                Wv2Log.LogW "Wv2Pane.NetCaptureStop: フィルタ解除に失敗 " & _
                            CStr(parts(0)) & " / " & NetCtxName(CLng(parts(1))) & _
                            " hr=&H" & Hex(hr)
                allOk = False
            End If
        Next v
    End If
    Set m_netFilters = New Collection

    ' --- (2) ハンドラを外す ---
    hr = View_RemoveWebResourceRequested()
    If hr <> 0 Then allOk = False

    m_netOn = False
    NetCaptureStop = allOk

    Wv2Log.LogI "Wv2Pane.NetCaptureStop: 捕捉終了 (総発火 " & m_netTotal & " 件 / " & _
                "手元 " & m_netCount & " 件 / 溢れ " & m_netDropped & " 件)"
End Function


' ============================================================
' NetCaptureOn (N-1) - 捕捉中かどうか
' ============================================================
Public Property Get NetCaptureOn() As Boolean
    NetCaptureOn = m_netOn
End Property


' ============================================================
' NetLogCount / NetLogTotal / NetLogDropped (N-1)
'
'   NetLogCount   : 今リングに入っている件数 (取り出せる数)
'   NetLogTotal   : 捕捉開始からの★総発火数★ (溢れた分も数える)
'   NetLogDropped : ★リングから溢れて捨てた件数★
'
'   ★どちらも「捕捉開始からの通算」で、NetLogDrain では 0 に戻らない★ (N-1b)
'   取り出したのは中身だけ。0 に戻したいときは NetLogClear を呼ぶ。
'
'   ★NetLogCount と NetLogTotal が食い違っていたら取りこぼしている★
'   件数を黙って切らないための組 (D-6 の LastAllTruncated と同じ判断)。
' ============================================================
Public Property Get NetLogCount() As Long
    NetLogCount = m_netCount
End Property

Public Property Get NetLogTotal() As Long
    NetLogTotal = m_netTotal
End Property

Public Property Get NetLogDropped() As Long
    NetLogDropped = m_netDropped
End Property


' ============================================================
' NetLogCapacity (N-1) - リングの容量 (件)
'   ★変えると中身は捨てられる★ (作り直すため)。捕捉を始める前に決めること。
' ============================================================
Public Property Get NetLogCapacity() As Long
    NetLogCapacity = m_netCap
End Property

Public Property Let NetLogCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetLogCapacity: " & n & " は上限 " & NET_CAP_MAX & " を超えるので丸める"
        n = NET_CAP_MAX
    End If
    m_netCap = n
    NetRingAlloc
    Wv2Log.LogD "Wv2Pane.NetLogCapacity = " & m_netCap & " (中身は作り直したので空)"
End Property


' ============================================================
' NetLogClear (N-1) - リングを空にする
'   ★経過 ms の基準もここで打ち直す★ (次の 1 件が +0ms から始まる)
'   ★通算カウンタ (NetLogTotal / NetLogDropped) も 0 に戻る★
'   中身だけ取り出したいなら NetLogDrain を使うこと (通算は保たれる)。
' ============================================================
Public Sub NetLogClear()
    NetRingAlloc
End Sub


' ============================================================
' NetLogLine (N-1) - 流さずに 1 件だけ読む
'
'   idx は★古い順に 1 起点★。範囲外なら空文字。
'   1 行は vbTab 区切りで 5 項目:
'     通し番号 / 経過ms / メソッド / 種別名 / URL
'
'   ★ドレインすると消えるので、判定に使うときはこちらで読む★
'   (N-5 で PowerShell コードを吐くときの入口もここになる)
' ============================================================
Public Property Get NetLogLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Property

    NetLogLine = m_netSeq(k) & vbTab & m_netMs(k) & vbTab & _
                 m_netMethod(k) & vbTab & NetCtxName(m_netCtx(k)) & vbTab & m_netUri(k)
End Property


' ============================================================
' NetLogDrain (N-1) - 溜まった分を Wv2Log へ流して空にする
'
'   ★取り出したらリングは空になる★ (D-4 の SpaProbeDrainLog と同じ契約)。
'   ★判定はログファイルで読む★ (設計原則105)。イミディエイトは配管ログで流れる。
'
'   戻り値: 流した件数。0 なら何も溜まっていなかった。
' ============================================================
Public Function NetLogDrain() As Long
    Dim i As Long
    Dim n As Long

    n = m_netCount
    If n = 0 Then
        Wv2Log.LogI "  --- ネットワーク要求: 溜まっていない (総発火 " & m_netTotal & " 件) ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- ネットワーク要求 " & n & " 件 (総発火 " & m_netTotal & " 件) ---"
    Wv2Log.LogI "        #   経過ms  メソッド 種別       URL"
    For i = 1 To n
        Wv2Log.LogI "      " & NetLogFormat(i)
    Next i

    If m_netDropped > 0 Then
        ' ★通算で数えている★ ので、溢れたまま何度もドレインすると毎回出る。
        '   出続けるのが鬱陶しいのではなく、★取りこぼしたままだという事実★ が
        '   消えないだけ。消したいなら NetLogClear で数え直す。
        Wv2Log.LogW "  ★捕捉開始から " & m_netDropped & " 件が溢れて捨てられている★ " & _
                    "NetLogCapacity を増やすか、こまめに NetLogDrain すること。"
    End If

    NetLogDrain = n
    ' ★中身だけ空にする (通算は保つ)★ 論点3 案Y
    NetRingEmpty
End Function


' ============================================================
' View_AddWebResourceRequested (N-1、ICoreWebView2 vtable 55)
'
'   ★index 55 は tools/vtable.py で WebView2.h から確定させた値★ (設計原則115)。
'   既存の配線からの逆算では 54 になっていた (OpenDevToolsWindow の数え落とし)。
'
'   実装は View2_AddDOMContentLoaded を逐語的な見本にしている。違いは
'   「m_pView2 ではなく m_pView」「vtable 64/65 ではなく 55/56」だけ。
'   pVTableObj = h.Slot + VTABLE_OBJ_OFFSET は同じく必ず通す (設計原則45)。
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View_AddWebResourceRequested() As Long
    If m_closed Then
        View_AddWebResourceRequested = &H80004003   ' E_POINTER
        Exit Function
    End If
    If m_pView = 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: m_pView が未取得"
        View_AddWebResourceRequested = &H80004003
        Exit Function
    End If

    ' 既に登録済みなら何もしない
    If Not (m_persistentHandlers(HK_WebResourceRequested) Is Nothing) Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: 既に登録済みなのでスキップ"
        View_AddWebResourceRequested = 0
        Exit Function
    End If

    Dim h As ComCallbackHandler
    Set h = AcquireHandlerFor(HK_WebResourceRequested, Me)
    If h Is Nothing Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: ハンドラの Acquire に失敗"
        View_AddWebResourceRequested = &H8000FFFF   ' E_UNEXPECTED
        Exit Function
    End If

    Dim pVTableObj As LongPtr
    pVTableObj = h.Slot + VTABLE_OBJ_OFFSET

    Dim token As LongLong
    Dim hr As Long
    hr = dcf(m_pView, 55, "View.add_WebResourceRequested", pVTableObj, VarPtr(token))

    If hr <> 0 Then
        Debug.Print "Wv2Pane.View_AddWebResourceRequested: 失敗 hr=&H" & Hex(hr)
        Set h = Nothing
        View_AddWebResourceRequested = hr
        Exit Function
    End If

    m_persistentTokens(HK_WebResourceRequested) = token
    Set m_persistentHandlers(HK_WebResourceRequested) = h
    Debug.Print "Wv2Pane.View_AddWebResourceRequested OK: token=" & token & _
                ", handler.slot=" & h.Slot

    View_AddWebResourceRequested = 0
End Function


' ============================================================
' View_RemoveWebResourceRequested (N-1、ICoreWebView2 vtable 56)
'   未登録なら no-op (S_OK で返す)。
' ============================================================
Public Function View_RemoveWebResourceRequested() As Long
    If m_persistentHandlers(HK_WebResourceRequested) Is Nothing Then
        View_RemoveWebResourceRequested = 0   ' 未登録は no-op
        Exit Function
    End If

    Dim hr As Long
    If m_pView <> 0 Then
        hr = dcf(m_pView, 56, "View.remove_WebResourceRequested", _
                 m_persistentTokens(HK_WebResourceRequested))
        If hr <> 0 Then
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: 解除失敗 hr=&H" & Hex(hr)
        Else
            Debug.Print "Wv2Pane.View_RemoveWebResourceRequested OK"
        End If
    Else
        ' m_pView が無いので remove を撃つ手段がない。参照だけ切る。
        Debug.Print "Wv2Pane.View_RemoveWebResourceRequested: m_pView が 0 なので参照のみ解放"
    End If

    Set m_persistentHandlers(HK_WebResourceRequested) = Nothing
    m_persistentTokens(HK_WebResourceRequested) = 0^

    View_RemoveWebResourceRequested = hr
End Function


' ============================================================
' View_AddWebResourceRequestedFilter (N-1、ICoreWebView2 vtable 57)
'
'   ★フィルタを 1 つも登録しないとイベントは 1 件も来ない★
'   種別ごとに 1 回ずつ呼ぶ。NetCaptureStart が既定の 3 種別ぶん撃つ。
'
'   uriFilter       : URL のパターン。* と ? が使える ("*" ですべて)
'   resourceContext : COREWEBVIEW2_WEB_RESOURCE_CONTEXT の値 (0=ALL, 1=DOCUMENT, ...)
'
'   戻り値: HRESULT (S_OK = 0 で成功)
' ============================================================
Public Function View_AddWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_AddWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_AddWebResourceRequestedFilter = dcf(m_pView, 57, _
                                             "View.AddWebResourceRequestedFilter", _
                                             StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_RemoveWebResourceRequestedFilter (N-1、ICoreWebView2 vtable 58)
'   ★登録したときと同じ (uri, context) の組で呼ばないと外れない★
' ============================================================
Public Function View_RemoveWebResourceRequestedFilter( _
    ByVal uriFilter As String, _
    ByVal resourceContext As Long) As Long

    If m_pView = 0 Then
        View_RemoveWebResourceRequestedFilter = &H80004003
        Exit Function
    End If
    View_RemoveWebResourceRequestedFilter = dcf(m_pView, 58, _
                                                "View.RemoveWebResourceRequestedFilter", _
                                                StrPtr(uriFilter), resourceContext)
End Function


' ============================================================
' View_OnWebResourceRequested (N-1)
'   ComCallbackHandler から動的バインディング経由で呼ばれる。Public Sub 必須。
'
'   ★他の受け口と違ってエラーを再送出しない★ (事後報告)
'     WebResourceRequested は通信の途中で同期的に呼ばれるので、ここで
'     Err.Raise すると★通信の最中にエラーダイアログが出る★。しかも 1 ページで
'     数十～数百件来るので、1 件壊れただけで操作が続けられなくなる。
'     そこで握り潰す代わりに★リングに 1 件「(err)」として書き残す★。
'     握り潰したことは NetLogDrain の出力に必ず現れる。
' ============================================================
Public Sub View_OnWebResourceRequested(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim errNum As Long
    Dim errDsc As String

    m_inCallback = m_inCallback + 1
    On Error GoTo GuardExit
    View_OnWebResourceRequested_Core pSender, pArgs

GuardExit:
    errNum = Err.Number
    errDsc = Err.Description
    If m_inCallback > 0 Then m_inCallback = m_inCallback - 1
    If errNum <> 0 Then
        On Error Resume Next
        NetPush "(err)", -1, "★受け口でエラー " & errNum & " " & errDsc & "★"
    End If
End Sub


' ============================================================
' View_OnWebResourceRequested_Core (N-1、Private)
'
'   ★記録だけしてすぐ返す (論点5)★
'     ここで Wv2Log も Debug.Print も撃たない。EvalSync は論外 (仕様事実14 / 20)。
'     dcf / GetStringProperty に渡す funcName も空にして、失敗しても
'     イミディエイトに何も出さない (失敗はリングに書き残すので消えない)。
'
'   args の vtable (tools/vtable.py で確定):
'     3 = get_Request(ICoreWebView2WebResourceRequest**)
'     7 = get_ResourceContext(COREWEBVIEW2_WEB_RESOURCE_CONTEXT*)
'   request の vtable:
'     3 = get_Uri(LPWSTR*)   5 = get_Method(LPWSTR*)
'
'   ★get_Request は AddRef 済みで返る★ ので必ず Release する (COM 規約)。
'   1 ページで数百回通る経路なので、ここを漏らすと目に見えて漏れる。
' ============================================================
Private Sub View_OnWebResourceRequested_Core(ByVal pSender As LongPtr, ByVal pArgs As LongPtr)
    Dim pReq As LongPtr
    Dim hr   As Long
    Dim uri  As String
    Dim mth  As String
    Dim ctx  As Long

    If pArgs = 0 Then
        NetPush "(err)", -1, "★args が 0★"
        Exit Sub
    End If

    ' 種別 (取得できなければ 0 = ALL が返る。get_ResourceContext は本来 ALL を
    ' 返さないので、ログに ALL が並んでいたらこの取得が失敗している合図)
    ctx = GetLongProperty(pArgs, 7, "")

    hr = dcf(pArgs, 3, "", VarPtr(pReq))
    If hr <> 0 Or pReq = 0 Then
        NetPush "(err)", ctx, "★get_Request 失敗 hr=&H" & Hex(hr) & "★"
        Exit Sub
    End If

    uri = GetStringProperty(pReq, 3, "")
    mth = GetStringProperty(pReq, 5, "")

    NetPush mth, ctx, uri

    ' --- N-2: 詳細 (ヘッダとボディ) ---
    '   ★既定 OFF★ なので、ここは NetDetailOn = True のときしか走らない。
    '   pReq を Release する前に読む (args も pReq も Invoke の間しか有効でない)。
    If m_detOn Then NetReadDetail pReq, mth, uri

    ComRelease pReq
End Sub


' ============================================================
' NetReadDetail (N-2、Private) - 1 要求ぶんの詳細を読んでリングに置く
'
'   ★ハンドラの中で走る★ (論点1 案A)。args は Invoke の間しか有効でないので、
'   ここで同期的に読むしかない。だから上限で必ず切る。
'
'   ★ボディを読むのは GET 以外だけ (論点7 案α)★
'     GET にボディはほぼ無い。既定でこれだけ外すと大半の無駄が消える。
'     URL でさらに絞りたいときは NetDetailUriFilter。
' ============================================================
Private Sub NetReadDetail(ByVal pReq As LongPtr, _
                          ByVal mth As String, _
                          ByVal uri As String)
    Dim headers  As String
    Dim cType    As String
    Dim nBytes   As Long
    Dim trunc    As Boolean
    Dim isText   As Boolean
    Dim bodyText As String
    Dim hexHead  As String
    Dim buf()    As Byte

    ' URL で絞る (空なら絞らない)
    If LenB(m_detUriFilter) > 0 Then
        If InStr(1, uri, m_detUriFilter, vbTextCompare) = 0 Then Exit Sub
    End If

    headers = NetReadHeaders(pReq)
    cType = NetHeaderValue(headers, "content-type")

    ' ★GET はボディを読まない★
    If UCase$(mth) <> "GET" Then
        If NetReadBody(pReq, nBytes, trunc, buf) Then
            If nBytes > 0 Then
                isText = NetIsTextType(cType)
                If isText Then
                    bodyText = Utf8BytesToString(buf, nBytes)
                Else
                    hexHead = NetHexHead(buf, nBytes)
                End If
            End If
        End If
    End If

    NetDetailPush m_netTotal, mth, uri, headers, cType, _
                  nBytes, trunc, isText, bodyText, hexHead
End Sub


' ============================================================
' NetRingAlloc (N-1、Private) - リングを作り直して空にする
'   容量を変えたときと、捕捉を始め直すときに通る。
' ============================================================
Private Sub NetRingAlloc()
    If m_netCap < 1 Then m_netCap = NET_CAP_DEFAULT

    ReDim m_netSeq(0 To m_netCap - 1)
    ReDim m_netMs(0 To m_netCap - 1)
    ReDim m_netCtx(0 To m_netCap - 1)
    ReDim m_netMethod(0 To m_netCap - 1)
    ReDim m_netUri(0 To m_netCap - 1)

    m_netHead = 0
    m_netCount = 0
    m_netTotal = 0
    m_netDropped = 0
    m_netT0 = GetTickCount()
End Sub


' ============================================================
' NetRingEmpty (N-1b、Private) - ★中身だけ★ 空にする
'   通算カウンタ (m_netTotal / m_netDropped) と経過 ms の基準は触らない。
'   NetLogDrain が「取り出したぶんを捨てる」ために使う。
' ============================================================
Private Sub NetRingEmpty()
    m_netHead = 0
    m_netCount = 0
End Sub


' ============================================================
' NetPush (N-1、Private) - 1 件をリングに置く
'   ★溢れたら古いものを上書きし、捨てた数を数える★ (黙って切らない)
' ============================================================
Private Sub NetPush(ByVal mth As String, ByVal ctx As Long, ByVal uri As String)
    m_netTotal = m_netTotal + 1

    m_netSeq(m_netHead) = m_netTotal
    m_netMs(m_netHead) = NetElapsedMs()
    m_netCtx(m_netHead) = ctx
    m_netMethod(m_netHead) = mth
    m_netUri(m_netHead) = uri

    m_netHead = m_netHead + 1
    If m_netHead >= m_netCap Then m_netHead = 0

    If m_netCount < m_netCap Then
        m_netCount = m_netCount + 1
    Else
        m_netDropped = m_netDropped + 1
    End If
End Sub


' ============================================================
' NetElapsedMs (N-1、Private) - 捕捉開始からの経過 ms
'   GetTickCount は Long なので約 49.7 日で一周する。差は Double に上げて
'   計算し、負になったら 2^32 を足して戻す。
' ============================================================
Private Function NetElapsedMs() As Long
    Dim d As Double
    d = CDbl(GetTickCount()) - CDbl(m_netT0)
    If d < 0 Then d = d + 4294967296#
    If d > 2147483647# Then d = 2147483647#
    NetElapsedMs = CLng(d)
End Function


' ============================================================
' NetRingIndex (N-1、Private) - 古い順 1 起点 → リング上の位置
'   範囲外なら -1。
' ============================================================
Private Function NetRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_netCount Then
        NetRingIndex = -1
        Exit Function
    End If

    k = m_netHead - m_netCount + (idx - 1)
    Do While k < 0
        k = k + m_netCap
    Loop
    NetRingIndex = k Mod m_netCap
End Function


' ============================================================
' NetLogFormat (N-1、Private) - 1 件を桁を揃えた 1 行にする
' ============================================================
Private Function NetLogFormat(ByVal idx As Long) As String
    Dim k As Long
    k = NetRingIndex(idx)
    If k < 0 Then Exit Function

    NetLogFormat = Right$("     " & m_netSeq(k), 5) & "  " & _
                   Right$("       " & m_netMs(k), 7) & "  " & _
                   Left$(m_netMethod(k) & "        ", 8) & " " & _
                   Left$(NetCtxName(m_netCtx(k)) & "          ", 10) & " " & _
                   m_netUri(k)
End Function


' ============================================================
' NetCtxName (N-1、Private) - COREWEBVIEW2_WEB_RESOURCE_CONTEXT の名前
'   ★値は tools/vtable.py --enum COREWEBVIEW2_WEB_RESOURCE_CONTEXT で確定★
' ============================================================
Private Function NetCtxName(ByVal ctx As Long) As String
    Select Case ctx
        Case 0:  NetCtxName = "ALL"
        Case 1:  NetCtxName = "DOCUMENT"
        Case 2:  NetCtxName = "STYLESHEET"
        Case 3:  NetCtxName = "IMAGE"
        Case 4:  NetCtxName = "MEDIA"
        Case 5:  NetCtxName = "FONT"
        Case 6:  NetCtxName = "SCRIPT"
        Case 7:  NetCtxName = "XHR"
        Case 8:  NetCtxName = "FETCH"
        Case 9:  NetCtxName = "TEXTTRACK"
        Case 10: NetCtxName = "EVENTSRC"
        Case 11: NetCtxName = "WEBSOCKET"
        Case 12: NetCtxName = "MANIFEST"
        Case 13: NetCtxName = "SIGNEDEX"
        Case 14: NetCtxName = "PING"
        Case 15: NetCtxName = "CSPREPORT"
        Case 16: NetCtxName = "OTHER"
        Case Else: NetCtxName = "CTX" & ctx
    End Select
End Function


' ============================================================
' NetCtxFromName (N-1、Private) - 名前 (または数字) → 種別の値
'   分からなければ -1。★数字も受ける★ SDK が種別を増やしても直さずに済む。
' ============================================================
Private Function NetCtxFromName(ByVal ctxToken As String) As Long
    Dim t As String

    t = UCase$(Trim$(ctxToken))
    t = Replace$(t, "_", "")
    t = Replace$(t, " ", "")

    If LenB(t) = 0 Then
        NetCtxFromName = -1
        Exit Function
    End If

    If IsNumeric(t) Then
        NetCtxFromName = CLng(t)
        Exit Function
    End If

    Select Case t
        Case "ALL":                             NetCtxFromName = 0
        Case "DOCUMENT", "DOC":                 NetCtxFromName = 1
        Case "STYLESHEET", "CSS":               NetCtxFromName = 2
        Case "IMAGE", "IMG":                    NetCtxFromName = 3
        Case "MEDIA":                           NetCtxFromName = 4
        Case "FONT":                            NetCtxFromName = 5
        Case "SCRIPT", "JS":                    NetCtxFromName = 6
        Case "XHR", "XMLHTTPREQUEST":           NetCtxFromName = 7
        Case "FETCH":                           NetCtxFromName = 8
        Case "TEXTTRACK":                       NetCtxFromName = 9
        Case "EVENTSOURCE", "EVENTSRC":         NetCtxFromName = 10
        Case "WEBSOCKET", "WS":                 NetCtxFromName = 11
        Case "MANIFEST":                        NetCtxFromName = 12
        Case "SIGNEDEXCHANGE", "SIGNEDEX":      NetCtxFromName = 13
        Case "PING":                            NetCtxFromName = 14
        Case "CSPVIOLATIONREPORT", "CSPREPORT": NetCtxFromName = 15
        Case "OTHER":                           NetCtxFromName = 16
        Case Else:                              NetCtxFromName = -1
    End Select
End Function


' ============================================================
' ★N-2 : リクエストのヘッダとボディ★
'
'   N-1 の一覧 (URL / メソッド / 種別) の上に、★1 本を再現するのに要るもの★を足す。
'
'   ■ ★詳細はまるごと既定 OFF (論点8 案κ)★
'     N-1 の一覧は「軽くて常に使える道具」として完成している。そこに
'     ヘッダのイテレータ (1 要求あたり COM 呼び出し 60 回前後) とボディ読み出しを
'     既定で足すと、その良さを壊す。`NetDetailOn = True` にしたときだけ働く。
'
'   ■ ★読むと通信が壊れうる (論点2)★
'     `get_Content` が返す `IStream` は位置を持つ。読めば位置が進むので、
'     そのあと WebView2 が同じストリームから送ると★空のボディが飛ぶ★。
'     読み終わったら必ず `Seek(0, STREAM_SEEK_SET)` で先頭へ戻す。
'     ★これは「たぶん大丈夫」で済ませない★ ―― 検証では POST 先が返す本文と
'     突き合わせて、ボディが壊れずに届いたことを数える (設計原則112)。
'
'   ■ ★秘密は既定で伏せる (論点3 案M)★
'     `Cookie` / `Authorization` などはログファイルに平文で残さない。
'     ★リング上には生のまま持ち、出すときに伏せる★ ので、必要なら
'     `NetRedact = False` で見られる。既定は安全側。
'
'   ■ 追加した Public メンバ
'     NetDetailOn        As Boolean  ★既定 False★ 詳細を取るか
'     NetDetailCapacity  As Long     詳細リングの容量 (既定 50)
'     NetBodyMaxBytes    As Long     ボディの上限 (既定 65536)
'     NetRedact          As Boolean  ★既定 True★ 秘密を伏せるか
'     AddRedactHeader(name) / ClearRedactHeaders()
'     NetDetailUriFilter As String   詳細を取る URL の部分一致 (空なら全部)
'     NetDetailCount     As Long
'     NetDetailDropped   As Long     ★溢れて捨てた件数★ (黙って切らない)
'     NetDetailLine(i)   As String   seq/メソッド/URL/バイト数/切った/種類
'     NetDetailHeaders(i) As String  ヘッダ (伏せ字を適用して返す)
'     NetDetailBody(i)   As String   ボディ本文 (テキストのときだけ中身)
'     NetDetail(i)       As Boolean  1 本を Wv2Log へ出す
'     NetDetailAll()     As Long     全部出す
'     NetDetailClear()
'
'   ■ vtable index (★tools/vtable.py で確定★、設計原則115)
'       ICoreWebView2WebResourceRequest::get_Content  →IStream**   7
'       ICoreWebView2WebResourceRequest::get_Headers                9
'       ICoreWebView2HttpRequestHeaders::GetIterator                8
'       ...HttpHeadersCollectionIterator::GetCurrentHeader          3
'       ...::get_HasCurrentHeader / ::MoveNext                    4 / 5
'       IStream::Read / Seek        (Windows SDK objidl.h より)   3 / 5
' ============================================================


' ============================================================
' NetDetailOn (N-2) - 詳細 (ヘッダとボディ) を取るか
'   ★既定 False★ 一覧だけなら N-1 のまま軽い。
' ============================================================
Public Property Get NetDetailOn() As Boolean
    NetDetailOn = m_detOn
End Property

Public Property Let NetDetailOn(ByVal onOff As Boolean)
    m_detOn = onOff
    Wv2Log.LogD "Wv2Pane.NetDetailOn = " & onOff & _
                IIf(onOff, " (ヘッダとボディも読む)", " (一覧だけ)")
End Property


' ============================================================
' NetDetailCapacity (N-2) - 詳細リングの容量 (件)
'   ★一覧とは別のリング (論点5 案S)★ 一覧は「順番と全体像」、詳細は
'   「これを再現したい 1 本」で用途が違う。詳細のために一覧を削らない。
'   ★変えると中身は捨てられる★
' ============================================================
Public Property Get NetDetailCapacity() As Long
    NetDetailCapacity = m_detCap
End Property

Public Property Let NetDetailCapacity(ByVal n As Long)
    If n < 1 Then n = 1
    If n > NET_DET_CAP_MAX Then
        Wv2Log.LogW "Wv2Pane.NetDetailCapacity: " & n & " は上限 " & _
                    NET_DET_CAP_MAX & " を超えるので丸める"
        n = NET_DET_CAP_MAX
    End If
    m_detCap = n
    NetDetailRingAlloc
    Wv2Log.LogD "Wv2Pane.NetDetailCapacity = " & m_detCap & " (中身は作り直したので空)"
End Property


' ============================================================
' NetBodyMaxBytes (N-2) - ボディを読む上限 (バイト)
'   ★ハンドラの中で同期的に読む★ ので、青天井にすると通信が詰まる。
'   上限で切ったことは NetDetailLine の「切った」欄に出る (黙って切らない)。
' ============================================================
Public Property Get NetBodyMaxBytes() As Long
    NetBodyMaxBytes = m_detBodyMax
End Property

Public Property Let NetBodyMaxBytes(ByVal n As Long)
    If n < 0 Then n = 0
    If n > NET_BODY_MAX_LIMIT Then n = NET_BODY_MAX_LIMIT
    m_detBodyMax = n
    Wv2Log.LogD "Wv2Pane.NetBodyMaxBytes = " & m_detBodyMax
End Property


' ============================================================
' NetRedact (N-2) - 秘密のヘッダを伏せるか
'   ★既定 True★ Cookie / Authorization などをログに平文で残さない。
'   ★リング上には生のまま持っている★ ので、False にすればその場で見られる。
' ============================================================
Public Property Get NetRedact() As Boolean
    NetRedact = m_detRedact
End Property

Public Property Let NetRedact(ByVal onOff As Boolean)
    m_detRedact = onOff
    If Not onOff Then
        Wv2Log.LogW "Wv2Pane.NetRedact = False ★Cookie などがログに平文で残る★ " & _
                    "ログを人に渡す前に必ず確かめること。"
    Else
        Wv2Log.LogD "Wv2Pane.NetRedact = True (秘密のヘッダは伏せる)"
    End If
End Property


' ============================================================
' AddRedactHeader / ClearRedactHeaders (N-2)
'   伏せる対象を足す / 既定リストごと空にする。名前は大小を問わない。
'   既定リストは NetRedactDefaults 参照。
' ============================================================
Public Sub AddRedactHeader(ByVal headerName As String)
    Dim t As String
    t = LCase$(Trim$(headerName))
    If LenB(t) = 0 Then Exit Sub
    If InStr(1, m_detRedactList, "|" & t & "|") > 0 Then Exit Sub
    m_detRedactList = m_detRedactList & t & "|"
    Wv2Log.LogD "Wv2Pane.AddRedactHeader: " & t
End Sub

Public Sub ClearRedactHeaders()
    m_detRedactList = "|"
    Wv2Log.LogW "Wv2Pane.ClearRedactHeaders: ★伏せる対象が 1 つも無くなった★ " & _
                "Cookie も平文で出る。"
End Sub


' ============================================================
' NetDetailUriFilter (N-2) - 詳細を取る URL の部分一致
'   空なら全部。★ヘッダのイテレータは安くない★ ので、狙いが決まっているなら
'   ここで絞ると軽い。
' ============================================================
Public Property Get NetDetailUriFilter() As String
    NetDetailUriFilter = m_detUriFilter
End Property

Public Property Let NetDetailUriFilter(ByVal pattern As String)
    m_detUriFilter = pattern
    Wv2Log.LogD "Wv2Pane.NetDetailUriFilter = [" & pattern & "]" & _
                IIf(LenB(pattern) = 0, " (絞らない)", "")
End Property


' ============================================================
' NetDetailCount (N-2) - 今リングに入っている詳細の件数
' ============================================================
Public Property Get NetDetailCount() As Long
    NetDetailCount = m_detCount
End Property


' ============================================================
' NetDetailDropped (N-2) - 詳細リングから溢れて捨てた件数
'   ★黙って切らない★ 一覧の NetLogDropped と同じ役目。
'   NetDetailClear / NetDetailCapacity の変更で 0 に戻る。
' ============================================================
Public Property Get NetDetailDropped() As Long
    NetDetailDropped = m_detDropped
End Property


' ============================================================
' NetDetailLine (N-2) - 詳細 1 本の見出しを読む (古い順・1 起点)
'   vbTab 区切りで 6 項目:
'     通し番号 / メソッド / URL / ボディのバイト数 / 切ったか / Content-Type
'   ★ヘッダとボディの中身は含まない★ (NetDetailHeaders / NetDetailBody で取る)
' ============================================================
Public Property Get NetDetailLine(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailLine = m_detSeq(k) & vbTab & m_detMethod(k) & vbTab & m_detUri(k) & _
                    vbTab & m_detBytes(k) & vbTab & CStr(m_detTrunc(k)) & _
                    vbTab & m_detCType(k)
End Property


' ============================================================
' NetDetailHeaders (N-2) - 詳細 1 本のヘッダを読む
'   1 行 1 ヘッダで vbLf 区切り。★NetRedact = True なら秘密は伏せて返す★
' ============================================================
Public Property Get NetDetailHeaders(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailHeaders = NetRedactApply(m_detHeaders(k))
End Property


' ============================================================
' NetDetailBody (N-2) - 詳細 1 本のボディを読む
'   ★テキストと判定できたときだけ中身が入る★ (論点4 案P)。
'   バイナリなら空文字。バイト数は NetDetailLine の 4 項目目を見ること。
' ============================================================
Public Property Get NetDetailBody(ByVal idx As Long) As String
    Dim k As Long
    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Property

    NetDetailBody = m_detBody(k)
End Property


' ============================================================
' NetDetail (N-2) - 詳細 1 本を Wv2Log へ出す
'   ★一覧 (NetLogDrain) には混ぜない (論点6 案V)★ 1 要求で数十行になるので、
'   混ぜると一覧が読めなくなる (設計原則105 の趣旨)。
'   戻り値: True なら出した。
' ============================================================
Public Function NetDetail(ByVal idx As Long) As Boolean
    Dim k As Long
    Dim i As Long
    Dim lines As Variant

    k = NetDetailRingIndex(idx)
    If k < 0 Then Exit Function

    Wv2Log.LogI "  --- 詳細 #" & m_detSeq(k) & " ---"
    Wv2Log.LogI "      " & m_detMethod(k) & " " & m_detUri(k)

    If LenB(m_detHeaders(k)) > 0 Then
        lines = Split(NetRedactApply(m_detHeaders(k)), vbLf)
        Wv2Log.LogI "      [ヘッダ " & (UBound(lines) - LBound(lines) + 1) & " 個]"
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [ヘッダ] 取れていない"
    End If

    If m_detBytes(k) = 0 Then
        Wv2Log.LogI "      [ボディ] 無し"
    ElseIf m_detIsText(k) Then
        Wv2Log.LogI "      [ボディ] " & m_detBytes(k) & " バイト (テキスト" & _
                    IIf(m_detTrunc(k), "・★上限で切った★", "") & ")"
        lines = Split(Replace$(m_detBody(k), vbCrLf, vbLf), vbLf)
        For i = LBound(lines) To UBound(lines)
            Wv2Log.LogI "        " & lines(i)
        Next i
    Else
        Wv2Log.LogI "      [ボディ] " & m_detBytes(k) & " バイト (バイナリ" & _
                    IIf(m_detTrunc(k), "・★上限で切った★", "") & _
                    ") 先頭: " & m_detHex(k)
    End If

    NetDetail = True
End Function


' ============================================================
' NetDetailAll (N-2) - 溜まっている詳細を全部 Wv2Log へ出す
'   ★明示的に呼んだときだけ★ 出る。戻り値は出した件数。
'   ★出しても消えない★ (NetLogDrain と違う。判定に使いたいので残す)
' ============================================================
Public Function NetDetailAll() As Long
    Dim i As Long

    If m_detCount = 0 Then
        Wv2Log.LogI "  --- 詳細: 溜まっていない ---"
        Exit Function
    End If

    Wv2Log.LogI "  --- 詳細 " & m_detCount & " 件 ---"
    If m_detDropped > 0 Then
        Wv2Log.LogW "  ★詳細リングから溢れて捨てたものが " & m_detDropped & " 件ある★ " & _
                    "NetDetailCapacity を増やすか、NetDetailUriFilter で絞ること。"
    End If
    For i = 1 To m_detCount
        NetDetail i
    Next i
    NetDetailAll = m_detCount
End Function


' ============================================================
' NetDetailClear (N-2) - 詳細リングを空にする
' ============================================================
Public Sub NetDetailClear()
    NetDetailRingAlloc
End Sub


' ============================================================
' NetReadHeaders (N-2、Private) - 要求のヘッダを 1 本の文字列にする
'
'   ICoreWebView2HttpRequestHeaders::GetIterator (8) でイテレータを取り、
'   GetCurrentHeader (3) / get_HasCurrentHeader (4) / MoveNext (5) で回す。
'
'   ★ハンドラの中で走る★ ので上限 (NET_HDR_MAX) で切る。
'   ★GetCurrentHeader は LPWSTR を 2 本返す★ ので、どちらも CoTaskMemFree する。
'
'   戻り値: "name: value" を vbLf で繋いだもの。取れなければ空文字。
' ============================================================
Private Function NetReadHeaders(ByVal pReq As LongPtr) As String
    Dim pHeaders As LongPtr
    Dim pIter    As LongPtr
    Dim hr       As Long
    Dim n        As Long
    Dim pName    As LongPtr
    Dim pValue   As LongPtr
    Dim moreFlag As Long
    Dim sb       As String

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 9, "", VarPtr(pHeaders))          ' get_Headers
    If hr <> 0 Or pHeaders = 0 Then Exit Function

    hr = dcf(pHeaders, 8, "", VarPtr(pIter))         ' GetIterator
    If hr <> 0 Or pIter = 0 Then
        ComRelease pHeaders
        Exit Function
    End If

    Do While n < NET_HDR_MAX
        If Not GetBoolProperty(pIter, 4, "") Then Exit Do   ' get_HasCurrentHeader

        pName = 0
        pValue = 0
        hr = dcf(pIter, 3, "", VarPtr(pName), VarPtr(pValue))   ' GetCurrentHeader
        If hr = 0 Then
            If LenB(sb) > 0 Then sb = sb & vbLf
            sb = sb & PtrToString(pName) & ": " & PtrToString(pValue)
            n = n + 1
        End If
        If pName <> 0 Then CoTaskMemFree pName
        If pValue <> 0 Then CoTaskMemFree pValue
        If hr <> 0 Then Exit Do

        moreFlag = 0
        hr = dcf(pIter, 5, "", VarPtr(moreFlag))     ' MoveNext(BOOL*)
        If hr <> 0 Or moreFlag = 0 Then Exit Do
    Loop

    If n >= NET_HDR_MAX Then
        sb = sb & vbLf & "(★" & NET_HDR_MAX & " 個で打ち切った★)"
    End If

    ComRelease pIter
    ComRelease pHeaders
    NetReadHeaders = sb
End Function


' ============================================================
' NetReadBody (N-2、Private) - 要求のボディを読む
'
'   ICoreWebView2WebResourceRequest::get_Content (7) が返す IStream を
'   IStream::Read (3) で読み、★読み終わったら Seek (5) で先頭へ戻す★。
'
'   ★戻さないと通信が壊れる (論点2)★
'     IStream は位置を持つ。読んだままにすると、そのあと WebView2 が
'     同じストリームから送るときに★空のボディが飛ぶ★。
'
'   引数 (すべて出力):
'     outBytes  : 読めたバイト数 (上限で切られていれば上限値)
'     outTrunc  : 上限で切ったか
'     outBuf()  : 読んだバイト列 (0 起点、outBytes 個が有効)
'
'   戻り値: True ならストリームがあって読み終えた (0 バイトでも True)。
'           False は「ボディが無い」(GET など) か取得に失敗した。
' ============================================================
Private Function NetReadBody(ByVal pReq As LongPtr, _
                             ByRef outBytes As Long, _
                             ByRef outTrunc As Boolean, _
                             ByRef outBuf() As Byte) As Boolean
    Dim pStream As LongPtr
    Dim hr      As Long
    Dim cap     As Long
    Dim got     As Long
    Dim chunk   As Long
    Dim cbRead  As Long

    outBytes = 0
    outTrunc = False

    If pReq = 0 Then Exit Function

    hr = dcf(pReq, 7, "", VarPtr(pStream))           ' get_Content
    If hr <> 0 Or pStream = 0 Then Exit Function     ' ボディ無し (GET など)

    cap = m_detBodyMax
    If cap < 1 Then
        ' 上限 0 = 読まない。それでも位置は動かしていないので戻す必要はない。
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ReDim outBuf(0 To cap - 1)

    Do While got < cap
        chunk = cap - got
        If chunk > NET_BODY_CHUNK Then chunk = NET_BODY_CHUNK
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(outBuf(got)), chunk, VarPtr(cbRead))  ' Read
        If hr <> 0 Then Exit Do
        If cbRead <= 0 Then Exit Do
        got = got + cbRead
    Loop

    ' ★まだ残っているか★ を 1 バイトだけ試して確かめる (黙って切らないため)
    If got >= cap Then
        Dim probe(0 To 0) As Byte
        cbRead = 0
        hr = dcf(pStream, 3, "", VarPtr(probe(0)), 1&, VarPtr(cbRead))
        If hr = 0 And cbRead > 0 Then outTrunc = True
    End If

    ' ★★ 読んだぶんを必ず戻す ★★ (論点2 案X)
    '   Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER* plibNewPosition)
    '   dlibMove は 8 バイトの値渡し、dwOrigin は 4 バイト、最後は NULL 可。
    hr = dcf(pStream, 5, "", 0^, STREAM_SEEK_SET, 0^)
    If hr <> 0 Then
        ' ★戻せなかった★ これは通信を壊した可能性があるので黙って進まない。
        ' ハンドラの中なのでログは撃たず、ボディの中身に印を残す。
        outTrunc = True
        outBytes = got
        ComRelease pStream
        NetReadBody = True
        Exit Function
    End If

    ComRelease pStream
    outBytes = got
    NetReadBody = True
End Function


' ============================================================
' NetDetailPush (N-2、Private) - 詳細を 1 件リングに置く
'   ★溢れたら古いものを上書きする★ (一覧と同じ作法)。
' ============================================================
Private Sub NetDetailPush(ByVal seqNo As Long, _
                          ByVal mth As String, _
                          ByVal uri As String, _
                          ByVal headers As String, _
                          ByVal cType As String, _
                          ByVal nBytes As Long, _
                          ByVal trunc As Boolean, _
                          ByVal isText As Boolean, _
                          ByVal bodyText As String, _
                          ByVal hexHead As String)
    m_detSeq(m_detHead) = seqNo
    m_detMethod(m_detHead) = mth
    m_detUri(m_detHead) = uri
    m_detHeaders(m_detHead) = headers
    m_detCType(m_detHead) = cType
    m_detBytes(m_detHead) = nBytes
    m_detTrunc(m_detHead) = trunc
    m_detIsText(m_detHead) = isText
    m_detBody(m_detHead) = bodyText
    m_detHex(m_detHead) = hexHead

    m_detHead = m_detHead + 1
    If m_detHead >= m_detCap Then m_detHead = 0
    If m_detCount < m_detCap Then
        m_detCount = m_detCount + 1
    Else
        m_detDropped = m_detDropped + 1
    End If
End Sub


' ============================================================
' NetDetailRingAlloc / NetDetailRingIndex (N-2、Private)
' ============================================================
Private Sub NetDetailRingAlloc()
    If m_detCap < 1 Then m_detCap = NET_DET_CAP_DEFAULT

    ReDim m_detSeq(0 To m_detCap - 1)
    ReDim m_detMethod(0 To m_detCap - 1)
    ReDim m_detUri(0 To m_detCap - 1)
    ReDim m_detHeaders(0 To m_detCap - 1)
    ReDim m_detCType(0 To m_detCap - 1)
    ReDim m_detBytes(0 To m_detCap - 1)
    ReDim m_detTrunc(0 To m_detCap - 1)
    ReDim m_detIsText(0 To m_detCap - 1)
    ReDim m_detBody(0 To m_detCap - 1)
    ReDim m_detHex(0 To m_detCap - 1)

    m_detHead = 0
    m_detCount = 0
    m_detDropped = 0
End Sub

Private Function NetDetailRingIndex(ByVal idx As Long) As Long
    Dim k As Long

    If idx < 1 Or idx > m_detCount Then
        NetDetailRingIndex = -1
        Exit Function
    End If

    k = m_detHead - m_detCount + (idx - 1)
    Do While k < 0
        k = k + m_detCap
    Loop
    NetDetailRingIndex = k Mod m_detCap
End Function


' ============================================================
' NetRedactApply (N-2、Private) - 秘密のヘッダを伏せる
'   ★出すときに伏せる★ リング上は生のまま持っているので、
'   NetRedact = False にすればその場で見られる (論点3 案M)。
' ============================================================
Private Function NetRedactApply(ByVal headers As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String
    Dim out   As String

    If LenB(headers) = 0 Then Exit Function
    If Not m_detRedact Then
        NetRedactApply = headers
        Exit Function
    End If

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If InStr(1, m_detRedactList, "|" & nm & "|") > 0 Then
                lines(i) = Left$(CStr(lines(i)), p) & " <伏せた " & _
                           (Len(CStr(lines(i))) - p - 1) & " 文字>"
            End If
        End If
        If LenB(out) > 0 Then out = out & vbLf
        out = out & CStr(lines(i))
    Next i

    NetRedactApply = out
End Function


' ============================================================
' NetIsTextType (N-2、Private) - Content-Type がテキスト系か
'   ★判定できないときは False (バイナリ扱い)★ 安全側。
'   Content-Type が無いときは、フォーム送信の既定である
'   application/x-www-form-urlencoded とみなさず False にする
'   (それらしい値を書き残さない = 設計原則111)。
' ============================================================
Private Function NetIsTextType(ByVal cType As String) As Boolean
    Dim t As String
    t = LCase$(cType)
    If LenB(t) = 0 Then Exit Function

    If InStr(1, t, "text/") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "json") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "xml") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "x-www-form-urlencoded") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "javascript") > 0 Then NetIsTextType = True: Exit Function
    If InStr(1, t, "charset=") > 0 Then NetIsTextType = True: Exit Function
End Function


' ============================================================
' NetHexHead (N-2、Private) - バイナリの先頭を 16 進で見せる
' ============================================================
Private Function NetHexHead(ByRef b() As Byte, ByVal cb As Long) As String
    Dim i As Long
    Dim n As Long
    Dim s As String

    n = cb
    If n > NET_HEX_HEAD Then n = NET_HEX_HEAD
    For i = 0 To n - 1
        s = s & Right$("0" & Hex$(b(i)), 2) & " "
    Next i
    If cb > n Then s = s & "..."
    NetHexHead = RTrim$(s)
End Function


' ============================================================
' NetHeaderValue (N-2、Private) - 読んだヘッダから 1 つ取り出す
'   名前は大小を問わない。無ければ空文字。
' ============================================================
Private Function NetHeaderValue(ByVal headers As String, ByVal headerName As String) As String
    Dim lines As Variant
    Dim i     As Long
    Dim p     As Long
    Dim nm    As String

    If LenB(headers) = 0 Then Exit Function

    lines = Split(headers, vbLf)
    For i = LBound(lines) To UBound(lines)
        p = InStr(1, CStr(lines(i)), ":")
        If p > 1 Then
            nm = LCase$(Trim$(Left$(CStr(lines(i)), p - 1)))
            If nm = LCase$(headerName) Then
                NetHeaderValue = Trim$(Mid$(CStr(lines(i)), p + 1))
                Exit Function
            End If
        End If
    Next i
End Function


' ============================================================
' NetRedactDefaults (N-2、Private) - 既定で伏せるヘッダ
'   ★ここに無いものは平文で出る★ 足したいときは AddRedactHeader。
' ============================================================
Private Function NetRedactDefaults() As String
    NetRedactDefaults = "|cookie|set-cookie|authorization|proxy-authorization|" & _
                        "x-api-key|x-auth-token|x-csrf-token|api-key|"
End Function
