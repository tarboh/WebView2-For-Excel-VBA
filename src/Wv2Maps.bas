Attribute VB_Name = "Wv2Maps"
''''''''''''''''''''''''''''''''''
' --- Wv2Maps.bas  K-4d 段階 (MapsOpen の各段階に時刻を刻む) ---
'
'   ★ロジックは 1 行も変えていない。ログ行を足しただけ。★
'
'   なぜ要るか: MapsOpen は実測 17 秒かかるのに、Wv2Tests 側で中身を
'   段階的に再現すると ★3.9 秒で終わる★。1/4 しか再現できていない。
'   ★13 秒をどこで使っているのかが分からない★ ので、本物を直接測る。
'
'   これは 2 つの現象の共通の入口:
'     現象1 MapsOpen がのんびり待っている (17 秒)
'     現象2 MapsOpen の後だとフォームを閉じるのが遅い (5 秒)
'   ★Test_K4_Step の再現では現象2 が一度も出なかった★ ので、
'   本物との差がそのまま原因に繋がっている見込みが高い。
''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''
' --- Wv2Maps.bas  D-7d 段階 (実用: 住所 → 座標 / 進捗と中断) ---
'
'   ★D 軸で作った部品を、業務でそのまま使える形に畳んだ最初の例。★
'   D-4d で Google Maps の住所検索が実サイトで通ることを実証したので、
'   その手順を 1 本の関数にした。
'
'   ■ D-7 で足したもの (進捗と中断)
'     MapsCancel    … 中断を要求する / 今の要求を読む (Property Get / Let)
'     MapsCanceled  … 直前の呼び出しが中断で終わったか (読み取り専用)
'     MapsCountRows … 処理する行数を先に数える (分母を出す口)
'
'     ★バッチ実行中に Esc を押すと止まる★
'     進捗は Application.StatusBar に出す。
'
'   ■ ★ステータスバーの戻し方 (D-7d の実測)★
'     ★Application.StatusBar = False では戻らない★ 文字列 "FALSE" が
'     表示されたまま残る (画面にも FALSE と出る)。実測:
'
'       False        → [FALSE] 型=String    ← 値として文字列化される
'       Empty        → [False] 型=Boolean   ← ★これが正解★
'       空文字 ""     → [False] 型=Boolean
'       vbNullString → [False] 型=Boolean
'       CVar(False)  → [FALSE] 型=String
'
'     よく言われる「False を入れれば Excel に制御が戻る」は★この環境では逆★。
'     Test_D7_StatusBar が同じ測定をいつでも再現できる。
'
'   ■ ★Esc をどう捕まえるか (D-7 → D-7b → D-7c と 2 回作り直した)★
'
'     ★実機で 2 回とも前提が外れた。経緯を残す。★
'
'     D-7  : GetAsyncKeyState だけで見た → ★一度も発火しなかった★。
'            VBA ランタイムの Esc は焦点に依存せず届くので、こちらが読むより先に
'            「コードの実行が中断されました」で止まっていた。
'     D-7b : EnableCancelKey = xlErrorHandler にして Esc をエラー18 で受けた
'            → ★ハンドラに届かなかった★。エラー18 が発生するのは主に
'            ★WebView2 の COM コールバック (View_On*) の中★で、そこは呼び出し元が
'            VBA ではない (サンク経由) ため On Error GoTo に遡らない。
'            Wv2Pane のガードが再送出し、COM 境界で未処理ダイアログになった。
'     D-7c : ★EnableCancelKey = xlDisabled★ にして、エラー18 自体を発生させない。
'            実行は止まらないので、次に MapsPump が回ったとき
'            ★GetAsyncKeyState が拾う★。経路が 1 本になる。
'
'     ★なぜ break を避けたいのか★ 仕様事実20 のクラッシュ窓そのものだから。
'     WebView2 のイベントが動いている最中の break は AV になりうる。
'     使いやすさだけでなく安全の問題でもある。
'
'     ★失うもの★ バッチ実行中は Excel の Esc / Ctrl+Break が効かない。
'     ただし待ちはすべてタイムアウトを持ち (EvalSync / WaitFor / WaitSettled)、
'     MapsPump は 0.2～0.3 秒ごとに必ず回るので、無限には固まらない。
'
'     ★戻し忘れると Excel 全体で Esc が効かなくなる★ ので、すべての出口で戻す。
'     そのために MapsOpen / MapsGeocode は「ラッパ + Core」の形にしてある
'     (Core は出口が多いので、後始末はラッパが 1 箇所で引き受ける)。
'
'   ■ 使い方
'     Dim p As Wv2Pane, lat As Double, lng As Double, nm As String
'     Set p = Wv2Maps.MapsOpen(UserForm1.CurrentBrowser)
'     If Wv2Maps.MapsGeocode(p, "東京都千代田区丸の内1-9-1", lat, lng, nm) Then
'         Debug.Print lat & "," & lng & " " & nm
'     Else
'         Debug.Print "失敗: " & Wv2Maps.MapsLastError
'     End If
'
'     ★同じ Pane を使い回せる★ ので、住所を次々に処理できる (バッチ)。
'     タブを開き直すより速い。
'
'   ■ なぜ Maps 固有の知識をここに隔離するか
'     セレクタも URL の形も★Google の都合で予告なく変わる★ (設計原則75)。
'     製品コア (Wv2Pane / Wv2Element) に混ぜると、壊れたときにどこを直せばよいか
'     分からなくなる。**変わるものは 1 箇所に閉じ込める。**
'
'   ■ ★完了の判定に何を使っているか★
'     `ArmUrlSignal` は使わない。2 件目以降は既に /place/ にいるので、
'     arm した瞬間に当たってしまう (D-4c で文書化した罠)。代わりに
'     ★検索前の URL を覚えておき、/place/ と !3d を含む別の URL になるまで待つ★。
'     「URL が変わったら完了」では早すぎる (最初は /search/ に変わるだけで、
'     座標は前の場所のまま)。詳しくは MapsWaitPlace のコメント。
'     URL の取得は COM 経由 (View_GetSource) なので EvalSync を使わず軽い。
'     URL が変わったあとで WaitSettled を掛け、DOM が落ち着くのを待つ。
'
'   ■ 座標をどこから取るか
'     URL の !3d<緯度>!4d<経度> を第一候補にする (★場所そのものの座標★)。
'     無ければ /@<緯度>,<経度>,<ズーム> (地図の中心) に落とす。
'     ★Val を使う★ CDbl はロケール依存で、小数点が , の環境で壊れる。
'     Val は常に . を小数点として読み、数字でない文字で止まる。
'
'   ■ 戻り値の規約 (設計原則93 を踏襲)
'     True                       … /place/ に着いた = 1 件に確定した
'     False + MapsLastError      … 理由:
'       no-pane / no-searchbox / no-search-trigger … 構造が変わった (要 Test_D4_Dom)
'       timeout-url              … 時間内に URL が変わらなかった
'       not-found                … 見つからなかった (★検索しても地図が動かない★)。
'                                  ★このとき座標は返さない★ 前の検索の残りかすなので
'       ambiguous                … ★候補が複数★ (/search/ に着いたが地図は動いた)。
'                                  ★pickFirstIfAmbiguous:=True なら候補の 1 件目に
'                                  入って確定させる (MapsLastPicked = True になる)★
'                                  このとき outLat / outLng には地図の中心が入る
'       no-coords                … URL から座標を読めなかった
'       canceled                 … ★Esc または MapsCancel で中断された★ (D-7)
''''''''''''''''''''''''''''''''''
Option Explicit


' --- 直前の失敗理由 ---
Private m_lastError As String

' --- D-6b: 直前の結果が★候補一覧から選んだもの★かどうか ---
'   ★確定と「候補から選んだ」を同じ顔で並べない★ (設計原則111 と同じ筋)。
'   精度が違うものが同じ ok として混ざると、後から見分けられなくなる。
Private m_lastPicked As Boolean

' --- D-7: 中断 (Esc) と進捗 ---
'   ★フォーカスに依存しない手段を選ぶ★ Application.EnableCancelKey (Esc /
'   Ctrl+Break) は Excel に焦点が無いと届かない。バッチ中の焦点は WebView2 側に
'   あるので、キーの状態を直接読む GetAsyncKeyState を使う。
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long) As Integer

Private Const MAPS_VK_ESCAPE As Long = &H1B&

' 中断が要求されたか。★外から立てることもできる★ (Wv2Maps.MapsCancel = True)
Private m_cancel As Boolean

' 直前の呼び出しが中断で終わったか
Private m_canceled As Boolean

' バッチ (MapsGeocodeSheet) の実行中か。
'   ★入口で中断状態を捨てるのは「一番外側」だけ★ 単発の MapsGeocode も入口で
'   捨てるので、これが無いと 1 行処理するたびに要求が消えてバッチが止まらない。
Private m_inBatch As Boolean

' D-7d: MapsCheckCancel が呼ばれた回数。★数えられるものは数える★ (設計原則112)
'   Esc を拾えないとき、原因が「呼ばれていない」のか「呼んでも見えない」のかは
'   推測では分けられない。1 回のバッチで何回ポーリングできているかを実測する。
Private m_checkCount As Long

' D-7b: arm する前の Application.EnableCancelKey。★必ず戻す★
'   xlDisabled は 0 なので「保存していない」と見分けが付かないが、
'   0 のときは xlInterrupt (既定) に戻す ― 砦を戻す方が安全なので。
Private m_savedCancelKey As Long

' 候補一覧のリンクの候補 (D-6b)。★href で拾うのを先に★ class は変わりやすい。
Private Const MAPS_CAND_SELECTORS As String = _
    "a[href*='/maps/place/']" & vbLf & "div[role='feed'] a[href*='/place/']" & vbLf & "a.hfpxzc"

' 検索ボックスの候補。★id は自動生成に変わりうる (仕様事実59)★ ので name / 構造を先に。
'   ★区切りは vbLf★ Const では Chr$() が使えない (定数式でないため)。
'   CSS セレクタに改行は現れないので衝突しない。
Private Const MAPS_BOX_SELECTORS As String = _
    "input[name='q']" & vbLf & "#searchboxinput" & vbLf & "form input[type='text']"

' 検索実行の候補。aria-label は UI 言語で変わるので英語も並べる。
Private Const MAPS_GO_SELECTORS As String = _
    "button[aria-label='検索']" & vbLf & "#searchbox-searchbutton" & vbLf & "button[aria-label='Search']"


' ============================================================
' MapsOpen - Google マップのタブを開いて、操作できる状態まで待つ
'
'   戻り値: 使える Wv2Pane。失敗したら Nothing (理由は MapsLastError)
'
'   ★タブを開いた直後は EvalSync が no-view で失敗する★ ので、
'   先に JS が通るようになるまで待つ (D-4d の教訓)。
' ============================================================
Public Function MapsOpen(ByVal b As Wv2Browser, _
                         Optional ByVal timeoutSec As Single = 30) As Wv2Pane
    Dim p As Wv2Pane
    Dim solo As Boolean
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    ' --- D-7b: 単発で呼ばれたときだけ arm する (バッチ側が既に arm 済み) ---
    solo = Not m_inBatch
    If solo Then
        MapsResetCancel
        MapsArmCancelKey
        On Error GoTo Cleanup
    End If

    Set p = MapsOpenCore(b, timeoutSec)

Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled にしたので通常は来ない。arm する前後の窓のために残す。
        Set p = Nothing
        m_cancel = True
        m_canceled = True
        m_lastError = "canceled"
        Wv2Log.LogW "Wv2Maps.MapsOpen: ★Esc で中断された★"
        errNo = 0
    End If

    If solo Then MapsDisarmCancelKey
    Set MapsOpen = p

    ' ★他のエラーは握り潰さない★ 後始末だけしてそのまま上げる
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function


' ============================================================
' MapsOpenCore (D-7b、Private) - MapsOpen の中身
'   ★出口が多いので、EnableCancelKey の後始末はラッパに任せる★
' ============================================================
Private Function MapsOpenCore(ByVal b As Wv2Browser, _
                              ByVal timeoutSec As Single) As Wv2Pane
    Dim p As Wv2Pane
    Dim el As Wv2Element
    Dim t0 As Single
    Dim tStep As Single      ' K-4d: 段階ごとの計測用
    Dim evalTries As Long    ' K-4d: EvalSync を何回撃ったか

    m_lastError = ""

    If b Is Nothing Then
        m_lastError = "no-browser"
        Exit Function
    End If

    tStep = Timer
    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    Wv2Log.LogD "[MapsOpen] AddTabWithUrl … " & _
                Format$(MapsElapsed(tStep), "0.00") & " 秒"   ' K-4d
    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ' --- JS が通るようになるまで待つ ---
    t0 = Timer
    tStep = Timer
    Do
        evalTries = evalTries + 1   ' K-4d
        p.EvalSync "1", 3
        If p.LastEvalOk Then Exit Do
        If MapsCheckCancel() Then           ' D-7
            m_lastError = "canceled"
            m_canceled = True
            Exit Function
        End If
        If MapsElapsed(t0) >= timeoutSec Then
            m_lastError = "pane-not-ready"
            Exit Function
        End If
        MapsPump 0.3
    Loop
    Wv2Log.LogD "[MapsOpen] JS が通るまで … " & _
                Format$(MapsElapsed(tStep), "0.00") & " 秒 (EvalSync " & _
                evalTries & " 回)"   ' K-4d

    ' --- 検索ボックスが出るまで待つ ---
    tStep = Timer
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, timeoutSec)
    Wv2Log.LogD "[MapsOpen] 検索ボックス待ち … " & _
                Format$(MapsElapsed(tStep), "0.00") & " 秒 " & _
                IIf(el Is Nothing, "(外れ)", "(当たり)") & _
                " レジストリ " & p.ElementCount & " 個"   ' K-4d
    If el Is Nothing Then
        m_lastError = "no-searchbox"
        Exit Function
    End If

    ' --- ★地図が位置を持つまで待つ★ ---
    '   開いた直後の URL は https://www.google.com/maps で座標を含まない。
    '   MapsGeocode は「検索して地図が動いたか」で見つかったかを判定するので、
    '   ★比較の基準になる座標が無いと判定できない★。1～2 秒で /@ が付くので待つ。
    '   付かなくても致命的ではない (その場合は保守的に not-found 扱いになる)。
    t0 = Timer
    tStep = Timer
    Do
        If InStr(1, p.View_GetSource(), "/@") > 0 Then Exit Do
        If MapsCheckCancel() Then Exit Do   ' D-7 (致命的ではないので抜けるだけ)
        If MapsElapsed(t0) >= 8 Then
            Wv2Log.LogW "Wv2Maps.MapsOpen: URL に座標が入らない (" & _
                        p.View_GetSource() & ")"
            Exit Do
        End If
        MapsPump 0.3
    Loop

    ' --- ★Maps 固有の知識をここで効かせる★ ---
    '   検索は /search?tbm=map を開いたまま結果を流し続ける (仕様事実58)。
    '   静穏判定から外さないと、1 件ごとに StaleInflightMs (既定 10 秒) を
    '   丸ごと待つことになる。名指しで外せば 1 秒で落ち着く。
    '   ★これが「変わるものを 1 箇所に閉じ込める」ことの利得★
    Wv2Log.LogD "[MapsOpen] /@ 待ち … " & _
                Format$(MapsElapsed(tStep), "0.00") & " 秒"   ' K-4d

    p.AddIgnoreNetwork "/search?tbm=map"

    Wv2Log.LogI "Wv2Maps.MapsOpen: 準備できた (" & p.View_GetSource() & ")"
    Set MapsOpenCore = p
End Function


' ============================================================
' MapsGeocode - 住所を検索して座標を取る
'
'   引数:
'     p           : MapsOpen が返した Pane (使い回せる)
'     addressText : 検索する住所
'     outLat / outLng : 緯度・経度
'     outName     : Maps が返した正規化後の名前 (例 〒100-0005 東京都千代田区…)
'     timeoutSec  : URL が変わるのを待つ上限
'
'   戻り値: True = 1 件に確定した / False = 理由は MapsLastError
' ============================================================
Public Function MapsGeocode(ByVal p As Wv2Pane, _
                            ByVal addressText As String, _
                            ByRef outLat As Double, _
                            ByRef outLng As Double, _
                            ByRef outName As String, _
                            Optional ByVal timeoutSec As Single = 20, _
                            Optional ByVal pickFirstIfAmbiguous As Boolean = False) As Boolean
    Dim r As Boolean
    Dim solo As Boolean
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String

    ' --- D-7b: 単発で呼ばれたときだけ arm する (バッチ側が既に arm 済み) ---
    '   ★入口で中断状態を捨てる★ のも単発のときだけ (設計原則106)。
    '   バッチの中で捨てると 1 行ごとに要求が消えて止まらなくなる。
    solo = Not m_inBatch
    If solo Then
        MapsResetCancel
        MapsArmCancelKey
        On Error GoTo Cleanup
    End If

    r = MapsGeocodeCore(p, addressText, outLat, outLng, outName, _
                        timeoutSec, pickFirstIfAmbiguous)

Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled にしたので通常は来ない。arm する前後の窓のために残す。
        '   判定できない値は書き残さない (設計原則111)。
        r = False
        outLat = 0
        outLng = 0
        outName = ""
        m_cancel = True
        m_canceled = True
        m_lastError = "canceled"
        Wv2Log.LogW "Wv2Maps.MapsGeocode: ★Esc で中断された★ (" & addressText & ")"
        errNo = 0
    End If

    If solo Then MapsDisarmCancelKey
    MapsGeocode = r

    ' ★他のエラーは握り潰さない★ 後始末だけしてそのまま上げる
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function


' ============================================================
' MapsGeocodeCore (D-7b、Private) - MapsGeocode の中身
'   ★出口が多いので、EnableCancelKey の後始末はラッパに任せる★
' ============================================================
Private Function MapsGeocodeCore(ByVal p As Wv2Pane, _
                                 ByVal addressText As String, _
                                 ByRef outLat As Double, _
                                 ByRef outLng As Double, _
                                 ByRef outName As String, _
                                 ByVal timeoutSec As Single, _
                                 ByVal pickFirstIfAmbiguous As Boolean) As Boolean
    Dim el As Wv2Element
    Dim prevUrl As String
    Dim newUrl As String
    Dim cands As Collection

    m_lastError = ""
    m_lastPicked = False
    outLat = 0
    outLng = 0
    outName = ""

    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    prevUrl = p.View_GetSource()

    ' --- 住所を書き込む (D-3。★ネイティブ setter 経由なので SPA に伝わる★) ---
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, 10)
    If MapsCheckCancel() Then GoTo Canceled     ' D-7
    If el Is Nothing Then
        m_lastError = "no-searchbox"
        Exit Function
    End If

    el.value = addressText
    If Not el.LastOk Then
        m_lastError = "write-failed: " & el.LastError
        Exit Function
    End If

    ' --- 検索を実行する ---
    Set el = MapsFindFirst(p, MAPS_GO_SELECTORS, 3)
    If el Is Nothing Then
        ' ボタンが無ければ Enter を合成する
        p.EvalSync "(function(){var e=document.querySelector('input[name=""q""]');" & _
                   "if(!e){return 0;}e.focus();" & _
                   "e.dispatchEvent(new KeyboardEvent('keydown'," & _
                   "{key:'Enter',code:'Enter',keyCode:13,which:13,bubbles:true}));" & _
                   "return 1;})()"
        If Not p.LastEvalOk Then
            m_lastError = "no-search-trigger"
            Exit Function
        End If
    Else
        If Not el.Click() Then
            m_lastError = "click-failed: " & el.LastError
            Exit Function
        End If
    End If

    ' --- ★場所が確定するまで待つ★ (arm を使わない理由はヘッダー参照) ---
    If Not MapsWaitPlace(p, prevUrl, timeoutSec, newUrl) Then
        ' --- 確定しなかった ---
        '   ★候補一覧が出ているかを直接数える★ (D-6 の QuerySelectorAll)
        '     0 件      … 何も見つからなかった。座標は返さない
        '     1 件以上  … 候補が複数。中心座標には意味がある
        '   ★以前は「検索して地図が動いたか」で見ていたが、同じ検索語を続けて
        '     引くと地図が動かず、候補が出ているのに not-found になった★
        '     (D-6b の実機検証で発覚)。間接的な推測をやめて直接数える。

        ' ★中断は「見つからなかった」より先に見る★ (D-7)
        If MapsCheckCancel() Then GoTo Canceled

        If InStr(1, newUrl, "/@") = 0 Then
            m_lastError = "timeout-url"
            Exit Function
        End If

        Set cands = MapsFindCandidates(p)
        Wv2Log.LogI "Wv2Maps.MapsGeocode: 候補リンク " & cands.Count & " 件"

        If cands.Count = 0 Then
            m_lastError = "not-found"
            Wv2Log.LogW "Wv2Maps.MapsGeocode: 見つからない (" & addressText & _
                        ") 候補一覧も出ていないので座標は返さない"
            Exit Function
        End If

        If pickFirstIfAmbiguous Then
            If MapsPickFirst(p, cands, timeoutSec, newUrl) Then
                m_lastPicked = True
                m_lastError = ""
                Wv2Log.LogI "Wv2Maps.MapsGeocode: ★候補の 1 件目を採った★ (" & _
                            addressText & ")"
                GoTo Resolved
            End If
            Wv2Log.LogW "Wv2Maps.MapsGeocode: 候補を選べなかった (" & addressText & ")"
            If MapsCheckCancel() Then GoTo Canceled     ' D-7
        End If

        MapsExtractLatLng newUrl, outLat, outLng
        outName = MapsTitleOf(p)
        m_lastError = "ambiguous"
        Wv2Log.LogW "Wv2Maps.MapsGeocode: 候補が複数 (" & addressText & _
                    ") 中心座標を返す"
        Exit Function
    End If

' ★確定した★ ここから先は「最初から確定していた場合」と
' 「候補の 1 件目に入った場合」で同じ処理をする。
Resolved:

    ' --- DOM が落ち着くのを待つ (D-4。居座る要求は既定で足切りされる) ---
    p.WaitSettled 10
    Wv2Log.LogD "Wv2Maps.MapsGeocode: " & p.LastSettleInfo

    outName = MapsTitleOf(p)

    If Not MapsExtractLatLng(newUrl, outLat, outLng) Then
        m_lastError = "no-coords"
        Exit Function
    End If

    MapsGeocodeCore = True
    Exit Function

' --- D-7: ★中断された★ ---
'   ★判定できない値は書き残さない★ (設計原則111)。座標も名前も空のまま返す。
Canceled:
    outLat = 0
    outLng = 0
    outName = ""
    m_lastError = "canceled"
    m_canceled = True
    Wv2Log.LogW "Wv2Maps.MapsGeocodeCore: ★中断された★ (" & addressText & ")"
End Function


' ============================================================
' MapsTitleOf (Private) - タイトルから ' - Google マップ' を落とす
'   ★正規化後の住所が手に入る★ 表記ゆれの整形にも使える。
' ============================================================
Private Function MapsTitleOf(ByVal p As Wv2Pane) As String
    Dim t As String

    t = p.View_GetDocumentTitle()
    If InStr(1, t, " - ") > 0 Then
        t = Left$(t, InStrRev(t, " - ") - 1)
    End If

    MapsTitleOf = t
End Function


' ============================================================
' MapsLastPicked (D-6b) - 直前の結果が候補一覧から選んだものか
'
'   ★True なら「その住所そのもの」ではなく「候補の 1 件目」★
'   確定 (True + MapsLastPicked = False) と精度が違うので、記録に残すときは
'   区別できるようにすること。MapsGeocodeSheet は状態欄に ok(候補1) と書く。
' ============================================================
Public Property Get MapsLastPicked() As Boolean
    MapsLastPicked = m_lastPicked
End Property


' ============================================================
' MapsFindCandidates (D-6b、Private) - 候補一覧のリンクを掴む
'
'   ★候補が「出ているかどうか」を直接数えるための口★
'   0 件なら「何も見つからなかった」、1 件以上なら「候補が複数」。
'   セレクタは候補方式 (設計原則107)。href で拾うものを先に置く。
' ============================================================
Private Function MapsFindCandidates(ByVal p As Wv2Pane) As Collection
    Dim sels As Variant
    Dim i As Long
    Dim els As Collection

    Set MapsFindCandidates = New Collection
    sels = Split(MAPS_CAND_SELECTORS, vbLf)

    For i = LBound(sels) To UBound(sels)
        Set els = p.QuerySelectorAll(CStr(sels(i)), 20)
        If els.Count > 0 Then
            Wv2Log.LogD "Wv2Maps.MapsFindCandidates: " & els.Count & _
                        " 件 [" & sels(i) & "]"
            Set MapsFindCandidates = els
            Exit Function
        End If
    Next i
End Function


' ============================================================
' MapsPickFirst (D-6b、Private) - 候補一覧の 1 件目に入る
'
'   ★D-6 の QuerySelectorAll がここで効く★ 候補のリンクをまとめて掴み、
'   文書順の 1 件目をクリックして、場所が確定するまで待つ。
'
'   戻り値: True = 確定した (outUrl に確定後の URL)
' ============================================================
Private Function MapsPickFirst(ByVal p As Wv2Pane, _
                               ByVal cands As Collection, _
                               ByVal timeoutSec As Single, _
                               ByRef outUrl As String) As Boolean
    Dim el As Wv2Element
    Dim prevUrl As String

    If cands Is Nothing Then Exit Function
    If cands.Count = 0 Then Exit Function

    prevUrl = p.View_GetSource()
    Set el = cands(1)   ' ★文書順の 1 件目★ (D-6 で順序が保たれることを確認済み)

    If Not el.Click() Then
        Wv2Log.LogW "Wv2Maps.MapsPickFirst: クリックできない err=" & el.LastError
        Exit Function
    End If

    MapsPickFirst = MapsWaitPlace(p, prevUrl, timeoutSec, outUrl)
End Function


' ============================================================
' MapsLastError - 直前の失敗理由
' ============================================================
Public Property Get MapsLastError() As String
    MapsLastError = m_lastError
End Property


' ============================================================
' MapsWaitPlace (Private) - ★場所が確定するまで★ 待つ
'
'   ★「URL が変わったら完了」では早すぎる★ (D-5 の実機検証で踏んだ)
'   Maps は検索を押すとまず /maps/search/<検索語>/@<前の中心> に変わり、
'   そのあとで /maps/place/<名前>/@<新座標>/data=!3d<緯度>!4d<経度> に落ち着く。
'   最初の変化で掴むと、★名前は新しいのに座標は前のまま★ という一件ずれになる。
'   実際 3 件続けて検索したら 2 件目と 3 件目の座標が 1 つ前のものになった。
'
'   したがって待つのは ★/place/ と !3d の両方を含む URL★。
'   !3d<緯度>!4d<経度> は場所が確定して初めて付く。
'
'   戻り値: True = 確定した (outUrl にその URL)
'           False = 時間切れ (outUrl には現在の URL。呼び出し側が候補複数か判断する)
'
'   ★同じ住所を 2 回続けて検索した場合★ URL が変わらないので時間切れになるが、
'   その時点で既に /place/ + !3d なら確定として扱う (前回の結果と同じで正しい)。
' ============================================================
Private Function MapsWaitPlace(ByVal p As Wv2Pane, _
                               ByVal prevUrl As String, _
                               ByVal timeoutSec As Single, _
                               ByRef outUrl As String) As Boolean
    Dim t0 As Single
    Dim u As String

    t0 = Timer
    Do
        u = p.View_GetSource()

        ' D-7: 中断されたら待たない。False で返して呼び出し側に判断させる。
        If MapsCheckCancel() Then
            outUrl = u
            Exit Function
        End If

        If u <> prevUrl And MapsIsPlaceUrl(u) Then
            outUrl = u
            MapsWaitPlace = True
            Exit Function
        End If

        If MapsElapsed(t0) >= timeoutSec Then
            outUrl = u
            ' 同じ住所を引き直した場合はここに来る。確定形ならそれでよい。
            MapsWaitPlace = MapsIsPlaceUrl(u)
            Exit Function
        End If

        MapsPump 0.2
    Loop
End Function

Private Function MapsIsPlaceUrl(ByVal u As String) As Boolean
    If InStr(1, u, "/place/") = 0 Then Exit Function
    If InStr(1, u, "!3d") = 0 Then Exit Function
    MapsIsPlaceUrl = True
End Function

' ============================================================
' MapsExtractLatLng (Private) - URL から座標を抜く
'
'   第一候補: !3d<緯度>!4d<経度>   ★場所そのものの座標★
'   第二候補: /@<緯度>,<経度>      地図の中心
'
'   ★Val を使う★ CDbl はロケール依存。Val は常に . を小数点として読み、
'   数字でない文字 (! や ,) で止まる。
' ============================================================
Private Function MapsExtractLatLng(ByVal url As String, _
                                   ByRef outLat As Double, _
                                   ByRef outLng As Double) As Boolean
    Dim i As Long
    Dim j As Long

    i = InStr(1, url, "!3d")
    j = InStr(1, url, "!4d")
    If i > 0 And j > 0 Then
        outLat = Val(Mid$(url, i + 3))
        outLng = Val(Mid$(url, j + 3))
        If MapsLatLngOk(outLat, outLng) Then
            MapsExtractLatLng = True
            Exit Function
        End If
    End If

    i = InStr(1, url, "/@")
    If i > 0 Then
        outLat = Val(Mid$(url, i + 2))
        j = InStr(i, url, ",")
        If j > 0 Then outLng = Val(Mid$(url, j + 1))
        If MapsLatLngOk(outLat, outLng) Then
            MapsExtractLatLng = True
            Exit Function
        End If
    End If

    outLat = 0
    outLng = 0
End Function

Private Function MapsLatLngOk(ByVal la As Double, ByVal lo As Double) As Boolean
    If la = 0 And lo = 0 Then Exit Function
    If la < -90 Or la > 90 Then Exit Function
    If lo < -180 Or lo > 180 Then Exit Function
    MapsLatLngOk = True
End Function


' ============================================================
' MapsFindFirst (Private) - 候補セレクタを順に試す
'   ★実サイトのセレクタは推測で書かない (設計原則107)★
'   当たったものをログに残すので、次に壊れたとき何が変わったか分かる。
' ============================================================
Private Function MapsFindFirst(ByVal p As Wv2Pane, _
                               ByVal selectorList As String, _
                               ByVal timeoutSec As Single) As Wv2Element
    Dim cands As Variant
    Dim i As Long
    Dim el As Wv2Element

    cands = Split(selectorList, vbLf)

    For i = LBound(cands) To UBound(cands)
        If i = LBound(cands) Then
            Set el = p.WaitFor(CStr(cands(i)), timeoutSec)
        Else
            Set el = p.QuerySelector(CStr(cands(i)))
        End If

        If Not el Is Nothing Then
            Set MapsFindFirst = el
            Exit Function
        End If
    Next i

    Wv2Log.LogW "Wv2Maps.MapsFindFirst: どの候補も外れた [" & _
                Replace$(selectorList, vbLf, " / ") & "]"
End Function


' ============================================================
' MapsPump / MapsElapsed (Private)
'   ★Timer は深夜 0 時に 0 へ戻る (仕様事実53)★ ので経過は補正して測る。
' ============================================================
Private Sub MapsPump(ByVal waitSec As Single)
    Dim t0 As Single

    t0 = Timer
    Do
        DoEvents
        ' ★D-7: 中断の検知はここ 1 箇所に置く★ すべての待ちがここを通るので、
        '   待ちループごとに書き足さなくてよい。抜けたあとの判断は呼び出し側。
        If MapsCheckCancel() Then Exit Do
        If MapsElapsed(t0) >= waitSec Then Exit Do
    Loop
End Sub

Private Function MapsElapsed(ByVal sinceTimer As Single) As Single
    Dim d As Single

    d = Timer - sinceTimer
    If d < 0 Then d = d + 86400
    MapsElapsed = d
End Function

' ============================================================
' MapsCancel (D-7) - 中断を要求する / 今の要求を読む
'
'   ★Esc でも止まるが、外から立てる口も残す★ (イミディエイトから止めたいとき、
'   将来フォームに中止ボタンを付けるとき)。DoEvents が回っているので届く。
'     Wv2Maps.MapsCancel = True
'
'   ★立てっぱなしにしても次回の実行は殺さない★ 次の MapsOpen / MapsGeocode /
'   MapsGeocodeSheet の入口で捨てられる (設計原則106: 前の状態を残さない)。
' ============================================================
Public Property Get MapsCancel() As Boolean
    MapsCancel = m_cancel
End Property

Public Property Let MapsCancel(ByVal newValue As Boolean)
    m_cancel = newValue
End Property


' ============================================================
' MapsCanceled (D-7) - 直前の呼び出しが中断で終わったか
'
'   ★戻り値の意味は変えない★ MapsGeocodeSheet は今までどおり
'   「ok になった行数」を返す。中断したかどうかはこちらで見る
'   (MapsLastError も "canceled" になる)。
' ============================================================
Public Property Get MapsCanceled() As Boolean
    MapsCanceled = m_canceled
End Property


' ============================================================
' MapsCountRows (D-7) - 処理する行数を先に数える
'
'   ★分母が要る★ 本体は「空セルが出たら終わり」なので、同じ規則で 1 回走査する。
'   セルを読むだけなので速い。0 件ならタブを開かずに帰れる。
'   呼び出し側が自前で進捗を出したいときのために Public にしてある。
' ============================================================
Public Function MapsCountRows(ByVal targetSheet As Object, _
                              Optional ByVal firstRow As Long = 2, _
                              Optional ByVal addrCol As Long = 1) As Long
    Dim r As Long
    Dim n As Long

    If targetSheet Is Nothing Then Exit Function

    r = firstRow
    Do
        If Len(Trim$(CStr(targetSheet.Cells(r, addrCol).value))) = 0 Then Exit Do
        n = n + 1
        r = r + 1
    Loop

    MapsCountRows = n
End Function


' ============================================================
' MapsCheckCancel / MapsResetCancel (D-7、Private)
'
'   ★bit15 (今押されている) と bit0 (前回以降に押された) の OR で見る★
'   押しっぱなし判定だけだと、WaitSettled のように MapsPump が回らない数秒の
'   空白で取りこぼす。bit0 があれば、その空白の間に押した 1 回も拾える。
'   GetAsyncKeyState はどちらのビットも返すので、0 でなければ中断とみなす。
'
'   ★開始時に 1 回読み捨てる★ さもないと直前の押下 (検証を打つ前に押した Esc
'   など) を持ち越して、始まった瞬間に中断する。
'
'   ★副作用★ Esc はブラウザ操作でも押されるキー。バッチ実行中に押せば止まる。
'   無人で回す前提なので実害は小さいと判断した (D-7 論点3)。
' ============================================================
Private Function MapsCheckCancel() As Boolean
    Dim k As Integer

    m_checkCount = m_checkCount + 1   ' D-7d: 実測用

    If m_cancel Then
        MapsCheckCancel = True
        Exit Function
    End If

    k = GetAsyncKeyState(MAPS_VK_ESCAPE)
    If k <> 0 Then
        m_cancel = True
        MapsCheckCancel = True
        Wv2Log.LogI "Wv2Maps: ★Esc で中断が要求された★"
    End If
End Function

Private Sub MapsResetCancel()
    Dim k As Integer

    m_cancel = False
    m_canceled = False
    m_checkCount = 0
    k = GetAsyncKeyState(MAPS_VK_ESCAPE)   ' ★直前の押下を読み捨てる★
End Sub


' ============================================================
' MapsArmCancelKey / MapsDisarmCancelKey (D-7b、Private)
'
'   ★Wv2Maps に入っている間だけ VBA ランタイムの Esc を止める★
'   これをしないと、VBA が先に Esc を掴んで break する (D-7 の実機で確認)。
'   ★break は仕様事実20 のクラッシュ窓★ でもあるので、塞ぐ価値は二重にある。
'
'   ★なぜ xlErrorHandler ではなく xlDisabled なのか (D-7c)★
'   xlErrorHandler だと Esc はエラー18 になるが、★発生場所が主に
'   WebView2 の COM コールバックの中★ で、そこは呼び出し元が VBA ではないので
'   こちらの On Error GoTo に遡らない (D-7b の実機で確認)。
'   xlDisabled ならエラー自体が起きず、実行が止まらないので、
'   ★次に MapsPump が回ったときに GetAsyncKeyState が拾える★。
'
'   ★必ず戻す★ 戻し忘れると Excel 全体で Esc も Ctrl+Break も効かなくなる。
'   入れ子で arm すると保存値が xlDisabled になってしまうが、
'   m_inBatch で「一番外側だけが arm する」ことを保証しているので起きない。
' ============================================================
Private Sub MapsArmCancelKey()
    m_savedCancelKey = Application.EnableCancelKey
    Application.EnableCancelKey = xlDisabled
End Sub

Private Sub MapsDisarmCancelKey()
    ' 保存値が xlDisabled (0) でも xlInterrupt に戻す ― 砦を戻す方が安全
    If m_savedCancelKey = 0 Then m_savedCancelKey = xlInterrupt
    Application.EnableCancelKey = m_savedCancelKey
    m_savedCancelKey = 0
End Sub


' ============================================================
' MapsCheckCount (D-7d) - 直前のバッチで中断検知を何回呼べたか
'
'   ★1 件 15 秒に対してこれが数十回しかないなら、待ちの大半は Wv2Pane 側の
'   ループ (EvalSync / WaitFor / WaitSettled) で、そこに検知点が無いということ。★
' ============================================================
Public Property Get MapsCheckCount() As Long
    MapsCheckCount = m_checkCount
End Property


' ============================================================
' MapsRawEscState (D-7d) - GetAsyncKeyState の生の値
'
'   ★宣言を 1 箇所に保つための観測口★ 検証側で同じ Declare を書くと
'   宣言が 2 つになる。プローブ (Test_D7_Key) はここを叩く。
'   戻り値: 0 = 押されていない / それ以外は生のビット
'           &H8000 相当 (負の値) = 今押されている / 1 = 前回以降に押された
' ============================================================
Public Function MapsRawEscState() As Long
    MapsRawEscState = GetAsyncKeyState(MAPS_VK_ESCAPE)
End Function


' ============================================================
' MapsProgress (D-7、Private) - 進捗をステータスバーに出す
'
'   ★シートは処理中に埋まっていくが、分母と残りが見えない★ ので、
'   件数と中止の手段をステータスバーに出す。終わったら必ず False で戻す。
' ============================================================
Private Sub MapsProgress(ByVal doneRows As Long, ByVal totalRows As Long, _
                         ByVal okRows As Long, ByVal ngRows As Long)
    Application.StatusBar = "住所→座標 " & doneRows & "/" & totalRows & _
                           " (ok " & okRows & " / 失敗 " & ngRows & ")" & _
                           "  ―  中止は Esc"
End Sub


' ============================================================
' MapsGeocodeSheet (D-5b) - シートの住所列をまとめて座標にする
'
'   引数:
'     targetSheet : 対象のワークシート
'     firstRow    : 開始行 (既定 2。1 行目は見出しの想定)
'     addrCol     : 住所の列番号 (既定 1 = A 列)
'     outCol      : 書き出す先頭列 (既定 2 = B 列)
'     skipDone    : 状態が ok で始まる行を飛ばすか (既定 True)
'     pickFirst   : 候補が複数のとき 1 件目を採るか (既定 False)
'
'   書き出す並び (outCol から 4 列):
'     outCol   : 緯度
'     outCol+1 : 経度
'     outCol+2 : 正規化後の住所 (Maps が返した名前)
'     outCol+3 : 状態 (ok / ★ok(候補1)★ / ambiguous / not-found / timeout-url ...)
'                ★ok(候補1) は「候補一覧の 1 件目を採った」★ 確定より精度が落ちる
'
'   戻り値: ★ok になった行数★ (処理した行数ではない)
'
'   ■ ★候補が複数 (ambiguous) のときの扱い★
'     **中心座標を書いたうえで状態欄に ambiguous と記す。**
'     黙って飛ばすと「処理したのに空」の理由が分からず、黙って確定扱いにすると
'     ★間違った座標が混ざる★。データは残して人が判断できる形にした。
'
'   ■ ★途中で止めても続きから再開できる★
'     skipDone = True なら状態が ok の行を飛ばす。100 件の途中で止めても
'     もう一度呼べば残りだけを処理する。やり直したいときは状態列を消す。
'
'   ■ 進捗と中断 (D-7)
'     1 行ごとに書き込んで DoEvents を回すので、★処理中の画面で埋まっていく★。
'     ログにも 1 行ずつ残る。加えて★ステータスバーに 12/100 の形で出す★。
'     ★止めたくなったら Esc を押す★ (押しっぱなしでなくてよい)。
'     中断は★行の境界で効く★ので、書きかけの行は残らない。
'     中断したかどうかは MapsCanceled で分かる (MapsLastError も canceled)。
'     ★中断した行には何も書かないので、skipDone のまま呼び直せば続きから進む★。
'
'   使用例:
'     Wv2Maps.MapsGeocodeSheet ActiveSheet
'     Wv2Maps.MapsGeocodeSheet Sheets("住所録"), 3, 2, 5   ' 3 行目から、B 列の住所を E 列以降へ
' ============================================================
Public Function MapsGeocodeSheet(ByVal targetSheet As Object, _
                                 Optional ByVal firstRow As Long = 2, _
                                 Optional ByVal addrCol As Long = 1, _
                                 Optional ByVal outCol As Long = 2, _
                                 Optional ByVal skipDone As Boolean = True, _
                                 Optional ByVal pickFirst As Boolean = False) As Long
    Dim p As Wv2Pane
    Dim r As Long
    Dim addr As String
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim doneCount As Long
    Dim total As Long
    Dim totalRows As Long
    Dim errNo As Long
    Dim errDesc As String
    Dim errSrc As String
    Dim tStart As Single

    m_lastError = ""
    tStart = Timer

    If targetSheet Is Nothing Then
        m_lastError = "no-sheet"
        Exit Function
    End If

    ' --- D-7: ★一番外側なのでここで中断状態を捨てる★ ---
    MapsResetCancel

    ' --- D-7: ★分母を先に数える★ 0 件ならタブを開かずに帰る ---
    totalRows = MapsCountRows(targetSheet, firstRow, addrCol)
    If totalRows = 0 Then
        m_lastError = "no-rows"
        Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: 処理する行が無い"
        Exit Function
    End If
    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: ★全 " & totalRows & " 件★ (中止は Esc)"

    ' --- D-7b: ★ここから先はすべて Cleanup を通る★ ---
    '   m_inBatch は MapsOpen より前に立てる (MapsOpen 側で arm させないため)。
    m_inBatch = True
    MapsArmCancelKey
    On Error GoTo Cleanup

    Set p = MapsOpen(UserForm1.CurrentBrowser)
    If p Is Nothing Then
        Wv2Log.LogE "Wv2Maps.MapsGeocodeSheet: Maps を開けない (" & m_lastError & ")"
        GoTo Cleanup
    End If

    r = firstRow
    Do
        ' --- D-7: ★中断は行の境界で効かせる★ 書きかけの行を残さない ---
        If MapsCheckCancel() Then
            m_canceled = True
            Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: ★中断された★ (" & _
                        r & " 行目の手前まで)"
            Exit Do
        End If

        addr = Trim$(CStr(targetSheet.Cells(r, addrCol).value))
        If Len(addr) = 0 Then Exit Do

        total = total + 1
        MapsProgress total, totalRows, doneCount, (total - 1) - doneCount

        If skipDone And Left$(CStr(targetSheet.Cells(r, outCol + 3).value), 2) = "ok" Then
            Wv2Log.LogI "  [" & r & "] 済みなので飛ばす: " & addr
            doneCount = doneCount + 1
        Else
            ok = MapsGeocode(p, addr, lat, lng, nm, 20, pickFirst)

            ' --- D-7: ★中断された行には何も書かない★ (設計原則111) ---
            '   状態欄に canceled と書くと、後から見て「そういう結果だった」
            '   ように読める。処理していないのだから空のままにする。
            If m_canceled Then
                total = total - 1
                Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: ★中断された★ (" & _
                            r & " 行目の処理中)"
                Exit Do
            End If

            If ok Then
                targetSheet.Cells(r, outCol).value = lat
                targetSheet.Cells(r, outCol + 1).value = lng
                targetSheet.Cells(r, outCol + 2).value = nm
                ' ★確定と「候補から選んだ」を書き分ける★ (精度が違うものを同じ顔で
                '   並べない)。再開時の判定は Left$(..., 2) = "ok" で見る。
                If MapsLastPicked Then
                    targetSheet.Cells(r, outCol + 3).value = "ok(候補1)"
                Else
                    targetSheet.Cells(r, outCol + 3).value = "ok"
                End If
                doneCount = doneCount + 1
            Else
                ' ★候補が複数のときも座標 (地図の中心) は書き残す★
                If lat <> 0 Or lng <> 0 Then
                    targetSheet.Cells(r, outCol).value = lat
                    targetSheet.Cells(r, outCol + 1).value = lng
                End If
                If Len(nm) > 0 Then targetSheet.Cells(r, outCol + 2).value = nm
                targetSheet.Cells(r, outCol + 3).value = MapsLastError
            End If

            Wv2Log.LogI "  [" & r & "] " & addr & " → " & _
                        IIf(ok, lat & "," & lng, "★" & MapsLastError & "★")
        End If

        DoEvents
        r = r + 1
    Loop

' --- D-7b: ★ここが唯一の出口★ 正常終了もエラーもここを通る ---
Cleanup:
    errNo = Err.Number
    errDesc = Err.Description
    errSrc = Err.source
    On Error GoTo 0

    If errNo = 18 Then
        ' D-7c: xlDisabled にしたので通常は来ない。arm する前後の窓のために残す。
        m_cancel = True
        m_canceled = True
        Wv2Log.LogW "Wv2Maps.MapsGeocodeSheet: ★Esc で中断された★ (" & _
                    r & " 行目の処理中)"
        errNo = 0
    End If

    ' --- ★必ず戻す★ ステータスバーも EnableCancelKey も残すと Excel 全体が汚れる ---
    MapsDisarmCancelKey
    m_inBatch = False

    ' ★D-7d: False ではなく Empty★ False だと "FALSE" と表示されたまま残る
    Application.StatusBar = Empty
    If m_canceled Then m_lastError = "canceled"

    Wv2Log.LogD "Wv2Maps.MapsGeocodeSheet: StatusBar 復帰後 = [" & _
                CStr(Application.StatusBar) & "] 型=" & _
                TypeName(Application.StatusBar)

    ' ★D-7d: 中断検知を何回呼べたか★ 拾えない原因の切り分け材料 (設計原則112)
    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: ★中断検知 " & m_checkCount & " 回 / " & _
                Format$(MapsElapsed(tStart), "0.0") & " 秒★"

    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: " & total & " 行中 " & _
                doneCount & " 行が ok" & IIf(m_canceled, " ★中断した★", "")
    MapsGeocodeSheet = doneCount

    ' ★他のエラーは握り潰さない★
    If errNo <> 0 Then Err.Raise errNo, errSrc, errDesc
End Function

