Attribute VB_Name = "Wv2Maps"
''''''''''''''''''''''''''''''''''
' --- Wv2Maps.bas  D-5 段階 (実用: 住所 → 座標) ---
'
'   ★D 軸で作った部品を、業務でそのまま使える形に畳んだ最初の例。★
'   D-4d で Google Maps の住所検索が実サイトで通ることを実証したので、
'   その手順を 1 本の関数にした。
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
'                                  このとき outLat / outLng には地図の中心が入る
'       no-coords                … URL から座標を読めなかった
''''''''''''''''''''''''''''''''''
Option Explicit


' --- 直前の失敗理由 ---
Private m_lastError As String

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
    Dim el As Wv2Element
    Dim t0 As Single

    m_lastError = ""

    If b Is Nothing Then
        m_lastError = "no-browser"
        Exit Function
    End If

    Set p = b.AddTabWithUrl("https://www.google.com/maps")
    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ' --- JS が通るようになるまで待つ ---
    t0 = Timer
    Do
        p.EvalSync "1", 3
        If p.LastEvalOk Then Exit Do
        If MapsElapsed(t0) >= timeoutSec Then
            m_lastError = "pane-not-ready"
            Exit Function
        End If
        MapsPump 0.3
    Loop

    ' --- 検索ボックスが出るまで待つ ---
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, timeoutSec)
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
    Do
        If InStr(1, p.View_GetSource(), "/@") > 0 Then Exit Do
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
    p.AddIgnoreNetwork "/search?tbm=map"

    Wv2Log.LogI "Wv2Maps.MapsOpen: 準備できた (" & p.View_GetSource() & ")"
    Set MapsOpen = p
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
                            Optional ByVal timeoutSec As Single = 20) As Boolean
    Dim el As Wv2Element
    Dim prevUrl As String
    Dim newUrl As String
    Dim prevLat As Double
    Dim prevLng As Double
    Dim tmpLat As Double
    Dim tmpLng As Double

    m_lastError = ""
    outLat = 0
    outLng = 0
    outName = ""

    If p Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    prevUrl = p.View_GetSource()
    MapsExtractLatLng prevUrl, prevLat, prevLng

    ' --- 住所を書き込む (D-3。★ネイティブ setter 経由なので SPA に伝わる★) ---
    Set el = MapsFindFirst(p, MAPS_BOX_SELECTORS, 10)
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
        ' --- 確定しなかった。★地図が動いたかで意味が変わる★ ---
        '   動いた   … 候補が複数 (広域が指定された等)。中心座標には意味がある
        '   動かない … 何も見つからなかった。★中心座標は前の検索の残りかす★ なので
        '              書き残してはいけない (D-5b の実機検証で踏んだ。でたらめな住所に
        '              1 つ前の行の座標が付いた)
        If InStr(1, newUrl, "/@") = 0 Then
            m_lastError = "timeout-url"
            Exit Function
        End If

        MapsExtractLatLng newUrl, tmpLat, tmpLng

        If prevLat = 0 And prevLng = 0 Then
            ' ★比較の基準が無い★ 動いたのか元からそこだったのか判定できないので、
            '   保守的に座標を返さない。★怪しい座標を書き残す方が害が大きい★
            m_lastError = "not-found"
            Wv2Log.LogW "Wv2Maps.MapsGeocode: 確定せず、比較の基準も無い (" & _
                        addressText & ") 座標は返さない"

        ElseIf MapsSameSpot(tmpLat, tmpLng, prevLat, prevLng) Then
            m_lastError = "not-found"
            Wv2Log.LogW "Wv2Maps.MapsGeocode: 見つからない (" & addressText & _
                        ") 地図が動いていないので座標は返さない"

        Else
            outLat = tmpLat
            outLng = tmpLng
            outName = MapsTitleOf(p)
            m_lastError = "ambiguous"
            Wv2Log.LogW "Wv2Maps.MapsGeocode: 候補が複数 (" & addressText & _
                        ") 中心座標を返す url=" & Left$(newUrl, 70)
        End If
        Exit Function
    End If

    ' --- DOM が落ち着くのを待つ (D-4。居座る要求は既定で足切りされる) ---
    p.WaitSettled 10
    Wv2Log.LogD "Wv2Maps.MapsGeocode: " & p.LastSettleInfo

    outName = MapsTitleOf(p)

    If Not MapsExtractLatLng(newUrl, outLat, outLng) Then
        m_lastError = "no-coords"
        Exit Function
    End If

    MapsGeocode = True
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

' ============================================================
' MapsSameSpot (Private) - 2 つの座標が同じ地点か
'   ★検索して地図が動いたかの判定に使う★ 動いていなければ、その座標は
'   「今回の検索の結果」ではなく「前の検索の残りかす」である。
' ============================================================
Private Function MapsSameSpot(ByVal la1 As Double, ByVal lo1 As Double, _
                              ByVal la2 As Double, ByVal lo2 As Double) As Boolean
    If la2 = 0 And lo2 = 0 Then Exit Function   ' 比較相手が無い
    If Abs(la1 - la2) > 0.000001 Then Exit Function
    If Abs(lo1 - lo2) > 0.000001 Then Exit Function
    MapsSameSpot = True
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
' MapsGeocodeSheet (D-5b) - シートの住所列をまとめて座標にする
'
'   引数:
'     targetSheet : 対象のワークシート
'     firstRow    : 開始行 (既定 2。1 行目は見出しの想定)
'     addrCol     : 住所の列番号 (既定 1 = A 列)
'     outCol      : 書き出す先頭列 (既定 2 = B 列)
'     skipDone    : 状態が ok の行を飛ばすか (既定 True)
'
'   書き出す並び (outCol から 4 列):
'     outCol   : 緯度
'     outCol+1 : 経度
'     outCol+2 : 正規化後の住所 (Maps が返した名前)
'     outCol+3 : 状態 (ok / ambiguous / no-searchbox / timeout-url ...)
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
'   ■ 進捗
'     1 行ごとに書き込んで DoEvents を回すので、★処理中の画面で埋まっていく★。
'     ログにも 1 行ずつ残る。
'
'   使用例:
'     Wv2Maps.MapsGeocodeSheet ActiveSheet
'     Wv2Maps.MapsGeocodeSheet Sheets("住所録"), 3, 2, 5   ' 3 行目から、B 列の住所を E 列以降へ
' ============================================================
Public Function MapsGeocodeSheet(ByVal targetSheet As Object, _
                                 Optional ByVal firstRow As Long = 2, _
                                 Optional ByVal addrCol As Long = 1, _
                                 Optional ByVal outCol As Long = 2, _
                                 Optional ByVal skipDone As Boolean = True) As Long
    Dim p As Wv2Pane
    Dim r As Long
    Dim addr As String
    Dim lat As Double
    Dim lng As Double
    Dim nm As String
    Dim ok As Boolean
    Dim doneCount As Long
    Dim total As Long

    m_lastError = ""

    If targetSheet Is Nothing Then
        m_lastError = "no-sheet"
        Exit Function
    End If

    Set p = MapsOpen(UserForm1.CurrentBrowser)
    If p Is Nothing Then
        Wv2Log.LogE "Wv2Maps.MapsGeocodeSheet: Maps を開けない (" & m_lastError & ")"
        Exit Function
    End If

    r = firstRow
    Do
        addr = Trim$(CStr(targetSheet.Cells(r, addrCol).value))
        If Len(addr) = 0 Then Exit Do

        total = total + 1

        If skipDone And CStr(targetSheet.Cells(r, outCol + 3).value) = "ok" Then
            Wv2Log.LogI "  [" & r & "] 済みなので飛ばす: " & addr
            doneCount = doneCount + 1
        Else
            ok = MapsGeocode(p, addr, lat, lng, nm)

            If ok Then
                targetSheet.Cells(r, outCol).value = lat
                targetSheet.Cells(r, outCol + 1).value = lng
                targetSheet.Cells(r, outCol + 2).value = nm
                targetSheet.Cells(r, outCol + 3).value = "ok"
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

    Wv2Log.LogI "Wv2Maps.MapsGeocodeSheet: " & total & " 行中 " & _
                doneCount & " 行が ok"
    MapsGeocodeSheet = doneCount
End Function
