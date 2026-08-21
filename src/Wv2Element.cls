VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2Element"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' --- Wv2Element.cls  D-3 段階 (書き込みと操作) ---
'
'   D-3 で増えた公開メンバは 4 つ:
'
'     Value = v                  Property Let  値を書き込む (★戻り値が無い★)
'     Click()             As Boolean           クリックする
'     SetAttribute(n, v)  As Boolean           属性を書き込む
'     LastInfo            As String            直前の操作が選んだ経路の名前
'
'   ■ 成否の見方 (D-2 の LastOk / LastError 規約をそのまま踏襲、論点8)
'     Click / SetAttribute … 戻り値 (Boolean) で足りる。理由が要るときは LastError。
'     Value =              … ★Property Let は戻り値を返せない★ ので、
'                            書き込みが通ったかは★必ず LastOk を見る★。
'
'       el.Value = "abc"
'       If Not el.LastOk Then Debug.Print el.LastError
'
'     LastError に入る値は D-2 と同じ (stale / no-pane / in-callback / timeout(n s) /
'     JS の例外メッセージ)。
'
'   ■ ★成功は「JS が例外なく走った」であって「意図した効果が出た」ではない★
'     disabled なボタンを Click しても例外は出ないので True が返る。readonly な
'     input への書き込みも同じ。効果が要るときは★読み戻して確かめる★こと。
'
'   ■ Value = がやっていること (論点6: ★SPA 対応の核心★)
'     素直に e.value = x とすると、React / Vue などのフレームワークは
'     値の変更に★気づかない★。フレームワークは要素インスタンスに独自の
'     value setter (value tracker) を被せていて、そこを通った代入は
'     「自分が知っている値」として記録されてしまうため。
'
'     そこで D-3 は次の順で撃つ:
'       (1) e.constructor.prototype から value の★ネイティブ setter★ を取り出し、
'           あればそれを e に対して呼ぶ (= tracker を迂回する)。LastInfo = 'setter'
'           取り出せない要素なら素の代入に落とす。            LastInfo = 'direct'
'       (2) input イベントを発火 (React 系がこれを見る)
'       (3) change イベントを発火 (jQuery 系がこれを見る)
'     ★両方撃つ★のが論点6 の合意。素の HTML では change が 1 回余分に飛ぶが、
'     実害よりも「反応しないページがある」事故の方が高くつくと判断した。
'
'   ■ Click の形 (論点7 案A)
'     e.click() があればそれを使う (LastInfo = 'click')。無ければ
'     MouseEvent('click') を bubbles つきで dispatch する (LastInfo = 'dispatch')。
'     ★mousedown / mouseup は撃たない★ ドラッグ系 UI が要るようになってから考える。
'     なお発火するイベントの isTrusted は false になる (JS から作った合成イベント
'     なので原理的にそうなる)。isTrusted を見て弾くページには効かない。
'
'   ■ 使用例
'     Dim p As Wv2Pane, el As Wv2Element
'     Set p = UserForm1.GetActivePane
'     Set el = p.QuerySelector("input[name='q']")
'     el.Value = "検索語"
'     If el.LastOk Then p.QuerySelector("button[type='submit']").Click
'
'   ★D-3 の時点では「変化し終わるのを待つ」手段はまだ無い★
'     Click の後に DOM が書き変わるのを待つ WaitFor / WaitGone は、
'     論点3 の前提を実測してから入れる (Wv2Tests.Test_D3_Probe_Promise)。
'''''''''''''''''''''''''''''''''''
' --- Wv2Element.cls  D-2 段階 (要素ハンドルの読み取り操作) ---
'
'   ★D 軸 (DOM / ページ自動制御) の中核クラス。D-2 で新規追加。★
'
'   ■ 役割
'     JS 側のレジストリ (window.__wv2.els) に登録された 1 個の DOM 要素を、
'     VBA 側から読み取るためのハンドル。実体は次の 3 つだけを持つ:
'
'       m_pane   … どの Wv2Pane (= どのタブ) の要素か
'       m_handle … JS 側 window.__wv2.els[] の添字
'       m_gen    … 取得したときのレジストリ世代 (JS が発行した文字列)
'
'     ★DOM の実体は保持しない。★ 値が要るたびに JS へ問い合わせる。
'     したがって 1 プロパティの読み取りにつき ExecuteScript が 1 往復する。
'     ループで大量に読むと遅いので、まとめて取りたい場合は
'     Wv2Pane.EvalSync で JS 式を 1 本書いた方が速い (D-2 の設計上の割り切り)。
'
'   ■ 生成方法
'     ★New で作らない。★ Wv2Pane.GetElementById / Wv2Pane.QuerySelector が返す。
'     直接 New しても m_pane が Nothing なので何も読めない。
'
'   ■ 公開メンバ (D-2 は★読み取りのみ★。書き込みと Click は D-3)
'     TagName            As String   要素のタグ名 (大文字。例: INPUT)
'     InnerText          As String   表示されるテキスト
'     InnerHTML          As String   内側の HTML (仕様事実30 の \u003C を復号済み)
'     Value              As String   input / textarea / select の値。無ければ空文字
'     GetAttribute(name) As String   属性値。属性が無ければ空文字
'     IsStale            As Boolean  ハンドルが無効になっていないか (往復 1 回)
'     Handle             As Long     JS 側の添字 (デバッグ用)
'     Generation         As String   取得時の世代 (デバッグ用)
'     LastOk             As Boolean  直前の読み取りが成功したか
'     LastError          As String   直前の失敗理由
'
'   ■ 「空文字が返った」の読み方 (★重要★)
'     読み取り系は失敗しても空文字を返すので、空文字の意味は LastOk で分ける:
'
'       LastOk = True   … ★正常★ 本当に空だった (テキストが無い / 属性が無い)
'       LastOk = False  … ★失敗★ 理由は LastError:
'                           stale        レジストリが作り直された (ページ遷移等)
'                           no-pane      Pane 参照が無い (New で作った場合)
'                           in-callback  イベントハンドラ内から呼ばれた (仕様事実14)
'                           timeout(n s) JS が返らなかった (仕様事実32)
'                           その他        EvalSync の LastEvalError がそのまま入る
'
'   ■ 世代 (Generation) の考え方 (D-2 論点2 案b)
'     世代は★JS が発行する★ (window.__wv2.gen)。ページ遷移で JS の実行環境ごと
'     消えるので、次のアクセスでレジストリが作り直され、必ず別の世代になる。
'     VBA 側でナビゲーションイベントを数える方式 (案a) は、SPA の pushState では
'     DOM が残るのにイベントは飛び、逆にページ内 JS が要素を差し替えても
'     イベントは飛ばないため、実態とズレる。よって★判定の権威は常に JS 側★に置く。
'
'     Wv2Pane が持つ世代キャッシュ (CurrentDomGen) は「往復する前に stale と
'     分かる場合だけ省く」高速パスにしか使わない。キャッシュが古い場合は
'     必ず JS に問い合わせる側に倒れるので、誤判定は起きない。
'
'   ■ 寿命
'     ★Class_Terminate では何もしない。★ JS 側スロットの解放はしない (D-2 論点3)。
'     Terminate がいつ・どの文脈で走るかを制御できないため、そこから EvalSync を
'     撃つと D-1 の in-callback ガードや Ctrl_Close 済みの Pane を踏みやすい。
'     レジストリはページ遷移でまるごと消えるので、実害は「同一ページ内で
'     大量に検索すると JS 側の配列が伸びる」だけ。手動で捨てたいときは
'     Wv2Pane.ClearElementRegistry を呼ぶ。本格的な寿命管理は D-4 で扱う。
''''''''''''''''''''''''''''''''''
Option Explicit


' --- 内部状態 ---
'   m_pane   : 所属する Wv2Pane への参照 (Wv2Pane 側は Wv2Element を保持しないので
'              循環参照にはならない)
'   m_handle : JS 側 window.__wv2.els[] の添字 (0 始まり)
'   m_gen    : 取得時のレジストリ世代 (JS が発行した文字列)
Private m_pane   As Wv2Pane
Private m_handle As Long
Private m_gen    As String

' --- 直前の読み取り結果 (D-1 の LastEvalOk / LastEvalError と同じ流儀) ---
Private m_lastOk    As Boolean
Private m_lastError As String

' --- 直前の★操作★が選んだ経路の名前 (D-3。'setter' / 'direct' / 'click' 等) ---
Private m_lastInfo  As String


Private Sub Class_Initialize()
    Set m_pane = Nothing
    m_handle = -1
    m_gen = ""
    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
End Sub


Private Sub Class_Terminate()
    ' ★意図的に何もしない★ (D-2 論点3 案a)。上のヘッダーコメントを参照。
    Set m_pane = Nothing
End Sub


' ============================================================
' InitElement (Friend、Wv2Pane からのみ呼ばれる)
'   Wv2Pane.GetElementById / QuerySelector が、JS から受け取った
'   ハンドルと世代を流し込むための唯一の入口。
' ============================================================
Friend Sub InitElement(ByVal ownerPane As Wv2Pane, _
                       ByVal elemHandle As Long, _
                       ByVal elemGen As String)
    Set m_pane = ownerPane
    m_handle = elemHandle
    m_gen = elemGen
    m_lastOk = True
    m_lastError = ""
    m_lastInfo = ""
End Sub


' ============================================================
' Handle / Generation (デバッグ用の素の値)
' ============================================================
Public Property Get Handle() As Long
    Handle = m_handle
End Property

Public Property Get Generation() As String
    Generation = m_gen
End Property


' ============================================================
' LastOk / LastError (直前の読み取りの成否)
'
'   ★空文字が返ったときの解釈に必ず使う★
'     LastOk = True  → 本当に空だった (正常)
'     LastOk = False → 失敗。LastError に理由が入る
' ============================================================
Public Property Get LastOk() As Boolean
    LastOk = m_lastOk
End Property

Public Property Get LastError() As String
    LastError = m_lastError
End Property


' ============================================================
' LastInfo (D-3) - 直前の★操作★が選んだ経路の名前
'
'   Value =      : 'setter' (ネイティブ setter 経由) / 'direct' (素の代入)
'   Click        : 'click'  (e.click())             / 'dispatch' (合成イベント)
'   SetAttribute : 'ok'
'
'   ★読み取り系 (TagName 等) を呼ぶと空文字に戻る★ 「直前の操作」の記録なので、
'   書き込みの経路を見たいときは★書いた直後に★読むこと。
'   診断用であって、成否の判定には使わない (それは LastOk / 戻り値の役目)。
' ============================================================
Public Property Get LastInfo() As String
    LastInfo = m_lastInfo
End Property


' ============================================================
' TagName / InnerText / InnerHTML / Value (D-2 の読み取り 4 種)
'
'   いずれも 1 回の呼び出しにつき ExecuteScript が 1 往復する。
'
'   ★InnerHTML は仕様事実30 (< が \u003C で届く) を必ず踏む★
'     Wv2Json.JsonUnescape が復号するので、VBA 側では素の HTML が得られる。
'
'   ★Value は input / textarea / select 以外では空文字★
'     それらの要素以外に value プロパティは無く、JS 側で undefined になるため。
'     このとき LastOk は True (エラーではない)。
' ============================================================
Public Property Get TagName() As String
    TagName = ReadProp("e.tagName")
End Property

Public Property Get InnerText() As String
    InnerText = ReadProp("e.innerText")
End Property

Public Property Get InnerHTML() As String
    InnerHTML = ReadProp("e.innerHTML")
End Property

Public Property Get value() As String
    value = ReadProp("e.value")
End Property


' ============================================================
' GetAttribute (D-2)
'   属性値を取り出す。属性が無ければ空文字 + LastOk = True。
'
'   使用例:
'     el.GetAttribute("href")
'     el.GetAttribute("data-id")
'
'   ★getAttribute と .value の違い★
'     <input value="初期値"> をユーザーが書き換えたあと、
'       .Value               → 現在の値 (書き換え後)
'       .GetAttribute("value") → 属性の値 (初期値のまま)
'     になる。用途で使い分けること。
' ============================================================
Public Function GetAttribute(ByVal attrName As String) As String
    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        GetAttribute = ""
        Exit Function
    End If
    GetAttribute = ReadProp("e.getAttribute(" & m_pane.JsQuote(attrName) & ")")
End Function


' ============================================================
' Value (D-3、Property Let) - 値を書き込む
'
'   ★Property Let なので戻り値が無い★ 成否は必ず LastOk を見ること (論点8)。
'   ★フレームワークに気づかせるための 3 手順★ (ヘッダー参照):
'     ネイティブ setter ← input イベント ← change イベント の順に撃つ
'
'   使用例:
'     el.Value = "検索語"
'     If Not el.LastOk Then Debug.Print "書き込み失敗: " & el.LastError
'
'   ★input / textarea / select 以外に撃つとどうなるか★
'     div などは value プロパティを持たないので、ネイティブ setter が取れず
'     素の代入に落ちる (LastInfo = 'direct')。JS 的には独自プロパティが 1 個
'     生えるだけで例外にはならないので、★LastOk は True になる★。
'     「効果が出たか」は読み戻して確かめること。
' ============================================================
Public Property Let value(ByVal newValue As String)
    Dim jsVal As String
    Dim js As String

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Property
    End If

    jsVal = m_pane.JsQuote(newValue)

    js = "var d=Object.getOwnPropertyDescriptor(e.constructor.prototype,'value');" & _
         "if(d&&d.set){d.set.call(e," & jsVal & ");rv='setter';}" & _
         "else{e.value=" & jsVal & ";rv='direct';}" & _
         "e.dispatchEvent(new Event('input',{bubbles:true}));" & _
         "e.dispatchEvent(new Event('change',{bubbles:true}));"

    ExecAction js, "Value"
End Property


' ============================================================
' Click (D-3) - クリックする
'
'   戻り値: True = JS 側で例外なく走った / False = 失敗 (理由は LastError)
'
'   e.click() があればそれを使い (LastInfo = 'click')、無ければ
'   MouseEvent('click') を bubbles + cancelable つきで dispatch する
'   (LastInfo = 'dispatch')。どちらも★合成イベントなので isTrusted は false★。
'
'   ★戻り値 True は「押した結果どうなったか」を保証しない★
'     disabled なボタンでも例外は出ないので True が返る。ページ遷移や DOM の
'     書き換えが起きるかは別問題で、それを待つ手段は D-3 の時点ではまだ無い。
'
'   使用例:
'     If Not el.Click() Then Debug.Print el.LastError
' ============================================================
Public Function Click() As Boolean
    Dim js As String

    js = "if(typeof e.click==='function'){e.click();rv='click';}" & _
         "else{e.dispatchEvent(new MouseEvent('click'," & _
         "{bubbles:true,cancelable:true,view:window}));rv='dispatch';}"

    Click = ExecAction(js, "Click")
End Function


' ============================================================
' SetAttribute (D-3) - 属性を書き込む
'
'   戻り値: True = 書き込めた / False = 失敗 (理由は LastError)
'
'   ★GetAttribute と .Value の違いは D-2 のヘッダー参照★
'     SetAttribute("value", ...) は★属性★を変えるだけで、ユーザーが編集済みの
'     input の★現在値★は変わらない。値を変えたいなら .Value = を使うこと。
'
'   属性名が不正なとき (空文字など) は JS が InvalidCharacterError を投げるので、
'   False + LastError にその名前とメッセージが入る。
'
'   使用例:
'     el.SetAttribute "data-id", "42"
' ============================================================
Public Function SetAttribute(ByVal attrName As String, _
                             ByVal attrValue As String) As Boolean
    Dim js As String

    SetAttribute = False

    If m_pane Is Nothing Then
        m_lastOk = False
        m_lastError = "no-pane"
        m_lastInfo = ""
        Exit Function
    End If

    js = "e.setAttribute(" & m_pane.JsQuote(attrName) & "," & _
         m_pane.JsQuote(attrValue) & ");rv='ok';"

    SetAttribute = ExecAction(js, "SetAttribute")
End Function

' ============================================================
' IsStale (D-2)
'   このハンドルがもう使えなくなっていないかを返す。
'
'   True になる条件:
'     ・ページ遷移などで JS 側のレジストリが作り直された (世代が変わった)
'     ・添字の位置に要素が入っていない (通常は起こらない)
'   判定できなかった場合 (通信失敗など) も安全側に倒して True を返す。
'   このとき LastOk = False なので、区別が必要なら LastError を見ること。
'
'   ★DOM から取り外された要素は stale にならない★
'     世代が同じで配列に残っていれば False を返す。detached な要素も
'     JS 上は有効なオブジェクトなので、読み取り自体は成功する。
' ============================================================
Public Property Get IsStale() As Boolean
    Dim st As String
    Dim stale As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        IsStale = True
        Exit Property
    End If

    stale = m_pane.Elem_CheckStale(m_handle, m_gen, st)

    If Len(st) = 0 Then
        m_lastOk = True
    Else
        m_lastError = st
    End If

    IsStale = stale
End Property


' ============================================================
' ReadProp (内部共通: プロパティ 1 個を読む)
'   valueExpr は JS 側で変数 e (= 対象要素) を使える式。
' ============================================================
Private Function ReadProp(ByVal valueExpr As String) As String
    Dim v As String
    Dim st As String
    Dim ok As Boolean

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ReadProp = ""

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    ok = m_pane.Elem_Read(m_handle, m_gen, valueExpr, v, st)

    If ok Then
        m_lastOk = True
        ReadProp = v
    Else
        m_lastError = st
    End If
End Function


' ============================================================
' ExecAction (D-3、内部共通: 操作を 1 回撃つ)
'   actionStmts は JS 側で変数 e (対象要素) を使える★文★。
'   経路の名前を残したいときは JS 側で変数 rv に代入する。
'   ReadProp と対になる書き込み側の窓口で、状態の更新の仕方も揃えてある。
' ============================================================
Private Function ExecAction(ByVal actionStmts As String, _
                            ByVal apiName As String) As Boolean
    Dim v As String
    Dim st As String

    m_lastOk = False
    m_lastError = ""
    m_lastInfo = ""
    ExecAction = False

    If m_pane Is Nothing Then
        m_lastError = "no-pane"
        Exit Function
    End If

    If m_pane.Elem_Exec(m_handle, m_gen, actionStmts, v, st) Then
        m_lastOk = True
        m_lastInfo = v
        ExecAction = True
    Else
        m_lastError = st
        Wv2Log.LogW "Wv2Element." & apiName & ": 失敗 (handle=" & m_handle & _
                    ") err=" & st
    End If
End Function

