VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Wv2SettingsBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''''''''''''''''''''''''''''''''''
' --- Wv2SettingsBridge.cls  第9.29 段階 (プリセット外の任意テンプレート + custom の永続化) ---
'
'   第9.29 の追加事項 (設定画面から任意の検索 URL を入れられるようにする):
'     ★記法は接頭辞方式のまま (論点1 案a)★
'       プリセットと同じ「末尾にエンコード済み検索語を連結するテンプレート」。
'       %s プレースホルダ方式 (Chrome/Firefox 風) は NormalizeUrl / PreviewUrlFor /
'       既存プリセット定義 / 保存済み値の互換をまとめて触ることになるため、
'       独立した別段階に切り出す。
'
'     ★検証は最小限 (論点2 案a)★
'       空でない + http:// または https:// で始まる、のみ。検索テンプレートは
'       ?q= で終わる★不完全な URL★なので、通常の URL 妥当性判定に掛けると
'       正しい入力が落ちる。判定は NormalizeCustomTemplate に一元化する。
'
'     ★保存形式は 2 キー・template 行は custom のときだけ (論点3 案a-1)★
'         engine=custom
'         template=https://example.com/search?q=
'       プリセット選択時は engine=bing の 1 行のみ (第9.20 と同一)。
'       ★常に 2 行書かない理由★ 将来プリセットの URL をコード側で直したとき、
'         保存値が古い URL を復活させる事故になる。「プリセットは名前が正、
'         URL の定義はコードが正」を保つ。
'       ※副作用: custom → プリセットへ切り替えると template 行が消えるので、
'         入力欄に覚えていたカスタム URL も消える (仕様)。
'
'     - 新規 Public (JS 公開、論点5 案a):
'         SetCustomEngine(template) As String
'           検証 → Browser へ適用 → 保存。戻り値は確定したエンジン名。
'           不正なら "" (Browser も保存も一切触らない)。
'           ★入力がプリセットと同一 URL だった場合は "custom" ではなく
'             そのプリセット名を返し、プリセットとして保存する★
'             (TemplateToName の逆引きに素直に従う。設定画面もそのカードを光らせる)
'         GetCustomTemplate() As String
'           入力欄の初期値。現在 custom ならその URL、でなければ保存済み
'           custom、無ければ ""。
'         PreviewUrlForTemplate(template, terms) As String
'           ★適用前★のライブプレビュー用。Browser の状態を一切変えない。
'           エンコードを JS 側で自作せず VBA を通すので本番と完全同一。
'
'     - 新規 Public (VBA 側):
'         LoadCustomTemplate() As String  (保存済み custom テンプレート。検証済み)
'         Debug_SaveSettings(engineName, template)  (検証用ラッパ)
'
'     - 既存の改造:
'         LoadEngineName  → 汎用の LoadValue(keyName) に一般化し、薄いラッパへ。
'                           (engine / template の 2 キーを同じパーサで読むため)
'         LoadEngineTemplate → engine=custom なら template 行を返す (検証済み。
'                           不正なら "" で、Wv2Browser は既定 Google のままになる)
'         SaveEngineName  → SaveSettingsFile(engineName, template) の薄いラッパへ。
'                           ★ログの主語が SaveSettingsFile に変わった★
'         ※Wv2Browser 側は「テンプレート文字列 or 空」を受け取るだけなので無変更
'           (第9.20 の「Browser はエンジン名の語彙を知らない」設計を維持)。
'
' --- Wv2SettingsBridge.cls  第9.20 段階 (検索エンジン選択の永続化) ---
'
'   第9.20 の追加事項 (選択したエンジンを次回起動へ持ち越す):
'     ★保存先は外部ファイル (APPDATA) ? 非表示シートではない★
'       当初 Excel の非表示シートを設定媒体にする案 (案D) も検討したが、
'       それだと Sheets(...) という Excel オブジェクトモデルに依存し、Word 等
'       他ホストへの移植性 (このプロジェクトの生命線) を壊す。よって外部ファイル
'       (%APPDATA%\Wv2Browser\settings.txt) を保存先にする。I/O は VBA 組み込みの
'       Open / Print / Line Input / Dir / MkDir のみで、Win32 も ADODB も使わない
'       (エンジン名は ASCII なので UTF-8 処理は不要)。ホスト非依存を保つ。
'
'     ★責務はこのブリッジに統合 (別モジュールを切らない)★
'       このクラスは既に「検索エンジン設定の司令塔」(プリセット定数 SE_*・
'       名前?テンプレート解決・SetEngine/GetEngine) を握っている。永続化は
'       「その設定を起動間で保持する」責務なので、同じクラスに集約するのが
'       ドメイン凝集として自然。Wv2Store.bas のような別モジュールは切らない
'       (モジュール数を増やさない)。
'
'     ★フォーマットは 1 行 key=value テキスト★
'       例: engine=bing 。今はエンジン 1 項目だが、行を足せば「プリセット外
'       テンプレート」等の将来項目がそのまま 2 行目として乗る。読み込みは行を
'       '=' で分割し engine キーの値だけ拾う。未知キーは無視。
'
'     ★保存/復元のタイミング★
'       ・保存: SetEngine が解決に成功した直後に SaveEngineName を内部呼び。
'         永続化の失敗 (書込権限なし等) は握りつぶし、本番動作は止めない。
'       ・復元: Wv2Browser.Class_Initialize が一時ブリッジを New し
'         LoadEngineTemplate() を呼ぶ (案イ)。値があれば m_searchTemplate を
'         それで上書き、無ければ既定 (Google) のまま。Browser 誕生時に確定する
'         ので、UserForm 経由でも直起動でも一貫して前回値で始まる。
'
'     ★公開した口★
'       LoadEngineName()     As String  保存済みエンジン名 ("bing" 等)。無ければ ""。
'       LoadEngineTemplate() As String  保存済みエンジンのテンプレート文字列。
'                                        無効/無ければ "" (呼び手が既定へ落とす)。
'       Debug_SaveEngineName(name)       検証用: SaveEngineName の薄い公開ラッパ。
'
' --- Wv2SettingsBridge.cls  第9.18 段階 (新規) ---
'
'   ★このクラスは何か★
'     設定画面 (設定タブの HTML/JS) から検索エンジンを操作するための「薄い口」。
'     ICoreWebView2.AddHostObjectToScript でこのインスタンスを JS に公開し、
'     JS から window.chrome.webview.hostObjects.sync.Settings.SetEngine("bing")
'     のように VBA メソッドを直接呼び出せるようにする (postMessage コマンドの
'     文字列パースを介さない直接 RPC 経路)。
'
'   ★なぜ Wv2Browser を丸ごと渡さずブリッジを立てるか (部品2-b)★
'     Wv2Browser を直接 AddHostObjectToScript すると、設定画面の JS から
'     CloseTab / MoveTab などタブ操作まで叩けてしまう (JS に晒す面が広すぎる)。
'     このブリッジは検索エンジン設定に必要な最小 API だけを公開し、内部で
'     Wv2Browser への参照を 1 本持つ。「JS に晒す面」をクラス境界で構造的に
'     最小化する。設定画面が触れるのは検索エンジンだけで、タブ操作には手が届かない。
'
'   ★プリセット解決の集約 (DRY)★
'     "google"/"bing"/"duckduckgo"(別名 "ddg")/"yahoo" ? テンプレート文字列 の
'     相互変換をこのクラスに集約する。9.17 まで UserForm1 が持っていた
'     プリセット定数 (SE_*) と名前解決ロジックはここへ移設し、UserForm1 の
'     UseSearchEngine もこのブリッジ経由に変えた。名前解決が 1 箇所になり、
'     イミディエイトからも設定画面からも同じ経路を通る。
'
'   ★Instancing について★
'     旧 WebView2 実装 (実機検証済み) では Instancing = 1 - Private のクラスでも
'     AddHostObjectToScript が機能した (VBA が生成する IDispatch を OS に公開
'     できれば足りる)。よって本クラスも既定の Private のままでよい。
'     PublicNotCreatable は不要 (かつてそう思い込んでいたが誤りだった)。
'
'   ★公開メソッド (JS から呼ばれる)★
'     GetEngine() As String          現在のエンジン名を返す ("google" 等)。
'                                     どのプリセットにも一致しなければ "custom"。
'     SetEngine(name) As String      プリセット名を解決して Browser に設定。
'                                     実際に設定されたエンジン名を返す。
'     PreviewUrl(terms) As String    指定語をいまのエンジンで検索したときの URL。
'                                     設定画面のライブプレビュー用。
'     EngineList() As String         プリセット定義を JSON 配列で返す。
'                                     設定画面がカードを描くために使う。
'
'   ★VBA→JS の戻り値★
'     旧実装で確認済みのとおり、これらの戻り値はそのまま JS 側の関数戻り値と
'     して受け取れる (hostObjects.sync 経由なら同期的に返る)。
' ============================================================
Option Explicit

' --- 検索エンジンのプリセット (第9.17 で UserForm1 にあったものを移設) ---
'   いずれも「クエリ前置型テンプレート」。末尾にエンコード済み検索語を連結する。
'   DuckDuckGo の /?q= 型・Yahoo の ?p= 型もキー名が違うだけで前置テンプレート
'   1 本で表現できる (第9.17 で確認済み)。
Private Const SE_GOOGLE     As String = "https://www.google.com/search?q="
Private Const SE_BING       As String = "https://www.bing.com/search?q="
Private Const SE_DUCKDUCKGO As String = "https://duckduckgo.com/?q="
Private Const SE_YAHOO      As String = "https://search.yahoo.com/search?p="

' --- 第9.20 段階: 永続化 (検索エンジン選択の保存先) ---
'   %APPDATA%\Wv2Browser\settings.txt に 1 行 key=value で書く。
'   I/O は VBA 組み込みのみ (Win32/ADODB 不使用)。Word 等への移植性を保つ。
Private Const STORE_SUBFOLDER As String = "Wv2Browser"
Private Const STORE_FILENAME  As String = "settings.txt"
Private Const STORE_KEY_ENGINE As String = "engine"
' 第9.29: custom のときだけ書く 2 キー目 (テンプレート URL そのもの)
Private Const STORE_KEY_TEMPLATE As String = "template"

' --- 内部が握る Browser 参照 (設定を反映する先) ---
Private m_browser As Wv2Browser


' ============================================================
' BindBrowser (VBA 側から呼ぶ。JS には公開しない口)
'
'   ブリッジに操作対象の Wv2Browser を結び付ける。設定タブを開いて
'   AddHostObjectToScript する直前に Browser が呼ぶ。
'   ★ライフタイム★ ブリッジ→Browser の一方向参照のみ。Browser はブリッジを
'     WithEvents で持たず、設定タブ (Pane) 経由で hostObject として WebView2 に
'     渡すだけ。設定タブが閉じれば WebView2 側の参照も消え、ブリッジは解放される。
' ============================================================
Public Sub BindBrowser(ByVal br As Wv2Browser)
    Set m_browser = br
End Sub


' ============================================================
' ResolveTemplate (内部ヘルパー: 名前 → テンプレート)
'
'   プリセット名を検索テンプレート文字列へ解決する。大文字小文字・前後空白は
'   無視。未知の名前は空文字を返す (呼び出し側が既定へフォールバックする)。
' ============================================================
Private Function ResolveTemplate(ByVal engineName As String) As String
    Dim key As String
    key = LCase$(Trim$(engineName))
    Select Case key
        Case "google":            ResolveTemplate = SE_GOOGLE
        Case "bing":              ResolveTemplate = SE_BING
        Case "duckduckgo", "ddg": ResolveTemplate = SE_DUCKDUCKGO
        Case "yahoo":             ResolveTemplate = SE_YAHOO
        Case Else:                ResolveTemplate = ""
    End Select
End Function


' ============================================================
' TemplateToName (内部ヘルパー: テンプレート → 名前)
'
'   現在のテンプレート文字列がどのプリセットに一致するかを逆引きする。
'   どれにも一致しなければ "custom" を返す (設定画面はこれを「プリセット外」
'   として扱える)。
' ============================================================
Private Function TemplateToName(ByVal template As String) As String
    Select Case template
        Case SE_GOOGLE:     TemplateToName = "google"
        Case SE_BING:       TemplateToName = "bing"
        Case SE_DUCKDUCKGO: TemplateToName = "duckduckgo"
        Case SE_YAHOO:      TemplateToName = "yahoo"
        Case Else:          TemplateToName = "custom"
    End Select
End Function


' ============================================================
' GetEngine (JS 公開: 現在のエンジン名)
'
'   いま Browser に設定されている検索テンプレートを名前へ逆引きして返す。
'   設定画面が開いたときに「現在どれが選ばれているか」を初期表示するのに使う。
'   Browser 未結合なら空文字。
' ============================================================
Public Function GetEngine() As String
    If m_browser Is Nothing Then
        GetEngine = ""
        Exit Function
    End If
    GetEngine = TemplateToName(m_browser.Debug_SearchEngine())
End Function


' ============================================================
' SetEngine (JS 公開: エンジンを設定)
'
'   プリセット名を解決して Browser.SearchEngine へ設定する。未知の名前は
'   既定 (Google) に落とす。空文字/空白のみに対する最終ガードは Browser 側
'   Property Let が担保するので、ここでは名前解決に専念する (二重ガードにしない)。
'   戻り値は「実際に設定されたエンジン名」(設定画面が選択状態を確定するのに使う)。
' ============================================================
Public Function SetEngine(ByVal engineName As String) As String
    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetEngine: Browser 未結合のため無視 (" & engineName & ")"
        SetEngine = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)

    If Len(template) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetEngine: 未知のエンジン名 '" & engineName & _
                    "'。既定 (Google) にします。"
        template = SE_GOOGLE
    End If

    m_browser.SearchEngine = template
    SetEngine = TemplateToName(m_browser.Debug_SearchEngine())
    Debug.Print "Wv2SettingsBridge.SetEngine: '" & engineName & "' → '" & SetEngine & "'"

    ' 第9.20: 解決に成功したエンジン名を永続化する (次回起動へ持ち越し)。
    '   保存失敗は握りつぶす (本番動作を止めない)。custom は保存しない
    '   (プリセット外は現状 UI から選べないため。将来テンプレート保存を足す)。
    If SetEngine <> "custom" And Len(SetEngine) > 0 Then
        SaveEngineName SetEngine
    End If
End Function


' ============================================================
' NormalizeCustomTemplate (内部ヘルパー: カスタムテンプレートの検証、第9.29)
'
'   受け付ける条件は 2 つだけ (論点2 案a):
'     ・前後空白を落として長さが 0 でない
'     ・http:// または https:// で始まる (大文字小文字は無視)
'   通れば Trim 済みの文字列を、通らなければ "" を返す。
'   ★これ以上厳しくしない★ 検索テンプレートは ?q= で終わる不完全な URL なので、
'     一般的な URL 妥当性判定に掛けると正しい入力まで落ちる。
' ============================================================
Private Function NormalizeCustomTemplate(ByVal template As String) As String
    NormalizeCustomTemplate = ""

    Dim t As String
    t = Trim$(template)
    If Len(t) = 0 Then Exit Function

    Dim lo As String
    lo = LCase$(t)
    If Left$(lo, 7) = "http://" Or Left$(lo, 8) = "https://" Then
        NormalizeCustomTemplate = t
    End If
End Function


' ============================================================
' SetCustomEngine (JS 公開: プリセット外の任意テンプレートを設定、第9.29)
'
'   設定画面のカスタム入力欄の「適用」から呼ばれる。
'     戻り値 "custom"      : カスタムとして適用・保存した
'     戻り値 "bing" 等     : 入力がプリセットと同一 URL だったので、そちらとして
'                            適用・保存した (設定画面はそのカードを光らせる)
'     戻り値 ""            : 不正な入力。★Browser も保存ファイルも一切触っていない★
'
'   ★関数名を右辺で読まない★ VBA では再帰呼び出しになりうるので、確定した
'     エンジン名は一時変数 (nm) に取ってから最後に一度だけ代入する。
' ============================================================
Public Function SetCustomEngine(ByVal template As String) As String
    SetCustomEngine = ""

    If m_browser Is Nothing Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: Browser 未結合のため無視"
        Exit Function
    End If

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: 不正なテンプレートなので無視 ('" & _
                    template & "')"
        Exit Function
    End If

    m_browser.SearchEngine = t

    Dim nm As String
    nm = TemplateToName(m_browser.Debug_SearchEngine())

    If nm = "custom" Then
        SaveSettingsFile "custom", t
    Else
        ' 入力がプリセットと同一 URL だった場合。custom として抱え込まず、
        ' プリセットとして保存する (template 行は書かない = 論点3 案a-1)。
        SaveSettingsFile nm, ""
        Debug.Print "Wv2SettingsBridge.SetCustomEngine: 入力がプリセットと同一だったので '" & _
                    nm & "' として保存"
    End If

    SetCustomEngine = nm
    Debug.Print "Wv2SettingsBridge.SetCustomEngine: 適用 '" & t & "' -> '" & nm & "'"
End Function


' ============================================================
' GetCustomTemplate (JS 公開: カスタム入力欄の初期値、第9.29)
'
'   優先順位:
'     1. いま Browser に入っているテンプレートが custom ならそれ
'     2. 保存ファイルに残っている custom テンプレート
'     3. どちらも無ければ ""
'   Browser 未結合でも 2 が効くので、設定画面は必ず何かしら表示できる。
' ============================================================
Public Function GetCustomTemplate() As String
    Dim t As String
    t = ""

    If Not m_browser Is Nothing Then
        Dim cur As String
        cur = m_browser.Debug_SearchEngine()
        If TemplateToName(cur) = "custom" Then t = cur
    End If

    If Len(t) = 0 Then t = LoadCustomTemplate()

    GetCustomTemplate = t
End Function


' ============================================================
' PreviewUrl (JS 公開: 検索 URL のプレビュー)
'
'   指定した検索語を「いまのエンジン」で検索したときに実際に飛ぶ URL を返す。
'   設定画面でエンジンカードにマウスを乗せると生成 URL がライブで見える、
'   といった演出に使う (hostObjects.sync の同期呼び出しだからこそ自然にできる)。
'   Browser の NormalizeUrl を通すので、本番の検索フォールバックと同一結果。
' ============================================================
Public Function PreviewUrl(ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrl = ""
        Exit Function
    End If
    PreviewUrl = m_browser.Debug_NormalizeUrl(terms)
End Function


' ============================================================
' PreviewUrlFor (JS 公開: 指定エンジンでの検索 URL プレビュー)
'
'   「いまのエンジン」ではなく、指定した名前のエンジンで検索したらどの URL に
'   なるかを返す (エンジンを実際に切り替えずにプレビューだけ見る)。設定画面が
'   各カードのプレビュー URL を一覧表示するのに使う。
'   Browser の状態を変えないよう、テンプレート連結を自前で行う (エンコードは
'   Browser 経由で本番と同一にする)。
' ============================================================
Public Function PreviewUrlFor(ByVal engineName As String, ByVal terms As String) As String
    If m_browser Is Nothing Then
        PreviewUrlFor = ""
        Exit Function
    End If

    Dim template As String
    template = ResolveTemplate(engineName)
    If Len(template) = 0 Then template = SE_GOOGLE

    PreviewUrlFor = template & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' PreviewUrlForTemplate (JS 公開: テンプレート直指定のプレビュー、第9.29)
'
'   カスタム入力欄に打っている★最中★の文字列で、実際に飛ぶ URL を組んで返す。
'   まだ適用していないので Browser の状態には一切触れない (プレビュー専用)。
'   検証を通らない文字列なら "" を返す (設定画面はこれを見てエラー表示する)。
'   ★エンコードは Wv2Url.UrlEncodeComponent★ JS 側で encodeURIComponent を
'     自作すると本番 (BuildSearchUrl) と結果がズレうるので、必ず VBA を通す。
' ============================================================
Public Function PreviewUrlForTemplate(ByVal template As String, ByVal terms As String) As String
    PreviewUrlForTemplate = ""

    Dim t As String
    t = NormalizeCustomTemplate(template)
    If Len(t) = 0 Then Exit Function

    PreviewUrlForTemplate = t & Wv2Url.UrlEncodeComponent(terms)
End Function


' ============================================================
' EngineList (JS 公開: プリセット定義を JSON で返す)
'
'   設定画面がカードを動的に描くための定義を JSON 配列で返す。
'   各要素は {"id":名前, "label":表示名, "template":テンプレート}。
'   プリセットを増やすときはこの 1 箇所とヘルパー 2 つを直せば設定画面に
'   自動で反映される (HTML 側にエンジン名を重複させない)。
'
'   ★JSON 文字列はここで直接組む★ 値は固定の英数 URL とラベルのみで、
'     エスケープが要る文字を含まないため、Wv2Json を介さず素直に連結する。
' ============================================================
Public Function EngineList() As String
    Dim s As String
    s = "["
    s = s & "{""id"":""google"",""label"":""Google"",""template"":""" & SE_GOOGLE & """},"
    s = s & "{""id"":""bing"",""label"":""Bing"",""template"":""" & SE_BING & """},"
    s = s & "{""id"":""duckduckgo"",""label"":""DuckDuckGo"",""template"":""" & SE_DUCKDUCKGO & """},"
    s = s & "{""id"":""yahoo"",""label"":""Yahoo"",""template"":""" & SE_YAHOO & """}"
    s = s & "]"
    EngineList = s
End Function


' ============================================================
' SettingsPath (内部: 保存ファイルのフルパスを組む純関数)
'
'   %APPDATA%\Wv2Browser\settings.txt を返す。APPDATA が取れない環境では
'   空文字を返し、呼び手が保存/読込をスキップできるようにする。
' ============================================================
Private Function SettingsPath() As String
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then
        SettingsPath = ""
        Exit Function
    End If
    SettingsPath = base & "\" & STORE_SUBFOLDER & "\" & STORE_FILENAME
End Function


' ============================================================
' EnsureFolder (内部: 保存フォルダが無ければ作る)
'
'   %APPDATA%\Wv2Browser\ を用意する。既存なら何もしない。失敗は呼び手が
'   握りつぶす前提で、ここでは On Error Resume Next で素通しする。
' ============================================================
Private Sub EnsureFolder()
    Dim base As String
    base = Environ$("APPDATA")
    If Len(base) = 0 Then Exit Sub

    Dim folder As String
    folder = base & "\" & STORE_SUBFOLDER

    On Error Resume Next
    If Len(Dir$(folder, vbDirectory)) = 0 Then
        MkDir folder
    End If
    On Error GoTo 0
End Sub


' ============================================================
' LoadEngineName / LoadCustomTemplate / LoadValue (公開: 保存値を読む)
'
'   ★第9.29 で LoadValue(keyName) に一般化した★ engine と template の 2 キーを
'   同じパーサで読むため。LoadEngineName は薄いラッパになった。
'
'   settings.txt を 1 行ずつ読み、<key>=<value> の <value> を返す。
'   ★値に = が含まれてよい★ 最初の = だけを区切りに使うので、
'   template=https://example.com/search?q= のような値も欠けずに読める。
'   ファイル/フォルダ欠如・読取り失敗・キー不在はすべて "" を返す
'   (呼び手が既定へフォールバックする)。名前の正当性 (プリセットか否か) は
'   ここでは判定しない (LoadEngineTemplate が担う)。
' ============================================================
Public Function LoadEngineName() As String
    LoadEngineName = LoadValue(STORE_KEY_ENGINE)
End Function


' --- 第9.29: 保存済みのカスタムテンプレート (検証を通らない値は "" に落とす) ---
Public Function LoadCustomTemplate() As String
    LoadCustomTemplate = NormalizeCustomTemplate(LoadValue(STORE_KEY_TEMPLATE))
End Function


' --- 第9.29: settings.txt から 1 キーの値を読む汎用パーサ ---
Private Function LoadValue(ByVal keyName As String) As String
    LoadValue = ""

    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Function

    On Error GoTo done
    If Len(Dir$(path)) = 0 Then Exit Function   ' ファイルが無い

    Dim fno As Integer
    fno = FreeFile
    Open path For Input As #fno

    Dim ln As String, eqPos As Long, k As String, v As String
    Do Until EOF(fno)
        Line Input #fno, ln
        eqPos = InStr(ln, "=")
        If eqPos > 0 Then
            k = LCase$(Trim$(Left$(ln, eqPos - 1)))
            v = Trim$(Mid$(ln, eqPos + 1))
            If k = LCase$(keyName) Then
                LoadValue = v
                ' 続けて読む必要はないが、最後の定義を優先するため break しない。
            End If
        End If
    Loop

    Close #fno
    Exit Function

done:
    ' 読取り失敗時: 開いていれば閉じ、"" を返す。
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    LoadValue = ""
End Function


' ============================================================
' LoadEngineTemplate (公開: 保存済みエンジンのテンプレート文字列を返す)
'
'   LoadEngineName で名前を読み、ResolveTemplate でテンプレートへ解決して返す。
'   保存が無い/未知の名前 (custom や壊れた値) の場合は "" を返し、呼び手
'   (Wv2Browser.Class_Initialize) が既定 (Google) のままにできるようにする。
'   ★名前?テンプレート解決をこのクラスに閉じたまま、Browser には「テンプレート
'     文字列 or 空」だけを渡す★ Browser はエンジン名の語彙を知らずに済む。
' ============================================================
Public Function LoadEngineTemplate() As String
    Dim name As String
    name = LoadEngineName()
    If Len(name) = 0 Then
        LoadEngineTemplate = ""
        Exit Function
    End If

    ' 第9.29: custom なら template 行を返す (NormalizeCustomTemplate 済み)。
    '   壊れた template 行や行そのものが無い場合は "" になり、呼び手
    '   (Wv2Browser.Class_Initialize) は既定 (Google) のままになる。
    If LCase$(Trim$(name)) = "custom" Then
        LoadEngineTemplate = LoadCustomTemplate()
        Exit Function
    End If

    LoadEngineTemplate = ResolveTemplate(name)   ' 未知なら "" が返る
End Function


' ============================================================
' SaveEngineName (内部: エンジン名を保存する)
'
'   settings.txt に engine=<name> の 1 行を書く (現状 1 項目なので丸ごと上書き)。
'   フォルダが無ければ EnsureFolder で作る。書込失敗は握りつぶす (永続化の
'   失敗で本番を止めない)。SetEngine が解決成功後に呼ぶ。
'   ★将来 2 項目目以降を足すときは、ここを「既存を読んで該当行だけ差し替え」
'     に育てる。今はエンジン 1 項目なので全書き換えで足りる。
' ============================================================
Private Sub SaveEngineName(ByVal engineName As String)
    SaveSettingsFile engineName, ""
End Sub


' ============================================================
' SaveSettingsFile (内部: settings.txt を書き出す、第9.29)
'
'   engine=<n> を必ず書き、template が非空のときだけ template=<URL> を続けて書く
'   (論点3 案a-1)。プリセット選択時は template を "" で呼ぶので 1 行だけになり、
'   ★custom から プリセットへ切り替えると template 行は自然に消える★。
' ============================================================
Private Sub SaveSettingsFile(ByVal engineName As String, ByVal template As String)
    Dim path As String
    path = SettingsPath()
    If Len(path) = 0 Then Exit Sub

    EnsureFolder

    On Error GoTo fail
    Dim fno As Integer
    fno = FreeFile
    Open path For Output As #fno
    Print #fno, STORE_KEY_ENGINE & "=" & engineName
    If Len(template) > 0 Then
        Print #fno, STORE_KEY_TEMPLATE & "=" & template
    End If
    Close #fno
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: 保存 '" & engineName & "'" & _
                IIf(Len(template) > 0, " / template=" & template, "") & " -> " & path
    Exit Sub

fail:
    On Error Resume Next
    Close #fno
    On Error GoTo 0
    Debug.Print "Wv2SettingsBridge.SaveSettingsFile: 保存失敗 (握りつぶし) '" & engineName & "'"
End Sub


' ============================================================
' Debug_SaveEngineName (検証用: SaveEngineName の公開ラッパ)
'
'   SaveEngineName は Private なので、純ロジック検証 (Test_9_20_Persistence) が
'   保存→読込の往復を確かめられるよう薄い公開口を用意する。本番経路では使わない。
' ============================================================
Public Sub Debug_SaveEngineName(ByVal engineName As String)
    SaveEngineName engineName
End Sub


' ============================================================
' Debug_SaveSettings (検証用: SaveSettingsFile の公開ラッパ、第9.29)
'
'   Test_9_29_Custom / Test_9_20_Persistence が「engine + template を対で
'   退避して書き戻す」ために使う。本番経路では使わない。
' ============================================================
Public Sub Debug_SaveSettings(ByVal engineName As String, ByVal template As String)
    SaveSettingsFile engineName, template
End Sub


