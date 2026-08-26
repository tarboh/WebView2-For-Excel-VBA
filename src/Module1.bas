Attribute VB_Name = "Module1"
''''''''''''''''''''''''''''''''''
' --- Module1.bas  K-4d 段階 (検証の入口をここに置く) ---
'
'   ★VBE で直に書いたコードは import.ps1 のたびに消える★
'   src\ が正なので、書き戻すとブック側は src\Module1.bas の内容で上書きされる。
'   検証でよく打つものは★ここに置く★こと。
''''''''''''''''''''''''''''''''''
Option Explicit

' ★★★ 日本語名は IME が要る = クラッシュの引き金 (仕様事実67) ★★★
'   WebView2 を起動した Excel では、イミディエイトに IME で打つと落ちうる。
'   ★英数字の別名を用意してあるのでそちらを使うこと★ (Boot / BootNoQuiet / BootMaps)。
'   日本語名は残してあるが、打つと IME が起動して危ない。

Sub Boot()
    実行用
End Sub

Sub BootNoQuiet()
    実行用2
End Sub

Sub BootMaps()
    実行用3
End Sub


' 起動一式。★毎回ここから★ (ログを 1 回分に閉じるため LogStart を含む)
Sub 実行用()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Full
End Sub

' 解放前の静穏化を切る (K-4c)。★遅い解放を再現したいときだけ★
'   既定は True (= 閉じる前に全タブを about:blank へ飛ばす)。
Sub 実行用2()
    UserForm1.CurrentBrowser.QuietOnShutdown = False
End Sub

' Maps を開く (本物の MapsOpen。K-4d の計測ログが出る)
Sub 実行用3()
    Wv2Maps.MapsOpen UserForm1.CurrentBrowser
End Sub

' ============================================================
' ★クラッシュの切り分け (K-5)★
'
'   症状: Wv2Log.LogStart のあと何かログを書くと、そのあと
'         ★イミディエイトで日本語入力オンにして 1 文字打つと Excel が落ちる★。
'         日本語入力オフなら平気。日本語をコピペするのも平気。
'
'   ★実行前に必ず保存すること。落ちる前提の実験。★
'
'   使い方: 1 本実行 → イミディエイトで日本語入力オンにして 1 文字打つ
'           → 落ちるかどうかを見る。★WebView2 もフォームも使わない★
'
'     CrashTest1 … LogStart だけ (ヘッダー 3 行が書かれる)
'     CrashTest2 … LogStart + LogI 1 行     ← ★最小再現の候補★
'     CrashTest3 … CrashTest2 + LogStop     ← 落ちなければ★開いたファイルが原因★
'     CrashTest4 … LogStart + LogW (WARN は自動で Close→Open する)
'     CrashTest5 … LogEcho = False で LogI  ← 落ちなければ★Debug.Print が原因★
' ============================================================
Sub CrashTest1()
    Wv2Log.LogStart
    Debug.Print "--- CrashTest1 済。日本語入力オンで 1 文字打ってみてください ---"
End Sub

Sub CrashTest2()
    Wv2Log.LogStart
    Wv2Log.LogI "テスト"
    Debug.Print "--- CrashTest2 済。日本語入力オンで 1 文字打ってみてください ---"
End Sub

Sub CrashTest3()
    Wv2Log.LogStart
    Wv2Log.LogI "テスト"
    Wv2Log.LogStop            ' ★ファイルを閉じる★
    Debug.Print "--- CrashTest3 済 (ファイルは閉じた)。1 文字打ってみてください ---"
End Sub

Sub CrashTest4()
    Wv2Log.LogStart
    Wv2Log.LogW "テスト"      ' WARN は LogFlush (Close → Open) を伴う
    Debug.Print "--- CrashTest4 済。1 文字打ってみてください ---"
End Sub

Sub CrashTest5()
    Wv2Log.LogStart
    Wv2Log.LogEcho = False    ' ★Debug.Print を出さない★
    Wv2Log.LogI "テスト"
    Wv2Log.LogEcho = True
    Debug.Print "--- CrashTest5 済 (ログはファイルだけ)。1 文字打ってみてください ---"
End Sub

' ============================================================
' ★第 2 弾: Wv2Log すら関係ないかもしれない (K-5)★
'
'   CrashTest1～5 (Wv2Log だけ) では★一度も落ちなかった★。
'   落ちるのは LogStart の★あとに UserForm1.Show を通したとき★。
'   逆順 (Show が先、LogStart が最後) では落ちない。
'
'   → 疑いは ★「ファイルを開いたままモードレスフォームを表示する」★ こと。
'     LogStart はログファイルを Open して★開きっぱなしにする★。
'
'   ★実行前に必ず保存すること。★
'   1 本実行 → イミディエイトで日本語入力オンにして 1 文字打つ → 落ちるか見る。
' ============================================================

' [6] パターン2 の Sub 版 (最小再現の確認)
Sub CrashTest6()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
End Sub

' [7] 順序を逆にするだけ (落ちなければ★順序が効く★と確定)
Sub CrashTest7()
    UserForm1.Show vbModeless
    Wv2Log.LogStart
End Sub

' [8] ★Wv2Log を使わず、ただファイルを開いたまま Show する★
'     落ちれば ★Wv2Log は無罪★。VBA の素の挙動ということになる。
'     開いたファイルは Excel を閉じれば解放される (.tmp が残っても無害)。
Sub CrashTest8()
    Dim crashPath As String
    Dim n As Long

    crashPath = Environ$("APPDATA") & "\Wv2Browser\crashtest.tmp"
    n = FreeFile
    Open crashPath For Binary Access Write Shared As #n
    ' ★わざと閉じない★
    Debug.Print "CrashTest8: ファイルを開いたまま Show します (" & crashPath & ")"
    UserForm1.Show vbModeless
End Sub

' [9] ★ファイルは開かず、Dir でログフォルダを列挙してから Show★
'     LogStart は RotateOldLogs で Dir / Kill を使う。そちらが原因かを見る。
Sub CrashTest9()
    Dim f As String
    Dim n As Long

    f = Dir(Environ$("APPDATA") & "\Wv2Browser\logs\*.log")
    Do While Len(f) > 0
        n = n + 1
        f = Dir
    Loop
    Debug.Print "CrashTest9: ログを " & n & " 本列挙してから Show します"
    UserForm1.Show vbModeless
End Sub

' ============================================================
' ★第 3 弾: Full と Spa の差はどこか (K-5)★
'
'   実測: StartWebView2_Full → ★落ちる★ / StartWebView2_Spa → ★落ちない★
'   差は 3 つ:
'     (a) TabBar / NavBar の有無 (★NavBar にはアドレスバーの input がある★)
'     (b) 中身が実サイトかローカル HTML か (bing には検索ボックスがある)
'     (c) WebView2 コントローラの枚数 (Full は 5 枚、Spa は 1 枚)
'
'   ★実行前に必ず保存すること。★
'   実行 → コードウィンドウに触らずイミディエイトで日本語入力オンにして 1 文字。
' ============================================================

' [10] Spa + ★実サイトのタブを 1 枚★ (TabBar / NavBar は無し)
'      落ちれば → ★(b) 中身が犯人★ (入力欄を持つページ)。UI 部品は無罪
'      落ちなければ → ★(a) TabBar / NavBar が犯人★ の線が濃くなる
Sub CrashTest10()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Spa
    UserForm1.CurrentBrowser.AddTabWithUrl "https://www.bing.com"
    Debug.Print "CrashTest10: Spa + bing (UI 部品なし)。1 文字打ってみてください"
End Sub

' [11] Spa + ★about:blank を 3 枚★ (枚数だけ増やす。入力欄は無い)
'      落ちれば → ★(c) コントローラの枚数が犯人★
'      落ちなければ → 枚数は無関係
Sub CrashTest11()
    Dim i As Long

    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Spa
    For i = 1 To 3
        UserForm1.CurrentBrowser.AddTabWithUrl "about:blank"
    Next i
    Debug.Print "CrashTest11: Spa + about:blank ×3 (入力欄なし)。1 文字打ってみてください"
End Sub

' [12] 対照: Full そのまま (★確実に落ちるはず★)
'      他のテストが「落ちない」ことに意味を持たせるための基準
Sub CrashTest12()
    Wv2Log.LogStart
    UserForm1.Show vbModeless
    UserForm1.StartWebView2_Full
    Debug.Print "CrashTest12: Full (対照)。1 文字打ってみてください"
End Sub
