Attribute VB_Name = "Wv2Log"
''''''''''''''''''''''''''''''''''
' --- Wv2Log.bas  K-1 段階 (デバッグログのファイル出力) ---
'
'   ★なぜ要るか★
'     D-2 の検証で、イミディエイトウィンドウの保持行数が足りず前半のログが流れて
'     確認できない事態が起きた。ハンドルの連番から健全性は確定できたが、
'     ★流れたことに気づけない★のが本質的な危険。D 軸は 1 テストで数十往復するため
'     必ず再発する。以降のすべての検証の土台。
'
'   ★何を持つか★
'     ・LogE / LogW / LogI / LogD : レベル別の 1 行出力
'     ・LogPath                   : 現在のログファイルのフルパス (Property Get)
'     ・LogLevel                  : しきい値 (Property Get/Let)。既定は LOG_DEBUG
'     ・LogEcho                   : Debug.Print への併記 (Property Get/Let)。既定は True
'     ・LogStart / LogStop / LogFlush : 明示制御
'
'   ★設計の要点 (K-1 で 9 論点を一括合意)★
'     ・出力先は %APPDATA%\Wv2Browser\logs\。settings.txt と同じ根に揃える。
'       ブックと同じ場所に置くと OneDrive の同期と衝突しうるので避けた
'     ・★起動ごとに 1 本★ (wv2_yyyymmdd_hhnnss.log)。古いものは 20 本残して自動削除。
'       D 軸は「1 テスト = 数十往復」なので、テスト 1 回分が 1 ファイルに閉じているのが効く
'     ・★LogPrint が内部で Debug.Print も撃つ★ 呼び出し側は 1 行だけ。
'       これで Debug.Print "..." -> Wv2Log.LogD "..." の機械的な 1 対 1 置換で移行でき、
'       イミディエイトの見え方は変わらない。LogEcho = False で切れる
'     ・★ファイルは開きっぱなし★ 毎行書く。ERROR / WARN は書いた直後に強制フラッシュ
'       ファイルは Shared で開く。★Lock Write では ADODB も外部エディタも
'       開けなかった (K-1 の実機検証で実測)★ 書き手はこのプロセスだけなので、
'       ロックを掛けずに読み手を通す方が要件に合う。
'       (閉じて開き直す)。メモリに溜めて終了時に書く方式は、仕様事実20 でクラッシュが
'       現実的な領域なので採らない
'     ・★再入ガードはカウンタ + 保留キュー★ (設計原則78 と同じ型)。再入中の行を
'       捨てると「流れたことに気づけない」という K-1 の目的そのものを裏切るので、
'       積んでおいて抜けたときに書く
'     ・★行に連番を振る★ 欠番を見れば「流れた」ことに気づける。K-1 の動機に直結する
'
'   ★文字コード★
'     ログファイルは UTF-8 / BOM なし / CRLF。UTF-8 変換は本モジュール内で
'     手書きしている (ADODB.Stream を毎行作ると重いため)。実行時の String には
'     Web ページ由来の任意の Unicode が入りうるので、サロゲートペアも正しく畳む。
'     ※ VBA の「ソース」が CP932 に縛られること (仕様事実35) とは別の話。
'
'   ★Wv2Tests.WriteUtf8NoBom は再利用しない★
'     Wv2Tests は「検証専用。移植時に外せる」位置づけ (本番実体 12 本に入っていない)。
'     製品コードがそこに依存してはいけない。
''''''''''''''''''''''''''''''''''

Option Explicit

' ===== ログレベル =====
Public Const LOG_ERROR As Long = 1
Public Const LOG_WARN As Long = 2
Public Const LOG_INFO As Long = 3
Public Const LOG_DEBUG As Long = 4

Private Const LOG_KEEP_FILES As Long = 20
Private Const LOG_PREFIX As String = "wv2_"

Private m_logNum As Long
Private m_logPath As String
Private m_logLevel As Long
Private m_logEcho As Boolean
Private m_logDepth As Long
Private m_logPending As String
Private m_logSeq As Long
Private m_logInit As Boolean


' ============================================================
' LogE / LogW / LogI / LogD
'
'   レベル別の 1 行出力。呼び出し側はこれだけ書けばよい。
'   ERROR と WARN は書いた直後に強制フラッシュする。
' ============================================================
Public Sub LogE(ByVal logMsg As String)
    LogWrite LOG_ERROR, logMsg
End Sub

Public Sub LogW(ByVal logMsg As String)
    LogWrite LOG_WARN, logMsg
End Sub

Public Sub LogI(ByVal logMsg As String)
    LogWrite LOG_INFO, logMsg
End Sub

Public Sub LogD(ByVal logMsg As String)
    LogWrite LOG_DEBUG, logMsg
End Sub


' ============================================================
' LogPath / LogLevel / LogEcho
' ============================================================
Public Property Get LogPath() As String
    LogPath = m_logPath
End Property

Public Property Get LogLevel() As Long
    If Not m_logInit Then InitDefaults
    LogLevel = m_logLevel
End Property

Public Property Let LogLevel(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < LOG_ERROR Then
        m_logLevel = LOG_ERROR
    ElseIf logArg > LOG_DEBUG Then
        m_logLevel = LOG_DEBUG
    Else
        m_logLevel = logArg
    End If
End Property

Public Property Get LogEcho() As Boolean
    If Not m_logInit Then InitDefaults
    LogEcho = m_logEcho
End Property

Public Property Let LogEcho(ByVal logArg As Boolean)
    If Not m_logInit Then InitDefaults
    m_logEcho = logArg
End Property


' ============================================================
' InitDefaults
'
'   モジュール変数の既定値。標準モジュールには Class_Initialize が無いので、
'   最初にアクセスされたときに 1 度だけ通す。
' ============================================================
Private Sub InitDefaults()
    If m_logInit Then Exit Sub
    m_logInit = True
    m_logLevel = LOG_DEBUG
    m_logEcho = True
    m_logSeq = 0
    m_logDepth = 0
    m_logPending = ""
End Sub


' ============================================================
' LogStart
'
'   新しいログファイルを開く。既に開いていれば閉じてから開き直す。
'   通常は明示で呼ばなくてよい (最初の LogD で自動的に開く)。
'   検証を仕切り直したいときに使う。
' ============================================================
Public Sub LogStart()
    On Error Resume Next
    LogStop
    On Error GoTo 0

    InitDefaults
    m_logSeq = 0

    Dim logFolder As String
    logFolder = LogFolderPath()
    If Len(logFolder) = 0 Then Exit Sub

    RotateOldLogs logFolder

    m_logPath = logFolder & LOG_PREFIX & Format$(Now, "yyyymmdd_hhnnss") & ".log"

    On Error GoTo eh
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1

    Dim logHead As String
    logHead = "=== Wv2Log 開始 " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & " ===" & vbCrLf & _
              "=== path = " & m_logPath & " ===" & vbCrLf & _
              "=== 行頭の 6 桁は連番。欠番があればログが落ちている ===" & vbCrLf
    PutBytes Utf8FromString(logHead)

    Debug.Print "Wv2Log.LogStart: " & m_logPath
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogStart: 失敗 (" & Err.Number & ") " & Err.Description
    m_logNum = 0
    m_logPath = ""
End Sub


' ============================================================
' LogStop
' ============================================================
Public Sub LogStop()
    If m_logNum = 0 Then Exit Sub
    On Error Resume Next
    Close #m_logNum
    On Error GoTo 0
    m_logNum = 0
End Sub


' ============================================================
' LogFlush
'
'   閉じて開き直すことで、OS のバッファをディスクへ落とす。
'   ERROR / WARN のときは LogWrite が自動で呼ぶ。
' ============================================================
Public Sub LogFlush()
    If m_logNum = 0 Then Exit Sub
    If Len(m_logPath) = 0 Then Exit Sub

    On Error GoTo eh
    Close #m_logNum
    m_logNum = FreeFile
    Open m_logPath For Binary Access Write Shared As #m_logNum
    Seek #m_logNum, LOF(m_logNum) + 1
    Exit Sub
eh:
    Debug.Print "Wv2Log.LogFlush: 失敗 (" & Err.Number & ") " & Err.Description
    m_logNum = 0
End Sub


' ============================================================
' LogWrite  (中核)
'
'   ・しきい値で早期 return
'   ・Debug.Print への併記 (LogEcho)
'   ・★再入ガード★ 書き込み中に再入したら保留キューへ積み、抜けるときに流す
' ============================================================
Private Sub LogWrite(ByVal logArg As Long, ByVal logMsg As String)
    If Not m_logInit Then InitDefaults
    If logArg > m_logLevel Then Exit Sub

    m_logSeq = m_logSeq + 1

    Dim logLine As String
    logLine = Right$("00000" & CStr(m_logSeq), 6) & " " & NowStamp() & _
              " [" & LevelTag(logArg) & "] " & logMsg

    If m_logEcho Then Debug.Print logLine

    ' --- 再入していたら積むだけ (設計原則78 と同じ型) ---
    If m_logDepth > 0 Then
        m_logPending = m_logPending & logLine & vbCrLf
        Exit Sub
    End If

    m_logDepth = m_logDepth + 1
    On Error Resume Next

    If m_logNum = 0 Then LogStart
    If m_logNum <> 0 Then
        PutBytes Utf8FromString(logLine & vbCrLf)

        ' --- 再入中に積まれた分を流す ---
        If Len(m_logPending) > 0 Then
            Dim logDrain As String
            logDrain = m_logPending
            m_logPending = ""
            PutBytes Utf8FromString(logDrain)
        End If
    End If

    On Error GoTo 0
    m_logDepth = m_logDepth - 1

    If logArg <= LOG_WARN Then LogFlush
End Sub


' ============================================================
' Debug_SetLogDepth  (検証専用)
'
'   再入ガードを外から動かすためのフック。Wv2Pane.Debug_SetInCallback と同じ型。
'   1 以上にすると LogWrite は保留キューへ積むだけになり、0 に戻したあとの
'   次の 1 行でまとめて流れる。★製品コードからは呼ばない★
' ============================================================
Public Sub Debug_SetLogDepth(ByVal logArg As Long)
    If Not m_logInit Then InitDefaults
    If logArg < 0 Then logArg = 0
    m_logDepth = logArg
End Sub

' ============================================================
' PutBytes
' ============================================================
Private Sub PutBytes(ByRef outBytes() As Byte)
    On Error Resume Next
    If m_logNum = 0 Then Exit Sub
    Put #m_logNum, , outBytes
End Sub


' ============================================================
' LevelTag
' ============================================================
Private Function LevelTag(ByVal logArg As Long) As String
    Select Case logArg
        Case LOG_ERROR: LevelTag = "ERR "
        Case LOG_WARN:  LevelTag = "WARN"
        Case LOG_INFO:  LevelTag = "INFO"
        Case Else:      LevelTag = "DBG "
    End Select
End Function


' ============================================================
' NowStamp
'
'   yyyy-mm-dd hh:nn:ss.mmm。ミリ秒は Timer から取る
'   (Now は秒までしか無いため。秒の境界で 1 桁ずれることがあるが、
'    ログの前後関係を読むには十分)。
' ============================================================
Private Function NowStamp() As String
    Dim logMs As Long
    logMs = CLng((Timer - Int(Timer)) * 1000)
    If logMs > 999 Then logMs = 999
    If logMs < 0 Then logMs = 0
    NowStamp = Format$(Now, "yyyy-mm-dd hh:nn:ss") & "." & Right$("00" & CStr(logMs), 3)
End Function


' ============================================================
' LogFolderPath
'
'   %APPDATA%\Wv2Browser\logs\ を作って返す。失敗したら空文字。
' ============================================================
Private Function LogFolderPath() As String
    On Error GoTo eh
    Dim logFolder As String
    logFolder = Environ$("APPDATA")
    If Len(logFolder) = 0 Then Exit Function
    If Right$(logFolder, 1) <> "\" Then logFolder = logFolder & "\"

    logFolder = logFolder & "Wv2Browser\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    logFolder = logFolder & "logs\"
    If Dir(logFolder, vbDirectory) = "" Then MkDir logFolder

    LogFolderPath = logFolder
    Exit Function
eh:
    Debug.Print "Wv2Log.LogFolderPath: 失敗 (" & Err.Number & ") " & Err.Description
End Function


' ============================================================
' RotateOldLogs
'
'   wv2_*.log を新しい順に LOG_KEEP_FILES 本残して、それより古いものを消す。
'   ファイル名がタイムスタンプなので、名前順 = 時刻順になる。
'   ★Dir は状態を持つので、先に全部集めてから消す★
' ============================================================
Private Sub RotateOldLogs(ByVal logFolder As String)
    On Error GoTo eh

    Dim logNames() As String
    Dim logCount As Long
    ReDim logNames(0 To 255)
    logCount = 0

    Dim logName As String
    logName = Dir(logFolder & LOG_PREFIX & "*.log")
    Do While Len(logName) > 0
        If logCount > UBound(logNames) Then ReDim Preserve logNames(0 To logCount + 255)
        logNames(logCount) = logName
        logCount = logCount + 1
        logName = Dir
    Loop

    If logCount <= LOG_KEEP_FILES Then Exit Sub

    ' --- 名前順に並べる (単純な挿入ソート。本数が少ないので十分) ---
    Dim logIdx As Long
    Dim logLow As Long
    Dim logHigh As String
    For logIdx = 1 To logCount - 1
        logHigh = logNames(logIdx)
        logLow = logIdx - 1
        Do While logLow >= 0
            If logNames(logLow) <= logHigh Then Exit Do
            logNames(logLow + 1) = logNames(logLow)
            logLow = logLow - 1
        Loop
        logNames(logLow + 1) = logHigh
    Next logIdx

    Dim logKept As Long
    logKept = logCount - LOG_KEEP_FILES
    For logIdx = 0 To logKept - 1
        On Error Resume Next
        Kill logFolder & logNames(logIdx)
        On Error GoTo 0
    Next logIdx

    Debug.Print "Wv2Log.RotateOldLogs: " & logKept & " 本を削除 (" & LOG_KEEP_FILES & " 本を保持)"
    Exit Sub
eh:
    Debug.Print "Wv2Log.RotateOldLogs: 失敗 (" & Err.Number & ") " & Err.Description
End Sub


' ============================================================
' Utf8FromString
'
'   String (UTF-16 コードユニット列) を UTF-8 バイト列にする。
'   ★サロゲートペアを 4 バイトに畳む★ 実行時の String には Web ページ由来の
'   任意の Unicode が入りうるため。
'   ADODB.Stream を毎行作ると重いので手書きしている。
' ============================================================
Private Function Utf8FromString(ByVal logText As String) As Byte()
    Dim outBytes() As Byte
    Dim logLen As Long
    Dim logIdx As Long
    Dim logCount As Long
    Dim logCode As Long
    Dim logLow As Long

    logLen = Len(logText)
    If logLen = 0 Then
        ReDim outBytes(0 To 0)
        Utf8FromString = outBytes
        Exit Function
    End If

    ' 最悪 1 文字 4 バイト
    ReDim outBytes(0 To logLen * 4 - 1)
    logCount = 0
    logIdx = 1

    Do While logIdx <= logLen
        logCode = AscW(Mid$(logText, logIdx, 1))
        ' ★AscW は 0x8000 以上を負の Integer で返す (仕様事実16)★
        If logCode < 0 Then logCode = logCode + 65536

        ' --- サロゲートペア (上位 D800-DBFF + 下位 DC00-DFFF) ---
        If logCode >= &HD800& And logCode <= &HDBFF& And logIdx < logLen Then
            logLow = AscW(Mid$(logText, logIdx + 1, 1))
            If logLow < 0 Then logLow = logLow + 65536
            If logLow >= &HDC00& And logLow <= &HDFFF& Then
                logCode = &H10000 + (logCode - &HD800&) * &H400 + (logLow - &HDC00&)
                logIdx = logIdx + 1
            End If
        End If

        If logCode < &H80 Then
            outBytes(logCount) = CByte(logCode)
            logCount = logCount + 1
        ElseIf logCode < &H800 Then
            outBytes(logCount) = CByte(&HC0 + (logCode \ &H40))
            outBytes(logCount + 1) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 2
        ElseIf logCode < &H10000 Then
            outBytes(logCount) = CByte(&HE0 + (logCode \ &H1000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 3
        Else
            outBytes(logCount) = CByte(&HF0 + (logCode \ &H40000))
            outBytes(logCount + 1) = CByte(&H80 + ((logCode \ &H1000) Mod &H40))
            outBytes(logCount + 2) = CByte(&H80 + ((logCode \ &H40) Mod &H40))
            outBytes(logCount + 3) = CByte(&H80 + (logCode Mod &H40))
            logCount = logCount + 4
        End If

        logIdx = logIdx + 1
    Loop

    ReDim Preserve outBytes(0 To logCount - 1)
    Utf8FromString = outBytes
End Function

