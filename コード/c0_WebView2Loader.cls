VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c0_WebView2Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c0_WebView2Loader

Option Explicit

'Environmentは各アプリに１つあれば十分。
'１つのEnvironmentから複数のController/WebView2のセットを生産できる。

'作成したCOMオブジェクト（ハンドラー）を保持しておくためのポインター
'※Class_Terminate時にGlobalAllocで確保したメモリ領域を解放する時にアドレスが必要
Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pObject As LongPtr

Sub DebugLoader()
    Dim hr As Long
    ' 第4引数に 0 (NULL) を渡すと、本来はエラーコードが返るはずです。
    ' これで Excel 自体が落ちるなら、Declare 文の定義（引数の型）が間違っています。
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr("C:\Temp\VBA_WebView2"), 0, 0)
    Debug.Print "Result: " & hr
End Sub

'Environmentを作成するための関数
Public Sub CreateWebView2Environment()

    ' 1. vTableに関数アドレスを格納
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf Handler_Invoke)

    ' 2. メモリ確保（vTable用）
    pVTable = GlobalAlloc(GPTR, (UBound(vTable) + 1) * LenB(vTable(0)))
    CopyMemory ByVal pVTable, vTable(0), (UBound(vTable) + 1) * LenB(vTable(0))

    ' 3. メモリ確保（オブジェクト本体用）
    pObject = GlobalAlloc(GPTR, LenB(pVTable))
    CopyMemory ByVal pObject, pVTable, LenB(pVTable)

    ' 4. WebView2Environmentの作成開始
    ' 1. 書き込み可能なフォルダパスを用意（Null終端のUnicode文字列）
    Dim userDataPath As String
    userDataPath = "C:\Temp\VBA_WebView2" ' フォルダは自動で作られます

    ' フォルダが存在することを確認（念のため）
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    ' 2. API呼び出し（第2引数にパスのポインタを渡す）
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, pObject)

    If hr <> S_OK Then
        MsgBox "Failed to start initialization: " & hr & vbCrLf & _
               "Check if WebView2Loader.dll is in the same folder as this Workbook."
    End If

    DoEvents

End Sub


