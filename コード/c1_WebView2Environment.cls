VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c1_WebView2Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class WebView2Environment

Option Explicit

Public Event CreateCoreWebView2ControllerCompleted()

Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pControllerHandler As LongPtr

Public Function CreateWebView2Controller(pEnvironment As LongPtr)
    ' 1. 最新の関数アドレスを取得（毎回実行が必要）
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf ControllerHandler_Invoke)

    ' 2. メモリが未確保なら確保する
    If pVTable = 0 Then
        pVTable = GlobalAlloc(GPTR, 4 * LenB(vTable(0)))
        pControllerHandler = GlobalAlloc(GPTR, LenB(pVTable))
        ' オブジェクト本体(pControllerHandler)にVTable(pVTable)の住所を焼き付けるのは一度でOK
        CopyMemory ByVal pControllerHandler, pVTable, LenB(pVTable)
    End If

    ' 3. VTableの中身（最新のアドレス一覧）をメモリに同期
    ' これにより、再コンパイルで関数が移動していても、pVTableの中身が最新化される
    CopyMemory ByVal pVTable, vTable(0), 4 * LenB(vTable(0))

    ' 4. コントローラー作成
    Dim hr As Long
    hr = CallCreateController(pEnvironment, TargetHwnd, pControllerHandler)
    
    Debug.Print "CreateCoreWebView2Controller Call Result: " & hr
End Function

Public Sub CreateCoreWebView2ControllerCompleted()
    RaiseEvent CreateCoreWebView2ControllerCompleted
End Sub
