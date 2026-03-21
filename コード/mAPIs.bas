Attribute VB_Name = "mAPIs"
Option Explicit

' WebView2Loader.dll のエントリポイント ※c0のCreateWebView2Environmentメソッド内で実行
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

'自作COMオブジェクト（ハンドラー）をメモリ上に設置するための宣言
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Const GPTR As Long = &H40 'GlobalAllocのuFlags用定数。「固定メモリ」として確保する　というフラグ
Public Const S_OK As Long = 0
Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

#If Win64 Then
    Public Const vbLongPtr = vbLongLong
#Else
    Public Const vbLongPtr = vbLong
#End If

' 文字列の長さを取得する（Unicode版）
Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Public Declare PtrSafe Function SysAllocString Lib "oleaut32.dll" (ByVal pOleChar As LongPtr) As LongPtr
' BSTRメモリを解放する
Public Declare PtrSafe Sub SysFreeString Lib "oleaut32.dll" (ByVal bstr As LongPtr)

Public Declare PtrSafe Function MoveWindow Lib "user32" (ByVal hwnd As LongPtr, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long

Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" ( _
    ByVal hWndParent As LongPtr, _
    ByVal hWndChildAfter As LongPtr, _
    ByVal lpszClass As String, _
    ByVal lpszWindow As String) As LongPtr

Public Declare PtrSafe Function DispCallFunc Lib "oleaut32.dll" ( _
    ByVal pvInstance As LongPtr, _
    ByVal oVft As LongPtr, _
    ByVal cc As Long, _
    ByVal vtReturn As Integer, _
    ByVal cArgs As Long, _
    ByRef rgvt As Integer, _
    ByRef rgpvarg As LongPtr, _
    ByRef pvargResult As Variant) As Long

' 定数
Public Const CC_STDCALL As Long = 4

' 標準モジュールにて
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

' 定数
Public Const LOGPIXELSX As Long = 88 ' 横方向のDPI
Public Const LOGPIXELSY As Long = 90 ' 縦方向のDPI

Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByRef Source As LongPtr, ByVal Length As LongPtr)



Public Sub PutMemPtr(ByVal pDest As LongPtr, ByVal pSrc As LongPtr)
    RtlMoveMemory pDest, pSrc, 8
End Sub
