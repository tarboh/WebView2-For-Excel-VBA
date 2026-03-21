VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c4_Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c4_Handler
Option Explicit

Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pHandler As LongPtr
Private m_CallbackAddr As LongPtr ' 呼び出し先（Invoke）のアドレス

Public Token As LongPtr
Private m_ParentPtr As LongPtr ' 親（c2_WebView2）のObjPtr

' 初期化：関数アドレスを渡して土地を作る
Public Sub CreateVTble(ByVal funcAddr As LongPtr, ByVal ParentPtr As LongPtr)
    m_ParentPtr = ParentPtr ' 親の住所を控えておく
    m_CallbackAddr = funcAddr
    
    ' 名簿（VTable）を作成
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = m_CallbackAddr ' 実際の処理を行う関数
    
    ' 土地を確保
    pVTable = GlobalAlloc(GPTR, 4 * LenB(pVTable))
    pHandler = GlobalAlloc(GPTR, LenB(pVTable))
    
    ' 紐付け
    CopyMemory ByVal pVTable, vTable(0), 4 * LenB(pVTable)
    CopyMemory ByVal pHandler, pVTable, LenB(pVTable)
End Sub

' 再コンパイル対策：最新のアドレスで名簿を更新する
Public Sub Refresh()
    ' ※funcAddrはAddressOfなので毎回最新を取得して渡す必要がある
    ' ここでは設計上、外部から呼び出し直す形にする
End Sub

' WebView2に渡すためのポインタを公開
Public Property Get Pointer() As LongPtr
    Pointer = pHandler
End Property

Public Property Get ParentPtr() As LongPtr
    ParentPtr = m_ParentPtr
End Property

' 終了時に確実にメモリを掃除する
Private Sub Class_Terminate()
    If pHandler <> 0 Then GlobalFree pHandler
    If pVTable <> 0 Then GlobalFree pVTable
    Debug.Print "Handler Memory Released."
End Sub
