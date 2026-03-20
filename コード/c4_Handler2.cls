VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c4_Handler2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Public Class c4_Handler2
Option Explicit

Public Event Invoked(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr)

' --- API 宣言 (x64用) ---
Private Declare PtrSafe Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal flAllocationType As Long, ByVal flProtect As Long) As LongPtr
Private Declare PtrSafe Function VirtualFree Lib "kernel32" (ByVal lpAddress As LongPtr, ByVal dwSize As LongPtr, ByVal dwFreeType As Long) As Long
Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByVal Source As LongPtr, ByVal Length As LongPtr)

' VirtualAlloc 定数
Private Const MEM_COMMIT As Long = &H1000
Private Const MEM_RESERVE As Long = &H2000
Private Const PAGE_EXECUTE_READWRITE As Long = &H40
Private Const MEM_RELEASE As Long = &H8000

' --- 内部変数 ---
Private mVTablePtr As LongPtr      ' 捏造VTable(関数ポインタの配列)のアドレス
Private mFakeObjPtr As LongPtr     ' WebView2に渡す「オブジェクトのフリ」をするアドレス
Private mThunkAddr As LongPtr      ' マシン語サンク(通訳)のアドレス
Public Token As LongPtr            ' WebView2解除用トークン

' 解放用サイズ定数
Private Const VTABLE_SIZE As Long = 32 ' IUnknown(3) + Invoke(1) = 4個 * 8byte
Private Const THUNK_SIZE As Long = 32  ' マシン語領域

Public ppWebView2 As LongPtr

' --- クラス内部変数に追加 ---
Private mThunkTimer As LongPtr     ' タイマー専用サンクのアドレス
Private m_TimerID As LongPtr       ' タイマー停止用

' --- 外部公開用：WebView2に渡すポインタ ---
Public Property Get Pointer() As LongPtr
    Pointer = mFakeObjPtr
End Property

' ==========================================================
'  本番：WebView2から叩かれるメソッド
' ==========================================================
' サンクによってレジスタが整えられているため、シグネチャ通りに受け取れます。
' pThis: 捏造オブジェクト(mFakeObjPtr), sender: WebView2本体, args: イベント引数
Public Function cls_Invoke(ByVal pThis As LongPtr, ByVal sender As LongPtr, ByVal args As LongPtr) As Long
    On Error Resume Next ' 外部呼び出しのため安全策
    
    RaiseEvent Invoked(pThis, sender, args)
'    Debug.Print "=== WebView2 Event Fired ==="
'    Debug.Print "Sender Ptr:", sender
'    Debug.Print "Args Ptr:", args
'    Debug.Print "ppWebView2:", ppWebView2
    ' ここで Me (このクラス自身) のプロパティ等にも安全にアクセス可能です。
    ' 例: Debug.Print Me.Token
    
    cls_Invoke = 0 ' S_OK
End Function

' ==========================================================
'  初期化：メモリ構築とマシン語インジェクション
' ==========================================================
Private Sub Class_Initialize()
    ' 1. メモリ確保 (VTable, Thunk, FakeInstance)
    mVTablePtr = VirtualAlloc(0, VTABLE_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    mThunkAddr = VirtualAlloc(0, THUNK_SIZE, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    mFakeObjPtr = VirtualAlloc(0, 8, MEM_COMMIT Or MEM_RESERVE, PAGE_EXECUTE_READWRITE)
    
    ' 2. クラスメソッド "cls_Invoke" の実体アドレスを取得
    Dim pTargetFunc As LongPtr, offset As Long
    pTargetFunc = GetClassMethodPtr(Me, "cls_Invoke", offset) ' ※既存の取得関数を使用
    
    ' 3. x64用 マシン語サンク作成
    ' 内容: WebView2の引数(RCX, RDX, R8)を一つ右(RDX, R8, R9)にずらし、RCXにこのクラスのMeをセットしてジャンプ
    Dim thunkBytes(31) As Byte
    Dim i As Long: i = 0
    
    ' mov r9, r8   (args -> 第4引数へ)
    thunkBytes(i) = &H4D: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HC1: i = i + 3
    ' mov r8, rdx  (sender -> 第3引数へ)
    thunkBytes(i) = &H49: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HD0: i = i + 3
    ' mov rdx, rcx (pThis -> 第2引数へ)
    thunkBytes(i) = &H48: thunkBytes(i + 1) = &H89: thunkBytes(i + 2) = &HCA: i = i + 3
    ' mov rcx, [ObjPtr(Me)] (Me -> 第1引数へ注入)
    thunkBytes(i) = &H48: thunkBytes(i + 1) = &HB9: i = i + 2
    RtlMoveMemory VarPtr(thunkBytes(i)), VarPtr(ObjPtr(Me)), 8: i = i + 8
    ' mov rax, [pTargetFunc] (ジャンプ先セット)
    thunkBytes(i) = &H48: thunkBytes(i + 1) = &HB8: i = i + 2
    RtlMoveMemory VarPtr(thunkBytes(i)), VarPtr(pTargetFunc), 8: i = i + 8
    ' jmp rax (転送)
    thunkBytes(i) = &HFF: thunkBytes(i + 1) = &HE0
    
    RtlMoveMemory mThunkAddr, VarPtr(thunkBytes(0)), THUNK_SIZE
    
    ' 4. VTable 構築 (Index 0-2 は標準モジュールのダミー、3 にサンク)
    ' ※Handler_QI/AddRef/Release は標準モジュールに定義されている前提
    Dim pQI As LongPtr, pAddRef As LongPtr, pRelease As LongPtr
    pQI = GetAddr(AddressOf Handler_QueryInterface)
    pAddRef = GetAddr(AddressOf Handler_AddRef)
    pRelease = GetAddr(AddressOf Handler_Release)
    
    RtlMoveMemory mVTablePtr + (0 * 8), VarPtr(pQI), 8
    RtlMoveMemory mVTablePtr + (1 * 8), VarPtr(pAddRef), 8
    RtlMoveMemory mVTablePtr + (2 * 8), VarPtr(pRelease), 8
    RtlMoveMemory mVTablePtr + (3 * 8), VarPtr(mThunkAddr), 8
    
    ' 5. FakeInstance の先頭に VTable へのポインタを書き込む
    RtlMoveMemory mFakeObjPtr, VarPtr(mVTablePtr), 8
    
    Debug.Print "Handler Ready. FakeObjPtr: " & Hex(mFakeObjPtr)
End Sub

Private Sub Class_Terminate()
    ' メモリ解放
    If mVTablePtr <> 0 Then VirtualFree mVTablePtr, 0, MEM_RELEASE
    If mThunkAddr <> 0 Then VirtualFree mThunkAddr, 0, MEM_RELEASE
    If mFakeObjPtr <> 0 Then VirtualFree mFakeObjPtr, 0, MEM_RELEASE
    Debug.Print "Handler Memory Released."
End Sub


' --- タイマー専用の受け皿メソッド ---
' 本来: (hwnd, msg, id, time) -> 4つ
' サンク後: (Me, hwnd, msg, id, time) -> 5つ
Public Function cls_TimerProc(ByVal pMe As LongPtr, ByVal hwnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long) As Long
    Debug.Print "--- Timer Fired! ---"
    Debug.Print "TimeCount:", dwTime
    cls_TimerProc = 0
End Function
