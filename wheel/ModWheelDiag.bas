Attribute VB_Name = "ModWheelDiag"
'==============================================================================
' ModWheelDiag.bas
'   ホイールメッセージがどのhwndに配送されているかを特定する診断ユーティリティ。
'==============================================================================
Option Explicit

Private Declare PtrSafe Function GetCursorPos Lib "user32" (ByRef lpPoint As POINTAPI) As Long
Private Declare PtrSafe Function WindowFromPoint Lib "user32" ( _
    ByVal xPoint As Long, ByVal yPoint As Long) As LongPtr
Private Declare PtrSafe Function GetAncestor Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal gaFlags As Long) As LongPtr
Private Declare PtrSafe Function GetParent Lib "user32" ( _
    ByVal hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function GetClassNameA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowTextA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal lpString As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetAsyncKeyState Lib "user32" ( _
    ByVal vKey As Long) As Integer
Private Declare PtrSafe Function GetWindowLongPtrA Lib "user32" ( _
    ByVal hwnd As LongPtr, ByVal nIndex As Long) As LongPtr

Private Const GA_ROOT As Long = 2
Private Const GA_ROOTOWNER As Long = 3
Private Const GWLP_WNDPROC As Long = -4

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare PtrSafe Function SetCursorPos Lib "user32" ( _
    ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
    ByVal hwnd As LongPtr, ByRef lpRect As RECT) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'==============================================================================
' 指定 hwnd の中央にカーソルを移動してから階層ダンプ
'   使い方: DumpAtHwnd &H2D0A5C  (UserForm の hwnd を16進で渡す)
'==============================================================================
Public Sub DumpAtHwnd(ByVal targetHwnd As LongPtr)
    Dim rc As RECT
    GetWindowRect targetHwnd, rc
    Dim cx As Long, cy As Long
    cx = (rc.Left + rc.Right) \ 2
    cy = (rc.Top + rc.Bottom) \ 2
    SetCursorPos cx, cy
    Debug.Print "=== Cursor moved to center of hwnd 0x" & Hex(targetHwnd) & _
                " at (" & cx & "," & cy & ") ==="
    DumpWindowHierarchyUnderCursor
End Sub

'==============================================================================
' カーソル下のウィンドウ階層を調べる
'==============================================================================
Public Sub DumpWindowHierarchyUnderCursor()
    Dim pt As POINTAPI
    GetCursorPos pt
    Debug.Print "Cursor: (" & pt.x & ", " & pt.y & ")"
    Debug.Print String(60, "-")

    Dim h As LongPtr
    h = WindowFromPoint(pt.x, pt.y)

    Dim level As Long: level = 0
    Do While h <> 0
        Debug.Print Space(level * 2) & "[" & level & "] hwnd=0x" & Hex(h) & _
                    " cls=" & GetCls(h) & _
                    " cap=""" & GetCap(h) & """" & _
                    " wndproc=0x" & Hex(GetWindowLongPtrA(h, GWLP_WNDPROC))
        Dim parent As LongPtr
        parent = GetParent(h)
        If parent = 0 Then Exit Do
        h = parent
        level = level + 1
        If level > 10 Then Exit Do
    Loop

    Debug.Print String(60, "-")
    Debug.Print "GA_ROOT:      0x" & Hex(GetAncestor(WindowFromPoint(pt.x, pt.y), GA_ROOT))
    Debug.Print "GA_ROOTOWNER: 0x" & Hex(GetAncestor(WindowFromPoint(pt.x, pt.y), GA_ROOTOWNER))
End Sub

Private Function GetCls(ByVal h As LongPtr) As String
    Dim buf As String
    buf = String(256, Chr(0))
    Dim n As Long: n = GetClassNameA(h, buf, 256)
    GetCls = Left$(buf, n)
End Function

Private Function GetCap(ByVal h As LongPtr) As String
    Dim buf As String
    buf = String(256, Chr(0))
    Dim n As Long: n = GetWindowTextA(h, buf, 256)
    GetCap = Left$(buf, n)
End Function

'==============================================================================
' 使い方:
'   1. UserForm を表示
'   2. マウスカーソルを UserForm 上のコントロール (TextBox等) の上に置く
'   3. Alt+F11 で VBE を表示
'   4. イミディエイトで DumpWindowHierarchyUnderCursor を実行
'      (但しマウスを動かす前に実行する必要があるので VBE 側から時間差で呼ぶ)
'
'   もしくは: UserForm の KeyDown などからこの Sub を呼ぶ
'==============================================================================


