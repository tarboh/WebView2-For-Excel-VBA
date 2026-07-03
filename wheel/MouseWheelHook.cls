VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MouseWheelHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'==============================================================================
' .cls  (方針3対応版)
'
'   各インスタンスは BuildInstanceThunk で生成された自分専用サンクを持つ。
'   StartHook では最初に EnsureDispatcherInstalled を呼び、スタブ +37 の
'   パッチとディスパッチャ生成を (プロセス一度だけ) 行う。
'==============================================================================
Option Explicit

' Win64 定数が環境によって Empty 扱いされる問題を避けるため明示的に再定義
#Const IS_X64 = Win64

#If VBA7 Then
    Private Declare PtrSafe Function SetWindowsHookExA Lib "user32" ( _
        ByVal idHook As Long, ByVal lpfn As LongPtr, _
        ByVal hMod As LongPtr, ByVal dwThreadId As Long) As LongPtr
    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
        ByVal hhk As LongPtr) As Long
    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
        ByVal hhk As LongPtr, ByVal nCode As Long, _
        ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
    Private Declare PtrSafe Function GetModuleHandleA Lib "kernel32" ( _
        ByVal lpModuleName As LongPtr) As LongPtr
#Else
    Private Enum LongPtr: [_]: End Enum
#End If

#If IS_X64 Then
    Private Const NullPtr As LongLong = 0^
#Else
    Private Const NullPtr As Long = 0&
#End If

Private Const WH_MOUSE_LL As Long = 14

Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const WM_MOUSEHWHEEL As Long = &H20E

Private Type MSLLHOOKSTRUCT
    pt_x As Long
    pt_y As Long
    mouseData As Long
    flags As Long
    time As Long
    dwExtraInfo As LongPtr
End Type

'==============================================================================
' VTable インデックス
'   ***MouseProc を最初の Public メソッドとして宣言すること***
'   IUnknown(3) + IDispatch(4) = 7 の次のスロット = 7
'==============================================================================
Private Enum VTableIndex
    MouseProc_Slot = 7
End Enum

Private m_hHook As LongPtr
Private m_pThunk As LongPtr
Private m_ThunkSize As LongPtr

Public Event WheelScroll(ByVal delta As Long, ByVal x As Long, ByVal y As Long, ByRef bHandled As Boolean)

'==============================================================================
' MouseProc (vtable 経由でディスパッチャから呼ばれる)
'==============================================================================
Public Function MouseProc(ByVal nCode As Long, _
                          ByVal wParam As LongPtr, _
                          ByVal lParam As LongPtr) As LongPtr
    On Error Resume Next

    If nCode < 0 Then
        MouseProc = CallNextHookEx(m_hHook, nCode, wParam, lParam)
        Exit Function
    End If

    If wParam = WM_MOUSEWHEEL Or wParam = WM_MOUSEHWHEEL Then
        Dim info As MSLLHOOKSTRUCT
        ModHookSupport.CopyMemory VarPtr(info), lParam, LenB(info)

        Dim delta As Long
        delta = (info.mouseData \ &H10000) And &HFFFF&
        If delta >= &H8000& Then delta = delta - &H10000

        Dim bHandled As Boolean: bHandled = False
        RaiseEvent WheelScroll(delta, info.pt_x, info.pt_y, bHandled)

        If bHandled Then
            MouseProc = 1
            Exit Function
        End If
    End If

    MouseProc = CallNextHookEx(m_hHook, nCode, wParam, lParam)
End Function

Public Sub StartHook()
    If m_hHook <> NullPtr Then Exit Sub

    ' ディスパッチャとスタブパッチは初回のみ
    ModHookSupport.EnsureDispatcherInstalled VTableIndex.MouseProc_Slot

    m_pThunk = ModHookSupport.BuildInstanceThunk(ObjPtr(Me), m_ThunkSize)

    Dim hMod As LongPtr
    hMod = GetModuleHandleA(NullPtr)
    m_hHook = SetWindowsHookExA(WH_MOUSE_LL, m_pThunk, hMod, 0)
    If m_hHook = NullPtr Then
        ModHookSupport.FreeThunk m_pThunk
        m_pThunk = NullPtr
        Err.Raise 7, "StartHook", "SetWindowsHookEx failed"
    End If
End Sub

Public Sub StopHook()
    If m_hHook <> NullPtr Then
        UnhookWindowsHookEx m_hHook
        m_hHook = NullPtr
    End If
    If m_pThunk <> NullPtr Then
        ModHookSupport.FreeThunk m_pThunk
        m_pThunk = NullPtr
    End If
End Sub

Private Sub Class_Terminate()
    StopHook
End Sub


