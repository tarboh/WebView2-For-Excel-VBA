VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8355.001
   OleObjectBlob   =   "UserForm1.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_Wheel As UserFormWheel
Private m_DiagDone As Boolean

Private Sub UserForm_Activate()
    Set m_Wheel = New UserFormWheel
    m_Wheel.Attach Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If Not m_Wheel Is Nothing Then m_Wheel.Detach
    Set m_Wheel = Nothing
End Sub

' テスト用 TextBox1 の上でマウス動かした時に一度だけ診断
Private Sub TextBox1_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, _
                                ByVal x As Single, ByVal y As Single)
    If m_DiagDone Then Exit Sub
    m_DiagDone = True
    DumpWindowHierarchyUnderCursor
End Sub

Private Sub UserForm_Click()
    ' フォーカスをUserForm自身に持っていく
    Debug.Print "UserForm clicked, hwnd=0x" & Hex(ModWheelDiag_GetHwnd(Me))
End Sub
