VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SafeTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


#Const Windows = (Mac = 0)
#If Windows Then
    #If VBA7 Then
        Private Declare PtrSafe Function SetTimer Lib "user32" (ByVal hwnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
    #Else
        Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
    #End If
#End If

Private m_form As New TimerForm
Private m_uElapse As Long
Private m_data As Variant

Public Event TimerCall(ByVal wTime As Long, ByRef varData As Variant)

Public Sub StartTimer(ByVal uElapse As Long, Optional ByRef varData As Variant)
    If uElapse = m_uElapse Then Exit Sub
    Set m_form = Nothing 'Kills existing timer, if any
    If uElapse = 0 Then Exit Sub
    '
    m_uElapse = uElapse
    If IsObject(varData) Then Set m_data = varData Else m_data = varData
#If Windows Then
    SetTimer m_form.Controls.Add("Forms.Frame.1").[_GethWnd], ObjPtr(Me) _
           , uElapse, GetTimerProc(Me)
#End If
End Sub

'Do not move. This must be the 8th position in virtual table
Private Sub TimerProc(ByVal wMsg As Long, ByVal nIDEvent As Long, ByVal wTime As Long)
    On Error Resume Next
    RaiseEvent TimerCall(wTime, m_data)
    On Error GoTo 0
End Sub

