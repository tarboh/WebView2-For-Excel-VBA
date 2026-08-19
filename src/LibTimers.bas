Attribute VB_Name = "LibTimers"
Option Explicit

#Const Windows = (Mac = 0)
#Const x64 = Win64
#Const x32 = (x64 = 0)

#If VBA7 = 0 Then
    Public Enum LongPtr: [_]: End Enum
    Private Enum LONG_PTR: [_]: End Enum
#End If

#If x64 Then
    Private Const NullPtr As LongLong = 0^
    Private Const PtrSize = 8
#Else
    Private Const NullPtr As Long = 0&
    Private Const PtrSize = 4
#End If

Private Enum SAFEARRAY_FEATURES
    FADF_AUTO = &H1
    FADF_FIXEDSIZE = &H10
End Enum
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY_1D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As LongPtr
    rgsabound0 As SAFEARRAYBOUND
End Type

Private Type PointerAccessor
    arr() As LongPtr
    sa As SAFEARRAY_1D
End Type

Private Sub EntryPoint(): End Sub
Private Sub DummyASM(): End Sub 'Custom assembly bytes

Public Function GetTimerProc(ByVal st As SafeTimer) As LongPtr
    If st Is Nothing Then Exit Function
    Static pa As PointerAccessor
    Dim aPtr As LongPtr
    '
    If pa.sa.cDims = 0 Then
        pa.sa.cDims = 1
        pa.sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        pa.sa.cbElements = PtrSize
        pa.sa.cLocks = 1
        MemLongPtr(VarPtr(pa)) = VarPtr(pa.sa)
    End If
    '
    pa.sa.pvData = ObjPtr(st)
    pa.sa.rgsabound0.cElements = 1
    '
#If x32 Then
    pa.sa.pvData = pa.arr(0) + PtrSize * 8
    Dim tProcPtr As Long: tProcPtr = pa.arr(0) 'SafeTimer.TimerProc
#End If
    '
    'Note that VBA does the work for us:
    ' - memory is allocated and managed by VBA
    ' - Break mode is handled by VBA. 'EBMode' is found at:
    '   * EntryPoint+37 (x64)
    '   * EntryPoint+10 (x32)
    '   i.e. in Break mode (EBMode = 2), TimerProc call is skipped
    'We simply swap the nIDEvent argument from the 3rd to the 1st position
    ' so that the correct class instance is called
    GetTimerProc = VBA.Int(AddressOf EntryPoint)
    aPtr = VBA.Int(AddressOf DummyASM)
    pa.sa.pvData = aPtr
#If x64 Then
    If (pa.arr(0) And &HFFFFFF) <> &HC1894C Then
                                  pa.arr(0) = &HC1894C   '4C89C1   MOV RCX,R8              ;nIDEvent (instance)
        pa.sa.pvData = aPtr + 3:  pa.arr(0) = &H18B48    '488B01   MOV RAX,QWORD PTR [RCX] ;vtbl
        pa.sa.pvData = aPtr + 6:  pa.arr(0) = &H55&      '55       PUSH RBP
        pa.sa.pvData = aPtr + 7:  pa.arr(0) = &HEC8B48   '488BEC   MOV RBP,RSP
        pa.sa.pvData = aPtr + 10: pa.arr(0) = &H20EC8348 '4883EC20 SUB RSP,0x20
        pa.sa.pvData = aPtr + 14: pa.arr(0) = &H4050FF   'FF5040   CALL QWORD PTR [RAX+40] ;SafeTimer.TimerProc
        pa.sa.pvData = aPtr + 17: pa.arr(0) = &H20C48348 '4883C420 ADD RSP,0x20
        pa.sa.pvData = aPtr + 21: pa.arr(0) = &H5D&      '5D       POP RBP
        pa.sa.pvData = aPtr + 22: pa.arr(0) = &HC3&      'C3       RET
    End If
    pa.sa.pvData = GetTimerProc + 55
#Else
    If pa.arr(0) <> &HC24448B Then
                                  pa.arr(0) = &HC24448B  '8B44240C MOV EAX,DWORD PTR [ESP+0C] ;nIDEvent (instance)
        pa.sa.pvData = aPtr + 4:  pa.arr(0) = &H4244489  '89442404 MOV DWORD PTR [ESP+04],EAX ;replace hWnd
        pa.sa.pvData = aPtr + 8:  pa.arr(0) = &HB8&      'B8       MOV EAX,...
        pa.sa.pvData = aPtr + 9:  pa.arr(0) = tProcPtr   '                                    ;SafeTimer.TimerProc
        pa.sa.pvData = aPtr + 13: pa.arr(0) = &HE0FF&    'FFE0     JMP EAX
        pa.sa.pvData = aPtr + 15: pa.arr(0) = &HE0FF&    '33C0     XOR EAX,EAX                ;Not needed / never reached
        pa.sa.pvData = aPtr + 17: pa.arr(0) = &H10C2&    'C21000   RET 0010                   ;Not needed / never reached
    End If
    pa.sa.pvData = GetTimerProc + 22
#End If
    pa.arr(0) = aPtr
    pa.sa.rgsabound0.cElements = 0
    pa.sa.pvData = NullPtr
End Function
Private Property Let MemLongPtr(ByVal addr As LongPtr, ByVal newValue As LongPtr)
    Dim pa(0 To 0) As PointerAccessor
    With pa(0)
        .sa.cDims = 1
        .sa.cLocks = 1
        .sa.fFeatures = FADF_AUTO Or FADF_FIXEDSIZE
        .sa.pvData = addr
        .sa.rgsabound0.cElements = 1
        WritePtrNatively pa, VarPtr(.sa)
        .arr(0) = newValue
        .sa.rgsabound0.cElements = 0
        .sa.pvData = NullPtr
    End With
End Property
'https://github.com/WNKLER/RefTypes/discussions/3#discussion-8595790
Private Sub WritePtrNatively(ByRef ptrs() As LONG_PTR, ByVal Ptr As LongPtr)
    ptrs(0) = Ptr
End Sub

