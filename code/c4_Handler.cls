VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c4_Handler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''Public Class c4_Handler
'Option Explicit
'
'Private vTable(3) As LongPtr
'Private pVTable As LongPtr
'Private pHandler As LongPtr
'Private m_CallbackAddr As LongPtr ' Invoke target address
'
'Public token As LongPtr
'Private m_ParentPtr As LongPtr ' ObjPtr of the parent (c2_WebView2)
'Public Namae As String
'
'' Initialization: Pass the function address to set up the environment
'Public Sub CreateVTble(ByVal funcAddr As LongPtr, ByVal ParentPtr As LongPtr)
'    m_ParentPtr = ParentPtr ' Store the parent address‚­
'    m_CallbackAddr = funcAddr
'
'    ' Set up the VTable (Virtual Function Table)
'    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
'    vTable(1) = GetAddr(AddressOf Handler_AddRef)
'    vTable(2) = GetAddr(AddressOf Handler_Release)
'    vTable(3) = m_CallbackAddr ' Target function for actual execution
'
'    ' Allocate memory
'    pVTable = GlobalAlloc(GPTR, 4 * LenB(pVTable))
'    pHandler = GlobalAlloc(GPTR, LenB(pVTable))
'
'    ' Map structures / Link pointers
'    CopyMemory ByVal pVTable, vTable(0), 4 * LenB(pVTable)
'    CopyMemory ByVal pHandler, pVTable, LenB(pVTable)
'End Sub
'
'' State reset handling: Update VTable with the latest function address
'Public Sub Refresh()
'    ' * AddressOf requires retrieving fresh pointers each time state resets.
'    ' Designed to be re-invoked from external callers.
'End Sub
'
'' Expose the pointer to be passed to WebView2
'Public Property Get Pointer() As LongPtr
'    Pointer = pHandler
'End Property
'
'Public Property Get ParentPtr() As LongPtr
'    ParentPtr = m_ParentPtr
'End Property
'
'' Clean up memory on class termination
'Private Sub Class_Terminate()
'    If pHandler <> 0 Then GlobalFree pHandler
'    If pVTable <> 0 Then GlobalFree pVTable
'    Debug.Print "c4_Handler Memory Released."
'End Sub
