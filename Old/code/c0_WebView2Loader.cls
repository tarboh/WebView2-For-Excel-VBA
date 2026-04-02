VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c0_WebView2Loader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Standard Class: c0_WebView2Loader
Option Explicit

' Note:
' One Environment is sufficient per application instance.
' Multiple Controller/WebView2 pairs can be spawned from a single Environment.

' Pointers to hold the dynamically created COM object (Handler).
' Required during Class_Terminate to free memory allocated via GlobalAlloc.
Private vTable(3) As LongPtr
Private pVTable As LongPtr
Private pObject As LongPtr

Public Sub DebugLoader()
    Dim hr As Long
    ' Test API by passing 0 (NULL) to the 4th argument.
    ' If Excel crashes here, the Declare statement (argument types) is likely incorrect.
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr("C:\Temp\VBA_WebView2"), 0, 0)
    Debug.Print "DebugLoader Result: " & hr
End Sub

' Initializes the WebView2 Environment
Public Sub CreateWebView2Environment()

    ' 1. Store function pointers in the vTable array
    vTable(0) = GetAddr(AddressOf Handler_QueryInterface)
    vTable(1) = GetAddr(AddressOf Handler_AddRef)
    vTable(2) = GetAddr(AddressOf Handler_Release)
    vTable(3) = GetAddr(AddressOf Handler_Invoke)

    ' 2. Allocate heap memory for the vTable
    pVTable = GlobalAlloc(GPTR, (UBound(vTable) + 1) * LenB(vTable(0)))
    CopyMemory ByVal pVTable, vTable(0), (UBound(vTable) + 1) * LenB(vTable(0))

    ' 3. Allocate heap memory for the Object instance itself (points to pVTable)
    pObject = GlobalAlloc(GPTR, LenB(pVTable))
    CopyMemory ByVal pObject, pVTable, LenB(pVTable)

    ' 4. Start creating the WebView2 Environment
    
    ' Prepare a writable user data folder path (Null-terminated Unicode string)
    Dim userDataPath As String
    userDataPath = "C:\Temp\VBA_WebView2" ' Folder will be created automatically if missing

    ' Ensure base directory exists
    If Dir("C:\Temp", vbDirectory) = "" Then MkDir "C:\Temp"

    ' Call API (pass the string pointer to the 2nd argument, and custom object to the 4th)
    Dim hr As Long
    hr = CreateCoreWebView2EnvironmentWithOptions(0, StrPtr(userDataPath), 0, pObject)

    If hr <> S_OK Then
        MsgBox "Failed to initialize WebView2 Environment. HRESULT: " & hr & vbCrLf & _
               "Ensure WebView2Loader.dll is in the same directory as this workbook.", vbCritical, "WebView2 Initialization Error"
    End If

    DoEvents

End Sub

