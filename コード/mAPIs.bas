Attribute VB_Name = "mAPIs"
Option Explicit

' WebView2Loader.dll entry point (Called from c0.CreateWebView2Environment)
Public Declare PtrSafe Function CreateCoreWebView2EnvironmentWithOptions Lib "WebView2Loader.dll" ( _
    ByVal browserExecutableFolder As LongPtr, _
    ByVal userDataFolder As LongPtr, _
    ByVal additionalBrowserArguments As LongPtr, _
    ByVal environmentCreatedHandler As LongPtr) As Long

' Declaration to allocate a custom COM object (Handler) in memory
Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As LongPtr)
Public Declare PtrSafe Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Public Declare PtrSafe Function GlobalFree Lib "kernel32" (ByVal hMem As LongPtr) As LongPtr
Public Const GPTR As Long = &H40 ' Flag for GlobalAlloc uFlags: Allocates fixed and zero-initialized memory
Public Const S_OK As Long = 0
Public Declare PtrSafe Sub CoTaskMemFree Lib "ole32.dll" (ByVal pv As LongPtr)

#If Win64 Then
    Public Const vbLongPtr = vbLongLong
#Else
    Public Const vbLongPtr = vbLong
#End If

' Gets the number of characters in a string (Unicode/UTF-16)
Public Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Public Declare PtrSafe Function SysAllocString Lib "oleaut32.dll" (ByVal pOleChar As LongPtr) As LongPtr
' Releases the BSTR memory
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

Public Const CC_STDCALL As Long = 4 ' Calling convention for the 3rd argument of DispCallFunc (__stdcall)

' Use in a Standard Module
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As LongPtr, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hwnd As LongPtr) As LongPtr
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hwnd As LongPtr, ByVal hdc As LongPtr) As Long

' Device capabilities indices for retrieving screen DPI (Used with GetDeviceCaps)
Public Const LOGPIXELSX As Long = 88 ' Logical pixels per inch (X-axis)
Public Const LOGPIXELSY As Long = 90 ' Logical pixels per inch (Y-axis)

Public Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" (ByVal Destination As LongPtr, ByRef Source As LongPtr, ByVal Length As LongPtr)


Public Sub PutMemPtr(ByVal pDest As LongPtr, ByVal pSrc As LongPtr)
    RtlMoveMemory pDest, pSrc, 8
End Sub
