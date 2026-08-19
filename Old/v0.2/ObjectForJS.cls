VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ObjectForJS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' JavaScript-callable method: Multiplies the passed value by 10 and returns it
Function Func1(ByVal value As Long) As Long
    Debug.Print "ObjectForJS.Func1 was called."
    Func1 = value * 10
End Function

