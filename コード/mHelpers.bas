Attribute VB_Name = "mHelpers"
Option Explicit

'DispCallFuncˆø”–³‚µ‚Å•¶š—ñ‚ğæ“¾
Public Function DCF_ˆø”–³‚µ‚Å•¶š—ñ‚ğæ“¾(pObj As LongPtr, vIndex As Long, strFuncName As String) As String
    Dim hr As Long, res As Variant, pStr As LongPtr
    Dim args(0) As Variant: args(0) = VarPtr(pStr)
    Dim argTypes(0) As Integer: argTypes(0) = vbLongPtr
    Dim argPtrs(0) As LongPtr: argPtrs(0) = VarPtr(args(0))
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            If pStr <> 0 Then
                DCF_ˆø”–³‚µ‚Å•¶š—ñ‚ğæ“¾ = PtrToStrW(pStr)
                Call CoTaskMemFree(pStr)
            Else
                Debug.Print strFuncName & "¬Œ÷A‚µ‚©‚µpStræ“¾¸”sApStr:" & pStr
            End If
        Else
            Debug.Print strFuncName & "_¸”sAres:" & res
        End If
    Else
        Debug.Print strFuncName & "_dispcallfunc¸”sAhr:" & hr
    End If
End Function
Public Function DCF_ˆø”1‚Â‚Å•¶š—ñ‚ğ“n‚·(pObj As LongPtr, vIndex As Long, strFuncName As String, ByVal str As String) As Long
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(str)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    Dim hr As Long, res As Variant
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
    If hr = 0 Then
        If res = 0 Then
            DCF_ˆø”1‚Â‚Å•¶š—ñ‚ğ“n‚· = res
        Else
            Debug.Print strFuncName & "_¸”sAres:" & res
        End If
    Else
        Debug.Print strFuncName & "_dispcallfunc¸”sAhr:" & hr
    End If
End Function

