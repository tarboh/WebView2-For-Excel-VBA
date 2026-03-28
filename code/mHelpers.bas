Attribute VB_Name = "mHelpers"
Option Explicit


' FUNCDESC structure (Retrieved via ITypeInfo::GetFuncDesc)
Private Type FUNCDESC
    memid As Long
    lprgscode As LongPtr
    lprgelemdescParam As LongPtr
    funckind As Long      ' 0 = Virtual, 1 = PureVirtual...
    invkind As Long       ' 1 = Method, 2 = PropGet, 4 = PropPut
    callconv As Long      ' 4 = CC_STDCALL
    cParams As Integer
    cParamsOpt As Integer
    oVft As Integer       ' <-- VTable offset in bytes
    wReserved1 As Integer
    varkind As Long
    resW32 As Long
End Type

' TYPEATTR structure for ITypeInfo
Private Type TYPEATTR
    GUID(15) As Byte
    lcid As Long
    dwReserved As Long
    memidConstructor As Long
    memidDestructor As Long
    lpstrSchema As LongPtr
    cbSizeInstance As Long
    typekind As Long
    cFuncs As Integer     ' Number of functions in the interface
        
    ' NOTE: If you only need the function count (cFuncs) for GetClassMethodPtr,
    ' you can safely ignore the rest of the members and copy only up to cFuncs using CopyMemory.
'    cVars As Integer
'    cImplTypes As Integer
'    cbSizeVft As Integer
'    cbAlignment As Integer
'    wTypeFlags As Integer
'    wMajorVerNum As Integer
'    wMinorVerNum As Integer
'    tdescAlias As TYPEDESC
'    idldescType As IDLDESC
End Type

' GUID Structure (16 bytes)
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

' IID_IDispatch: {00020400-0000-0000-C000-000000000046}
Public Function IID_IDispatch() As GUID
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
End Function

Public Function IsIDispatchSupported(ByVal pUnk As IUnknown) As Boolean
    
    Dim ppUnk As LongPtr
    Dim iid As GUID
    Dim iidDisp As LongPtr
    Dim hr As Long
    Dim res As Variant
    
    ppUnk = ObjPtr(pUnk)
    iid = IID_IDispatch()
    
    ' QueryInterface for IDispatch from Object (IUnknown)
    ' IUnknown::QueryInterface is VTable Index 0
    ' It takes two parameters:
    ' [in]  REFIID riid      -> VarPtr(iid)
    ' [out] void **ppvObject -> VarPtr(iidDisp)
    
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
    
    args(0) = VarPtr(iid) ' Address of the GUID structure
    argTypes(0) = vbLongPtr
    
    args(1) = VarPtr(iidDisp)   ' Address of the pointer variable to receive output
    argTypes(1) = vbLongPtr
    
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))
    
    ' Call Index 0 (QueryInterface)
    hr = DispCallFunc(ppUnk, 0, CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then ' S_OK
        If res = 0 Then
            IsIDispatchSupported = True
            ' QueryInterface increments the reference count, so we must Release it (also via DispCallFunc)
            ' Call Release(pDisp)
            hr = DispCallFunc(ppUnk, 2 * LenB(ppUnk), CC_STDCALL, vbLong, 0, argTypes(0), argPtrs(0), res)
            If hr = 0 Then
                ' res contains the reference count after Release. Can be safely ignored.
                Debug.Print "Release succeeded. Remaining reference count: " & res
            Else
                Debug.Print "Release failed. hr: " & hr
            End If
        Else
            IsIDispatchSupported = False
        End If
    Else
        IsIDispatchSupported = False
    End If
End Function

'DCF_GetStringNoArgs
Public Function DCF_GetStringNoArgs(pObj As LongPtr, vIndex As Long, strFuncName As String) As String
    Dim hr As Long, res As Variant, pStr As LongPtr
    Dim args(0) As Variant: args(0) = VarPtr(pStr)
    Dim argTypes(0) As Integer: argTypes(0) = vbLongPtr
    Dim argPtrs(0) As LongPtr: argPtrs(0) = VarPtr(args(0))
    
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            If pStr <> 0 Then
                DCF_GetStringNoArgs = PtrToStrW(pStr)
                Call CoTaskMemFree(pStr)
            Else
                ' Succeeded, but failed to retrieve the string pointer
                Debug.Print strFuncName & " succeeded, but failed to retrieve pStr. pStr: " & pStr
            End If
        Else
            Debug.Print strFuncName & " failed. res: " & res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
    End If
End Function

Public Function DCF_OneArgString(pObj As LongPtr, vIndex As Long, strFuncName As String, ByVal str As String) As Long
    ' Use StrPtr to pass the Unicode string pointer (LPCWSTR) to the COM method
    Dim nArgs(0) As Variant: nArgs(0) = StrPtr(str)
    Dim nTypes(0) As Integer: nTypes(0) = vbLongPtr
    Dim nPtrs(0) As LongPtr: nPtrs(0) = VarPtr(nArgs(0))
    Dim hr As Long, res As Variant
    
    hr = DispCallFunc(pObj, vIndex * LenB(pObj), CC_STDCALL, vbLong, 1, nTypes(0), nPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            DCF_OneArgString = res
        Else
            Debug.Print strFuncName & " failed. res: " & res
            DCF_OneArgString = res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
        DCF_OneArgString = hr
    End If
End Function

Public Function DCF_TwoArgsStringAndObject(pObj As LongPtr, vIndex As Long, strFuncName As String, str As String, obj As Object) As Long
    Dim hr As Long
    Dim res As Variant
    Dim vObj As Variant

    ' 1. Store the instance into a VARIANT type to wrap it as a formal IDispatch interface
    Set vObj = obj

    ' 2. Prepare arguments for DispCallFunc
    Dim args(1) As Variant
    Dim argTypes(1) As Integer
    Dim argPtrs(1) As LongPtr
 
    ' Arg 1: Unicode string pointer (LPCWSTR)
    args(0) = StrPtr(str)
    argTypes(0) = vbLongPtr ' 64-bit: vbLongLong(20), 32-bit: vbLong(3)
   
    ' Arg 2: Pointer to the VARIANT structure (VARIANT*)
    args(1) = VarPtr(vObj)
    argTypes(1) = vbLongPtr

    ' Build the pointer array of argument pointers
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(pObj, _
                      vIndex * LenB(pObj), _
                      CC_STDCALL, _
                      vbLong, _
                      2, _
                      argTypes(0), _
                      argPtrs(0), _
                      res)

    ' 3. Evaluate results
    If hr = S_OK Then
        If res = S_OK Then
            DCF_TwoArgsStringAndObject = res
        Else
            Debug.Print strFuncName & " failed. res: " & res
            DCF_TwoArgsStringAndObject = res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
        DCF_TwoArgsStringAndObject = hr
    End If
End Function

Public Function DCF_RegisterHandler(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, funcPtr As LongPtr, Optional HandlerName As String) As Long
    
    Dim Handler As c4_Handler: Set Handler = New c4_Handler
    Handler.CreateVTble funcPtr, WB2.ppWebView2
    WB2.Col_Handler.Add Handler
    Handler.Namae = HandlerName ' (Rename c4_Handler.Namae to HandlerName if refactoring properties)
    
    Dim pObj As LongPtr, token As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    
    args(0) = Handler.Pointer
    args(1) = VarPtr(token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))
    
    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            Handler.token = token
            RegisterInstance Handler.Pointer, WB2
            DCF_RegisterHandler = res
        Else
            Debug.Print strFuncName & " failed. res: " & res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
    End If
End Function

Public Function DCF_RegisterHandler2(WB2 As c3_WebView2, pObj As LongPtr, vTblIndex As Long, strFuncName As String, funcPtr As LongPtr, Optional HandlerName As String) As Long
    
    Dim Handler As c4_Handler: Set Handler = New c4_Handler
    Handler.CreateVTble funcPtr, pObj
    
    Debug.Print "CreateVTble Complete."
    Debug.Print "ParentPtr:" & Handler.ParentPtr
    Debug.Print "Pointer  :" & Handler.Pointer
    
    WB2.Col_Handler.Add Handler
    Handler.Namae = HandlerName ' (Rename c4_Handler.Namae to HandlerName if refactoring properties)
    m_InstanceMap.Add CStr(Handler.Pointer), WB2
    'RegisterInstance Handler.Pointer, WB2
    
    Dim token As LongPtr, hr As Long, res As Variant
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    
    args(0) = Handler.Pointer
    args(1) = VarPtr(token)
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))
    
    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            Handler.token = token
            On Error GoTo 0
            Debug.Print "a"
            Debug.Print "Exists:" & m_InstanceMap.Exists(CStr(Handler.Pointer))
            'RegisterInstance Handler.Pointer, WB2
            
            Debug.Print "Handler.Pointer" & Handler.Pointer
            Debug.Print "Exists:" & m_InstanceMap.Exists(CStr(Handler.Pointer))
            Debug.Print "b"
            DCF_RegisterHandler2 = res
        Else
            Debug.Print strFuncName & " failed. res: " & res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
    End If
End Function


Public Function DCF_RegisterHandlerWithString(WB2 As c3_WebView2, vTblIndex As Long, strFuncName As String, str As String, funcPtr As LongPtr) As Long
    
    Dim Handler As c4_Handler: Set Handler = New c4_Handler
    Handler.CreateVTble funcPtr, WB2.ppWebView2
    WB2.Col_Handler.Add Handler
    RegisterInstance Handler.Pointer, WB2

    Dim pObj As LongPtr, hr As Long, res As Variant
    pObj = WB2.ppWebView2
    Dim args(1) As Variant, argTypes(1) As Integer, argPtrs(1) As LongPtr
    
    ' Arg 1: Pointer to the String (LPCWSTR)
    args(0) = StrPtr(str)
    ' Arg 2: Pointer to the custom VTable (Handler)
    args(1) = Handler.Pointer
    
    argTypes(0) = vbLongPtr: argTypes(1) = vbLongPtr
    argPtrs(0) = VarPtr(args(0)): argPtrs(1) = VarPtr(args(1))

    hr = DispCallFunc(pObj, vTblIndex * LenB(pObj), CC_STDCALL, vbLong, 2, argTypes(0), argPtrs(0), res)
    
    If hr = 0 Then
        If res = 0 Then
            DCF_RegisterHandlerWithString = res
        Else
            Debug.Print strFuncName & " failed. res: " & res
        End If
    Else
        Debug.Print strFuncName & " DispCallFunc failed. hr: " & hr
    End If
End Function

Public Function remove_Handler(WebView2 As c3_WebView2, vTblIndex As Long, token As Long, HandlerName As String) As Long
    Dim ppWebView2 As LongPtr: ppWebView2 = WebView2.ppWebView2
    Dim hr As Long, res As Variant
    Dim vtoken As Variant: vtoken = token
    Dim args(0) As Variant: args(0) = vtoken
    Dim argsPtr(0) As Variant: argsPtr(0) = VarPtr(args(0))
    Dim argsType(0) As Integer: argsType(0) = vbLong
    hr = DispCallFunc(ppWebView2, vTblIndex * LenB(ppWebView2), CC_STDCALL, vbLong, 1, argsType(0), argsPtr(0), res)
    If hr = 0 Then
        If res = 0 Then
            remove_Handler = res
        Else
            Debug.Print HandlerName & " Failed. res:" & res
            remove_Handler = res
        End If
    Else
        Debug.Print HandlerName & " Failed. hr:" & hr
        remove_Handler = hr
    End If
End Function

Public Function GetClassMethodPtr(ByVal TargetObj As Object, ByVal methodName As String, ByRef vTableOffset As Long) As LongPtr
    If TargetObj Is Nothing Then Exit Function

    Dim pDisp As LongPtr: pDisp = ObjPtr(TargetObj)
    Dim pTInfo As LongPtr
    Dim hr As Long, res As Variant
    
    ' 1. Retrieve ITypeInfo by calling IDispatch::GetTypeInfo (VTable Index 4)
    ' HRESULT GetTypeInfo([in] UINT iTInfo, [in] LCID lcid, [out] ITypeInfo** ppTInfo)
    Dim args(2) As Variant
    args(0) = 0&
    args(1) = 0&
    args(2) = VarPtr(pTInfo)
    
    Dim argTypes(2) As Integer
    argTypes(0) = vbLong
    argTypes(1) = vbLong
    argTypes(2) = vbLongPtr
    
    Dim argPtrs(2) As LongPtr
    argPtrs(0) = VarPtr(args(0))
    argPtrs(1) = VarPtr(args(1))
    argPtrs(2) = VarPtr(args(2))
    
    hr = DispCallFunc(pDisp, 4 * LenB(pDisp), CC_STDCALL, vbLong, 3, argTypes(0), argPtrs(0), res)
    
    If hr <> 0 Or res <> 0 Then Exit Function
    
    ' 2. Retrieve TYPEATTR by calling ITypeInfo::GetTypeAttr (VTable Index 3)
    ' HRESULT GetTypeAttr([out] TYPEATTR** ppTypeAttr)
    Dim pTypeAttr As LongPtr
    Dim args_t(0) As Variant, argsType_t(0) As Integer, argsPtr_t(0) As LongPtr
    args_t(0) = VarPtr(pTypeAttr)
    argsType_t(0) = vbLongPtr
    argsPtr_t(0) = VarPtr(args_t(0))
    
    hr = DispCallFunc(pTInfo, 3 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_t(0), argsPtr_t(0), res)
    
    If hr <> 0 Or res <> 0 Then GoTo CleanUp_ITypeInfo
    
    ' --- Start Scanning Functions ---
    Dim uTypeAttr As TYPEATTR
    CopyMemory uTypeAttr, ByVal pTypeAttr, LenB(uTypeAttr)
    
    Dim i As Long
    Dim pFuncDesc As LongPtr
    Dim uFuncDesc As FUNCDESC
    Dim bstrName As String
    Dim pBstr As LongPtr
    
    For i = 0 To uTypeAttr.cFuncs - 1
    
        ' STEP 1: GetFuncDesc (VTable Index 5)
        ' HRESULT GetFuncDesc([in] UINT index, [out] FUNCDESC** ppFuncDesc)
        Dim args_Gf(1) As Variant, argTypes_Gf(1) As Integer, argPtrs_Gf(1) As LongPtr
        args_Gf(0) = i: args_Gf(1) = VarPtr(pFuncDesc)
        argTypes_Gf(0) = vbLong: argTypes_Gf(1) = vbLongPtr
        argPtrs_Gf(0) = VarPtr(args_Gf(0)): argPtrs_Gf(1) = VarPtr(args_Gf(1))
        
        hr = DispCallFunc(pTInfo, 5 * LenB(pTInfo), CC_STDCALL, vbLong, 2, argTypes_Gf(0), argPtrs_Gf(0), res)
        
        If hr = 0 And res = 0 Then
            CopyMemory uFuncDesc, ByVal pFuncDesc, LenB(uFuncDesc)
            
            ' STEP 2: GetDocumentation (VTable Index 12)
            ' HRESULT GetDocumentation([in] MEMBERID memid, [out] BSTR* pbstrName, ...)
            Dim args_Gd(4) As Variant, argTypes_Gd(4) As Integer, argPtrs_Gd(4) As LongPtr
            pBstr = 0
            
            args_Gd(0) = uFuncDesc.memid: argTypes_Gd(0) = vbLong
            args_Gd(1) = VarPtr(pBstr):    argTypes_Gd(1) = vbLongPtr
            
            ' Passing 0 (NULL pointer) directly to the stack for unused outputs
            args_Gd(2) = 0: argTypes_Gd(2) = vbLongPtr
            args_Gd(3) = 0: argTypes_Gd(3) = vbLongPtr
            args_Gd(4) = 0: argTypes_Gd(4) = vbLongPtr
            
            argPtrs_Gd(0) = VarPtr(args_Gd(0))
            argPtrs_Gd(1) = VarPtr(args_Gd(1))
            argPtrs_Gd(2) = VarPtr(args_Gd(2))
            argPtrs_Gd(3) = VarPtr(args_Gd(3))
            argPtrs_Gd(4) = VarPtr(args_Gd(4))
            
            hr = DispCallFunc(pTInfo, 12 * LenB(pTInfo), CC_STDCALL, vbLong, 5, argTypes_Gd(0), argPtrs_Gd(0), res)
                
            If hr = 0 And res = 0 Then
                ' Inject the BSTR pointer directly into the VBA String variable (it automatically takes ownership)
                CopyMemory ByVal VarPtr(bstrName), pBstr, LenB(pBstr)
                
                ' STEP 3: Compare names and calculate the VTable pointer
                If LCase$(bstrName) = LCase$(methodName) Then
                    Dim pVTable As LongPtr, pRealAddr As LongPtr
                    CopyMemory pVTable, ByVal pDisp, LenB(pVTable)
                    CopyMemory pRealAddr, ByVal (pVTable + uFuncDesc.oVft), LenB(pRealAddr)
                    
                    GetClassMethodPtr = pRealAddr
                    vTableOffset = uFuncDesc.oVft
                End If
            End If
            
            ' STEP 4: Release FUNCDESC memory (VTable Index 20)
            ' HRESULT ReleaseFuncDesc([in] FUNCDESC* pFuncDesc)
            Dim args_Rf(0) As Variant, argTypes_Rf(0) As Integer, argPtrs_Rf(0) As LongPtr
            args_Rf(0) = pFuncDesc: argTypes_Rf(0) = vbLongPtr: argPtrs_Rf(0) = VarPtr(args_Rf(0))
            Call DispCallFunc(pTInfo, 20 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argTypes_Rf(0), argPtrs_Rf(0), res)
        End If
        
        If GetClassMethodPtr <> 0 Then Exit For
    Next i
    
    ' CleanUp Part 1: Release TYPEATTR (ITypeInfo VTable Index 19)
    Dim args_r(0) As Variant, argsType_r(0) As Integer, argsPtr_r(0) As LongPtr
    args_r(0) = pTypeAttr
    argsType_r(0) = vbLongPtr
    argsPtr_r(0) = VarPtr(args_r(0))

    hr = DispCallFunc(pTInfo, 19 * LenB(pTInfo), CC_STDCALL, vbLong, 1, argsType_r(0), argsPtr_r(0), res)
    
CleanUp_ITypeInfo:
    ' CleanUp Part 2: Release ITypeInfo (IUnknown VTable Index 2)
    hr = DispCallFunc(pTInfo, 2 * LenB(pTInfo), CC_STDCALL, vbLong, 0, 0, 0, res)

End Function

Public Function dcf(ptr As LongPtr, vTblIndex As Long, funcName As String, ParamArray args() As Variant) As Long
    Debug.Print "dcf called for " & funcName
    Dim l As Long: l = LBound(args)
    Dim u As Long: u = UBound(args)
    Dim cnt As Long: cnt = u - l + 1
    Dim hr As Long, res As Variant
    Dim args_Type() As Integer
    Dim args_Ptr() As LongPtr
    If cnt > 0 Then
        ReDim args_Type(l To u): ReDim args_Ptr(l To u)
        Dim i As Long
        For i = l To u
            args_Type(i) = VarType(args(i))
            args_Ptr(i) = VarPtr(args(i))
            Debug.Print "args(" & i & ")", "Type:" & args_Type(i), "Value:" & args(i)
        Next
        hr = DispCallFunc(ptr, vTblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, args_Type(l), args_Ptr(l), res)
    Else
        hr = DispCallFunc(ptr, vTblIndex * LenB(ptr), CC_STDCALL, vbLong, cnt, 0, 0, res)
    End If
    If hr = 0 Then
        If res <> 0 Then
            Debug.Print funcName & " failed. res:" & res
        End If
        dcf = res
    Else
        Debug.Print funcName & " failed. hr:" & hr
        dcf = hr
    End If
End Function

Sub tttttttt()
    Call dcf(1, 1, "a")
End Sub

''' <summary>
''' Recursively creates directories if they do not exist, including deep subfolders.
''' </summary>
''' <param name="folderPath">The absolute path of the directory to create.</param>
Public Sub CreateDeepFolder(ByVal folderPath As String)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Do nothing if the folder already exists
    If fso.FolderExists(folderPath) Then Exit Sub
    
    ' Get the parent folder path
    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)
    
    ' If the parent folder does not exist, recursively call itself to create the parent first
    If Not fso.FolderExists(parentPath) Then
        CreateDeepFolder parentPath
    End If
    
    ' Finally, create the target folder
    fso.CreateFolder folderPath
    Debug.Print "Deep folder created: " & folderPath
    
    Set fso = Nothing
End Sub


''' <summary>
''' Parses a JSON string into a Scripting.Object using JScript runtime.
''' </summary>
Public Function ParseJSON(ByVal jsonString As String) As Object
    Dim html As Object
    Set html = CreateObject("htmlfile")
    
    ' Leverage htmlfile's JavaScript space to parse JSON
    html.parentWindow.execScript "function parse(json) { return JSON.parse(json); }", "JScript"
    
    ' Returns as a VBA Object (usable like a Dictionary)
    Set ParseJSON = html.parentWindow.Parse(jsonString)
End Function

''' <summary>
''' Decodes a Base64 string and converts it into a binary Byte array.
''' </summary>
Public Function Base64Decode(ByVal base64Str As String) As Byte()
    Dim byteLen As Long
    Dim skip As Long, outFlags As Long
    
    ' Calculate the required buffer size
    If CryptStringToBinary(StrPtr(base64Str), Len(base64Str), CRYPT_STRING_BASE64, 0, byteLen, skip, outFlags) = 0 Then
        Err.Raise vbObjectError + 1001, , "Failed to calculate Base64 buffer size."
    End If
    
    Dim bytes() As Byte
    ReDim bytes(byteLen - 1)
    
    ' Perform the actual Base64 decoding
    If CryptStringToBinary(StrPtr(base64Str), Len(base64Str), CRYPT_STRING_BASE64, VarPtr(bytes(0)), byteLen, skip, outFlags) = 0 Then
        Err.Raise vbObjectError + 1002, , "Failed to decode Base64 string."
    End If
    
    Base64Decode = bytes
End Function

''' <summary>
''' Saves a Byte array to a physical file (No ADODB.Stream reference required).
''' </summary>
Public Sub SaveBytesToFile(ByRef bytes() As Byte, ByVal filePath As String)
    Dim stream As Object
    Set stream = CreateObject("ADODB.Stream")
    
    With stream
        .Type = 1 ' adTypeBinary
        .Open
        .Write bytes
        .SaveToFile filePath, 2 ' adSaveCreateOverWrite
        .Close
    End With
End Sub
