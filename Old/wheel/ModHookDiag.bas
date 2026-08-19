Attribute VB_Name = "ModHookDiag"
'==============================================================================
' ModHookDiag.bas
'   EbModeAnchor スタブのバイト列をダンプして、EbMode 参照命令の実際の
'   位置を特定するための診断ユーティリティ。
'==============================================================================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub RtlMoveMemory Lib "kernel32" ( _
        ByVal Destination As LongPtr, ByVal Source As LongPtr, _
        ByVal Length As LongPtr)
#Else
    Private Enum LongPtr: [_]: End Enum
    Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
        ByVal Destination As Long, ByVal Source As Long, _
        ByVal Length As Long)
#End If

'==============================================================================
' 使い方:
'   Immediate ウィンドウで以下を実行:
'     DumpEbModeAnchor
'   → EbModeAnchor スタブの先頭から 128 バイトを 16 バイトずつ表示。
'
'   出力結果をそのままコピーして共有してください。
'==============================================================================
Public Sub DumpEbModeAnchor()
    Dim pAnchor As LongPtr
    pAnchor = VBA.Int(AddressOf ModHookSupport.EbModeAnchor)

    Debug.Print "EbModeAnchor @ " & PtrToHex(pAnchor)
    Debug.Print String(60, "-")

    Dim buf(0 To 127) As Byte
    RtlMoveMemory VarPtr(buf(0)), pAnchor, 128

    Dim i As Long, row As String, ascii As String
    For i = 0 To 127
        If i Mod 16 = 0 Then
            If i > 0 Then
                Debug.Print Right$("  " & Hex(i - 16), 4) & ": " & row & "  " & ascii
            End If
            row = ""
            ascii = ""
        End If
        row = row & Right$("0" & Hex(buf(i)), 2) & " "
        If buf(i) >= 32 And buf(i) < 127 Then
            ascii = ascii & Chr(buf(i))
        Else
            ascii = ascii & "."
        End If
    Next i
    Debug.Print Right$("  " & Hex(112), 4) & ": " & row & "  " & ascii
End Sub

'==============================================================================
' EbMode 参照候補となる MOV 命令を走査して列挙する。
' 正規の EbModeAnchor は本体が空なので、現れる MOV 命令のほとんどは
' EbMode/EbModeTop 系への参照のはず。
'==============================================================================
Public Sub ScanMovInstructions()
    Dim pAnchor As LongPtr
    pAnchor = VBA.Int(AddressOf ModHookSupport.EbModeAnchor)

    Dim buf(0 To 255) As Byte
    RtlMoveMemory VarPtr(buf(0)), pAnchor, 256

    Debug.Print "Scanning MOV [mem] patterns in EbModeAnchor..."
    Debug.Print String(60, "-")

    Dim i As Long
    i = 0
    Do While i < 250
        ' 48 8B 05 xx xx xx xx  MOV RAX, [RIP+disp32]
        If buf(i) = &H48 And buf(i + 1) = &H8B And buf(i + 2) = &H5 Then
            Dim disp As Long
            RtlMoveMemory VarPtr(disp), pAnchor + i + 3, 4
            Dim target As LongPtr: target = (pAnchor + i + 7) + disp
            Debug.Print "+" & i & ": 48 8B 05 (MOV RAX,[RIP+disp32]) -> " & PtrToHex(target)
            i = i + 7
        ' 48 A1 xx..xx (8 bytes)  MOV RAX, [abs64]
        ElseIf buf(i) = &H48 And buf(i + 1) = &HA1 Then
            Dim abs64 As LongPtr
            RtlMoveMemory VarPtr(abs64), pAnchor + i + 2, 8
            Debug.Print "+" & i & ": 48 A1 (MOV RAX,[abs64]) -> " & PtrToHex(abs64)
            i = i + 10
        ' 8B 05 xx xx xx xx  MOV EAX, [RIP+disp32]
        ElseIf buf(i) = &H8B And buf(i + 1) = &H5 Then
            Dim disp2 As Long
            RtlMoveMemory VarPtr(disp2), pAnchor + i + 2, 4
            Dim t2 As LongPtr: t2 = (pAnchor + i + 6) + disp2
            Debug.Print "+" & i & ": 8B 05 (MOV EAX,[RIP+disp32]) -> " & PtrToHex(t2)
            i = i + 6
        ' C3 RET  → スタブ終端 (probably)
        ElseIf buf(i) = &HC3 Then
            Debug.Print "+" & i & ": C3 (RET)"
            Exit Do
        Else
            i = i + 1
        End If
    Loop
End Sub

Private Function PtrToHex(ByVal p As LongPtr) As String
#If Win64 Then
    PtrToHex = "0x" & Right$(String(16, "0") & Hex(p), 16)
#Else
    PtrToHex = "0x" & Right$(String(8, "0") & Hex(p), 8)
#End If
End Function



