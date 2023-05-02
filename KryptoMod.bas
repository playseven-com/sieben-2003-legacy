Attribute VB_Name = "KryptoMod"
Option Explicit
Option Compare Text
Private Const Seperator As String = ", "

Function Encrypt(str As String, En As Boolean) As String
Dim StrLen As Long, StrRet As String, i As Integer, Var As Integer, Var2 As Integer, Zahl As Integer
Dim LastI As Integer, ENDE As Boolean, ii As Integer, codeB As String

On Error GoTo ERRHand
    ENDE = False
    LastI = 1
    StrRet = ""
    i = 0
    StrLen = Len(str)
    
    If Not En Then
        StrLen = StrLen \ 2
    End If
    
    If StrLen = ZERO Then Exit Function
    Var = (Int((StrLen / 11) * 9) Mod 5) + Sin(StrLen) * 9
    For i = ZERO To StrLen - 1
        Var2 = Cos(i) * 9 + Sin(i * 1.1) * 3
        'Debug.Print Var2
        If En Then
            Zahl = Asc(Mid$(str, i + 1, 1))
        Else
            Zahl = CInt(Hex2Long(Mid$(str, i * 2 + 1, 2)) Xor Abs(Var * Var2))
        End If
        If En Then
            codeB = Abs(Zahl - Var + Var2)
            StrRet = StrRet & Long2Hex(codeB Xor Abs(Var * Var2))
        Else
            StrRet = StrRet & Chr(Abs(Zahl - Var2 + Var))
        End If
        'Debug.Print Var, Var2
    Next
    Encrypt = StrRet
Exit Function
ERRHand:
If ErrorBox("888888" & En, Err) Then Resume Next
End Function

Private Function Long2Hex(ByVal l As Long) As String
'wandelt eine 4-Byte-Zahl in einen 8-Zeichen-HexDigitString
On Error GoTo ERRHand

Long2Hex = UCase$(Right$(String$(2, 48) & Hex(l), 2))

Exit Function
ERRHand:
If ErrorBox("Hex2Long", Err) Then Resume Next
End Function

Private Function Hex2Long(ByVal s As String) As Long
'wandelt HexDigit-String zu Long
On Error GoTo ERRHand

  Hex2Long = val("&h" & s)

Exit Function
ERRHand:
If ErrorBox("Hex2Long", Err) Then Resume Next
End Function

'Public Function XORstring2Hex(ByVal s1 As String, ByVal s2 As String) As String
''verknüpft zwei Strings gleicher(!) Länge
'Dim i As Long, le As Long
'Dim ret As String
'le = Len(s1)
'
'For i = 1 To le
'    ret = ret & Long2Hex(Asc(Mid$(s1, i, 1)) Xor Asc(Mid$(s2, i, 1)))
'Next
'XORstring2Hex = ret
'End Function




