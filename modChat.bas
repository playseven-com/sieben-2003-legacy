Attribute VB_Name = "modChat"
Option Explicit
Private Const sexten_k As Integer = 16384
Public Names() As String


Public Sub UpdateChat(msgType As ChatMsgType, ByVal sString As String, CallerFrm As Form)
On Error GoTo ERRHand
    With CallerFrm
    
        'Autoscroll the text
        .rtxtChat.SelStart = Len(.rtxtChat.Text)
        Select Case msgType
            Case SystemMsg
                If .chkShowSystemMsg = vbChecked Then
                    .rtxtChat.SelColor = vbMagenta
                    .rtxtChat.SelText = Time$ & cstrSystemLbl & sString & vbCrLf
                End If
                
            Case Else
                .rtxtChat.SelColor = &HC0C0C0
                .rtxtChat.SelText = Time$
                
                If InStr(InStr(1, sString, gstrDblDot, vbTextCompare) + 1, sString, AktuellerSpieler.SpielerName, vbTextCompare) > 0 Then
                    .rtxtChat.SelColor = vbYellow
                ElseIf InStr(1, sString, AktuellerSpieler.SpielerName, vbTextCompare) > 0 Then
                    .rtxtChat.SelColor = GRUEN
                Else
                    .rtxtChat.SelColor = vbWhite
                End If
                
                PlaySound SMPlayerChoose
                
                .rtxtChat.SelText = gstrSpace & gstrMinus & gstrSpace & sString & vbCrLf
                
        End Select
        If .chkLog = vbChecked Then Write2Log sString
        'Now limit the text in the window to be 16k
        If Len(.rtxtChat.Text) > sexten_k Then
            .rtxtChat.Text = Right$(.rtxtChat.Text, sexten_k)
        End If
        'Autoscroll the text
        .rtxtChat.SelStart = Len(.rtxtChat.Text)
    End With
Exit Sub
ERRHand:
If ErrorBox("UpdateChat:" & CallerFrm.Name, Err) Then Resume Next
End Sub

Public Function AutoUpdateName(ByVal rtxt As String) As String
Dim pos As Integer, ii As Integer
Dim found As Boolean
Dim s As String

Static base As String, LastBase As String
Static i As Integer


'position ermitteln ab der Name zu vervollständigen ist
pos = InStrRev(rtxt, gstrSpace, , vbTextCompare) + 1

s = Mid$(rtxt, pos)
'suche initieren
If i = 0 Then

    base = Mid$(rtxt, pos)
    LastBase = base

    i = LBound(Names)
    
ElseIf s = LastBase Then

    i = i + 1
    
Else

    i = 0
    base = vbNullString
    LastBase = vbNullString
    AutoUpdateName = rtxt
    Exit Function
    
End If

'nach Namen suchen
For ii = i To UBound(Names)
    If UCase$(Left$(Names(ii), Len(base))) = UCase$(base) Then
        found = True
        Exit For
    End If
Next

If found Then
    AutoUpdateName = Left$(rtxt, pos - 1) & Replace(rtxt, LastBase, Names(ii), pos)
    i = ii
Else
    AutoUpdateName = Left$(rtxt, pos - 1) & Replace(rtxt, LastBase, base, pos, 1, vbTextCompare)
    i = 0
    base = vbNullString
End If

End Function
