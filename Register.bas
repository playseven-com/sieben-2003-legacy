Attribute VB_Name = "Register"
Option Explicit
Option Compare Text

Private Const cstrRegistration = "Registration"
Public MyLic As New License

Public Const SEC_LICENSE As String = "License"
Public Const KEY_LICCODE As String = "LicCode"
Public Const KEY_LICCODE2 As String = "LicCode2"
Public Const KontoVerbindung = "Milosz Weckowski" & vbCrLf & "KtoNr: 730271943" & vbCrLf & "BLZ  : 10050000" & vbCrLf & "Berliner Sparkasse" & vbCrLf & "IBAN: DE26 1005 0000 0730 2719 43" & vbCrLf & "SWIFT: BELADEBE"
Public Const Zahlen = "7B093D1A183A72012E77682D0905331A683A0E1D2C1C6924050B161D63082E021B22"

Public Sub GetRegistry_IDs()

    '#############ClientID##############
    'ClientID aus DB
    AktuellerSpieler.ClientID = GetFromReg(cstrClientID)

    'wenn hier nix drin
    'alte ClientID aus Registry holen
    'wenn prog schon mal da war
    If AktuellerSpieler.ClientID = vbNullString Then
        AktuellerSpieler.ClientID = GetSetting(AppExeName, cstrDefault, cstrClientID)
    End If
    
    'wenn auch hier nix drin ist, dann ist es ein Erststart
    'und wir generieren eine neue
    If AktuellerSpieler.ClientID = vbNullString Then
        AktuellerSpieler.ClientID = GetGUID()
        SaveSetting AppExeName, cstrDefault, cstrClientID, AktuellerSpieler.ClientID
    End If
    'ermittelte ClientID nochmal in die Db schreiben
    Write2Reg cstrClientID, AktuellerSpieler.ClientID
    
    
    'RegCode aus der DB holen
    AktuellerSpieler.RegID = Encrypt(GetFromReg(KEY_LICCODE), False)
    
    'wenn unregistriert
    If AktuellerSpieler.RegID = vbNullString Then
        'nach gesichertem RegCode in Registry schauen
        AktuellerSpieler.RegID = GetSetting(AppExeName, cstrDefault, cstrRegID, cstrRegID)
        If AktuellerSpieler.RegID <> cstrRegID Then
            'wenn einer da war in die DB schreiben
            Write2Reg KEY_LICCODE, Encrypt(AktuellerSpieler.RegID, True)
        End If
    Else
        SaveSetting AppExeName, cstrDefault, cstrRegID, AktuellerSpieler.RegID
    End If
    
End Sub


Public Function IsRegistered() As Boolean

On Error GoTo ERRHand

GetRegistry_IDs
With MyLic
    .Init
    IsRegistered = .License_OK(AktuellerSpieler.RegID)
End With

bool_isRegistered = IsRegistered

Exit Function
ERRHand:
If ErrorBox("777777", Err) Then Resume Next
End Function

Public Function IsRegisteredFree() As Boolean
Dim s As String, s2 As String

On Error GoTo ERRHand

If Not bool_isRegistered Then
    s = GetSetting(AppExeName, cstrOptions, cstrRegistration)
    If s = gstrNullstr Then Exit Function
    s2 = Encrypt(Zahlen, False)
    If Len(s) <> Len(s2) Then Exit Function
    If s = s2 Then
        IsRegisteredFree = True
        'AktuellerSpieler.RegId = s
    End If
Else
    IsRegisteredFree = True
End If
Exit Function
ERRHand:
If ErrorBox("7777778", Err) Then Resume Next
End Function

Function chkRegister() As Boolean

On Error GoTo ERRHand
    If StartAnz >= MaxStartsWReg And Not IsRegistered Then
        RegisterApp frmSplash
    End If
Exit Function
ERRHand:
If ErrorBox("666666", Err) Then Resume Next
End Function

Public Function RegisterApp(frm As Form) As Boolean
Dim s As String

If Not IsRegistered Then
    frmLic.Show 1, frm
'    s = InputBox(ModText(0) & vbCr & ModText(1) & ModText(2), ModText(3))
'    SaveSetting AppExeName, cstrOptions, cstrRegistration, s
'    If StartAnz >= MaxStartsWReg And Not IsRegistered Then getWebRegKey
End If
RegisterApp = IsRegistered()

End Function
Public Function RegisterFree() As Boolean
Dim s As String

If Not IsRegistered Then
    s = InputBox(ModText(0) & vbCr & ModText(1) & ModText(2), ModText(3))
    SaveSetting App.EXEName, cstrOptions, cstrRegistration, s
    If StartAnz >= MaxStartsWReg And Not IsRegisteredFree Then getWebRegKey
End If
RegisterFree = IsRegisteredFree()

End Function

Public Sub getWebRegKey()
    SendMail "register7@playseven.com", ModText(4), ModText(5) & gstrSpace & AppInfo
End Sub

