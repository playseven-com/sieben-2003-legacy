Attribute VB_Name = "MSAgent"
Option Explicit
Option Compare Text

Private Const cstrDefault As String = "default"
Public myAgent As IAgentCtlCharacterEx
'Const DATAPATH = "genie.acs"

Public Enum AgentAnims
    Decline
    Confused
    Congratulate
    Writes
    Pleased
    Process
    Sad
    Surprised
    Explain
    Greet
End Enum

Public Enum AgentGesture
    ShowAt
    move2
End Enum

Public boolAgentTalkChat As Boolean
Private LastRequest As IAgentCtlRequest

Public Sub InitAgent()
Dim X As Long, Y As Long
Dim Antw As Long
On Error Resume Next
If useAgent Then
    If myAgent Is Nothing And frmMain_Loaded Then
        frmMain.Agent1.Characters.Load cstrDefault
        If Err.Number = -2147024894 Or Err.Number = -2147213311 Or Err.Number = -2147213304 Or Err.Number = -2147213289 Then
            Antw = MsgBox("No Agent installed." & vbCr & "Would You like to download the Characters ?", vbInformation + vbYesNo)
            If Antw = vbYes Then
                Go2URL "www.microsoft.com/msagent/downloads.htm#character"
            End If
            useAgent = False
            frmMain.men_useAgent.Enabled = False
            Exit Sub
        End If
        Set myAgent = frmMain.Agent1.Characters(cstrDefault)
        
    End If
    myAgent.Show
    
    myAgent.LanguageID = LangID '&H409 '&H415
    If myAgent.TTSModeID = vbNullString Then
        AgentThink ModText(28)
        frmMain.menAgentTalkChat.Visible = False
        boolAgentTalkChat = False
        Antw = AgentQuestion("Would You like to download the text to speech (TTS) engine in your language ?", "No TTS Engine installed.")
        If Antw = vbYes Then
            Go2URL "www.microsoft.com/msagent/downloads.htm#tts"
        End If
    Else
        frmMain.menAgentTTS.Visible = False
    End If

    SetAgentAsEnemy
    Y = (frmMain.Top + (frmMain.Height * 0.5)) \ Screen.TwipsPerPixelY
    X = (frmMain.Left + frmMain.Width * 0.3) \ Screen.TwipsPerPixelX
    myAgent.MoveTo X, Y - myAgent.Height, 300
    myAgent.Play "Greet"
End If

End Sub
Public Sub SetAgentAsEnemy()
If myAgent Is Nothing Then Exit Sub
    If Playermodus = singleplayer Then
        Gegner.SpielerName = myAgent.Name
    End If
End Sub

Public Sub AgentSpeak(str As String, Optional alternativeMsgB)

If Test Then Debug.Print str
If useAgent And Not myAgent Is Nothing And Not str = vbNullString And frmMain_Loaded Then
    myAgent.Speak str
Else
    If Not IsMissing(alternativeMsgB) And Not str = vbNullString Then
        If Not Test Then MsgBox str
    End If
End If

End Sub
Public Sub AgentThink(str As String)

If Test Then Debug.Print str
If useAgent And Not myAgent Is Nothing And Not str = vbNullString And frmMain_Loaded Then
    myAgent.Think str
End If

End Sub

Public Function AgentQuestion(str As String, Title As String) As Long
    If Not myAgent Is Nothing And frmMain_Loaded Then
        AgentQuestion = frmMain.AgentBalloon.MsgBalloon(str, vbYesNo, Title, myAgent)
    Else
        AgentQuestion = MsgBox(str, vbYesNo, Title)
    End If
End Function


Public Sub MoveAgentInForm(frm As Form, ByVal X As Long, ByVal Y As Long, gesture As AgentGesture)
On Error GoTo ERRHand

    If Not myAgent Is Nothing Then
'        If Not LastRequest Is Nothing And gesture = move2 Then
'            myAgent.Stop LastRequest
'            Debug.Print "AgentSopped"
'        End If
'        DoEvents
        Y = (frm.Top + Y) \ Screen.TwipsPerPixelY
        X = (frm.Left + X) \ Screen.TwipsPerPixelX
        
        'If Abs(myAgent.Left - x) > 1.5 * myAgent.Width Then
        If gesture = move2 Then
            Set LastRequest = myAgent.MoveTo(X - (myAgent.Width * 0.8), Y, 250)
        Else
            Set LastRequest = myAgent.GestureAt(X, Y + myAgent.Height) ' * 1.6)
        End If
   
    End If
    Sleep 300 + (100 * (6 - AktuellerSpieler.SpielerLevel))
    'myAgent.Stop
Exit Sub
ERRHand:
If ErrorBox("MoveAgentInForm", Err) Then Resume Next

End Sub
Public Sub AgentTakeStich()

Dim X As Long

On Error Resume Next

    If myAgent Is Nothing Then Exit Sub
    If Int(100 * Rnd + 1) >= 20 Then Exit Sub
    If myAgent.Left > (frmMain.Left + frmMain.picKarteStich(ONE).Left) \ Screen.TwipsPerPixelX Then
        X = frmMain.picKarteStich(ONE).Left
    Else
        X = frmMain.picKarteStich(8).Left + frmMain.picKarteStich(ONE).Width + (myAgent.Width * Screen.TwipsPerPixelX)
    End If
    MoveAgentInForm frmMain, X, frmMain.picKarteStich(ONE).Top + 111, move2
    
    'Sleep 2000
    
'Exit Sub
'ERRHand:
'If ErrorBox("AgentTakeStich", Err) Then Resume Next

End Sub
Public Sub AgentAnim(AnimIndex As AgentAnims)
'Animation abspielen
On Error Resume Next
Dim str As String
Select Case AnimIndex
    Case AgentAnims.Confused
        str = "Confused"
    Case AgentAnims.Congratulate
        str = "Congratulate"
    Case AgentAnims.Decline
        str = "Decline"
    Case AgentAnims.Explain
        str = "Explain"
    Case AgentAnims.Process
        str = "Process"
    Case AgentAnims.Pleased
        str = "Pleased"
    Case AgentAnims.Sad
        str = "Sad"
    Case AgentAnims.Surprised
        str = "Surprised"
    Case AgentAnims.Writes
        str = "Write"
    Case AgentAnims.Greet
        str = "Greet"
    Case Else
        'MsgBox "Ubekannte Animation"
End Select
If useAgent And Not myAgent Is Nothing Then
    myAgent.Play str
End If
End Sub

Public Sub DestroyAgent()
On Error GoTo ERRHand
    If myAgent Is Nothing Then Exit Sub
    If Not LastRequest Is Nothing Then myAgent.Stop LastRequest
    myAgent.Hide
'    useAgent = False
    frmMain.men_useAgent.Checked = False
    frmMain.Agent1.Characters.Unload cstrDefault

    Set myAgent = Nothing
    If Playermodus = singleplayer Then Gegner.SpielerName = cstrGegnerStandardName
    Debug.Print "Agent Destroyed"
Exit Sub
ERRHand:
If ErrorBox("DestroyAgent", Err) Then Resume Next
End Sub
