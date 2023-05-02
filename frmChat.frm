VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmChat 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   2640
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8730
   Icon            =   "frmChat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   8730
   Begin VB.PictureBox picLevel 
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   7290
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   19
      Top             =   1200
      Width           =   420
   End
   Begin GIF89LibCtl.Gif89a aniGifAvatar 
      Height          =   1155
      Left            =   5970
      OleObjectBlob   =   "frmChat.frx":08CA
      TabIndex        =   15
      ToolTipText     =   "www.avatarus.de"
      Top             =   510
      Width           =   1155
   End
   Begin VB.CommandButton cmdStartMPGame 
      Height          =   495
      Left            =   5910
      Picture         =   "frmChat.frx":090C
      Style           =   1  'Grafisch
      TabIndex        =   16
      Top             =   2070
      Width           =   495
   End
   Begin VB.CheckBox chkShowSystemMsg 
      Alignment       =   1  'Rechts ausgerichtet
      Caption         =   "Check1"
      Height          =   225
      Left            =   6450
      TabIndex        =   8
      Top             =   2310
      Value           =   1  'Aktiviert
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Check1"
      Height          =   195
      Left            =   5910
      TabIndex        =   6
      Top             =   1800
      Value           =   1  'Aktiviert
      Width           =   195
   End
   Begin RichTextLib.RichTextBox rtxtChat 
      Height          =   1545
      Left            =   60
      TabIndex        =   5
      Top             =   450
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   2725
      _Version        =   393217
      BackColor       =   4210752
      ScrollBars      =   2
      TextRTF         =   $"frmChat.frx":11D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   10020
      Top             =   60
   End
   Begin VB.CheckBox chkVoIP 
      BackColor       =   &H00C00000&
      Caption         =   "VoIP"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   8010
      TabIndex        =   2
      Top             =   2340
      Width           =   195
   End
   Begin VB.CommandButton cmdWhisper 
      Caption         =   "privat"
      Height          =   255
      Left            =   8760
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtSend 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   60
      MaxLength       =   777
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertikal
      TabIndex        =   0
      Top             =   2040
      Width           =   5775
   End
   Begin MSComctlLib.ListView lvMembers 
      Height          =   840
      Left            =   8730
      TabIndex        =   3
      Top             =   900
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   1482
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FlatScrollBar   =   -1  'True
      _Version        =   393217
      ForeColor       =   16777215
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Status"
         Object.Width           =   1411
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8850
      Top             =   1710
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483647
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   8421376
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":1251
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":2544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":28F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":4C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":7159
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":95A3
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChat.frx":BC53
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblGegnerLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "blind und zwei linke Hände"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   825
      Left            =   7830
      TabIndex        =   18
      Top             =   1110
      Width           =   855
   End
   Begin VB.Label lblGegnerName 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      Caption         =   "New Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   7260
      TabIndex        =   17
      Top             =   480
      Width           =   1425
   End
   Begin VB.Image ImgAvatar 
      BorderStyle     =   1  'Fest Einfach
      Height          =   1200
      Left            =   5940
      Stretch         =   -1  'True
      ToolTipText     =   "www.avatarus.de"
      Top             =   480
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label lbl_Cancel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8415
      TabIndex        =   12
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   225
   End
   Begin VB.Label lblMin 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   8070
      TabIndex        =   10
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   345
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Show System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6720
      TabIndex        =   9
      Top             =   2340
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "LOG"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   6180
      TabIndex        =   7
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label lblVoIp 
      BackColor       =   &H00000000&
      Caption         =   "VoIP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8220
      TabIndex        =   4
      Top             =   2340
      Width           =   525
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   8475
      TabIndex        =   13
      ToolTipText     =   "Exit"
      Top             =   60
      Width           =   225
   End
   Begin VB.Label lblMinback 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   8205
      TabIndex        =   11
      ToolTipText     =   "Exit"
      Top             =   90
      Width           =   120
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00000000&
      Caption         =   "  Private Chat"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8715
   End
   Begin VB.Shape VoiceShape 
      BackColor       =   &H000000FF&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H000000FF&
      Height          =   1275
      Left            =   5880
      Shape           =   1  'Quadrat
      Top             =   450
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "frmChat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
'  File:       frmChat.frm
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Implements DirectPlay8Event
Implements DirectPlayVoiceEvent8

Private HostConnected As Boolean
Private MPSpielAngefragt As Boolean

Private MPMixedGame As Kartenspiel
Private myText() As String


Public Enum ChatMsgType
    SystemMsg
    TextMsg
End Enum



Private Sub chkVoIP_Click()
'VoIP straten ?
    If Me.chkVoIP.Value = vbChecked Then
        ConnectVoice
    ElseIf chkVoIP.Value = vbUnchecked Then
        CleanUpVoice
        Me.lblVoIp.BackColor = vbBlack
    End If
    
End Sub

Private Sub ConnectVoice()
        Dim oSession As DVSESSIONDESC
        
        'First let's set up the DirectPlayVoice stuff since that's the point of this demo
        'After we've created the session and let's start
        'the DplayVoice server
        
        'Create our DPlayVoice Server
        If DPlayEventsForm.IsHost Then
            Set dvServer = dx.DirectPlayVoiceServerCreate
                
            'Set up the Session
            oSession.lBufferAggressiveness = DVBUFFERAGGRESSIVENESS_DEFAULT
            oSession.lBufferQuality = DVBUFFERQUALITY_MAX 'DVBUFFERQUALITY_DEFAULT
            oSession.lSessionType = DVSESSIONTYPE_PEER
            oSession.guidCT = vbNullString
            
            'Init and start the session
            dvServer.Initialize dpp, 0
            dvServer.StartSession oSession, 0
        End If
    
        Dim oSound As DVSOUNDDEVICECONFIG
        Dim oClient As DVCLIENTCONFIG
        'Now create a client as well (so we can both talk and listen)
        Set dvClient = dx.DirectPlayVoiceClientCreate
        'Now let's create a client event..
        dvClient.StartClientNotification Me
        dvClient.Initialize dpp, 0
        'Set up our client and sound structs
        oClient.lFlags = DVCLIENTCONFIG_AUTOVOICEACTIVATED Or DVCLIENTCONFIG_AUTORECORDVOLUME
        oClient.lBufferAggressiveness = DVBUFFERAGGRESSIVENESS_DEFAULT
        oClient.lBufferQuality = DVBUFFERQUALITY_MAX 'DVBUFFERQUALITY_DEFAULT
        oClient.lNotifyPeriod = 0
        oClient.lThreshold = DVTHRESHOLD_UNUSED
        oClient.lPlaybackVolume = DVPLAYBACKVOLUME_DEFAULT
        oSound.hwndAppWindow = Me.hWnd
        
        On Error Resume Next
        'Connect the client
        dvClient.Connect oSound, oClient, 0
        If Err.Number = DVERR_RUN_SETUP Then    'The audio tests have not been run on this
                                                'machine.  Run them now.
            'we need to run setup first
            Dim dvSetup As DirectPlayVoiceTest8
            
            Set dvSetup = dx.DirectPlayVoiceTestCreate
            dvSetup.CheckAudioSetup vbNullString, vbNullString, Me.hWnd, ZERO 'Check the default devices since that's what we'll be using
            If Err.Number = DVERR_COMMANDALREADYPENDING Then
                UpdateChat SystemMsg, myText(0) & vbCr & myText(1), Me
                GoTo RAUS
            End If
            If Err.Number = DVERR_USERCANCEL Then
                UpdateChat SystemMsg, myText(0) & vbCr & myText(3), Me
                GoTo RAUS
            End If
            Set dvSetup = Nothing
            dvClient.Connect oSound, oClient, 0
        ElseIf Err.Number <> ZERO And Err.Number <> DVERR_PENDING Then
            UpdateChat SystemMsg, myText(0) & vbCrLf & "Error:" & CStr(Err.Number), Me
            GoTo RAUS
        ElseIf Err.Number = ZERO Then
            UpdateChat SystemMsg, myText(27), Me
            
        End If
Exit Sub

RAUS:
Me.chkVoIP = vbUnchecked
Me.chkVoIP.Visible = False
End Sub

Private Sub cmdStartMPGame_Click()
'startet das Multiplayerspiel
If LaufendesSpiel Then
    MsgBox myText(4), vbInformation
ElseIf MPSpielAngefragt Then
    MsgBox myText(35)
Else
    MPSpielAngefragt = True
    If Me.lvMembers.SelectedItem.Key = "K" & glMyID Then
        MsgBox myText(6), vbOKOnly Or vbQuestion, myText(7)
        Exit Sub
    End If
    SendNetworkMessage MsgWhisper, StartGame, CLng(Mid$(Me.lvMembers.SelectedItem.Key, 2))
    UpdateChat SystemMsg, myText(36), Me
End If
End Sub

'Private Sub cmdWhisper_Click()
'
'    If lstUsers.ListIndex < Zero Then
'        MsgBox "You must select a user in the list before you can whisper to that person.", vbOKOnly Or vbInformation, "Select someone"
'        Exit Sub
'    End If
'
'    If lstUsers.ItemData(lstUsers.ListIndex) = Zero Then
'        MsgBox "Why are you whispering to yourself?", vbOKOnly Or vbInformation, "Select someone else"
'        Exit Sub
'    End If
'
'    If txtSend.Text = vbNullString Then
'        MsgBox "What's the point of whispering if you have nothing to say..", vbOKOnly Or vbInformation, "Enter text"
'        Exit Sub
'    End If
'
'    SendNetworkMessage MsgWhisper, txtSend.Text, lstUsers.ItemData(lstUsers.ListIndex)
'    txtSend.Text = vbNullString
'    'Send this message to the person you are whispering to
'    UpdateChat "**<" & gsUserName & ">** " & txtSend.Text
'
'End Sub

Public Sub SendNetworkMessage(msgType As vbDplayChatMsgType, MSG As String, UserId As Long)
    Dim lOffset As Long
    Dim oBuf() As Byte
    
    If Not HostConnected And Not DPlayEventsForm.IsHost Then Exit Sub
    
    If dpp Is Nothing Then
        MsgBox "DirectPlay is not initialized"
        Unload Me
        Exit Sub
    End If
    
    lOffset = NewBuffer(oBuf)
    AddDataToBuffer oBuf, msgType, LenB(msgType), lOffset
    AddStringToBuffer oBuf, MSG, lOffset
    
    Select Case msgType
        Case vbDplayChatMsgType.MsgWhisper
            dpp.SendTo UserId, oBuf, 0, DPNSEND_NOLOOPBACK + DPNSEND_GUARANTEED
        Case Else
            dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK + DPNSEND_GUARANTEED
    End Select
End Sub

Private Sub Form_Load()

   'We did choose to play a game
    gsUserName = AktuellerSpieler.SpielerName
    MPSpielAngefragt = False
    
    'Text für das form laden
    If Not ArrayIsFilled(myText()) Then LoadObjectText Me.Name, myText()
    
    If Not DPlayEventsForm Is Nothing Then
        If DPlayEventsForm.IsHost Then
            Me.lblCaption = Me.lblCaption & "- (HOST)"
        End If
    Else
        'raus hier
        Unload Me
        Exit Sub
    End If
    SetBackGround Me
    makeRoundEdges Me
    Me.cmdStartMPGame.ToolTipText = myText(8)
    If frmMain_Loaded Then
        Dock2Main
        'frmMain.Timer1.Enabled = True
        Me.Timer1.Enabled = True
    Else
        SetzFensterMittig Me
    End If
    
    ShowAvatar Me, App.path & cstrSubPathAvatars & Gegner.AvatarFileName
    Me.lblGegnerLevel = strPlayerLevel(Gegner.SpielerLevel)
    Me.lblGegnerName = Gegner.SpielerName
    Me.picLevel.Picture = Me.ImageList1.ListImages(Gegner.SpielerLevel + 1).Picture

    frmChat_Loaded = True
'    Debug.Print "Host started"
End Sub

Public Sub Dock2Main()
    Me.Move frmMain.Left, frmMain.Top - Me.Height
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = vbWhite
    Me.lblMin.ForeColor = vbWhite
End Sub


'Private Sub Form_Resize()
'Const RandAbstand_Y = 550
'Const StdHeight = 1450
'Const myHeight = 2505
'Const myWidth = 8730
'Const maxHeight = 5000
'Dim newY As Long
'
'If Not Me.WindowState = vbMinimized Then
'    If Me.Height < myHeight Then Me.Height = myHeight
'    If Me.Height > maxHeight Then Me.Height = maxHeight
'    If Me.Width <> myWidth Then Me.Width = myWidth
'
'    newY = Me.Height - RandAbstand_Y
'    Me.cmdStartMPGame.Top = newY
'    Me.txtSend.Top = newY
'    Me.chkVoIP.Top = newY
'    Me.Label1.Top = newY
'    Me.chkLog.Top = newY + 230
'    Me.Label3.Top = newY + 230
'    Me.chkShowSystemMsg.Top = newY + 170
'    Me.Label4.Top = newY + 50
'    newY = StdHeight + Me.Height - myHeight
'    Me.rtxtChat.Height = newY
'    Me.lvMembers.Height = newY
'    'Me.lvMembers.Left = Me.txtChat.Width + 50
'End If
'End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim Antw As Integer

    
If LaufendesSpiel Then
    If Not myAgent Is Nothing And frmMain_Loaded Then
        Antw = frmMain.AgentBalloon.MsgBalloon(myText(9) & vbCr & myText(10), vbQuestion + vbYesNo, myText(34), myAgent)
    Else
        Antw = MsgBox(myText(9) & vbCr & myText(10), vbCritical + vbYesNo)
    End If
    If Antw = vbYes Then
        Me.Timer1.Enabled = False
        LaufendesSpiel = False
        frmChat.SendNetworkMessage MsgSendSpielAbbruch, gstrNullstr, 0
        If ServerConnected And Not ServerEventsForm Is Nothing Then
            ServerEventsForm.SendMsg2Server Msg_PlayerLost
        End If
    Else
        Cancel = True
        Exit Sub
    End If

End If

If frmMain_Loaded Then Unload frmMain
AnimWindow Me, AW_HIDE + AW_BLEND

CleanUpP2P

MPSpielAngefragt = False
frmChat_Loaded = False

If ServerConnected Then
    AktuellerSpieler.Status = Idle
    ServerEventsForm.SendMsg2Server Msg_PlayerInfo
    ServerEventsForm.Show
Else
    frmSplash.Show
End If

End Sub

Private Sub UpdateList(ByVal lPlayerID As Long, fTalking As Boolean)
'    Dim lCount As Long
'    For lCount = lvMembers.ListItems.count To 1 Step -1
'        If lvMembers.ListItems.item(lCount).Key = "K" & CStr(lPlayerID) Then
'            'Change this guys status
'            If fTalking Then
'                lvMembers.ListItems.item(lCount).SubItems(1) = myText(11)
'            Else
'                lvMembers.ListItems.item(lCount).SubItems(1) = myText(12)
'            End If
'        End If
'    Next
If lPlayerID <> glMyID Then
    Me.VoiceShape.Visible = fTalking
Else
    If fTalking Then
        Me.lblVoIp.BackColor = vbRed
    Else
        Me.lblVoIp.BackColor = vbBlack
    End If
End If
End Sub

Public Sub UpdatePlayerList()
    'Get everyone who is currently in the session and add them if we don't have them currently.
    Dim lCount As Long
    Dim Player As DPN_PLAYER_INFO
    Dim Key As String
    Dim lItem As ListItem, sName As String

    ' Enumerate players
    For lCount = 1 To dpp.GetCountPlayersAndGroups(DPNENUM_PLAYERS)
        If Not (AmIInList(dpp.GetPlayerOrGroup(lCount))) Then 'Add this player
            
            Player = dpp.GetPeerInfo(dpp.GetPlayerOrGroup(lCount))
            sName = Player.Name
            If sName = vbNullString Then sName = "Unknown"
            If isBitSet(Player.lPlayerFlags, DPNPLAYER_LOCAL) Then
                glMyID = dpp.GetPlayerOrGroup(lCount)
            End If
                
            Key = "K" & CStr(dpp.GetPlayerOrGroup(lCount))
'            Debug.Print Key, sName, glMyID
            Set lItem = lvMembers.ListItems.Add(, Key, sName)
            lItem.SubItems(1) = myText(12)
            lItem.Selected = True
        End If
    Next lCount
End Sub

Private Function AmIInList(ByVal lPlayerID As Long) As Boolean
    Dim lCount As Long, fInThis As Boolean
    
    For lCount = lvMembers.ListItems.count To 1 Step -1
        If lvMembers.ListItems.item(lCount).Key = "K" & CStr(lPlayerID) Then
            fInThis = True
        End If
    Next
    AmIInList = fInThis
End Function

Private Sub RemovePlayer(ByVal lPlayerID As Long)
'    Dim lCount As Long
'
'    For lCount = lvMembers.ListItems.Count To 1 Step -1
'        If lvMembers.ListItems.Item(lCount).Key = "K" & CStr(lPlayerID) Then
'            lvMembers.ListItems.Remove lCount
'        End If
'    Next
    lvMembers.ListItems.Remove "K" & CStr(lPlayerID)
End Sub

Private Sub lbl_Cancel_Click()
    
    Unload Me
    
End Sub

Private Sub lbl_Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize - 3
End Sub

Private Sub lbl_Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = ROT
    Me.lblMin.ForeColor = vbWhite
End Sub

Private Sub lbl_Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize + 3
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub lblMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lblMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.FontSize = Me.lblMin.FontSize - 2
End Sub

Private Sub lblMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.ForeColor = vbRed
    Me.lbl_Cancel.ForeColor = vbWhite
End Sub

Private Sub lblMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.FontSize = Me.lblMin.FontSize + 2
End Sub

Private Sub Timer1_Timer()

'ermitteln ob sich hauptform bewget hat
Static X As Long
Static Y As Long
If frmMain_Loaded Then
    If frmMain.Top <> Y Or frmMain.Left <> X Then
        Y = frmMain.Top
        X = frmMain.Left
        Dock2Main
    End If
End If

End Sub
Private Sub txtSend_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If txtSend.Text <> vbNullString Then 'Make sure they are trying to send something
            'Send this message to everyone
            SendNetworkMessage MsgChat, txtSend.Text, 0
            UpdateChat TextMsg, "<" & gsUserName & ">" & gstrSpace & txtSend.Text, Me
            txtSend.Text = vbNullString
            KeyAscii = 0
            'dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK
        End If 'We won't set KeyAscii to Zero here, because if they are trying to
               'send blank data, we don't care about the ding for hitting enter on
               'an empty line
    ElseIf KeyAscii = vbKeyTab Then
        KeyAscii = 0
        txtSend.Text = AutoUpdateName(Me.txtSend.Text)
        txtSend.SelStart = Len(Me.txtSend.Text)
    End If

End Sub

Private Function GetName(ByVal lID As Long) As String
    
    GetName = Me.lvMembers.ListItems.item("K" & lID)
End Function

Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
    HostConnected = True
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)

UpdatePlayerList
UpdateChat SystemMsg, "->" & GetName(lPlayerID) & myText(13), Me
'    Dim dpPeer As DPN_PLAYER_INFO
'    Dim item As ListItem
'    dpPeer = dpp.GetPeerInfo(lPlayerID)
'
'    'Add this person to chat (even if it's me)
'    If (dpPeer.lPlayerFlags And DPNPLAYER_LOCAL) <> DPNPLAYER_LOCAL Then 'this isn't me, someone just joined
'
'        'If it's not me, include an ItemData
'        Gegner.Spielername = dpPeer.Name
'        Gegner.GlobalID = lPlayerID
'
'        Key = "K" & CStr(dpp.GetPlayerOrGroup(lCount))
'        Debug.Print Key
'        Set lItem = lvMembers.ListItems.Add(, Key, sName)
'        lItem.SubItems(1) = myText(12)
'    End If
End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)

'We only care when someone leaves.  When they join we will receive a 'MSGJoin'
'Remove this player from our list
UpdateChat SystemMsg, "<-" & GetName(lPlayerID) & myText(14), Me
RemovePlayer lPlayerID
End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)
    Dim dpPeer As DPN_PLAYER_INFO
    dpPeer = dpp.GetPeerInfo(lNewHostID)
    If (dpPeer.lPlayerFlags And DPNPLAYER_LOCAL) = DPNPLAYER_LOCAL Then 'I am the new host
        Me.lblCaption = Me.lblCaption & " (HOST)"
    End If
End Sub

Private Sub DirectPlay8Event_IndicateConnect(dpnotify As DxVBLibA.DPNMSG_INDICATE_CONNECT, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_IndicatedConnectAborted(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_InfoNotify(ByVal lMsgID As Long, ByVal lNotifyID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_Receive(dpnotify As DxVBLibA.DPNMSG_RECEIVE, fRejectMsg As Boolean)
    'process what msgs we receive.
    Dim lMsg As Long, lOffset As Long
    Dim dpPeer As DPN_PLAYER_INFO, sName As String
    Dim sChat As String
    Dim Antw As Long
    
    With dpnotify
        GetDataFromBuffer .ReceivedData, lMsg, LenB(lMsg), lOffset
        sName = GetName(.idSender)
        sChat = GetStringFromBuffer(.ReceivedData, lOffset)
        Select Case lMsg
            Case MsgChat
                UpdateChat TextMsg, "<" & sName & "> " & gstrSpace & sChat, Me
                If boolAgentTalkChat Then AgentSpeak sChat
            Case MsgWhisper
                
                If MPSpielAngefragt And sChat = StartGameOK Then
                    moveNewPos
                    Playermodus = multiplayer
                    Gegner.SpielerName = sName
                    UpdateChat SystemMsg, myText(30), Me
                    AnimWindow frmMain, AW_ACTIVATE + AW_SLIDE + AW_VER_POSITIVE
                    frmMain.Init
                ElseIf MPSpielAngefragt And sChat = StartGameNOK Then
                    UpdateChat TextMsg, sName & gstrSpace & myText(15), Me
                    MPSpielAngefragt = False
                ElseIf sChat = StartGame Then
                    UpdateChat SystemMsg, myText(31), Me
                    Antw = MsgBox(sName & myText(16) & vbCr & myText(17), vbQuestion + vbYesNo)
                    If Antw = vbYes Then
                        Me.SendNetworkMessage MsgWhisper, StartGameOK, .idSender
                        moveNewPos
                        Playermodus = multiplayer
                        Gegner.SpielerName = sName
                        AnimWindow frmMain, AW_ACTIVATE + AW_SLIDE + AW_VER_POSITIVE
                        frmMain.Init
                    Else
                        Me.SendNetworkMessage MsgWhisper, StartGameNOK, .idSender
                    End If
                Else
                    UpdateChat TextMsg, "**<" & sName & ">** " & gstrSpace & sChat, Me
                End If
            Case MsgSendMixedGame
                GetKartenSpielFromString (sChat)
                SpielErhalten = True
                UpdateChat SystemMsg, myText(18), Me
            Case MsgSendGeber
                Geber = CInt(sChat)
                UpdateChat SystemMsg, myText(19) & gstrSpace & IIf(Geber = Spieler, AktuellerSpieler.SpielerName, Gegner.SpielerName), Me
                frmMain.Kartenausbreiten
                frmMain.ersterZug
            Case MsgSendAbheben
                frmMain.picAbheben CInt(sChat)
                UpdateChat SystemMsg, myText(20), Me
            Case MsgSendCard
                Debug.Print "P2PNetz :" & sChat & " Karte geworfen"
                frmMain.picKarteComp_Wirf CInt(sChat)
            Case MsgSendStichEnde
                Debug.Print "P2PNetz : StichEnde"
                frmMain.StichEnde
            Case MsgSendSpielAbbruch
                UpdateChat SystemMsg, Gegner.SpielerName & myText(21), Me
                AgentSpeak Gegner.SpielerName & myText(21), True
                LaufendesSpiel = False
'                If frmMain_Loaded Then Unload frmMain
                
            Case MsgSendMeisterfehler
                UpdateChat SystemMsg, Gegner.SpielerName & myText(32), Me
                MeisterFehlerGegner = True
                frmMain.checkSieg
            Case MsgSendZeitueberschreitung
                ZeitUeberschreitungGegner = True
                UpdateChat SystemMsg, Gegner.SpielerName & myText(33), Me
                frmMain.checkSieg
        End Select
    End With
    
End Sub


Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
    If dpnotify.hResultCode = DPNERR_HOSTTERMINATEDSESSION Then
        MsgBox myText(23) & vbCr & myText(24), vbOKOnly Or vbInformation
    Else
        MsgBox myText(25), vbOKOnly Or vbInformation  ' & vbCr & myText(24)
    End If
    LaufendesSpiel = False
    HostConnected = False
End Sub

Sub moveNewPos()
    Dim i As Integer
    Dim Old_Y As Long
    Dim New_Y As Long
    Old_Y = Me.Top
    For i = 1 To frmMain.ScaleHeight \ 2 Step Screen.TwipsPerPixelX * 10
        New_Y = Old_Y - (i)
        If New_Y <= ZERO Then Exit For
        Me.Move Me.Left, New_Y
'        DoEvents
    Next
    frmMain.Dock2Chat
End Sub

Private Sub DirectPlayVoiceEvent8_ConnectResult(ByVal ResultCode As Long)
    Dim lTargets(0) As Long
    
    If ResultCode = ZERO Then
        lTargets(0) = DVID_ALLPLAYERS
        dvClient.SetTransmitTargets lTargets, 0

        'Update the list
        UpdatePlayerList
    Else
        UpdateChat SystemMsg, myText(0) & vbCr & myText(24) & vbCrLf & "Error:" & CStr(Err.Number), Me
        Me.chkVoIP.Value = 0
        'DPlayEventsForm.CloseForm Me
    End If
End Sub

Private Sub DirectPlayVoiceEvent8_CreateVoicePlayer(ByVal playerID As Long, ByVal flags As Long)
    'Someone joined, update the player list
    UpdatePlayerList
    UpdateChat SystemMsg, myText(29), Me

End Sub

Private Sub DirectPlayVoiceEvent8_DeleteVoicePlayer(ByVal playerID As Long)
    'Someone quit, remove them from the session
    'RemovePlayer playerID
    UpdatePlayerList
    UpdateChat SystemMsg, myText(28), Me

End Sub

Private Sub DirectPlayVoiceEvent8_DisconnectResult(ByVal ResultCode As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlayVoiceEvent8_HostMigrated(ByVal NewHostID As Long, ByVal NewServer As DxVBLibA.DirectPlayVoiceServer8)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlayVoiceEvent8_InputLevel(ByVal PeakLevel As Long, ByVal RecordVolume As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlayVoiceEvent8_OutputLevel(ByVal PeakLevel As Long, ByVal OutputVolume As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlayVoiceEvent8_PlayerOutputLevel(ByVal playerID As Long, ByVal PeakLevel As Long)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStart(ByVal playerID As Long)
    'Someone is talking, update the list
    UpdateList playerID, True
End Sub

Private Sub DirectPlayVoiceEvent8_PlayerVoiceStop(ByVal playerID As Long)
    'Someone stopped talking, update the list
    UpdateList playerID, False
End Sub

Private Sub DirectPlayVoiceEvent8_RecordStart(ByVal PeakVolume As Long)
    'I am talking, update the list
    UpdateList glMyID, True
End Sub

Private Sub DirectPlayVoiceEvent8_RecordStop(ByVal PeakVolume As Long)
    'I have quit talking, update the list
    UpdateList glMyID, False
End Sub

Private Sub DirectPlayVoiceEvent8_SessionLost(ByVal ResultCode As Long)
    'The voice session has exited, let's quit
    UpdateChat SystemMsg, myText(26), Me
    Me.chkVoIP = vbUnchecked
End Sub
