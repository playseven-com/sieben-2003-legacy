VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form ServerConnect 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   6840
   ClientLeft      =   5250
   ClientTop       =   3210
   ClientWidth     =   8745
   ControlBox      =   0   'False
   ForeColor       =   &H00C0C0C0&
   Icon            =   "ServerCon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   456
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   Begin VB.CheckBox chkShowSystemMsg 
      Caption         =   "Check1"
      Height          =   225
      Left            =   7920
      TabIndex        =   16
      Top             =   3270
      Value           =   1  'Aktiviert
      Width           =   195
   End
   Begin VB.CheckBox chkLog 
      Caption         =   "Check1"
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   2970
      Width           =   195
   End
   Begin RichTextLib.RichTextBox rtxtSend 
      Height          =   585
      Left            =   90
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6180
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   1032
      _Version        =   393217
      ScrollBars      =   2
      MaxLength       =   777
      TextRTF         =   $"ServerCon.frx":08CA
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
   Begin RichTextLib.RichTextBox rtxtChat 
      Height          =   3255
      Left            =   90
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2880
      Width           =   7755
      _ExtentX        =   13679
      _ExtentY        =   5741
      _Version        =   393217
      BackColor       =   4210752
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"ServerCon.frx":0945
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
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Spieloptionen"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1725
      Left            =   7020
      TabIndex        =   8
      Top             =   420
      Width           =   1605
      Begin VB.CheckBox chkNoHigherLevel 
         BackColor       =   &H00000000&
         Caption         =   "kein höheres Level"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   465
         Left            =   90
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1050
         Width           =   1455
      End
      Begin VB.OptionButton optSpielOption 
         BackColor       =   &H00000000&
         Caption         =   "Freundschaft"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   540
         Width           =   1455
      End
      Begin VB.OptionButton optSpielOption 
         BackColor       =   &H00000000&
         Caption         =   "Liga"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   90
         X2              =   1530
         Y1              =   930
         Y2              =   930
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7110
      Top             =   2190
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":09C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":1CB3
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":2061
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":43F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":68C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":8D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ServerCon.frx":B3C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox lstGames 
      BackColor       =   &H00404040&
      Columns         =   1
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1530
      Left            =   90
      TabIndex        =   2
      Top             =   390
      Width           =   6855
   End
   Begin VB.CommandButton cmdEnterIP 
      Caption         =   "Gegner IP eingeben"
      Height          =   315
      Left            =   1890
      TabIndex        =   1
      Top             =   2070
      Width           =   1755
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "S&erver suchen"
      Height          =   315
      Left            =   180
      TabIndex        =   0
      Top             =   2070
      Width           =   1575
   End
   Begin VB.Timer tmrExpire 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   7770
      Top             =   2310
   End
   Begin MSComctlLib.ListView lvPlayers 
      Height          =   2355
      Left            =   90
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   420
      Visible         =   0   'False
      Width           =   6885
      _ExtentX        =   12144
      _ExtentY        =   4154
      SortKey         =   2
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   16777215
      BackColor       =   4210752
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   9
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Text            =   "GUID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "LEVEL"
         Object.Width           =   2118
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Punkte"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "IP Adress"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Wertung"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Gewinnquote"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "AvatarFile"
         Object.Width           =   0
      EndProperty
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
      Left            =   7980
      TabIndex        =   18
      ToolTipText     =   "Minimize"
      Top             =   -30
      Width           =   345
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
      Left            =   8145
      TabIndex        =   19
      ToolTipText     =   "Exit"
      Top             =   30
      Width           =   120
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Show System"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   17
      Top             =   3180
      Width           =   555
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8160
      TabIndex        =   15
      Top             =   2970
      Width           =   555
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   7
      Top             =   60
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   210
      TabIndex        =   6
      Top             =   60
      Width           =   1815
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
      Left            =   8340
      TabIndex        =   3
      ToolTipText     =   "Exit"
      Top             =   -30
      Width           =   225
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
      Left            =   8415
      TabIndex        =   4
      ToolTipText     =   "Exit"
      Top             =   30
      Width           =   225
   End
   Begin VB.Label lblCaption 
      BackColor       =   &H00000000&
      Height          =   375
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   8775
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "pop"
      Visible         =   0   'False
      Begin VB.Menu mnuContact 
         Caption         =   "Auffordern"
      End
      Begin VB.Menu mnuStrich 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAddFriends 
         Caption         =   "zur Buddyliste hinzufügen"
      End
      Begin VB.Menu mnuDeleteFriend 
         Caption         =   "aus Buddyliste entfernen"
      End
      Begin VB.Menu mnuIgnore 
         Caption         =   "Ignorieren"
      End
      Begin VB.Menu mnuUnIgnore 
         Caption         =   "aus Ignoreliste entfernen"
      End
   End
End
Attribute VB_Name = "ServerConnect"
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
'  File:       DPlayCon.frm
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Sleep declare
'Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'GetTickCount declare
Private Declare Function GetTickCount Lib "kernel32" () As Long
'Declares for closing the form without waiting
'Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Const WM_CLOSE = &H10

'Host expire threshold constant
Private Const HOST_EXPIRE_THRESHHOLD As Long = 2000

Private Enum SearchingButton
    StartSearch
    StopSearch
End Enum

'Internal DirectX variables
Private moDPC As DirectPlay8Client
Private moDPA As DirectPlay8Address
Private moDX As DirectX8
Private moCallback As DirectPlay8Event

'App specific vars
Private msGuid As String
Private sUser As String
Private mlSearch As SearchingButton
Private sGameName As String
Private mlMax As Long
Private mlNumPlayers As Long
Private mfComplete As Boolean
Private mfHost As Boolean
Private mlEnumAsync As Long
Private mfGotEvent As Boolean
Private mfDoneWiz As Boolean

Private mfCanUnload As Boolean

'We need to keep track of the hosts we get
Private moHosts() As HostFound
Private mlHostCount As Long
'Declaration for our API
Private mfDoneEnum As Boolean
Private mfConnectComplete As Boolean

Private myText() As String

'We need to implement the Event model for DirectPlay so we can receive callbacks
Implements DirectPlay8Event

Private Function StartWizard(oDX As DirectX8, sGuid As String) As Boolean
    Dim lCount As Long, lIndex As Long
    Dim dpn As DPN_SERVICE_PROVIDER_INFO
    'Now we can start our connection
On Error GoTo ERRHand

    mfCanUnload = False
    mlSearch = StartSearch
    mlHostCount = -1
    
    'First we need to keep track of our Peer Object, and app guid
    Set moDX = oDX
    'Set moCallback = oCallback
    msGuid = sGuid
    'mlMax = lMaxPlayers
    
    'lIndex = GetSetting("VBDirectPlay", "Defaults", "SPListIndex", -1)
    If Not (moDPC Is Nothing) Then
        moDPC.RegisterMessageHandler Me
    End If
    sUser = AktuellerSpieler.SpielerName
    'Show this screen
    ChooseSP
    Me.Show vbModeless
    
    'We have this loop here rather than just displaying the form as a modal
    'dialog if we did just display the form as modal, it would not get a
    'button in the toolbar, since it would have a parent window that wasn't visible
    
    'By displaying the window modeless, and going into a loop we get to have our
    'icon on the taskbar, and keep the main form waiting until we are done in this form.
    Do While Not mfDoneWiz
        DoSleep 10 'Give other threads cpu time
    Loop
    'Now we can return our success (or failure)
    StartWizard = mfComplete
Exit Function
ERRHand:
If ErrorBox("ServerConnect:StartWizard", Err) Then Resume Next

End Function


Public Function StartClientConnectWizard(oDX As DirectX8, oDPC As DirectPlay8Client, sGuid As String) As Boolean
On Error GoTo ERRHand
    'Set moDPP = Nothing
    If oDPC Is Nothing Then GoTo ERRHand
    Set moDPC = oDPC
    'cmdCreate.Visible = False
    StartClientConnectWizard = StartWizard(oDX, sGuid)
Exit Function
ERRHand:
MsgBox ModText(16), vbInformation, "Internet ?"

End Function

'Public Sub CloseForm(oForm As Form)
'    'Anytime we need to close a form from within a DirectPlay callback
'    'we need to use this function.  The reason is that DirectPlay uses multiple
'    'threads to spawn all of it's messages back to the application.  However
'    'it cannot close down until all of it's threads have returned.
'    'If we attempt to simply call Unload Me in the callback, we will run into
'    'a deadlock instance, since the callback will be running on the DirectPlay
'    'thread waiting for the unload to finish, and the unload will be waiting
'    'for the DirectPlay thread to finish.
'
'    'PostMessage puts the message on the queue for our form and returns immediately
'    'allowing the thread to finish
'    PostMessage oForm.hWnd, WM_CLOSE, 0, 0
'End Sub
'
'Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
'    'The DoSleep function allows other threads to have a time slice
'    'and still keeps the main VB thread alive (since DPlay callbacks
'    'run on separate threads outside of VB).
'    Sleep lMilliSec
'
'    DoEvents
'End Sub

Private Sub chkLog_Click()
    SaveSetting AppExeName, cstrOptions, cstrMakeLog, Me.chkLog
End Sub

Private Sub chkNoHigherLevel_Click()
    SaveSetting AppExeName, cstrOptions, cstrNoHigherLevel, Me.chkNoHigherLevel
End Sub

Private Sub chkShowSystemMsg_Click()
    SaveSetting AppExeName, cstrOptions, cstrshowSystemMessage, Me.chkShowSystemMsg
End Sub

Private Sub cmdJoin_Click()
    Dim HostAddr As DirectPlay8Address
    Dim DeviceAddr As DirectPlay8Address
    
    Dim dpApp As DPN_APPLICATION_DESC
    
    'You must select a game before you try to join one
    If lstGames.ListIndex < ZERO Then
        AgentSpeak myText(5), True
        Exit Sub
    End If
    
    'Wenn wir nach servern suchen suche beenden
    If mlSearch = StartSearch Then cmdRefresh_Click
    
    'Lets join the server
    Dim pInfo As DPN_PLAYER_INFO
    'Set up my peer info
    pInfo.Name = sUser
    pInfo.lInfoFlags = DPNINFO_NAME
    
    If Not (moDPC Is Nothing) Then
        moDPC.SetClientInfo pInfo, DPNOP_SYNC
    End If
    mfDoneEnum = True

    With moHosts(lstGames.ItemData(lstGames.ListIndex)).AppDesc
        dpApp.guidApplication = .guidApplication
        dpApp.guidInstance = .guidInstance
        mlNumPlayers = .lMaxPlayers
    End With
    
    mfGotEvent = False
    mfConnectComplete = False
    'Lets get our host address
    If moHosts(lstGames.ItemData(lstGames.ListIndex)).AddressHost <> vbNullString Then
        Set HostAddr = moDX.DirectPlayAddressCreate
        HostAddr.BuildFromURL moHosts(lstGames.ItemData(lstGames.ListIndex)).AddressHost
    Else
        Set HostAddr = moDPA
    End If
    If moHosts(lstGames.ItemData(lstGames.ListIndex)).AddressDevice <> vbNullString Then
        Set DeviceAddr = moDX.DirectPlayAddressCreate
        DeviceAddr.BuildFromURL moHosts(lstGames.ItemData(lstGames.ListIndex)).AddressDevice
    Else
        Set DeviceAddr = moDPA
    End If
    If Not (moDPC Is Nothing) Then
        'Now we can join the selected session
        moDPC.Connect dpApp, HostAddr, DeviceAddr, DPNCONNECT_OKTOQUERYFORADDRESSING, ByVal 0&, 0
    End If
    
    
    Do While Not mfGotEvent 'Let's wait for our connectcomplete event
        DoSleep 10 'Give other threads cpu time
    Loop
    
'    ServerConnected = True
'    UpdateChat SystemMsg, "Serverconnection etabliert", Me
    '
'    If mfConnectComplete Then
'        'We've joined our game
'        mfComplete = True
'        mfHost = False
'        'Clean up our address
'        Set HostAddr = Nothing
'        Set DeviceAddr = Nothing
'        Set moDPA = Nothing
'        Unload Me
'    End If
End Sub

Private Sub ChooseSP()
    'Set up the address
On Error GoTo ERRHand
    Set moDPA = moDX.DirectPlayAddressCreate
    If Not (moDPC Is Nothing) Then
        moDPA.SetSP DP8SP_TCPIP
    End If
Exit Sub
ERRHand:
MsgBox ModText(16), vbInformation, "Internet ?"
CleanUpServer

End Sub



Private Sub cmdEnterIP_Click()
    Unload Me
End Sub

Private Sub cmdRefresh_Click()
Dim ServerAdr As DirectPlay8Address
Dim ServerUrl As String
    If mlSearch = StartSearch Then
        'Time to enum our hosts
        mfDoneEnum = False
        Dim Desc As DPN_APPLICATION_DESC
        Desc.guidApplication = msGuid
        
        Set ServerAdr = moDX.DirectPlayAddressCreate
        ServerUrl = Encrypt(gstrServerName, False)
        ServerAdr.SetSP DP8SP_TCPIP
        ServerAdr.AddComponentString "hostname", ServerUrl '"localhost"
        'ServerAdr.BuildFromURL ServerUrl

        Debug.Print ServerAdr.GetComponentString("hostname")
        If Not (moDPC Is Nothing) Then
            mlEnumAsync = moDPC.EnumHosts(Desc, ServerAdr, moDPA, INFINITE, 0, INFINITE, 0, ByVal 0&, 0)
        End If
        cmdRefresh.Caption = myText(7)
        mlSearch = StopSearch
        Me.tmrExpire.Enabled = True
        UpdateChat SystemMsg, "Suche Server ..... ", Me
        UpdateChat SystemMsg, "In der Betaphase ist der Server nur von 15:00 bis 24:00 Uhr erreichbar.", Me
    ElseIf mlSearch = StopSearch Then
        mfDoneEnum = True
        If Not (moDPC Is Nothing) Then
            If mlEnumAsync <> ZERO Then moDPC.CancelAsyncOperation mlEnumAsync, 0
        End If
        cmdRefresh.Caption = "Server suchen" 'myText(3)
        mlSearch = StartSearch
        Me.tmrExpire.Enabled = False
    End If
End Sub

Private Sub SwitchMode(mode As Boolean)

    Me.lvPlayers.Visible = mode
    Me.Label2.Visible = mode
    Me.rtxtSend.Visible = mode
    
    Me.lstGames.Visible = Not mode
    Me.cmdEnterIP.Visible = Not mode
    Me.cmdRefresh.Visible = Not mode
    Me.Label1.Visible = Not mode
    
End Sub
Private Sub AddHostsToListBox(oHost As DPNMSG_ENUM_HOSTS_RESPONSE)
    Dim lFound As Long
    
    'Here we will add a host that was found to our list box (or ignore it
    'if it's already been added)
    If mfDoneEnum Then Exit Sub
    If mlHostCount = -1 Then
        
        If LCase$(oHost.ApplicationDescription.guidApplication) <> LCase$(AppGuid) Then Exit Sub

        'We have no hosts already. Clear our list, and add this one to the list.
        UpdateChat SystemMsg, "Server gefunden", Me
        lstGames.Clear
        ReDim moHosts(0)
        moHosts(0).AppDesc = oHost.ApplicationDescription
        moHosts(0).AddressHost = oHost.AddressSenderUrl
        moHosts(0).AddressDevice = oHost.AddressDeviceUrl
        'Save the last time this host was found
        moHosts(0).TimeLastFound = GetTickCount
        With oHost.ApplicationDescription
            lstGames.AddItem .SessionName & " - " & CStr(.lCurrentPlayers) & "/" & CStr(.lMaxPlayers) & " - Latency:" & CStr(oHost.lRoundTripLatencyMS) & " ms"
        End With
        lstGames.ItemData(0) = 0
        mlHostCount = mlHostCount + 1
        lstGames.Selected(0) = True
        cmdJoin_Click
        SwitchMode True
    Else
        Dim lCount As Long
        Dim fFound As Boolean
        
        For lCount = ZERO To mlHostCount
            If moHosts(lCount).AppDesc.guidInstance = oHost.ApplicationDescription.guidInstance Then
                'Save the last time this host was found
                moHosts(lCount).TimeLastFound = GetTickCount
                fFound = True
                Exit For
            End If
        Next
        
        If Not fFound Then 'We need to add this to the list
            ReDim Preserve moHosts(mlHostCount + 1)
            moHosts(mlHostCount + 1).AppDesc = oHost.ApplicationDescription
            moHosts(mlHostCount + 1).AddressHost = oHost.AddressSenderUrl
            moHosts(mlHostCount + 1).AddressDevice = oHost.AddressDeviceUrl
            With oHost.ApplicationDescription
                lstGames.AddItem .SessionName & " - " & CStr(.lCurrentPlayers) & "/" & CStr(.lMaxPlayers) & " - Latency:" & CStr(oHost.lRoundTripLatencyMS) & " ms"
            End With
            'Save the last time this host was found
            moHosts(mlHostCount + 1).TimeLastFound = GetTickCount
            lstGames.ItemData(lstGames.ListCount - 1) = mlHostCount + 1
            mlHostCount = mlHostCount + 1
        Else 'We did find it, update the list
            For lFound = ZERO To lstGames.ListCount - 1
                With oHost.ApplicationDescription
                If lstGames.ItemData(lFound) = lCount Then 'This is it
                    lstGames.List(lFound) = .SessionName & " - " & CStr(.lCurrentPlayers) & "/" & CStr(.lMaxPlayers) & " - Latency:" & CStr(oHost.lRoundTripLatencyMS) & " ms"
                End If
                End With
            Next
        End If
    End If
End Sub


'We will handle all of the msgs here, and report them all back to the callback sub
'in case the caller cares what's going on
Private Sub DirectPlay8Event_AddRemovePlayerGroup(ByVal lMsgID As Long, ByVal lPlayerID As Long, ByVal lGroupID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_AppDesc(fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_AsyncOpComplete(dpnotify As DxVBLibA.DPNMSG_ASYNC_OP_COMPLETE, fRejectMsg As Boolean)
    If dpnotify.AsyncOpHandle = mlEnumAsync Then mlEnumAsync = 0
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_ConnectComplete(dpnotify As DxVBLibA.DPNMSG_CONNECT_COMPLETE, fRejectMsg As Boolean)
    mfGotEvent = True
    If dpnotify.hResultCode = DPNERR_SESSIONFULL Then 'Already too many people joined up
        AgentSpeak myText(8) & vbCr & myText(9), True
        'ShowPane CreateJoinGame
    Else
        'We got our connect complete event
        mfConnectComplete = True
        UpdateChat SystemMsg, "Server connected ... ", Me
        ServerConnected = True
        SendMsg2Server Msg_PlayerInfo

        'VB requires that we must implement *every* member of this interface
    End If
End Sub

Private Sub DirectPlay8Event_CreateGroup(ByVal lGroupID As Long, ByVal lOwnerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_CreatePlayer(ByVal lPlayerID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "CreatePlayer " & lPlayerID

End Sub

Private Sub DirectPlay8Event_DestroyGroup(ByVal lGroupID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_DestroyPlayer(ByVal lPlayerID As Long, ByVal lReason As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    Debug.Print "DestroyPlayer " & lPlayerID

End Sub

Private Sub DirectPlay8Event_EnumHostsQuery(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_QUERY, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_EnumHostsResponse(dpnotify As DxVBLibA.DPNMSG_ENUM_HOSTS_RESPONSE, fRejectMsg As Boolean)
    'Go ahead and add this to our list
    AddHostsToListBox dpnotify
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_HostMigrate(ByVal lNewHostID As Long, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
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
Dim lMsg As ServerMsgTypes, lOffset As Long
Dim Player As SpielerInfo
Dim dpPeer As DPN_PLAYER_INFO
Dim dpAdr As DirectPlay8Address
Dim item As ListItem
Dim i As Long, Antw As Integer
Dim str As String

On Error Resume Next
With dpnotify
    GetDataFromBuffer .ReceivedData, lMsg, LenB(lMsg), lOffset
    'sname = GetName(.idSender)
    'sChat = GetStringFromBuffer(.ReceivedData, lOffset)
    
    Select Case lMsg
                        
        Case Msg_EnumPlayersResponse
            Dim count As Long
            Dim WinPerc As Single
            GetDataFromBuffer .ReceivedData, count, LenB(count), lOffset
            Me.lvPlayers.ListItems.Clear
            For i = 1 To count
                
                Player.SpielerName = GetStringFromBuffer(.ReceivedData, lOffset)
                Player.SpielerLevel = CInt(GetStringFromBuffer(.ReceivedData, lOffset))
                Player.Points = GetStringFromBuffer(.ReceivedData, lOffset)
                Player.IP_Adress = GetStringFromBuffer(.ReceivedData, lOffset)
                Player.Status = CInt(GetStringFromBuffer(.ReceivedData, lOffset))
                Player.GlobalID = GetStringFromBuffer(.ReceivedData, lOffset)
                Player.SpielOption = CInt(GetStringFromBuffer(.ReceivedData, lOffset))
                Player.AvatarFileName = GetStringFromBuffer(.ReceivedData, lOffset)
                
                Set item = Me.lvPlayers.ListItems.Add(, "K" & Player.GlobalID, Player.SpielerName, , Player.SpielerLevel + 1)
                
                item.SubItems(1) = Player.GlobalID
                item.SubItems(2) = strPlayerLevel(Player.SpielerLevel)
                item.SubItems(3) = Player.Points
                item.SubItems(4) = Player.IP_Adress
                item.SubItems(5) = strPlayerStatus(Player.Status)
                item.SubItems(6) = strSpielOption(Player.SpielOption)
                item.SubItems(8) = Player.AvatarFileName
                
                WinPerc = GetMPGewinnQuote(Player)
                If WinPerc > -1 Then item.SubItems(7) = Format$(WinPerc, PercFormat)
                If WinPerc > 0.5 Then
                    item.ForeColor = vbYellow
                ElseIf WinPerc > -1 Then
                    item.ForeColor = vbRed
                End If
                'Namen in die Liste der Autozuvervollständigen Namen eintragen
                ReDim Preserve Names(1 To count)
                Names(i) = Player.SpielerName
                
                'Ist freund oder feind :-)
                Dim FI As FriendIgnore
                FI = IsInList(Player.GlobalID)
                If FI = FriendIgnore.FriendS Then
                    item.Bold = True
                ElseIf FI = FriendIgnore.Ignore Then
                    item.Ghosted = True
                Else
                    'nix
                End If
            Next
            UpdateChat SystemMsg, "Spielerliste erhalten", ServerEventsForm
            
            'Größe der Columns anpassen
            ColumnAutoSize Me.lvPlayers, 1, True
            Me.lvPlayers.ColumnHeaders(2).Width = 0
            ColumnAutoSize Me.lvPlayers, 3, True
            ColumnAutoSize Me.lvPlayers, 4, True
            Me.lvPlayers.ColumnHeaders(5).Width = 0
            ColumnAutoSize Me.lvPlayers, 6, True
            ColumnAutoSize Me.lvPlayers, 7, True
            
        Case Msg_PlayerInfo_OK
        
'            SendMsg2Server Msg_EnumPlayers
            UpdateChat SystemMsg, "Anmeldung OK", Me
'           nix, spielerliste wird von server zugeschickt
                        
        Case Msg_PlayerInfo_NOK
            SaveSetting AppExeName, cstrDefault, cstrRegID, cstrRegID
            bool_isRegistered = False
            str = "Der Server hat Ihre Registrierung abgelehnt !" & vbCr & _
                "Bitte registrieren Sie sich korrekt oder kontaktieren Sie das playseven-Team für Hilfe"
            UpdateChat SystemMsg, str, Me
            AgentSpeak str, True
            CleanUpServer
            'Me.GoUnload
            
        Case Msg_NoOtherPlayers
        
            str = "Currently there are no other players available on this server." & vbCr & _
                "If existing, please select an other server for playing or login again later."
            UpdateChat SystemMsg, str, Me
            AgentSpeak str, True
            
        Case Msg_StartGame
            
            If Not LaufendesSpiel Then
                Dim Key As String
                
                Key = GetStringFromBuffer(.ReceivedData, lOffset)
                
                If ((Me.chkNoHigherLevel.Value = vbChecked) And (getLevelFromStr(Me.lvPlayers.ListItems("K" & Key).SubItems(2)) > AktuellerSpieler.SpielerLevel)) Then
                    SendMsg2Server Msg_GameStarted_NOK, Key
                    UpdateChat SystemMsg, "Kontakt zu " & Me.lvPlayers.ListItems("K" & Key) & " wurde aufgrund höheren Levels abgelehnt.", Me
                    
                ElseIf AktuellerSpieler.SpielOption <> getSpielOption(Me.lvPlayers.ListItems("K" & Key).SubItems(6)) Then
                    SendMsg2Server Msg_GameStarted_NOK, Key
                    UpdateChat SystemMsg, "Kontakt zu " & Me.lvPlayers.ListItems("K" & Key) & " wurde aufgrund unterschiedlichen SpielOptionen abgelehnt.", Me
                
                ElseIf AktuellerSpieler.Status = PlayingMP Then
                    SendMsg2Server Msg_GameStarted_NOK, Key
                    UpdateChat SystemMsg, "Kontakt zu " & Me.lvPlayers.ListItems("K" & Key) & " wurde aufgrund laufendem Spiels abgelehnt.", Me
                ElseIf Me.lvPlayers.ListItems("K" & Key).Ghosted Then
                    SendMsg2Server Msg_GameStarted_NOK, Key
                    UpdateChat SystemMsg, "Kontakt zu " & Me.lvPlayers.ListItems("K" & Key) & " wurde aufgrund Ignorelist abgelehnt.", Me
                Else
                    'to do Gegnerinformation im Fragefenster anzeigen
                    PlaySound KissSound
                    Me.lvPlayers.ListItems("K" & Key).EnsureVisible
                    
                    Gegner.SpielerName = Me.lvPlayers.ListItems("K" & Key)
                    UpdateChat SystemMsg, Gegner.SpielerName & " möchte einen Spielraum betreten", Me
                    Load frmUserAnswer
                    
                    Antw = frmUserAnswer.Antwort
                    AgentSpeak Me.lvPlayers.ListItems("K" & Key) & " möchte Kontakt mit Ihnen aufnehmen"
                    If Antw = vbYes Then
                        StartP2PbyServer Key
                    Else
                        MakeHost = False
                        'SendMsg2Server IIf(antw = vbNo, Msg_GameStarted_NOK, Msg_NoAnswer), Key
                        SendMsg2Server Msg_GameStarted_NOK, Key
                    End If
                End If
            End If
            
        Case Msg_GameStarted_NOK, Msg_NoAnswer
        
            Key = GetStringFromBuffer(.ReceivedData, lOffset)
            If Gegner.GlobalID = Key Then
                Unload frmUserReq
                AgentSpeak Me.lvPlayers.ListItems("K" & Key) & _
                    IIf(lMsg = Msg_GameStarted_NOK, " möchte kein Kontakt.", " hat nicht geantwortet"), True
                Gegner.IP_Adress = vbNullString
            End If
            
        Case Msg_GameStarted_OK
        
            MakeHost = False
            AktuellerSpieler.Status = PlayingMP

            SendMsg2Server Msg_PlayerInfo
            Unload frmUserReq
            Me.Hide
                       
            Init_DPP Gegner.IP_Adress
            
        Case Msg_PopUpMsg
            
            str = GetStringFromBuffer(.ReceivedData, lOffset)
            AgentSpeak str, True
            UpdateChat SystemMsg, str, Me

        Case Msg_UpdateGame
        
            str = GetStringFromBuffer(.ReceivedData, lOffset)
            UpdateChat SystemMsg, "Neue Version vorhanden ! " & str, Me
            GetNewVersion (str)
            
        Case msg_gamestart_cancel
            Key = GetStringFromBuffer(.ReceivedData, lOffset)
            UpdateChat SystemMsg, Me.lvPlayers.ListItems("K" & Key) & " nimmt Kontaktversuch zurück.", Me
            'Frage fenster schliessen
            Unload frmUserAnswer
            Gegner.GlobalID = vbNullString
            
        Case Msg_Chat
        
            str = GetStringFromBuffer(.ReceivedData, lOffset)
            UpdateChat TextMsg, str, ServerEventsForm
        
        Case Msg_SystemMsg
        
            str = GetStringFromBuffer(.ReceivedData, lOffset)
            UpdateChat SystemMsg, str, ServerEventsForm
            
        Case Else
        
            AgentSpeak "ServerMessage Nr: " & lMsg & " nicht bekannt. Bitte updaten Sie das Spiel."

    End Select
End With
Exit Sub

ERRHand:
If Err.Number = 5 Then
    Resume Next
Else
    If ErrorBox("ServerConnect:DirectPlay8Event_Receive", Err) Then Resume Next
End If


End Sub

Private Sub StartP2PbyServer(Key As String)
MakeHost = True
AktuellerSpieler.Status = PlayingMP
Me.Hide

Gegner.SpielerName = Me.lvPlayers.ListItems("K" & Key)
Gegner.GlobalID = Me.lvPlayers.ListItems("K" & Key).SubItems(1)
Gegner.SpielerLevel = getPlayerLevel(Me.lvPlayers.ListItems("K" & Key).SubItems(2))
Gegner.AvatarFileName = Me.lvPlayers.ListItems("K" & Key).SubItems(8)

'HostStarten
Init_DPP "MakeHost"
SendMsg2Server Msg_GameStarted_OK, Key
SendMsg2Server Msg_PlayerInfo

End Sub


Private Sub DirectPlay8Event_SendComplete(dpnotify As DxVBLibA.DPNMSG_SEND_COMPLETE, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
End Sub

Private Sub DirectPlay8Event_TerminateSession(dpnotify As DxVBLibA.DPNMSG_TERMINATE_SESSION, fRejectMsg As Boolean)
    'VB requires that we must implement *every* member of this interface
    
    If dpnotify.hResultCode = DPNERR_HOSTTERMINATEDSESSION Then
        AgentSpeak "You have been kicked from the Server !", True
    Else
        AgentSpeak "Serversession has been lost. Please reconnect", True
    End If
    '    CleanUpServer
    SwitchMode False
    cmdRefresh_Click

End Sub

Private Sub Form_Load()
On Error GoTo ERRHand
    SetzFensterMittig Me
    SetBackGround Me
    HideTitleBar Me
    makeRoundEdges Me
    
    LoadObjectText "DPlayConnect", myText()
    Me.lvPlayers.ColumnHeaders(2).Width = 0
    AktuellerSpieler.Status = Idle
    
    Me.chkLog = GetSetting(AppExeName, cstrOptions, cstrMakeLog, vbUnchecked)
    Me.chkNoHigherLevel = GetSetting(AppExeName, cstrOptions, cstrNoHigherLevel, vbUnchecked)
    Me.chkShowSystemMsg = GetSetting(AppExeName, cstrOptions, cstrshowSystemMessage, vbUnchecked)
    Me.optSpielOption(GetSetting(AppExeName, cstrOptions, cstrSpielOption, ZERO)).Value = True
    
    strPlayerStatus(PlayerStatus.PlayingMP) = "Spielt MP"
    strPlayerStatus(PlayerStatus.PlayingSP) = "Spielt SP"
    strPlayerStatus(PlayerStatus.Idle) = "ready"
       
    strSpielOption(SpielOptionen.Liga) = "Liga"
    strSpielOption(SpielOptionen.Freundschaft) = "Freundschaft"
       
    Me.Caption = myText(4)
    Me.lvPlayers.Icons = Me.ImageList1
    Me.lvPlayers.SmallIcons = Me.ImageList1
    Me.cmdRefresh = True
'    AgentSpeak "In der Betaphase ist der Server nur von 15:00 bis 24:00 Uhr erreichbar.", True
    
   
Exit Sub
ERRHand:
If ErrorBox("ServerConnect:FormLoad", Err) Then Resume Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = vbWhite
    Me.lblMin.ForeColor = vbWhite
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not mfCanUnload Then Cancel = 1
    mfDoneWiz = True
'    Me.Hide
End Sub

Private Sub Form_Resize()
If Me.WindowState <> vbMinimized Then
    Me.Move Me.Left, Me.Top, 8835, 7215
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    Clean up our address

    If Not moDPA Is Nothing Then Set moDPA = Nothing
    If Not moDPC Is Nothing Then
        moDPC.Close
'        Set moDPC = Nothing
    End If
'    If Not moDX Is Nothing Then Set moDX = Nothing
    frmSplash.Show
End Sub




Private Sub lbl_Cancel_Click()
'    GoUnload
'
    mfComplete = True
    CleanUpServer
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

Private Sub lstGames_DblClick()
    cmdJoin_Click
End Sub

'Public Property Get IsHost() As Boolean
'    IsHost = mfHost
'End Property

Public Property Get SessionName() As String
    SessionName = sGameName
End Property

'Public Property Get UserName() As String
'    UserName = sUser
'End Property

Public Sub GoUnload()
    tmrExpire.Enabled = False
    mfCanUnload = True
    CloseForm Me
    Unload Me
End Sub

'Public Sub RegisterCallback(oCallback As DirectPlay8Event)
'    Set moCallback = oCallback
'End Sub

'Public Property Get NumPlayers() As Long
'    NumPlayers = mlNumPlayers
'End Property

Private Sub lvPlayers_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
With Me.lvPlayers
    .SortKey = ColumnHeader.Index - 1
    'SortOrder bestimmen Asc oder Desc
    If .SortOrder = lvwAscending Then
        .SortOrder = lvwDescending
    Else
        .SortOrder = lvwAscending
    End If
    'Sort anstossen
    .Sorted = True
    'Zeiger auf 1. Zeile und scrollen
    .ListItems(1).Selected = True
    .ListItems(1).EnsureVisible
End With
End Sub

Private Sub lvPlayers_DblClick()

    Gegner.GlobalID = Me.lvPlayers.SelectedItem.SubItems(1)
    
    'wenn man mit sich selbst spielen will--> raus
    If Gegner.GlobalID = AktuellerSpieler.GlobalID Then Exit Sub
    'wenn gegner schon netwzerk spielt --> raus
    If getPlayerStatus(Me.lvPlayers.SelectedItem.SubItems(5)) = PlayerStatus.PlayingMP Then Exit Sub
    'wenn gegner in andrer Option spielt --> raus
    If getSpielOption(Me.lvPlayers.SelectedItem.SubItems(6)) <> AktuellerSpieler.SpielOption Then
        AgentSpeak Me.lvPlayers.SelectedItem & " spielt in der Option " & Me.lvPlayers.SelectedItem.SubItems(6), True
        Exit Sub
    End If
    
    Gegner.IP_Adress = Me.lvPlayers.SelectedItem.SubItems(4)
    Gegner.SpielerName = Me.lvPlayers.SelectedItem
    Gegner.SpielerLevel = getPlayerLevel(Me.lvPlayers.SelectedItem.SubItems(2))
    Gegner.AvatarFileName = Me.lvPlayers.SelectedItem.SubItems(8)
    SendMsg2Server Msg_StartGame
    frmUserReq.GlobalID = Gegner.GlobalID
    frmUserReq.lblUserInfo = Gegner.SpielerName & " Level(" & Me.lvPlayers.SelectedItem.SubItems(2) & ")"
    frmUserReq.Show 1
End Sub

Private Sub lvPlayers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    
    Me.mnuAddFriends.Visible = True
    Me.mnuIgnore.Visible = True
    Me.mnuDeleteFriend.Visible = True
    Me.mnuUnIgnore.Visible = True
    
    If Me.lvPlayers.SelectedItem.Bold Then
        Me.mnuAddFriends.Visible = False
        Me.mnuUnIgnore.Visible = False
    ElseIf Me.lvPlayers.SelectedItem.Ghosted Then
        Me.mnuIgnore.Visible = False
        Me.mnuDeleteFriend.Visible = False
    Else
        Me.mnuUnIgnore.Visible = False
        Me.mnuDeleteFriend.Visible = False
    End If
    
    If Not Me.lvPlayers.SelectedItem.SubItems(1) = AktuellerSpieler.GlobalID Then Me.PopupMenu mnuPopUp
    
End If
End Sub

Private Sub mnuAddFriends_Click()
    If AlterList(Me.lvPlayers.SelectedItem.SubItems(1), Me.lvPlayers.SelectedItem, TableList.FriendS, adder) Then
        Me.lvPlayers.SelectedItem.Bold = True
    End If
End Sub

Private Sub mnuContact_Click()
    lvPlayers_DblClick
End Sub

Private Sub mnuDeleteFriend_Click()

If AlterList(Me.lvPlayers.SelectedItem.SubItems(1), Me.lvPlayers.SelectedItem, TableList.FriendS, Deleter) Then
    Me.lvPlayers.SelectedItem.Bold = False
End If

End Sub

Private Sub mnuIgnore_Click()
    If AlterList(Me.lvPlayers.SelectedItem.SubItems(1), Me.lvPlayers.SelectedItem, TableList.Ignore, adder) Then
        Me.lvPlayers.SelectedItem.Ghosted = True
    End If
End Sub

Private Sub mnuUnIgnore_Click()
    
    If AlterList(Me.lvPlayers.SelectedItem.SubItems(1), Me.lvPlayers.SelectedItem, TableList.Ignore, Deleter) Then
        Me.lvPlayers.SelectedItem.Ghosted = False
    End If

End Sub

Private Sub optSpielOption_Click(Index As Integer)
    AktuellerSpieler.SpielOption = Index
    SaveSetting AppExeName, cstrOptions, cstrSpielOption, Index
    If ServerConnected Then
        SendMsg2Server Msg_PlayerInfo
    End If
End Sub



Private Sub rtxtSend_GotFocus()
'In anderen controls TABStaop ausschalten
'Tab wird benötigt um Namenautovervollsändigen zu ermöglichen
    EnableTabStop False
End Sub

Private Sub rtxtSend_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    If rtxtSend.Text <> vbNullString Then 'Make sure they are trying to send something
        
        If Me.rtxtSend.Text = "/clear" Then
            Me.rtxtChat.Text = vbNullString
        Else
            'Send this message to everyone
            If Not ServerConnected Then Exit Sub
            SendMsg2Server Msg_Chat, rtxtSend.Text
            'UpdateChat TextMsg, "<" & gsUserName & ">" & gstrSpace & txtSend.Text
        End If
        
        rtxtSend.Text = vbNullString
        KeyAscii = 0
        'dpp.SendTo DPNID_ALL_PLAYERS_GROUP, oBuf, 0, DPNSEND_NOLOOPBACK
    End If
ElseIf KeyAscii = vbKeyTab Then
    KeyAscii = 0
    Me.rtxtSend.Text = AutoUpdateName(Me.rtxtSend.Text)
    Me.rtxtSend.SelStart = Len(Me.rtxtSend.Text)
End If

End Sub

Private Sub rtxtSend_LostFocus()
    EnableTabStop True
End Sub

Private Sub EnableTabStop(s As Boolean)
'In anderen controls TABStaop ein/ausschalten
Dim ctl As Control
On Error Resume Next

For Each ctl In Me
    ctl.TabStop = s
Next
End Sub

Private Sub tmrExpire_Timer()
    
    'We need to periodically expire the hosts that are in this list in case they are
    'no longer hosting or what have you.
    Dim lCount As Long, lIndex As Long
    Dim lInner As Long
    
    On Error GoTo LeaveSub 'If there are no hosts, just go
    For lCount = ZERO To UBound(moHosts)
        If (GetTickCount - moHosts(lCount).TimeLastFound) > HOST_EXPIRE_THRESHHOLD Then
            'Yup, this guy expired.. remove him from the list
            For lIndex = lstGames.ListCount - 1 To ZERO Step -1
                If lstGames.ItemData(lIndex) = lCount Then 'this is the one
                    lstGames.RemoveItem lIndex
                End If
            Next
            moHosts(lCount).AddressDevice = vbNullString
            moHosts(lCount).AddressHost = vbNullString
            'Now we need an internal loop to 'remove' all of the old hosts info
            For lInner = lCount + 1 To UBound(moHosts)
                moHosts(lInner - 1).AddressDevice = moHosts(lInner).AddressDevice
                moHosts(lInner - 1).AddressHost = moHosts(lInner).AddressHost
                moHosts(lInner - 1).AppDesc = moHosts(lInner).AppDesc
                moHosts(lInner - 1).TimeLastFound = moHosts(lInner).TimeLastFound
            Next
            'Now we need to decrement each of the remaining items in the listbox
            For lIndex = lstGames.ListCount - 1 To ZERO Step -1
                If lstGames.ItemData(lIndex) > lCount Then 'decrement this one
                    lstGames.ItemData(lIndex) = lstGames.ItemData(lIndex) - 1
                End If
            Next
            mlHostCount = mlHostCount - 1
            If UBound(moHosts) > ZERO Then
                ReDim Preserve moHosts(UBound(moHosts) - 1)
            Else
                Erase moHosts 'This will just erase the memory
            End If
        End If
    Next
LeaveSub:
End Sub
'Private Function GetWord(Rich As RichTextBox, ByVal X&, _
'                         ByVal Y&) As String
'  Dim pos&, P1&, P2&
'  Dim Char$
'  Dim MPointer As POINTAPI
'
'    '### Position des Textzeichens unter dem Mauszeiger auslesen
'    MPointer.X = X \ Screen.TwipsPerPixelX
'    MPointer.Y = Y \ Screen.TwipsPerPixelY
'    pos = SendMessage(Rich.hWnd, EM_CHARFROMPOS, 0&, MPointer)
'    If pos <= 0 Then Exit Function
'
'    '### Wortanfang finden
'    For P1 = pos To 1 Step -1
'      Char = Mid$(Rich.Text, P1, 1)
'      If Not CheckChar(Char) Then Exit For
'    Next P1
'    P1 = P1 + 1
'
'    '### Wortende finden
'    For P2 = pos To Len(Rich.Text)
'      Char = Mid$(Rich.Text, P2, 1)
'      If Not CheckChar(Char) Then Exit For
'    Next P2
'    P2 = P2 - 1
'
'    If P1 < P2 Then GetWord = Mid$(Rich.Text, P1, P2 - P1 + 1)
'End Function
'
'Private Function CheckChar(ByVal Char$) As Boolean
'  '### Testen auf Trennzeichen eines Wortes
'  If ((Char >= "0" And Char <= "9") Or _
'      (Char >= "a" And Char <= "z") Or _
'      (Char >= "A" And Char <= "Z") Or _
'      (InStr("ÄöüÄÖÜß", Char))) Then CheckChar = True
'End Function
Public Sub SendMsg2Server(msgType As ServerMsgTypes, Optional MSG)
Dim oMsg() As Byte, lOffset As Long
Dim str As String

lOffset = NewBuffer(oMsg)
AddDataToBuffer oMsg, msgType, LenB(msgType), lOffset

Select Case msgType
    
    Case Msg_PlayerInfo
        Dim BufferSize As Long
        AktuellerSpieler.IP_Adress = moDPA.GetURL
        AddStringToBuffer oMsg, AktuellerSpieler.GlobalID, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.IP_Adress, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.Points, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.SpielerLevel, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.SpielerName, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.ClientID, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.RegID, lOffset
        AddStringToBuffer oMsg, StartAnz, lOffset
        AddStringToBuffer oMsg, WindowsVersion, lOffset
        AddStringToBuffer oMsg, App.Major & gstrDot & App.Minor & gstrDot & App.Revision, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.Status, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.SpielOption, lOffset
        AddStringToBuffer oMsg, AktuellerSpieler.AvatarFileName, lOffset
        
        UpdateChat SystemMsg, "Sende Spielerinfo", Me
        
    Case Msg_GameStarted_NOK, Msg_GameStarted_OK, Msg_Chat
        AddStringToBuffer oMsg, MSG, lOffset
    Case msg_gamestart_cancel, Msg_NoAnswer
        AddStringToBuffer oMsg, MSG, lOffset
        Gegner.GlobalID = vbNullString
        Gegner.IP_Adress = vbNullString
    Case Msg_EnumPlayers, Msg_PlayerWon, Msg_PlayerLost
        'nix
    Case Msg_StartGame
        AddStringToBuffer oMsg, Gegner.GlobalID, lOffset
    Case Else
        MsgBox "Unbehandelter ServerMsgType. Please update"
End Select

dpc.Send oMsg, 0, DPNSEND_NOLOOPBACK + DPNSEND_GUARANTEED

End Sub

Public Sub GetNewVersion(Url As String)
Dim Antw As Integer
If Url = vbNullString Then Exit Sub

Antw = AgentQuestion("Eine neue Version ist verfügbar unter: " & vbCr & Url & " !" & vbCr & _
    " Möchten Sie diese jetzt downloaden ?" & vbCr & vbCr & _
    "(Sollte der automatische Download nicht funktionieren, besuchen Sie bitte " & vbCr & _
    "unsere Website und downloaden sich die neuste Version manuell)", "New Version ..")
If Antw = vbYes Then
    Go2URL Url
End If
End Sub
