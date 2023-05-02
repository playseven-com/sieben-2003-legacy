VERSION 5.00
Object = "{F5BE8BC2-7DE6-11D0-91FE-00C04FD701A5}#2.0#0"; "agentctl.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{593776BD-DE85-11D3-9C50-8790D659BC67}#6.0#0"; "MsgBalloon6.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   6615
   ClientLeft      =   3600
   ClientTop       =   810
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   8700
   Begin VB.Timer ShareWareTimer 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   8160
      Top             =   450
   End
   Begin VB.CommandButton cmdTip 
      BackColor       =   &H00FFFF00&
      Caption         =   "HELP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5880
      MaskColor       =   &H00000000&
      Style           =   1  'Grafisch
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   6240
      Visible         =   0   'False
      Width           =   525
   End
   Begin MsgBalloon6.Balloon AgentBalloon 
      Left            =   7650
      Top             =   960
      _ExtentX        =   820
      _ExtentY        =   767
      ButtonPicture   =   "Form1.frx":08CA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonsCaptions =   "*&Ok*&Cancel*&Abort*&Retry*&Ignore*&Yes*&No*"
      URLButtonPicture=   ""
   End
   Begin VB.Timer StichEndeTimer 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   7200
      Top             =   450
   End
   Begin VB.CommandButton cmdSiebenSuch 
      Caption         =   "7 finden"
      Height          =   375
      Left            =   5910
      TabIndex        =   2
      Top             =   480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CheckBox chkAudio 
      BackColor       =   &H00004000&
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
      Height          =   190
      Left            =   7850
      TabIndex        =   1
      Top             =   6390
      Value           =   1  'Aktiviert
      Width           =   215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   25
      Left            =   6750
      Top             =   450
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6990
      Top             =   900
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
            Picture         =   "Form1.frx":3DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":50A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5455
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":77E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9CBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C106
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":E7B6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox PictureDummy 
      Appearance      =   0  '2D
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   315
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   315
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Undurchsichtig
      Height          =   6135
      Left            =   5790
      Shape           =   4  'Gerundetes Rechteck
      Top             =   450
      Width           =   75
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   8
      Left            =   3345
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   7
      Left            =   2970
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   6
      Left            =   2595
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   5
      Left            =   2220
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   4
      Left            =   1845
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   52
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   51
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   50
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   49
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   48
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   47
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   46
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   45
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   44
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   43
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   42
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   41
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   40
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   39
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   38
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   37
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   36
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   35
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   34
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   33
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   32
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   31
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   30
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   29
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   28
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   27
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   26
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   25
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   24
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   23
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   22
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   21
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   20
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   19
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   18
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   17
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   16
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   15
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   14
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   13
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   12
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   11
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   10
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   9
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   8
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   7
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   6
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   5
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   4
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   3
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   2
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picHaufen 
      Height          =   1125
      Index           =   1
      Left            =   6000
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   2880
      Width           =   810
   End
   Begin VB.Image picKarte 
      Height          =   1125
      Index           =   0
      Left            =   130
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   810
   End
   Begin VB.Image picKarte 
      Height          =   1125
      Index           =   1
      Left            =   1370
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   810
   End
   Begin VB.Image picKarte 
      Height          =   1125
      Index           =   2
      Left            =   2600
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   810
   End
   Begin VB.Image picKarte 
      Height          =   1125
      Index           =   3
      Left            =   3820
      MousePointer    =   10  'Aufwärtspfeil
      Stretch         =   -1  'True
      Top             =   4680
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   3
      Left            =   1470
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   2
      Left            =   1095
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Label lblMinMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
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
      Index           =   0
      Left            =   7920
      TabIndex        =   38
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   300
   End
   Begin VB.Image picKarteStich 
      Height          =   1125
      Index           =   1
      Left            =   720
      MousePointer    =   1  'Pfeil
      Stretch         =   -1  'True
      Top             =   2655
      Visible         =   0   'False
      Width           =   810
   End
   Begin VB.Image picLevel 
      BorderStyle     =   1  'Fest Einfach
      Height          =   420
      Left            =   5250
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   420
   End
   Begin VB.Image picKarteComp 
      Height          =   1125
      Index           =   2
      Left            =   2640
      MousePointer    =   12  'Nicht ablegen
      Stretch         =   -1  'True
      Top             =   600
      Width           =   810
   End
   Begin VB.Image picKarteComp 
      Height          =   1125
      Index           =   0
      Left            =   120
      MousePointer    =   12  'Nicht ablegen
      Stretch         =   -1  'True
      Top             =   585
      Width           =   810
   End
   Begin VB.Image picKarteComp 
      Height          =   1125
      Index           =   1
      Left            =   1365
      MousePointer    =   12  'Nicht ablegen
      Stretch         =   -1  'True
      Top             =   585
      Width           =   810
   End
   Begin VB.Image picKarteComp 
      Height          =   1125
      Index           =   3
      Left            =   3810
      MousePointer    =   12  'Nicht ablegen
      Stretch         =   -1  'True
      Top             =   585
      Width           =   810
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      Visible         =   0   'False
      X1              =   4830
      X2              =   5730
      Y1              =   5250
      Y2              =   5250
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   6  'Innen ausgefüllt
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   4830
      X2              =   5730
      Y1              =   1110
      Y2              =   1110
   End
   Begin VB.Label lblScore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   2430
      TabIndex        =   36
      Top             =   6300
      Width           =   1305
   End
   Begin VB.Label lblHighScore 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1350
      TabIndex        =   35
      Top             =   6300
      Width           =   1005
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Highscore"
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
      Height          =   255
      Left            =   1440
      TabIndex        =   34
      Top             =   6090
      Width           =   885
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Score"
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
      Height          =   255
      Left            =   2850
      TabIndex        =   33
      Top             =   6090
      Width           =   495
   End
   Begin VB.Label lblLevel 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   3720
      TabIndex        =   32
      Top             =   6120
      Width           =   1185
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "???"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   30
      TabIndex        =   31
      Top             =   6090
      Width           =   1275
   End
   Begin VB.Label lblRPGegner 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Rundenpunkte Computer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1035
      Left            =   7530
      TabIndex        =   26
      Top             =   2130
      Width           =   1125
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRPunkteSpieler 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Rundenpunkte Spieler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   0
      Left            =   7530
      TabIndex        =   25
      Top             =   3255
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSpielerRundenPunkte 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   735
      Index           =   0
      Left            =   7545
      TabIndex        =   24
      Top             =   3585
      Width           =   975
   End
   Begin VB.Label lblComputerRundenPunkte 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Index           =   0
      Left            =   7515
      TabIndex        =   23
      Top             =   2415
      Width           =   975
   End
   Begin VB.Label cmdStichEnde 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Stich Ende"
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
      Left            =   645
      TabIndex        =   22
      Top             =   3930
      Visible         =   0   'False
      Width           =   3570
   End
   Begin VB.Label lblSpielerPunkte 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   585
      Index           =   0
      Left            =   4950
      TabIndex        =   18
      Top             =   4530
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblSpielerPunkte 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   585
      Index           =   1
      Left            =   4965
      TabIndex        =   20
      Top             =   4575
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblPunkteSpieler 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Punkte Spieler"
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
      Height          =   375
      Index           =   1
      Left            =   4800
      TabIndex        =   19
      Top             =   5340
      Visible         =   0   'False
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblComputerPunkte 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   705
      Index           =   0
      Left            =   4980
      TabIndex        =   15
      Top             =   1140
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPGegner 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Punkte Gegner"
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
      Height          =   555
      Left            =   4860
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Shape rectIndikator 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1305
      Index           =   0
      Left            =   45
      Top             =   4575
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape rectIndikator 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1305
      Index           =   1
      Left            =   1275
      Top             =   4575
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape rectIndikator 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Height          =   1305
      Index           =   2
      Left            =   2505
      Top             =   4575
      Visible         =   0   'False
      Width           =   1005
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
      Left            =   7590
      TabIndex        =   11
      ToolTipText     =   "Minimize"
      Top             =   0
      Width           =   315
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
      Left            =   7725
      TabIndex        =   12
      ToolTipText     =   "Exit"
      Top             =   60
      Width           =   120
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   " Audio"
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
      Left            =   8070
      TabIndex        =   3
      Top             =   6390
      Width           =   645
   End
   Begin VB.Label lblExit 
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
      Height          =   435
      Left            =   8310
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   405
   End
   Begin AgentObjectsCtl.Agent Agent1 
      Left            =   7620
      Top             =   450
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "playseven.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   6060
      TabIndex        =   8
      Top             =   90
      Width           =   1635
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "7 s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Index           =   1
      Left            =   1740
      TabIndex        =   5
      Top             =   1830
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "7 s"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Index           =   0
      Left            =   1800
      TabIndex        =   4
      Top             =   1890
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label lblExitBack 
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
      Left            =   8370
      TabIndex        =   10
      ToolTipText     =   "Exit"
      Top             =   60
      Width           =   225
   End
   Begin VB.Label lblComputerPunkte 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   705
      Index           =   1
      Left            =   5025
      TabIndex        =   16
      Top             =   1185
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblPGegner1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Punkte Gegner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   4920
      TabIndex        =   17
      Top             =   660
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label lblPunkteSpieler 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Punkte Spieler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   0
      Left            =   4860
      TabIndex        =   21
      Top             =   5385
      Visible         =   0   'False
      Width           =   1005
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblRPGegner1 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Rundenpunkte Computer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   855
      Left            =   7545
      TabIndex        =   29
      Top             =   2160
      Width           =   1155
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblComputerRundenPunkte 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   735
      Index           =   1
      Left            =   7560
      TabIndex        =   28
      Top             =   2445
      Width           =   975
   End
   Begin VB.Label lblRPunkteSpieler 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Rundenpunkte Spieler"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Index           =   1
      Left            =   7560
      TabIndex        =   30
      Top             =   3300
      Width           =   1095
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblSpielerRundenPunkte 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   765
      Index           =   1
      Left            =   7605
      TabIndex        =   27
      Top             =   3615
      Width           =   945
   End
   Begin VB.Shape shpPlayerInfo 
      FillStyle       =   0  'Ausgefüllt
      Height          =   585
      Left            =   0
      Top             =   6030
      Width           =   5745
   End
   Begin VB.Shape shpComputerPunkte 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1305
      Left            =   4830
      Shape           =   4  'Gerundetes Rechteck
      Top             =   540
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape shpSpielerPunkte 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1425
      Left            =   4830
      Shape           =   4  'Gerundetes Rechteck
      Top             =   4500
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Shape shpRundenPunkte 
      BorderColor     =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   2535
      Left            =   7440
      Shape           =   4  'Gerundetes Rechteck
      Top             =   1980
      Width           =   1245
   End
   Begin VB.Label lblMinMax 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "m"
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
      Index           =   1
      Left            =   7995
      TabIndex        =   39
      ToolTipText     =   "Exit"
      Top             =   60
      Width           =   300
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Sieben Version 2.0.215"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   435
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   8685
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1545
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   1800
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1545
      Index           =   0
      Left            =   80
      TabIndex        =   6
      Top             =   1870
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Shape rectIndikator 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      BorderColor     =   &H00FFFFFF&
      BorderStyle     =   5  'Strich-Punkt-Punkt
      BorderWidth     =   2
      Height          =   1305
      Index           =   3
      Left            =   3735
      Top             =   4590
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Shape rectStichIndikator 
      BackStyle       =   1  'Undurchsichtig
      BorderWidth     =   2
      Height          =   1330
      Left            =   630
      Top             =   2550
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Menu menGame 
      Caption         =   "Spiel"
      Begin VB.Menu menSpielerAuswahl 
         Caption         =   "Spieler &Auswahl"
         Shortcut        =   ^P
      End
      Begin VB.Menu menStartGame 
         Caption         =   "Spiel &neu starten"
         Shortcut        =   ^N
      End
      Begin VB.Menu menEndGame 
         Caption         =   "Spiel &beenden"
      End
      Begin VB.Menu menLine1 
         Caption         =   "-"
      End
      Begin VB.Menu menGetFreeRegKey 
         Caption         =   "freien Registrierungsschlüssel anfordern"
      End
      Begin VB.Menu menWriteFreeRegKey 
         Caption         =   "freien Registrierungsschlüssel eingeben"
      End
      Begin VB.Menu menWriteRegkey 
         Caption         =   "Registrierungsschlüssel eingeben"
      End
   End
   Begin VB.Menu menOptionen 
      Caption         =   "Optionen"
      Visible         =   0   'False
      Begin VB.Menu menSchmulen 
         Caption         =   "Kiebitzen"
         Shortcut        =   ^K
      End
      Begin VB.Menu menShowTips 
         Caption         =   "Tips beim Start anzeigen"
         Checked         =   -1  'True
      End
      Begin VB.Menu menAgentTips 
         Caption         =   "Agent gibt Tips"
         Checked         =   -1  'True
      End
      Begin VB.Menu menShowIndikator 
         Caption         =   "Kartenindikator anzeigen"
         Checked         =   -1  'True
         Shortcut        =   ^I
      End
      Begin VB.Menu menStatistik 
         Caption         =   "Statistik einblenden"
         Shortcut        =   ^E
      End
      Begin VB.Menu menLine2 
         Caption         =   "-"
      End
      Begin VB.Menu menChooseAgent 
         Caption         =   "Agent auswählen"
         Shortcut        =   ^W
      End
      Begin VB.Menu men_useAgent 
         Caption         =   "Agent anzeigen"
         Shortcut        =   ^A
      End
      Begin VB.Menu menAgentTTS 
         Caption         =   "Sprachausgabemodul für Agenten downloaden"
      End
      Begin VB.Menu menAgentDownload 
         Caption         =   "zus. Agenten downloaden"
      End
      Begin VB.Menu menAgentTalkChat 
         Caption         =   "Agent spricht Chat"
         Shortcut        =   ^T
      End
      Begin VB.Menu menLine3 
         Caption         =   "-"
      End
      Begin VB.Menu menAudio 
         Caption         =   "Audio"
         Checked         =   -1  'True
      End
      Begin VB.Menu menBackground 
         Caption         =   "Hintergrund"
         Begin VB.Menu menBackStandard 
            Caption         =   "Standard"
            Index           =   0
         End
         Begin VB.Menu menBackStandard 
            Caption         =   "Hintergründe downloaden"
            Index           =   1
         End
      End
      Begin VB.Menu menLine4 
         Caption         =   "-"
      End
      Begin VB.Menu menLanguage 
         Caption         =   "Sprache"
         Begin VB.Menu menLang 
            Caption         =   "Deutsch"
            Index           =   1
         End
      End
      Begin VB.Menu menBackTransparent 
         Caption         =   "Transparenz"
         Visible         =   0   'False
         Begin VB.Menu menBackTransparentProz 
            Caption         =   "0 %"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu menBackTransparentProz 
            Caption         =   "20%"
            Index           =   20
         End
         Begin VB.Menu menBackTransparentProz 
            Caption         =   "40%"
            Index           =   40
         End
         Begin VB.Menu menBackTransparentProz 
            Caption         =   "60%"
            Index           =   60
         End
         Begin VB.Menu menBackTransparentProz 
            Caption         =   "80%"
            Index           =   80
         End
      End
   End
   Begin VB.Menu menHelp 
      Caption         =   "?"
      Visible         =   0   'False
      Begin VB.Menu menForum 
         Caption         =   "&Forum"
      End
      Begin VB.Menu menTutorial 
         Caption         =   "&Tutorial"
         Shortcut        =   ^U
      End
      Begin VB.Menu menInfo 
         Caption         =   "&Info"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Schmulen As Boolean
Private StatistikUsed As Boolean
Private Const DemoTime = 15
Private Const picKarte_Top As Single = 0.704
'Private recIndikator_Top As Long
Private KarteGehoben As Boolean
Private Const SecundeEinheit As String = " s"
Private oldHeight As Long
Private Const myHeight = 6645
Private Tutorial As Boolean
Private TutorialStepCompleted(0 To 12) As Boolean
Private TutorialStepAktuell As Integer

Private myText() As String
Private cFormResizer As New clFormResizer

Private Sub Agent1_DblClick(ByVal CharacterID As String, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Integer, ByVal Y As Integer)
Debug.Print X, Y
If AktuellerSpieler.SpielerLevel < 4 And Playermodus = singleplayer Then cmdTip_Click
    
End Sub

Private Sub Agent1_DefaultCharacterChange(ByVal Guid As String)
    SetAgentAsEnemy
    AgentSpeak myText(57)
    Aktualisieren
End Sub

Private Sub Agent1_Hide(ByVal CharacterID As String, ByVal Cause As Integer)
    If Tutorial Then
        Unload Me
    Else
        DestroyAgent
    End If
End Sub

Private Sub Agent1_IdleStart(ByVal CharacterID As String)
Dim perc As Single
On Error GoTo ERRHand

If Not Tutorial And Playermodus = singleplayer And _
    ((StichBesitzer = Spieler And (StichAktion Mod 2 = ZERO)) Or _
    (StichBesitzer = Computer And (StichAktion Mod 2 = ONE))) Then
    
    If AktuellerSpieler.SpielerLevel = ZERO And AgentGivesTips And StichBesitzer <> ZERO Then
        AgentGivesTipp (KI_SpielKarte())
    ElseIf AktuellerSpieler.SpielerLevel = ONE And AgentGivesTips And StichBesitzer <> ZERO Then
        perc = Rnd(ONE)
        Select Case perc
            Case Is > 0.8
                AgentSpeak myText(129)
            Case Is > 0.8
                AgentSpeak myText(130)
        End Select
    ElseIf Not (StichBesitzer = ZERO And Geber = Spieler) Then
        'wenn Spieler dran ist
        perc = Rnd(ONE)
        Select Case perc
            Case Is > 0.96
                AgentSpeak myText(131)
            Case Is > 0.92
                AgentSpeak myText(132)
            Case Is > 0.9
                AgentSpeak myText(ZERO)
            Case Is > 0.88
                AgentSpeak myText(133)
            Case Is > 0.85
                AgentSpeak myText(134)
            Case Is > 0.83
                AgentSpeak myText(135)
            Case Is > 0.8
                AgentSpeak myText(136)
            Case Is > 0.77
                AgentSpeak myText(137)
            Case Is > 0.75
                AgentSpeak myText(138)
            Case Else
                If AktuellerSpieler.SpielerLevel < 3 And AgentGivesTips And StichBesitzer <> ZERO Then _
                    AgentGivesTipp (KI_SpielKarte())
        End Select
    End If
End If
Exit Sub
ERRHand:
If ErrorBox("Agent_Idle", Err) Then Resume Next
End Sub

Private Sub Agent1_Move(ByVal CharacterID As String, ByVal X As Integer, ByVal Y As Integer, ByVal Cause As Integer)
Dim perc As Single
    If Cause = 1 And val(frmStatistik.txtGelegteKarten) > 0 And Not Tutorial _
        And val(Me.lblComputerPunkte(ZERO).Caption) - val(Me.lblSpielerPunkte(ZERO).Caption) > 1 Then
        perc = Rnd(ONE)
        Select Case perc
            Case Is > 0.97
                AgentSpeak myText(56)
            Case Is > 0.93
                AgentSpeak myText(139)
            Case Is > 0.9
                AgentSpeak myText(140)
            Case Is > 0.87
                AgentSpeak myText(141)
            Case Is > 0.84
                AgentSpeak myText(142)
            Case Is > 0.8
                AgentSpeak myText(143)
            Case Is > 0.77
                AgentSpeak myText(144)
            Case Is > 0.75
                AgentSpeak myText(145)
        End Select
    End If
End Sub

Private Sub chkAudio_Click()
    Me.menAudio.Checked = IIf(Me.chkAudio.Value = vbChecked, True, False)
    AudioOn = Me.chkAudio.Value
    SaveSetting AppExeName, cstrOptions, cstrAudio, AudioOn
End Sub

Private Sub cmdSiebenSuch_Click()

Dim i As Integer
For i = KartenAnzahl To 1 Step -1
    If gemischtesSpiel.Karte(i).Bild = sieben Then Exit For
Next

picAbheben i
End Sub

Private Sub cmdStichEnde_Click()

Me.cmdStichEnde.Visible = False

If StichAktion = ZERO Then
    AgentSpeak myText(ZERO), True
    Exit Sub
End If
If StichAktion Mod 2 <> ZERO Then
    AgentSpeak myText(ONE), True
    Exit Sub
End If
#If Not Tiny Then
    If Playermodus = multiplayer Then frmChat.SendNetworkMessage MsgSendStichEnde, gstrNullstr, ZERO
#End If

SpielerHilfe -2

StichEnde
End Sub

Private Sub cmdStichEnde_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdStichEnde.Font.Size = 14
    If StichBesitzer = Spieler Then
        Me.cmdStichEnde.Alignment = ZERO
    Else
        Me.cmdStichEnde.Alignment = ONE
    End If
End Sub

Private Sub cmdStichEnde_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdStichEnde.ForeColor = IIf(StichBesitzer = Spieler, vbBlue, vbRed)
    If ShowIndikator Then ShowStichIndikator True
End Sub

Private Sub cmdStichEnde_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.cmdStichEnde.Font.Size = 16
    Me.cmdStichEnde.Alignment = 2
End Sub

Private Sub cmdTip_Click()
AgentGivesTipp (KI_SpielKarte())
End Sub

Private Sub AgentGivesTipp(KartenPosition As Integer, Optional past)
Dim perc As Single
On Error GoTo ERRHand
perc = Rnd(ONE)

If KartenPosition = -1 Then
    If IsMissing(past) Then
        Select Case perc
            Case Is > 0.75
                AgentSpeak myText(146) & gstrSpace & Me.cmdStichEnde.Caption, True
            Case Is > 0.5
                AgentSpeak myText(147) & gstrSpace & Me.cmdStichEnde.Caption, True
            Case Is > 0.25
                AgentSpeak Me.cmdStichEnde.Caption & myText(148), True
            Case Else
                AgentSpeak Me.cmdStichEnde.Caption & myText(149), True
        End Select
    Else
        AgentThink myText(175)
    End If
ElseIf KartenPosition <> 0 Then
    If IsMissing(past) Then
        Select Case perc
            Case Is > 0.75
                AgentSpeak myText(150) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText, True
            Case Is > 0.5
                AgentSpeak myText(151) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText & myText(152), True
            Case Is > 0.25
                AgentSpeak myText(153) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText, True
            Case Else
                AgentSpeak myText(154) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText, True
        End Select
        
    Else
        Select Case perc
            Case Is > 0.75
                AgentThink myText(155) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText & myText(156)
            Case Is > 0.5
                AgentThink myText(157) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText & myText(158)
            Case Is > 0.25
                AgentThink Me.picKarte(KartenPosition - 1).ToolTipText & myText(159)
            Case Else
                AgentThink myText(160) & gstrSpace & Me.picKarte(KartenPosition - 1).ToolTipText & myText(161)
        End Select
    End If
Else
    If IsMissing(past) Then
        Select Case perc
            Case Is > 0.75
                AgentSpeak myText(162), True
            Case Is > 0.5
                AgentSpeak myText(163), True
            Case Is > 0.25
                AgentSpeak myText(164), True
            Case Else
                AgentSpeak myText(165), True
        End Select
    End If
End If

If IsMissing(past) Then AgentSpeak KI_Erklaerung, True

Exit Sub
ERRHand:
If ErrorBox("AgentGivesTipp", Err) Then Resume Next
End Sub




Private Sub Form_Activate()
    Aktualisieren
End Sub

Private Sub Aktualisieren()
    Me.SetCaption
    Me.lblPGegner = myText(2) & gstrSpace & Gegner.SpielerName
    Me.lblPGegner1 = myText(2) & gstrSpace & Gegner.SpielerName
    Me.lblRPGegner = myText(3) & gstrSpace & Gegner.SpielerName
    Me.lblRPGegner1 = myText(3) & gstrSpace & Gegner.SpielerName
'    If Me.Top < -10000 And Me.Left < -10000 Then SetzFensterMittig Me
End Sub

Public Sub SetCaption()
Dim s As String
On Error Resume Next
    s = AppExeName & gstrSpace & App.Major & gstrDot & App.Minor & gstrDot & App.Revision
    Me.lblCaption = s
    Me.lblLevel.Caption = strPlayerLevel(AktuellerSpieler.SpielerLevel)
    Me.lblName.Caption = AktuellerSpieler.SpielerName
    Me.picLevel.Picture = Me.ImageList1.ListImages(AktuellerSpieler.SpielerLevel + 1).Picture
    'Me.Refresh
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 27 Then
    CloseAll
ElseIf Playermodus = singleplayer And KeyCode = 84 And Shift = 2 Then
    menStatistik_Click
ElseIf KeyCode = 32 And Shift = ZERO Then
    SetBackGround Me
Else
    Debug.Print KeyCode, Shift
End If
End Sub

Private Sub Form_Load()
    
    frmMain_Loaded = True
    cFormResizer.AutoResize = False
    If frmChat_Loaded Then
        #If Not Tiny Then
            Dock2Chat
            frmChat.Timer1.Enabled = True
        #End If
    Else
        SetzFensterMittig Me
    End If
    
'   die Größe betreffend
    Me.menGame.Visible = False
    oldHeight = Me.Height
    Me.Move Me.Left, Me.Top, Me.Width, myHeight

'    makeRoundEdges Me
'    HideTitleBar Me

    'Agenteninitialisierungen
    Me.men_useAgent.Checked = useAgent
    Me.menAgentTalkChat.Enabled = useAgent
    AgentGivesTips = GetSetting(AppExeName, cstrOptions, cstrAgentGivesTips, True)
    Me.menAgentTalkChat.Visible = (Playermodus = multiplayer)
    boolAgentTalkChat = CBool(GetSetting(AppExeName, cstrOptions, cstrAgentTalkChat, True))
    Me.menAgentTalkChat.Checked = boolAgentTalkChat
    
    If useAgent Then
        InitAgent
    Else
        Me.menAgentTTS.Visible = False
    End If
    
    
    Me.lblMsg(ZERO).BorderStyle = ZERO
    Me.lblMsg(ONE).BorderStyle = ZERO
    
'    Me.menBackTransparent.Visible = IsWin2000()
    
    Me.menShowTips.Checked = ShowTipAtStartup
    
    Me.menAudio.Checked = AudioOn
    Me.chkAudio.Value = -AudioOn
    
    ShowIndikator = CBool(IIf(GetSetting(AppExeName, cstrOptions, cstrShowIndikator) = gstrNullstr, True, GetSetting(AppExeName, cstrOptions, cstrShowIndikator)))
    
    'Texte laden
    If Not ArrayIsFilled(myText()) Then LoadObjectText Me.Name, myText()
    
'    menBackTransparentProz_Click ((GetSetting(AppExeName, cstrOptions, cstrTransparency, ZERO)))
    
    If IsRegistered Then
        Me.menGetFreeRegKey.Visible = False
        Me.menWriteFreeRegKey.Visible = False
        Me.menWriteRegkey.Visible = False
        Me.menLine1.Visible = False
        Me.ShareWareTimer.Enabled = False
    ElseIf IsRegisteredFree Then
        Me.menGetFreeRegKey.Visible = False
        Me.menWriteFreeRegKey.Visible = False
        Me.ShareWareTimer.Enabled = (StartAnz > (MaxStartsWReg * 11))
    Else
        If StartAnz > MaxStartsWReg Then
            AgentSpeak myText(166) & gstrSpace & DemoTime & myText(167), True
            If StartAnz < 11 * MaxStartsWReg Then AgentOffer
            Me.ShareWareTimer.Enabled = True
        End If
    End If
    
    getBackGrounds
    SetBackGround Me
        
    'Menu beschreiben
    Me.lblPunkteSpieler(ZERO) = myText(47)
    Me.lblPunkteSpieler(ONE) = myText(47)
    Me.lblRPunkteSpieler(ZERO) = myText(48)
    Me.lblRPunkteSpieler(ONE) = myText(48)
    Me.menBackground.Caption = myText(46)
    Me.menGame.Caption = myText(35)
    Me.menGetFreeRegKey.Caption = myText(38)
    Me.menOptionen.Caption = myText(40)
    Me.menSchmulen.Caption = myText(41)
    Me.menShowIndikator.Caption = myText(42)
    Me.menShowTips.Caption = myText(45)
    Me.menSpielerAuswahl.Caption = myText(36)
    Me.menStartGame.Caption = myText(37)
    Me.menStatistik.Caption = myText(43)
    Me.menWriteRegkey.Caption = myText(39)
    Me.men_useAgent.Caption = myText(168)
    Me.menAgentTTS.Caption = myText(169)
    Me.menAgentTalkChat.Caption = myText(170)
    Me.menAgentTips.Caption = myText(171)
    Me.menChooseAgent.Caption = myText(172)
    Me.menLanguage.Caption = myText(173)
    Me.menEndGame.Caption = myText(203)
    Me.menWriteFreeRegKey.Caption = myText(204)
    Me.menAgentDownload.Caption = myText(206)
    
    WriteMenLangs
   
    If Playermodus = multiplayer Then
'        Me.BorderStyle = vbBSNone
        Me.menAgentTips.Visible = False
        Me.cmdTip.Visible = False
        Me.lblExit.Visible = False
        Me.lblMin.Visible = False
        Me.lblExitBack.Visible = False
        Me.lblMinback.Visible = False
        
        Me.menSchmulen.Enabled = False
        Me.menEndGame.Enabled = False
        Me.menSpielerAuswahl.Enabled = False
        
'        Me.ShowInTaskbar = False
    End If
    'Now put up our system tray icon
    With sysIcon
        .cbSize = LenB(sysIcon)
        .hWnd = Me.hWnd
        .uFlags = NIF_DOALL
        .uCallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .sTip = myText(4) & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, sysIcon
    SetCaption

    'Beim ersten Start eines Spielers Tutorialmodus erfragen
    If Playermodus = singleplayer And AktuellerSpieler.Points = 0 _
        And AktuellerSpieler.SpielerLevel < 2 And useAgent Then
        If AgentQuestion(myText(178), myText(179)) = vbYes Then Tutorial = True
    End If
End Sub

Private Sub WriteMenLangs()
Dim i As Integer
    For i = LBound(strAvailableLangs) + 1 To UBound(strAvailableLangs)
        If i > 1 Then Load Me.menLang(i)
        Me.menLang(i).Caption = strAvailableLangs(i)
        Me.menLang(i).Checked = (strAvailableLangs(i) = StandardLanguage)
    Next
    
End Sub

Private Sub getBackGrounds()
Dim Name1 As String, pfad1 As String, LastbackG As String
Dim i As Integer, found As Boolean

On Error GoTo ERRHand

pfad1 = App.path & cstrSubPathBackGround
Name1 = Dir$(pfad1, vbDirectory)   ' Ersten Eintrag abrufen.
i = 2
Do While Name1 <> gstrNullstr   ' Schleife beginnen.
    
   ' Aktuelles und übergeordnetes Verzeichnis ignorieren.
   If Name1 <> gstrDot And Name1 <> ".." And Name1 <> "Standard" Then
      ' Mit bit-weisem Vergleich sicherstellen, daß Name1 ein
      ' Verzeichnis ist.
      If (GetAttr(pfad1 & Name1) And vbDirectory) = vbDirectory Then
         'Debug.Print Name1   ' Eintrag nur anzeigen, wenn es sich
         Load frmMain.menBackStandard(i)
         If i = 2 Then frmMain.menBackStandard(ONE).Visible = True
         frmMain.menBackStandard(i).Caption = Name1
         frmMain.menBackStandard(i).Checked = False
         i = i + 1
      End If   ' um ein Verzeichnis handelt.
   End If
   Name1 = Dir$   ' Nächsten Eintrag abrufen.
Loop

If i > 2 Then frmMain.menBackStandard(ONE).Caption = gstrMinus

LastbackG = GetSetting(AppExeName, cstrOptions, cstrLastBackGround)


For i = frmMain.menBackStandard.LBound To frmMain.menBackStandard.UBound
    If frmMain.menBackStandard(i).Caption = LastbackG Then
        frmMain.menBackStandard(i).Checked = True
        found = True
    End If
Next
If Not found Or LastbackG = gstrNullstr Then
    frmMain.menBackStandard(ZERO).Checked = True
End If

Exit Sub

ERRHand:
If ErrorBox("getBackGrounds", Err) Then Resume Next

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShellMsg As Long
    Dim i As Integer

    Me.cmdStichEnde.ForeColor = vbWhite
    Me.lblURL.ForeColor = vbWhite
    Me.lblExit.ForeColor = vbWhite
    Me.lblMin.ForeColor = vbWhite
    Me.lblMinMax(0).ForeColor = vbWhite
    
    ShowMenu False
'    ShowStichIndikator False

    ShellMsg = X \ Screen.TwipsPerPixelX
    Select Case ShellMsg
    Case WM_LBUTTONDBLCLK
        If Playermodus = multiplayer And Not LaufendesSpiel Then
            #If Not Tiny Then
                frmChat.Visible = True
            #End If
        Else
            'SetzFensterMittig Me
            Me.Visible = True
        End If
    Case WM_RBUTTONUP
        'Show the menu
        'If gfStarted Then mnuStart.Enabled = False
        PopupMenu Me.menOptionen
    End Select

    If KarteGehoben Then
        For i = ZERO To 3
            Me.picKarte(i).Move Me.picKarte(i).Left, Me.Height * picKarte_Top
            'Me.rectIndikator(i).Top = recIndikator_Top
        Next
        KarteGehoben = False
    End If

'    If X < Me.picKarteComp(one).ScaleTop Then
'        Me.menGame.Visible = True
'    Else
'
'    End If
'
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then PopupMenu Me.menOptionen
If Button = vbLeftButton And Me.WindowState <> vbMaximized Then MoveME Me
End Sub

Public Sub Dock2Chat()
On Error Resume Next
    #If Not Tiny Then
        If frmChat.Top < -10000 And frmChat.Left < -10000 Then Exit Sub
        Me.Move frmChat.Left, frmChat.Top + frmChat.Height
    #End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    frmMain_Loaded = False
    TimerEnde
    Tutorial = False
    DestroyAgent
    Shell_NotifyIcon NIM_DELETE, sysIcon
    Me.Timer1.Enabled = False
    frmStatistik.Timer1.Enabled = False
    Unload frmStatistik
    AnimWindow Me, AW_HIDE + AW_BLEND
    LaufendesSpiel = False
    If Playermodus = singleplayer Then AnimWindow frmSplash, AW_ACTIVATE + AW_BLEND
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub lblCaption_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ShowMenu True
End Sub
Private Sub ShowMenu(ShowIt As Boolean)
Static menuIsON As Boolean

If Me.WindowState = vbMinimized Or Me.WindowState = vbMaximized Then Exit Sub

If menuIsON <> ShowIt Then
    Me.menGame.Visible = ShowIt
    Me.menOptionen.Visible = ShowIt
    Me.menHelp.Visible = ShowIt
    If Me.Height > oldHeight Then oldHeight = Me.Height
    'Me.Refresh
    menuIsON = ShowIt

    Me.Move Me.Left, Me.Top, Me.Width, IIf(ShowIt, oldHeight, myHeight)

'    Debug.Print Me.Height
    'makeRoundEdges Me
End If
End Sub


Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.FontSize = Me.lblExit.FontSize - 2
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.ForeColor = vbRed
    Me.lblMin.ForeColor = vbWhite
    Me.lblMinMax(0).ForeColor = vbWhite
    Me.lblURL.ForeColor = vbWhite
    ShowMenu False
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.FontSize = Me.lblExit.FontSize + 2
End Sub

Private Sub lblMin_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub lblMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.FontSize = Me.lblMin.FontSize - 2
End Sub

Private Sub lblMin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.ForeColor = vbRed
    Me.lblExit.ForeColor = vbWhite
    Me.lblMinMax(0).ForeColor = vbWhite
    Me.lblURL.ForeColor = vbWhite
    ShowMenu False
End Sub

Private Sub lblMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMin.FontSize = Me.lblMin.FontSize + 2
End Sub


Private Sub SetResizer()
With cFormResizer
    .Initialize Me
    .SetExceptionCtrl Me.chkAudio, exNoScale + exScaleKeepAspect
    .SetExceptionCtrl Me.cmdStichEnde, exNoScaleFont
End With
End Sub


Private Sub lblMinMax_Click(Index As Integer)
ShowMenu False
SetResizer
If Me.WindowState = vbMaximized Then
    Me.lblMinMax(0).ToolTipText = "maximize"
    Me.WindowState = ZERO
    SetWindowPos frmStatistik.hWnd, HWND_NOTOPMOST, frmStatistik.Left, frmStatistik.Top, frmStatistik.Width, frmStatistik.Height, 3
Else
    Me.WindowState = vbMaximized 'normal
    Me.lblMinMax(0).ToolTipText = "normal"
    SetWindowPos frmStatistik.hWnd, HWND_TOPMOST, frmStatistik.Left, frmStatistik.Top, frmStatistik.Width, frmStatistik.Height, 3
End If

Me.AutoRedraw = (Me.WindowState = vbMaximized)
SetBackGround Me
cFormResizer.ResizeForm
'frmStatistik.ZOrder
End Sub

Private Sub lblMinMax_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinMax(0).FontSize = Me.lblMinMax(0).FontSize - 2
End Sub

Private Sub lblMinMax_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinMax(0).ForeColor = vbRed
    Me.lblExit.ForeColor = vbWhite
    Me.lblMin.ForeColor = vbWhite
    Me.lblURL.ForeColor = vbWhite
    
    ShowMenu False
End Sub

Private Sub lblMinMax_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinMax(0).FontSize = Me.lblMinMax(0).FontSize + 2
End Sub

Private Sub lblURL_Click()
    GoHome
End Sub

Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.Font.Size = Me.lblURL.Font.Size - 2
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.ForeColor = ROT
    Me.lblMin.ForeColor = vbWhite
    ShowMenu False
End Sub

Private Sub lblURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.Font.Size = Me.lblURL.Font.Size + 2
End Sub

Private Sub men_useAgent_Click()
    Me.men_useAgent.Checked = Not Me.men_useAgent.Checked
    useAgent = Me.men_useAgent.Checked
    SaveSetting AppExeName, cstrOptions, cstrUseAgent, Me.men_useAgent.Checked
    If useAgent Then
        InitAgent
        Me.men_useAgent.Caption = myText(176)
    Else
        If Not myAgent Is Nothing Then myAgent.Hide
        'DestroyAgent
        Me.men_useAgent.Caption = myText(168)
    End If
    
    Me.menAgentTalkChat.Enabled = useAgent
    
End Sub

Private Sub menAgentDownload_Click()
    Go2URL "http://www.msagentring.org/chars.htm"
End Sub

Private Sub menAgentTalkChat_Click()
    Me.menAgentTalkChat.Checked = Not Me.menAgentTalkChat.Checked
    boolAgentTalkChat = Me.menAgentTalkChat.Checked
    SaveSetting AppExeName, cstrOptions, cstrAgentTalkChat, boolAgentTalkChat
End Sub

Private Sub menAgentTips_Click()
    Me.menAgentTips.Checked = Not Me.menAgentTips.Checked
    AgentGivesTips = Me.menAgentTips.Checked
    SaveSetting AppExeName, cstrOptions, cstrAgentGivesTips, AgentGivesTips
End Sub

Private Sub menAgentTTS_Click()
Const LocalComponentUrl = "http://activex.microsoft.com/activex/controls/agent2/AgtX0407.exe"
Const TTSEngineUrl = "http://activex.microsoft.com/activex/controls/agent2/lhttsged.exe"
If BestehtVerbindung Then
    Go2URL "www.microsoft.com/msagent/downloads.htm"
    Go2URL LocalComponentUrl
    Go2URL TTSEngineUrl
Else
    AgentSpeak ModText(16), True
End If
End Sub

Private Sub menBackTransparent_Click()
    'MakeTransparent Me
End Sub

Private Sub menBackTransparentProz_Click(Index As Integer)
'Dim i As Integer
'
'SaveSetting AppExeName, cstrOptions, cstrTransparency, Index
'menBackTransparentProz(Index).Checked = True
'For i = menBackTransparentProz.LBound To menBackTransparentProz.ubound Step 20
'    If i <> Index Then menBackTransparentProz(i).Checked = False
'Next
'
'If Index <> 0 Then
'    SetWindowTranslucency Me.hWnd, (100 - Index) \ 100 * 255
'Else
'    ClearWindowTranslucency Me.hWnd
'End If
'RefreshAllCtls Me
End Sub

Private Sub menChooseAgent_Click()
    Me.Agent1.ShowDefaultCharacterProperties
End Sub

Private Sub menEndGame_Click()
    Unload Me
End Sub

'Private Sub menShowChat_Click()
'    #If Not Tiny Then
'        If Not Register Then Exit Sub
'        InitDPlay Me.Name
'    #End If
'    Me.menShowChat.Enabled = False
'End Sub

Private Sub menAudio_Click()
    Me.menAudio.Checked = Not Me.menAudio.Checked
    Me.chkAudio.Value = IIf(Me.menAudio.Checked, vbChecked, vbUnchecked)
End Sub

Private Sub menBackStandard_Click(Index As Integer)
Dim i As Integer

If Index = 1 Then
    If BestehtVerbindung Then
        Go2URL "http://www.playseven.com/downloads/BackGrounds.zip"
    Else
        AgentSpeak ModText(16)
    End If
    Exit Sub
End If

menBackStandard(Index).Checked = True

For i = menBackStandard.LBound To menBackStandard.UBound
    If i <> Index Then
        menBackStandard(i).Checked = False
    End If
Next
SetBackGround Me
SaveSetting AppExeName, cstrOptions, cstrLastBackGround, menBackStandard(Index).Caption
End Sub

Private Sub menForum_Click()
Dim rc As Integer
    Go2URL "http://www.playseven.de/modules.php?op=modload&name=XForum&file=index"
End Sub

Private Sub menGetFreeRegKey_Click()
    getWebRegKey
End Sub

Private Sub menInfo_Click()
    frmAbout.Show 1
End Sub

Private Sub menKartenAnz_Click(Index As Integer)
'Dim NeueAnzahl As Integer
'
'    If Index = ZERO Then
'        NeueAnzahl = 32
'    Else
'        NeueAnzahl = 52
'    End If
'    If KartenAnzahl <> NeueAnzahl Then
'        KartenAnzahl = NeueAnzahl
'        Init
'    End If
'
End Sub

Private Sub menLang_Click(Index As Integer)
Dim i As Integer
For i = Me.menLang.LBound To Me.menLang.UBound
    Me.menLang(i).Checked = False
Next
Me.menLang(Index).Checked = True
SaveSetting AppExeName, cstrOptions, cstrLanguage, Me.menLang(Index).Caption
AgentSpeak myText(174), True
End Sub

Private Sub menSchmulen_Click()
Dim i As Integer
    Schmulen = Not Schmulen
    Me.menSchmulen.Checked = Schmulen
    For i = 1 To UBound(HandkartenComputer)
        If Schmulen Then
            Me.picKarteComp(i - 1).ToolTipText = HandkartenComputer(i).Caption
        Else
            Me.picKarteComp(i - 1).ToolTipText = cstrXXX
        End If
    Next
    
End Sub



Private Sub menShowIndikator_Click()
    Me.menShowIndikator.Checked = Not Me.menShowIndikator.Checked
    ShowIndikator = Me.menShowIndikator.Checked
    SaveSetting AppExeName, cstrOptions, cstrShowIndikator, ShowIndikator
    MakeIndikator
    ShowStichIndikator
End Sub

Sub MakeIndikator()
Dim i As Integer

On Error GoTo ERRHand
For i = ZERO To 3
    If ShowIndikator Then
        Select Case HandkartenSpieler(i + 1).Bild
            Case 7, StichTrumpf
                Me.rectIndikator(i).BorderColor = &HAA00AA
                Me.rectIndikator(i).BackColor = vbMagenta
                Me.rectIndikator(i).BackStyle = 1

'                Me.picKarte(i).ToolTipText = myText(52)
            Case 10, 14
                Me.rectIndikator(i).BorderColor = 43690
                Me.rectIndikator(i).BackColor = vbYellow
                Me.rectIndikator(i).BackStyle = 1

'                Me.picKarte(i).ToolTipText = myText(51)
            Case 0
                Me.rectIndikator(i).BorderColor = SCHWARZ
                Me.rectIndikator(i).BackColor = vbBlack
                Me.rectIndikator(i).BackStyle = 0
            Case Else
                Me.rectIndikator(i).BorderColor = hellGRUEN
                Me.rectIndikator(i).BackStyle = 1
                Me.rectIndikator(i).BackColor = GRUEN
'                Me.picKarte(i).ToolTipText = myText(53)
        End Select
        Me.rectIndikator(i).Refresh
    End If
    Me.rectIndikator(i).Visible = ShowIndikator
Next
'Me.Refresh
Exit Sub
ERRHand:
If ErrorBox("MakeIndikator", Err) Then Resume Next
End Sub


Private Sub menShowTips_Click()
    Me.menShowTips.Checked = Not Me.menShowTips.Checked
    SaveSetting AppExeName, cstrOptions, cstrShowTips, Me.menShowTips.Checked
End Sub

Private Sub menSpielerAuswahl_Click()
    Unload Me
    AnimWindow frmSplash, AW_ACTIVATE + AW_BLEND
End Sub

Private Sub menStartGame_Click()
    Tutorial = False
    Init
End Sub

Private Sub menStatistik_Click()
    Dim i As Integer
    Dim Old_X As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.WindowState <> vbMaximized Then
        If Not StatistikUsed Then
            Old_X = Me.Left
            'frmStatistik.ScaleMode = vbTwips
            For i = 1 To frmStatistik.ScaleWidth \ 2 Step Screen.TwipsPerPixelX * 3
                Me.Move Old_X - (i)
    '            DoEvents
            Next
        End If
        frmStatistik.Dock2Main
    
    End If
    frmStatistik.Show
    AnimWindow frmStatistik, AW_BLEND + AW_ACTIVATE
    frmStatistik.Timer1.Enabled = True
    
    If Me.WindowState = vbMaximized Then SetWindowPos frmStatistik.hWnd, HWND_TOPMOST, frmStatistik.Left, frmStatistik.Top, frmStatistik.Width, frmStatistik.Height, 3
    
    Me.menStatistik.Enabled = False
    StatistikUsed = True
End Sub

Private Sub menTutorial_Click()
Dim rc As Integer

ShowMenu False
If Playermodus = singleplayer Then
    If Not useAgent Then
        useAgent = True
        InitAgent
    End If
    Tutorial = True
    Init
Else
    Go2URL "www.playseven.com/tutorial.html"
End If
End Sub

Private Sub menWriteFreeRegKey_Click()
    RegisterFree
End Sub

Private Sub menWriteRegkey_Click()
    RegisterApp Me
End Sub

Private Sub picHaufen_Click(Index As Integer)

If Tutorial Then Exit Sub

If Playermodus = multiplayer Then
    If Geber = Spieler Then
        HaufenEnable False
        Exit Sub
    Else
        
        #If Not Tiny Then
            If Not SpielErhalten Then
                WriteMsg myText(54)
                Exit Sub
            End If
            HaufenEnable False
            frmChat.SendNetworkMessage MsgSendAbheben, CStr(Index), ZERO
        #End If
    End If
End If
picAbheben Index
End Sub

Private Sub HaufenEnable(Switch As Boolean)
Dim i As Integer
For i = picHaufen.LBound To picHaufen.UBound
    picHaufen(i).Enabled = Switch
Next
End Sub


Public Sub picAbheben(Index As Integer)
Dim i As Long, X As Long, found As Boolean, str As String
Dim stp As Long, rng As Long

X = Index
SiebenGefunden = ZERO
If Not myAgent Is Nothing Then myAgent.Stop

WriteMsg gstrNullstr

If Me.WindowState = vbMaximized Then
    rng = Me.picKarteStich(ONE).Height / Screen.TwipsPerPixelY * 0.75
    stp = Me.picKarteStich(ONE).Height / Screen.TwipsPerPixelY / 50
Else
    rng = 50
    stp = ONE
End If

Do
    
    For i = 1 To rng Step stp
        Me.picHaufen(Index).Move Me.picHaufen(Index).Left, Me.picHaufen(Index).Top + IIf(Geber = Spieler, -i, i)
        'Me.picHaufen(Index).Refresh
    Next
    'Me.picHaufen(Index).ZOrder
    
    'gefundene Karte anzeigen
    Me.picHaufen(Index).Picture = gemischtesSpiel.Karte(Index).Pic
    Me.picHaufen(Index).Refresh
    If gemischtesSpiel.Karte(Index).Bild = 7 Then
        MakeMsg gemischtesSpiel.Karte(Index).Caption & gstrSpace & myText(7)
        SiebenGefunden = SiebenGefunden + 1
        found = True
    End If
    'Me.Refresh
    Sleep 1000
    
    If found Then
'       Karte wieder umdrehen
        Me.picHaufen(Index).Visible = False
        'Me.Refresh

        'gefundene Sieben an ziehenden Spieler
        If Geber = Computer Then
            Me.picKarte(SiebenGefunden - 1).Picture = Me.picHaufen(Index).Picture
            Me.picKarte(SiebenGefunden - 1).ToolTipText = gemischtesSpiel.Karte(Index).Caption
            Me.picKarte(SiebenGefunden - 1).Visible = True
            HandkartenSpieler(SiebenGefunden) = gemischtesSpiel.Karte(Index)
'            If ShowIndikator Then
'                Me.rectIndikator(SiebenGefunden - 1).BorderColor = BLAU
'                Me.rectIndikator(SiebenGefunden - 1).BorderColor = BLAU
'                Me.rectIndikator(SiebenGefunden - 1).Visible = True
'            End If
        Else
            Me.picKarteComp(SiebenGefunden - 1).Picture = Me.picHaufen(Index).Picture
            Me.picKarteComp(SiebenGefunden - 1).ToolTipText = gemischtesSpiel.Karte(Index).Caption
            Me.picKarteComp(SiebenGefunden - 1).Visible = True
            HandkartenComputer(SiebenGefunden) = gemischtesSpiel.Karte(Index)
        End If
        'Karte wieder umdrehen
        Me.picHaufen(Index).Picture = Me.picHaufen(52).Picture
    
    End If
    
    
    Index = Index - 1
    found = False
Loop While gemischtesSpiel.Karte(Index + 1).Bild = 7 And Index > ZERO
If X <> 1 Then
    lastCard = gemischtesSpiel.Karte(Index + 1)
    str = lastCard.Caption & gstrSpace & myText(8)
    MakeMsg str
    AgentSpeak str
    ShowTutorial 5
Else
    lastCard = NullKarte
End If
'unten liegende Karte wieder umdrehen
Me.picHaufen(Index + 1).Picture = Me.picHaufen(52).Picture

MakeMsg cstrLinie

gemischtesSpiel = KartenAbheben(gemischtesSpiel, X, SiebenGefunden)
KartenGeben

End Sub

Private Sub ShowPunkteAnzeige(Switch As Boolean)
'Me.shpPktComp.Visible = Switch
'Me.shpPktPlayer.Visible = Switch
Me.lblComputerPunkte(ZERO).Visible = Switch
Me.lblComputerPunkte(ONE).Visible = Switch
Me.lblSpielerPunkte(ZERO).Visible = Switch
Me.lblSpielerPunkte(ONE).Visible = Switch
Me.lblPGegner.Visible = Switch
Me.lblPGegner1.Visible = Switch
'Me.Line2.Visible = Switch
'Me.Line3.Visible = Switch
Me.lblPunkteSpieler(ZERO).Visible = Switch
Me.lblPunkteSpieler(ONE).Visible = Switch
'Me.shpPlayerInfo.Visible = Switch
Me.shpComputerPunkte.Visible = Switch
Me.shpSpielerPunkte.Visible = Switch
'Me.shpRundenPunkte.Visible = Switch
If AktuellerSpieler.SpielerLevel < 4 And Playermodus = singleplayer Then
    Me.cmdTip.Visible = Switch
Else
    Me.cmdTip.Visible = False
End If

End Sub
Private Sub ShowTutorial(Step As Integer)
If Tutorial And frmMain_Loaded Then
    TutorialStepAktuell = Step
    Select Case Step
        Case ZERO
            AgentSpeak myText(180) & gstrSpace & AktuellerSpieler.SpielerName & " !!!"
            AgentSpeak myText(71) & gstrSpace & Gegner.SpielerName & myText(72)
            DoSleep 10000
        Case 1
            AgentSpeak myText(73)
            AgentSpeak myText(74)
            AgentSpeak myText(75)
            AgentSpeak myText(76)
            DoSleep 15000
        Case 2
            AgentSpeak myText(77)
            AgentSpeak myText(78)
            DoSleep 10000
        Case 3
            AgentSpeak myText(79)
            AgentSpeak myText(80)
            AgentSpeak myText(81)
            DoSleep 10000
        Case 4
            AgentSpeak myText(82)
            AgentAnim Process
            AgentSpeak "Sim Saalabim !"
            DoSleep 10000
            cmdSiebenSuch_Click
        Case 5
            AgentSpeak myText(83)
            DoSleep 7500
        Case 6
            AgentSpeak myText(84)
            MoveAgentInForm Me, Me.picKarteComp(3).Left + Me.picKarteComp(3).Width, Me.picKarteComp(3).Top, move2
            MoveAgentInForm Me, Me.picKarteComp(ZERO).Left, Me.picKarteComp(ZERO).Top, ShowAt
            AgentSpeak myText(85)
            
            MoveAgentInForm Me, Me.shpComputerPunkte.Left + 2 * Me.shpComputerPunkte.Width, Me.shpComputerPunkte.Top, move2
            MoveAgentInForm Me, Me.shpComputerPunkte.Left, Me.shpComputerPunkte.Top, ShowAt
            AgentSpeak myText(86)
            AgentSpeak myText(87)
            
            MoveAgentInForm Me, Me.picHaufen(31).Left - Me.picHaufen(31).Width, Me.picHaufen(31).Top, move2
            MoveAgentInForm Me, Me.picHaufen(31).Left + Me.picHaufen(31).Width, Me.picHaufen(31).Top, ShowAt
            AgentSpeak myText(88)
            AgentSpeak myText(89) & gstrSpace & lastCard.Caption & myText(90)
            
            
            MoveAgentInForm Me, Me.shpSpielerPunkte.Left + 2 * Me.shpSpielerPunkte.Width, Me.shpSpielerPunkte.Top, move2
            MoveAgentInForm Me, Me.shpSpielerPunkte.Left, Me.shpSpielerPunkte.Top, ShowAt
            AgentSpeak myText(91)
            AgentSpeak myText(92)
            AgentSpeak myText(93)
            AgentSpeak myText(94)
            
            MoveAgentInForm Me, Me.picKarte(3).Left + Me.picKarte(3).Width, Me.picKarte(ONE).Top, move2
            MoveAgentInForm Me, Me.picKarte(ZERO).Left - Me.picKarte(ZERO).Width, Me.picKarte(ZERO).Top, ShowAt
            AgentSpeak myText(95)
                       
            AgentSpeak myText(96)
            AgentSpeak myText(97)
            AgentSpeak myText(98)
            AgentSpeak myText(99)

            DoSleep 45000
        Case 7
            MoveAgentInForm Me, Me.picHaufen(31).Left, Me.picHaufen(31).Top, move2
            AgentSpeak myText(100)
            AgentSpeak myText(101)
            AgentSpeak myText(102)
            AgentSpeak myText(103)
            cmdTip_Click
            DoSleep 10000
        Case 8
            MoveAgentInForm Me, Me.picKarteStich(5).Left, Me.picKarteStich(ONE).Top, move2
            MoveAgentInForm Me, Me.picKarteStich(ONE).Left, Me.picKarteStich(ONE).Top, ShowAt
            AgentSpeak myText(104)
            AgentSpeak myText(105)
            AgentSpeak myText(106)
            AgentSpeak myText(107)
            
            DoSleep 8000
        Case 9
            If Not TutorialStepCompleted(Step) Then
                TutorialStepCompleted(Step) = True
                AgentSpeak myText(108)
                DoSleep 1000
                AgentSpeak myText(109)
                DoSleep 2000
                picKarteComp_Wirf (ComputerAntwort(HandkartenComputer, 5, ZERO, ZERO, True) - 1)
            End If
        Case 10
            DoSleep 3000
            AgentSpeak myText(110)
            AgentSpeak myText(111)
            AgentSpeak myText(112)
            AgentSpeak myText(113)
            AgentSpeak myText(114) & IIf(StichBesitzer = Computer, myText(115), myText(116))
            AgentSpeak myText(117)
            AgentSpeak myText(118)
            cmdTip_Click
            'dosleep 15000
        Case 11
            AgentSpeak myText(119)
            AgentSpeak myText(120)
            AgentSpeak myText(121)
            AgentSpeak myText(122)
            MoveAgentInForm Me, Me.shpRundenPunkte.Left, Me.Line3.Y1, move2
            MoveAgentInForm Me, Me.cmdTip.Left, Me.cmdTip.Top, ShowAt
            AgentSpeak myText(123)
            AgentSpeak myText(124)
            If frmStatistik_Loaded Then
                AgentSpeak myText(125)
                MoveAgentInForm frmStatistik, frmStatistik.txtKarteGef(10).Left, frmStatistik.txtKarteGef(10).Top, move2
                MoveAgentInForm frmStatistik, frmStatistik.txtKarteGef(13).Left, frmStatistik.txtKarteGef(13).Top, ShowAt
                AgentSpeak myText(126)
            End If
            AgentSpeak myText(127)
            DoSleep 30500
            Tutorial = False
            Init
        Case 12
            AgentSpeak myText(128)
            AgentSpeak myText(118)
            cmdTip_Click
    End Select
    TutorialStepCompleted(Step) = True
'    Debug.Print "Step " & Step & " completed"
End If
End Sub

Public Sub Init(Optional NeueRunde)

Dim i As Integer

If KartenAnzahl = ZERO Then KartenAnzahl = 32
WriteMsg vbNullString
    
If Tutorial Then
    AgentGivesTips = True
    ShowIndikator = True
End If
Me.menShowIndikator.Checked = ShowIndikator
Me.menAgentTips.Checked = AgentGivesTips

#If Not Tiny Then
    SpielErhalten = False
    MeisterFehlerGegner = False
    ZeitUeberschreitungGegner = False
#End If
TimerEnde
MeisterFehler = False

ShowTutorial ZERO

If frmMain_Loaded Then
    If Not Tutorial Then AgentAnim Process
    
    For i = ZERO To 3
        Me.rectIndikator(i).Visible = False
        Me.picKarte(i).Visible = False
        Me.picKarteComp(i).Visible = False
        Me.picKarteStich(i + 1).Visible = False
        Me.picKarteStich((i + 1) * 2).Visible = False
    Next
    
    'Stichindikator ausblenden
    Me.rectStichIndikator.Visible = False
    
    'Kartenzähler zurücksetzen
    For i = frmStatistik.txtKarteGef.LBound To frmStatistik.txtKarteGef.UBound
        frmStatistik.txtKarteGef(i) = ZERO
    Next
    
    'PunkteAnzeige ausblenden
    ShowPunkteAnzeige False
    
    Me.cmdStichEnde.Visible = False
    
    'Stich ausblenden
    For i = 1 To 8
        Me.picKarteStich(i).Visible = False
    Next
    
    'Haufen ausblenden
    For i = 1 To 52
        Me.picHaufen(i).Visible = False
    Next
    
    Me.menShowIndikator.Visible = False
    
    setSpielerPunkte ZERO
    setComputerPunkte ZERO
    
    Me.lblScore = AktuellerSpieler.Points
    Me.lblHighScore = glHighScore
    'Me.lblScore.Refresh
    'Me.lblHighScore.Refresh
    
    'Me.Refresh
    
    
    'Ab Spiellevel 2, Spielhilfen ausschalten
    If AktuellerSpieler.SpielerLevel > 1 Then
        Me.menSchmulen.Checked = False
        Me.menSchmulen.Enabled = False
    End If
    If AktuellerSpieler.SpielerLevel > 2 And Not Test Then
        Me.menStatistik.Enabled = False
        Me.menStatistik.Checked = False
        If ShowIndikator Then menShowIndikator_Click
        Me.menShowIndikator.Enabled = False
        Me.menAgentTips.Visible = False
        Me.menTutorial.Visible = False
    End If
       
    frmStatistik.txtGelegteKarten = ZERO
    frmStatistik.txtKartenImStapel = KartenAnzahl
    frmStatistik.lstVerlauf.Clear
    
    StichAktion = ZERO
    StichBesitzer = ZERO
    StichTrumpf = -1
    
    Schmulen = Me.menSchmulen.Checked
    boolPlayerWon = False
    
    If Playermodus = singleplayer Then
        Me.menStartGame.Enabled = True
        Me.menSpielerAuswahl.Enabled = True
        
        'Kartenspiel erstellen
        KartenSpiel_Init
        'ein gemischtes Spiel erstellen
        gemischtesSpiel = KartenMischen(Kartenspiel)
    Else
        Me.menStartGame.Enabled = False
        Me.menSpielerAuswahl.Enabled = False
        
        #If Not Tiny Then
            If DPlayEventsForm.IsHost Then
                'Kartenspiel erstellen
                KartenSpiel_Init
                'ein gemischtes Spiel erstellen
                gemischtesSpiel = KartenMischen(Kartenspiel)
                'gemischtes Spiel an Clienten schicken
                sendMixedGame
                SpielErhalten = True
            Else
                'nix : Steuerung übernimmt Host über frmChat Events
            End If
        #End If
    End If
    
    If IsMissing(NeueRunde) Then
        setComputerRndPunkte ZERO
        Me.lblSpielerRundenPunkte(ZERO).Caption = ZERO
        Me.lblSpielerRundenPunkte(ONE).Caption = ZERO
        
        If Playermodus = multiplayer Then
            #If Not Tiny Then
                If DPlayEventsForm.IsHost Then
                    Geber = IIf(Int(2 * Rnd) = ZERO, Spieler, Computer)
                    sendGeber -Geber
                Else
                    'nix : Steuerung übernimmt Host über frmChat Events
                End If
            #End If
        Else
            If Tutorial Then
                Geber = Computer
            Else
                Geber = IIf(Int(2 * Rnd) = 0, Spieler, Computer)
            End If
        End If
        
    Else
        SetBackGround Me
    End If
End If

ShowTutorial 1

If frmMain_Loaded Then
    GemischtesSpielAktPosition = ZERO
    'debug.Print GemischtesSpielAktPosition
    
    SiebenGefunden = ZERO
    
    'Stich leeren
    ReDim Preserve Stich(ZERO)
    
End If

ShowTutorial 2

If Not Test Then
    WriteMsg myText(49), 4444
Else
    WriteMsg myText(49)
End If

Kartenausbreiten
WriteMsg vbNullString


ShowTutorial 3


If Geber = Spieler And IsMissing(NeueRunde) Then WriteMsg Gegner.SpielerName & gstrSpace & myText(9), 2222

ShowTutorial 4
ersterZug

'Statistik anzeigen in den unteren Levels
If Not StatistikUsed And AktuellerSpieler.SpielerLevel < 2 Then menStatistik_Click


Debug.Print "INIT Finished"
End Sub
Sub Kartenausbreiten()
'Karten zum Abheben ausbreiten
Dim i As Integer

Me.cmdStichEnde.Visible = False
For i = 1 To 52
    If i = 1 Then
        Me.picHaufen(i).Picture = RueckSeitePic
    Else
        Me.picHaufen(i).Picture = Me.picHaufen(i - 1).Picture
    End If
    Me.picHaufen(i).Enabled = True
    Me.picHaufen(i).ToolTipText = i
    Me.picHaufen(i).Move (Me.picHaufen(i).Width * 0.185185185185185 * i) + Me.Height * 2.29095074455899E-03, Me.Height * 0.37                                  '2580
    Me.picHaufen(i).ZOrder
    Me.picHaufen(i).Visible = IIf(i > KartenAnzahl, False, True)
Next
'Me.Refresh
End Sub


Public Sub ComputerSpieltFuerSpieler()
Dim Index As Integer

Index = KI_SpielKarte()
AgentGivesTipp Index

If Index > 0 Then
    pickarte_Click (Index - 1)
ElseIf Index = 0 Then
    pickarte_Click (SuchZufallsKarte(HandkartenSpieler).Position - 1)
Else
    cmdStichEnde_Click
End If
End Sub

Sub ersterZug()

'Me.Refresh

If Geber = Computer Then
    If Tutorial Then
        'nix
    Else
        If Test Then
            ComputerHebtAb
            ComputerSpieltFuerSpieler
        Else
            WriteMsg myText(50), 2500
            AgentSpeak myText(181)
        End If
    End If
    
Else
    PlayerSwitch Computer
    If Playermodus = singleplayer Then
        WriteMsg Gegner.SpielerName & gstrSpace & myText(55), 2222
        ComputerHebtAb
        'AgentSpeak "Ich komme."
        ComputerAktion
    Else
        WriteMsg Gegner.SpielerName & gstrSpace & myText(55)
    End If

End If
LaufendesSpiel = True
'SunkenPanel3D picKarteComp(zero)

End Sub

Sub KartenGeben()
'überprüfen ob die unterste Karte eine 7 ist
Dim X As Integer
Dim i As Integer

GemischtesSpielAktPosition = 1


'wenn 7 gezogen 7 an Spieler und dann wurde erst Computer geben
If SiebenGefunden > ZERO Then
    For X = ZERO To SiebenGefunden - 1
        If Geber = Computer Then
            If Test Then
                Me.picKarteComp(X).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Caption
                Me.picKarteComp(X).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Pic
            Else
                'Me.picKarteComp(x).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + x).Caption
                Me.picKarteComp(X).ToolTipText = cstrXXX
                Me.picKarteComp(X).Picture = RueckSeitePic
            End If
            Me.picKarteComp(X).Visible = True
            HandkartenComputer(X + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X)
        Else
            Me.picKarte(X).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Caption
            Me.picKarte(X).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Pic
            Me.picKarte(X).Visible = True
            HandkartenSpieler(X + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X)
            'Kartenindikator
            If ShowIndikator Then MakeIndikator
        End If
    Next
    GemischtesSpielAktPosition = GemischtesSpielAktPosition + SiebenGefunden
    'Debug.Print GemischtesSpielAktPosition
End If

'Karten geben
For X = ZERO To 7 - (SiebenGefunden * 2)
    If X Mod 2 = IIf(Geber = Computer, ZERO, 1) Then
        Me.picKarte(Int((X + SiebenGefunden * 2) / 2)).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Caption
        Me.picKarte(Int((X + SiebenGefunden * 2) / 2)).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Pic
        Me.picKarte(Int((X + SiebenGefunden * 2) / 2)).Visible = True
        HandkartenSpieler(Int((X + SiebenGefunden * 2) / 2) + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X)
        'Kartenindikator
        If ShowIndikator Then MakeIndikator
    Else
        If Test Then
            Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Caption
            Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Pic
        Else
            If Not Schmulen Then
                Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).ToolTipText = cstrXXX
            Else
                Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X).Caption
            End If
            Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).Picture = RueckSeitePic
        End If
        
        Me.picKarteComp(Int((X + SiebenGefunden * 2) / 2)).Visible = True

        HandkartenComputer(Int((X + SiebenGefunden * 2) / 2) + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition + X)
    End If
Next

'Anfang und Endposition im Haufen
GemischtesSpielAktPosition = GemischtesSpielAktPosition + X
'Debug.Print GemischtesSpielAktPosition

frmStatistik.txtKartenImStapel = KartenAnzahl - GemischtesSpielAktPosition + 1 - SiebenGefunden


'Geberhaufen initieren
For i = 1 To picHaufen.count - frmStatistik.txtKartenImStapel
    picHaufen(i).Visible = False
Next

For i = picHaufen.count - frmStatistik.txtKartenImStapel + 1 To picHaufen.count
    If i = 29 Then
        picHaufen(i).Picture = RueckSeitePic
    ElseIf picHaufen(i).Picture <> picHaufen(i - 1).Picture Then
        picHaufen(i).Picture = picHaufen(i - 1).Picture
    End If
    picHaufen(i).Move (Me.Width * 0.7), (i - 8) * (Me.picHaufen(i).Width * 0.09) * (KartenAnzahl \ 32) + (Me.Height * 0.07)
    picHaufen(i).ZOrder
    picHaufen(i).Enabled = False
    'picHaufen(i).ToolTipText = i
    picHaufen(i).Visible = True
Next

If frmStatistik.txtKartenImStapel Mod 2 = 1 Then Stop

For i = 1 To 4
    HandkartenComputer(i).Position = i
Next

StichAktion = 0
ErsterStichSpieler = True
ErsterStichComputer = True
frmStatistik.txtGelegteKarten = 0
Me.cmdStichEnde.Visible = False
Me.menShowIndikator.Visible = True
ShowPunkteAnzeige True

'Statistik ausblenden in oberen Spiellevels
If Not Test Then
    If AktuellerSpieler.SpielerLevel >= 3 And frmStatistik.Visible Then AnimWindow frmStatistik, AW_HIDE + AW_BLEND
End If

If Geber = Computer Then
    PlayerSwitch Spieler
    If Not Tutorial Then AgentSpeak myText(177)
End If
'Me.Refresh
'If AktuellerSpieler.SpielerLevel = 1 Then
ShowTutorial 6
ShowTutorial 7
End Sub

Sub KartenNachziehen()
Dim i As Integer
Dim AnzahlKartenziehen
Dim KartenGezogen

Dim X As Integer

frmStatistik.txtKartenImStapel = KartenAnzahl - GemischtesSpielAktPosition + 1 - SiebenGefunden
If frmStatistik.txtKartenImStapel Mod 2 = 1 Then Stop

If (frmStatistik.txtKartenImStapel \ 2) < 4 Then
    AnzahlKartenziehen = frmStatistik.txtKartenImStapel \ 2
Else
    AnzahlKartenziehen = StichAktion \ 2
End If


''Karten aus dem Geberhaufen entfernen
'For i = Me.picHaufen.Count - Me.txtGelegteKarten + StichAktion To Me.picHaufen.Count - (Me.txtGelegteKarten) + 1 Step -1
'    Me.picHaufen(i).Visible = False
'    'Debug.Print i
'
'Next


'Karten an Spieler
If StichBesitzer = Spieler Then
    GoSub SpielerZieht
    GoSub ComputerZieht
Else
    GoSub ComputerZieht
    GoSub SpielerZieht
End If

'Position merken für spätere Sortiereung
For i = 1 To 4
    HandkartenComputer(i).Position = i
Next


frmStatistik.txtKartenImStapel = KartenAnzahl - GemischtesSpielAktPosition + 1 - SiebenGefunden
'Me.Refresh

Exit Sub
SpielerZieht:
For i = ZERO To 3
    If Me.picKarte(i).ToolTipText = gstrNullstr Then
        If KartenGezogen >= AnzahlKartenziehen Then Exit For
        Me.picKarte(i).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition).Caption
        Me.picKarte(i).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition).Pic
        Me.picKarte(i).Visible = True
        
        'Karteninidkator anzeigen
        HandkartenSpieler(i + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition)
        GemischtesSpielAktPosition = GemischtesSpielAktPosition + 1
        KartenGezogen = KartenGezogen + 1
    End If
Next
If ShowIndikator Then MakeIndikator
KartenGezogen = 0
Return

ComputerZieht:
For i = ZERO To 3
    If HandkartenComputer(i + 1).Bild = ZERO Then
        If KartenGezogen >= AnzahlKartenziehen Then Exit For
        'Beim testen Karte nicht verdecken
        If Test Then
            Me.picKarteComp(i).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition).Caption
            Me.picKarteComp(i).Picture = gemischtesSpiel.Karte(GemischtesSpielAktPosition).Pic
        Else
            If Schmulen Then
                Me.picKarteComp(i).ToolTipText = gemischtesSpiel.Karte(GemischtesSpielAktPosition).Caption
            Else
                Me.picKarteComp(i).ToolTipText = cstrXXX
            End If
            
            Me.picKarteComp(i).Picture = Me.picHaufen(52).Picture
            
        End If
        
        Me.picKarteComp(i).Visible = True
        HandkartenComputer(i + 1) = gemischtesSpiel.Karte(GemischtesSpielAktPosition)
        GemischtesSpielAktPosition = GemischtesSpielAktPosition + 1
        'Debug.Print GemischtesSpielAktPosition
        
        KartenGezogen = KartenGezogen + 1
    End If
Next
KartenGezogen = 0
Return

End Sub



Public Sub picKarteComp_Wirf(Index As Integer)
'computer zieht Karte
Dim str As String

'wenn Karte schon gelegt wurde
If picKarteComp(Index).ToolTipText = gstrNullstr Then
    str = "Index Fehler in picKarteComp_Wirf. Karte nicht vorhanden!"
    str = str & vbCr & "Index = " & Index
    str = str & vbCr & "Karte = " & HandkartenComputer(Index + 1).Caption
    MsgBox str, vbExclamation
    SendMail "betatester@playseven.com", "Error in " & AppInfo, Replace(str, vbCr, mailCrLf)

    Exit Sub
    
End If


'wenn computer Stich verlängern will, aber nicht darf
If StichAktion Mod 2 = ZERO And StichAktion > 1 _
    And (HandkartenComputer(Index + 1).Bild <> StichTrumpf And HandkartenComputer(Index + 1).Bild <> 7) Then
    MsgBox Gegner.SpielerName & vbCr & myText(14)
    Exit Sub
End If

MoveAgentInForm Me, picKarteComp(Index).Left, picKarteComp(Index).Top, ShowAt

'Karte legen
StichAktion = StichAktion + 1
PlaySound Karte_legen
MakeMsg Gegner.SpielerName & ": " & HandkartenComputer(Index + 1).Caption


'gelegte Karte in den Stich
'im Stich anzeigen
Me.picKarteStich(StichAktion).Visible = True
Me.picKarteStich(StichAktion).Picture = HandkartenComputer(Index + 1).Pic


'gelegte Karte in den Stich
ReDim Preserve Stich(StichAktion)
Stich(StichAktion) = HandkartenComputer(Index + 1)
frmStatistik.txtGelegteKarten = CInt(frmStatistik.txtGelegteKarten) + 1
frmStatistik.txtKarteGef(HandkartenComputer(Index + 1).Bild) = frmStatistik.txtKarteGef(HandkartenComputer(Index + 1).Bild) + 1

'Bei erstem Stich Trumpf merken
If StichAktion = 1 Then
    StichTrumpf = Stich(ONE).Bild
    str = Mid(Stich(ONE).Caption, 1, InStr(1, Stich(ONE).Caption, gstrSpace, vbTextCompare)) & gstrSpace & myText(13)
    MakeMsg str
    StichBesitzer = Computer
End If


'wer hat den Stich übernommen
If (HandkartenComputer(Index + 1).Bild = StichTrumpf Or HandkartenComputer(Index + 1).Bild = 7) Then
    StichBesitzer = Computer
Else
    StichBesitzer = Spieler
End If

'anzeigen wer den Stich nehmen kann
If StichAktion Mod 2 = ZERO Then
    Me.cmdStichEnde.Visible = True
    If StichBesitzer = Computer Then
        MakeMsg Gegner.SpielerName & gstrSpace & myText(10)
        Me.cmdStichEnde.Caption = myText(11)
    Else
        MakeMsg AktuellerSpieler.SpielerName & myText(10)
        Me.cmdStichEnde.Caption = myText(12)
    End If
Else
    Me.cmdStichEnde.Visible = False
End If

'in den handkarten reseten
picKarteComp(Index).Picture = Nothing
picKarteComp(Index).ToolTipText = gstrNullstr
HandkartenComputer(Index + 1) = NullKarte
picKarteComp(Index).Visible = False


If ShowIndikator Then
    MakeIndikator
    ShowStichIndikator
End If

'TimerStart
PlayerSwitch Spieler
If Tutorial Then
    If StichAktion = 2 Then
        ShowTutorial 10
    Else
        ShowTutorial 12
    End If
End If
End Sub

Private Sub ShowStichIndikator(Optional Show)
Static old_index As Integer

If (ShowIndikator Or Not IsMissing(Show)) And StichAktion > ZERO Then
    If Not Me.rectStichIndikator.Visible Or old_index <> StichAktion Then
        Me.rectStichIndikator.Move Me.rectStichIndikator.Left, Me.rectStichIndikator.Top, Me.picKarteStich(StichAktion).Left - Me.picKarteStich(ONE).Left + Me.picKarteStich(ONE).Width + 2 * (Me.picKarteStich(1).Left - Me.rectStichIndikator.Left)
        Me.rectStichIndikator.BackColor = IIf(StichBesitzer = Computer, vbRed, vbBlue)
        Me.rectStichIndikator.BorderColor = IIf(StichBesitzer = Computer, hellROT, hellBLAU)
        old_index = StichAktion
    End If
    If IsMissing(Show) Then
        Me.rectStichIndikator.Visible = True
    Else
        If AktuellerSpieler.SpielerLevel <= 4 Then Me.rectStichIndikator.Visible = Show
    End If
Else
    Me.rectStichIndikator.Visible = False
End If
End Sub

Private Sub pickarte_Click(Index As Integer)
'Spieler zieht Karte
Dim str As String
Dim poss As Integer

If Tutorial Then
    If TutorialStepAktuell < 7 Or TutorialStepAktuell >= 11 Or Not TutorialStepCompleted(7) Then
        Exit Sub
    End If
End If

'wenn Karte schon gelegt wurde, dann nix und bitte nochmal
If Not Test Then
    If picKarte(Index).ToolTipText = gstrNullstr Then Exit Sub
Else
    If picKarte(Index).ToolTipText = gstrNullstr Then
        MsgBox "Index Fehler in picKarte_Click"
    End If
End If


'wenn Spieler Stich verlängern will, aber nicht darf
If StichAktion Mod 2 = ZERO And StichAktion > 1 _
    And (HandkartenSpieler(Index + 1).Bild <> StichTrumpf And HandkartenSpieler(Index + 1).Bild <> 7) Then
    If AktuellerSpieler.SpielerLevel >= 5 Then
        MeisterFehler = True
        checkSieg
        Exit Sub
    Else
        AgentSpeak myText(14) & IIf(Rnd(ONE) > 0.95, vbCr & myText(15), vbNullString), True
        Exit Sub
    End If
End If

TimerEnde

PlayerSwitch Computer

If Playermodus = multiplayer Then
    #If Not Tiny Then
        frmChat.SendNetworkMessage MsgSendCard, CStr(Index), ZERO
    #End If
Else
    'Spilerhilfe in Singleplayer
    SpielerHilfe Index
End If

'Karte legen
PlaySound Karte_legen
StichAktion = StichAktion + 1
MakeMsg AktuellerSpieler.SpielerName & ": " & HandkartenSpieler(Index + 1).Caption

'gelegte Karte in den Stich
ReDim Preserve Stich(StichAktion)
Stich(StichAktion) = HandkartenSpieler(Index + 1)

'im Stich anzeigen
Me.picKarteStich(StichAktion).Visible = True
Me.picKarteStich(StichAktion).Picture = picKarte(Index).Picture
Me.picKarteStich(StichAktion).Refresh

frmStatistik.txtGelegteKarten = CInt(frmStatistik.txtGelegteKarten) + 1
frmStatistik.txtKarteGef(HandkartenSpieler(Index + 1).Bild) = CInt(frmStatistik.txtKarteGef(HandkartenSpieler(Index + 1).Bild)) + 1

'Bei erstem Stich Trumpf merken
If StichAktion = 1 Then
    StichTrumpf = Stich(ONE).Bild
    str = Mid(Stich(ONE).Caption, 1, InStr(1, Stich(ONE).Caption, gstrSpace, vbTextCompare)) & myText(13)
    MakeMsg str
    StichBesitzer = Spieler
End If


'wer hat den Stich übernommen
If (HandkartenSpieler(Index + 1).Bild = StichTrumpf Or HandkartenSpieler(Index + 1).Bild = 7) Then
    StichBesitzer = Spieler
Else
    StichBesitzer = Computer
End If

'anzeigen wer den Stich nehmen kann
If StichAktion Mod 2 = ZERO Then
    If StichBesitzer = Spieler Then
'        Me.cmdStichEnde.Visible = True
        MakeMsg AktuellerSpieler.SpielerName & myText(10)
        Me.cmdStichEnde.Caption = myText(12)
    Else
        MakeMsg Gegner.SpielerName & myText(10)
        Me.cmdStichEnde.Caption = myText(11)
    End If
Else
    Me.cmdStichEnde.Visible = False
End If

'in den handkarten reseten
picKarte(Index).Picture = Nothing
picKarte(Index).Refresh
picKarte(Index).Visible = False

picKarte(Index).ToolTipText = gstrNullstr
HandkartenSpieler(Index + 1) = NullKarte

'Me.Refresh

If ShowIndikator Then
    MakeIndikator
    ShowStichIndikator
End If

If Playermodus = singleplayer Then
    If StichAktion Mod 2 = 1 Then
        If Tutorial Then
            If frmStatistik.txtGelegteKarten = 1 Then
                ShowTutorial 8
            Else
                picKarteComp_Wirf (ComputerAntwort(HandkartenComputer, 5, CInt(lblSpielerPunkte(ZERO).Caption), CInt(lblComputerPunkte(ZERO).Caption), ErsterStichComputer) - 1)
            End If
        Else
            picKarteComp_Wirf (ComputerAntwort(HandkartenComputer, AktuellerSpieler.SpielerLevel, CInt(lblSpielerPunkte(ZERO).Caption), CInt(lblComputerPunkte(ZERO).Caption), False) - 1)
        End If
    Else
        poss = ComputerVerlaengert(HandkartenComputer, Computer, CInt(Me.lblComputerPunkte(ZERO).Caption), CInt(Me.lblSpielerPunkte(ZERO).Caption), ErsterStichSpieler)
        If poss > -1 Then
            picKarteComp_Wirf poss - 1
        Else
            StichEnde
        End If
    End If
End If

If Test Then ComputerSpieltFuerSpieler

End Sub

Private Sub SpielerHilfe(GeworfeneKartePos As Integer)
Dim pos As Integer
    If (AktuellerSpieler.SpielerLevel = 1 And AgentGivesTips And frmStatistik.txtGelegteKarten < 31 _
        And ((CInt(Me.lblComputerPunkte(ZERO).Caption) - CInt(Me.lblSpielerPunkte(ZERO).Caption) >= 2) _
        Or (CInt(Me.lblComputerPunkte(ZERO).Caption) >= 3))) Or Tutorial Then
        
        pos = KI_SpielKarte()
        If pos <> GeworfeneKartePos + 1 Then
            If pos = -1 Then
                AgentGivesTipp pos, True
            ElseIf pos = 0 Then
                'nix --> computer hat kein rat
            ElseIf GeworfeneKartePos = -2 And pos <> -1 Then
                AgentGivesTipp pos, True
            ElseIf HandkartenSpieler(pos).Bild <> HandkartenSpieler(GeworfeneKartePos + 1).Bild Then
                'computer hätte anders geworfen
                AgentGivesTipp pos, True
            Else
                'nix -> spieler hat wie computer geworfen
            End If
        End If
    End If
End Sub

Sub PlayerSwitch(NaechsterSpieler As Integer)
Dim Switch As Boolean
If NaechsterSpieler = Computer Then
    Switch = True
Else
    Switch = False
    TimerStart
End If

'Me.cmdStichEnde.Visible = Not Switch
EnablePicKarten Not Switch
End Sub
Private Sub EnablePicKarten(Switch As Boolean)
Dim i As Integer
For i = Me.picKarte.LBound To Me.picKarte.UBound
    Me.picKarte(i).Enabled = Switch
Next
End Sub


Sub TimerEnde()
Me.StichEndeTimer.Enabled = False
Me.lblTime(ZERO).Visible = False
Me.lblTime(ONE).Visible = False
ZeitUeberschreitung = False
End Sub
Public Sub StichEnde()

Dim NaechsterSpieler As Integer
Dim i As Integer
Dim Antw As Integer
Dim SpannungsIndex As Integer
Dim HaufenPos As Integer
Dim perc As Single

If Tutorial Then
    If TutorialStepAktuell >= 8 And Not TutorialStepCompleted(8) Then
        Exit Sub
    End If
End If

TimerEnde
NaechsterSpieler = StichBesitzer
MakeMsg cstrLinie

'Karten in den Stich des Spielers legen
If StichBesitzer = Spieler Then
    If ErsterStichSpieler Then
        ReDim Preserve HaufenSpieler(StichAktion)
        ErsterStichSpieler = False
    Else
        ReDim Preserve HaufenSpieler(UBound(HaufenSpieler) + StichAktion)
    End If
    'Punkte zählen
    For i = 1 To UBound(Stich)
        HaufenSpieler(i) = Stich(i)
        If Stich(i).Bild = 10 Or Stich(i).Bild = Ass Then
            setSpielerPunkte CInt(Me.lblSpielerPunkte(ZERO).Caption) + 1
            SpannungsIndex = SpannungsIndex + 1
        End If
    Next
    'Audio
    If SpannungsIndex > ZERO Then
        PlaySound SpielerNimmtPunkte, -1300 + (1000 \ 8 * SpannungsIndex)
        
        If SpannungsIndex >= 3 Then AgentAnim Surprised
        
        'agent reaktion auf punktenehmen
        If SpannungsIndex >= 4 And Playermodus = singleplayer Then
            perc = Rnd(ONE)
            Select Case perc
                Case Is > 0.95
                    AgentSpeak myText(182)
                Case Is > 0.9
                    AgentSpeak "Grmpf !"
                Case Is > 0.85
                    AgentSpeak "Aaaaaaaahhh !"
                Case Is > 0.8
                    AgentSpeak "Mennno !"
                Case Is > 0.75
                    AgentSpeak myText(183)
                Case Is > 0.7
                    AgentSpeak myText(184)
                Case Is > 0.65
                    AgentSpeak myText(185)
                Case Is > 0.6
                    AgentSpeak myText(186)
                Case Is > 0.55
                    AgentSpeak myText(187)
                Case Is > 0.5
                    AgentSpeak myText(188)
                Case Is > 0.45
                    AgentSpeak myText(189)
                Case Is > 0.4
                    AgentSpeak myText(190)
            End Select
        
        End If
    Else
        PlaySound SpielerNimmt
    End If
Else
    AgentTakeStich
    
    If ErsterStichComputer Then
        ReDim Preserve HaufenComputer(StichAktion)
        ErsterStichComputer = False
    Else
        ReDim Preserve HaufenComputer(UBound(HaufenComputer) + StichAktion)
    End If
    'Punkte zählen
    For i = 1 To UBound(Stich)
        HaufenComputer(i) = Stich(i)
        If Stich(i).Bild = 10 Or Stich(i).Bild = Ass Then
            setComputerPunkte CInt(Me.lblComputerPunkte(ZERO).Caption) + 1
            SpannungsIndex = SpannungsIndex + 1
        End If
        'Anzeige
        'debug.Print , Me.picHaufen.Count - frmStatistik.txtGelegteKarten + i, 5750 + (UBound(HaufenComputer) - i) * 50

    Next
    'Audio
    If SpannungsIndex > ZERO Then
        PlaySound ComputerNimmtPunkte, -1000 + (1000 \ 8 * SpannungsIndex)
        
        If SpannungsIndex >= 3 Then AgentAnim Pleased
        'agent reaktion auf punktenehmen
        If SpannungsIndex >= 4 And Playermodus = singleplayer Then
            perc = Rnd(ONE)
            Select Case perc
                Case Is > 0.95
                    AgentSpeak myText(191)
                Case Is > 0.9
                    AgentSpeak myText(192)
                Case Is > 0.85
                    AgentSpeak myText(193)
                Case Is > 0.8
                    AgentSpeak myText(194)
                Case Is > 0.75
                    AgentSpeak myText(195)
                Case Is > 0.7
                    AgentSpeak myText(196)
                Case Is > 0.65
                    AgentSpeak myText(197)
                Case Is > 0.6
                    AgentSpeak myText(198)
                Case Is > 0.55
                    AgentSpeak myText(199)
                Case Is > 0.5
                    AgentSpeak myText(200)
                Case Is > 0.45
                    AgentSpeak myText(201)
                Case Is > 0.4
                    AgentSpeak myText(202)
            End Select
        End If
    Else
        PlaySound ComputerNimmt
    End If
End If





EnablePicKarten False
WaitTick 555 - 555 \ 5 * (AktuellerSpieler.SpielerLevel + 1) '+ UBound(Stich) * 66
'Stich verstecken
For i = 1 To StichAktion
    Me.picKarteStich(StichAktion - i + 1).Visible = False
    'Anzeige
    WaitTick 66 - 66 \ 5 * (AktuellerSpieler.SpielerLevel + 1)
    If StichBesitzer = Spieler Then
        HaufenPos = (UBound(HaufenSpieler) - StichAktion + i - 1) * (Me.picKarteStich(ONE).Width * 0.0617)
    Else
        HaufenPos = (UBound(HaufenComputer) - StichAktion + i - 1) * (Me.picKarteStich(ONE).Width * 0.0617)
    End If
    'Me.picHaufen(Me.picHaufen.count - frmStatistik.txtGelegteKarten + i).Visible = False
    Me.picHaufen(Me.picHaufen.count - frmStatistik.txtGelegteKarten + StichAktion - i + 1).ZOrder
    Me.picHaufen(Me.picHaufen.count - frmStatistik.txtGelegteKarten + StichAktion - i + 1).Move (Me.Width * 0.687) + HaufenPos, IIf(StichBesitzer = Spieler, Me.Height * 0.7287, Me.Height * 0.0779)
    Me.picHaufen(Me.picHaufen.count - frmStatistik.txtGelegteKarten + StichAktion - i + 1).Visible = True

Next

'For i = 1 To 8
'    Me.picHaufen(i).Visible = False
'Next
Me.rectStichIndikator.Visible = False

'wenn sieg vorzeitig feststeht
If CInt(frmStatistik.txtGelegteKarten.Text) < KartenAnzahl _
    And ((Not ErsterStichSpieler And Me.lblComputerPunkte(ZERO) > 4) _
    Or (Not ErsterStichComputer And Me.lblSpielerPunkte(ZERO) > 4)) Then
        checkSieg
        Exit Sub
End If

'wenn alle Karten gelegt, wer hat den da gewonnen
If frmStatistik.txtGelegteKarten = KartenAnzahl Then
    checkSieg
    Exit Sub
End If



StichTrumpf = -1
KartenNachziehen
'Me.rectStichIndikator.Visible = False
Me.cmdStichEnde.Visible = False

StichAktion = ZERO
StichBesitzer = ZERO


ReDim Preserve Stich(ZERO)
'Nächsten Spieler anzeigen
'wer macht nächsten zug ?
If NaechsterSpieler = Computer Then
    PlayerSwitch Computer
    If Playermodus = singleplayer And Not Tutorial Then ComputerAktion
Else
    PlayerSwitch Spieler
    If Test Then ComputerSpieltFuerSpieler
End If
ShowTutorial 11
End Sub




Sub ComputerHebtAb()

picHaufen_Click (Int(KartenAnzahl - 1) * Rnd + 1)

End Sub



Private Sub TimerStart()
On Error Resume Next

If AktuellerSpieler.SpielerLevel > 3 Then
    Me.lblTime(ONE).ForeColor = vbWhite
    Me.lblTime(ZERO) = gstrTimeInterval
    Me.lblTime(ONE) = gstrTimeInterval
    Me.lblTime(ZERO).Visible = True
    Me.lblTime(ONE).Visible = True
'    Me.lblTime(one).Top = 3150
'    Me.lblTime(zero).Top = 3210
    Me.StichEndeTimer.Enabled = True
End If

End Sub

Sub checkSieg()
'Überprüft wer gewonnen hat und gibt folgebefehle
'
On Error GoTo ERRHand

    Dim Antw As Integer
    If ZeitUeberschreitung Then
        TimerEnde
        PlaySound spielerFehler
        AgentAnim Confused
        #If Not Tiny Then
            If Playermodus = multiplayer Then frmChat.SendNetworkMessage MsgSendZeitueberschreitung, gstrNullstr, ZERO
        #End If
        
        AgentSpeak myText(17) & vbCr & myText(18), True
        Write2Log myText(17) & vbCr & myText(18)
        
        setComputerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + ONE
        Geber = Spieler
    ElseIf ZeitUeberschreitungGegner Then
        TimerEnde
        PlaySound spielerFehler
        AgentAnim Confused
        
        AgentSpeak Gegner.SpielerName & myText(69), True
        Write2Log Gegner.SpielerName & myText(69)
        
        setSpielerRndPunkte CInt(lblSpielerRundenPunkte(ZERO)) + ONE
        Geber = Computer
    ElseIf MeisterFehler Then
        TimerEnde
        PlaySound spielerFehler
        
        AgentAnim Decline

        #If Not Tiny Then
            If Playermodus = multiplayer Then frmChat.SendNetworkMessage MsgSendMeisterfehler, gstrNullstr, ZERO
        #End If
        
        AgentSpeak myText(14) & vbCr & myText(18), True
        Write2Log myText(14) & vbCr & myText(18)
        
        setComputerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + ONE
        Geber = Spieler
    ElseIf MeisterFehlerGegner Then
        TimerEnde
        PlaySound spielerFehler
        
        AgentAnim Decline
       
        AgentSpeak Gegner.SpielerName & myText(70), True
        Write2Log Gegner.SpielerName & myText(70)
        
        setSpielerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + ONE
        Geber = Computer
    'normaler Gewinn
    ElseIf Me.lblSpielerPunkte(ZERO) > Me.lblComputerPunkte(ZERO) Then
    
        If ErsterStichComputer Then
        '16er
            AgentSpeak Gegner.SpielerName & gstrSpace & myText(21) & vbCr & myText(24) & myText(22), True
            Write2Log Gegner.SpielerName & gstrSpace & myText(21) & vbCr & myText(24) & myText(22)
            
            setSpielerRndPunkte CInt(lblSpielerRundenPunkte(ZERO)) + 2
        Else
        'normal
            AgentSpeak myText(24), True
            Write2Log myText(24)
            
            setSpielerRndPunkte CInt(lblSpielerRundenPunkte(ZERO)) + ONE
        End If
        Geber = Computer
    '4 und letzter
    ElseIf Me.lblSpielerPunkte(ZERO) = Me.lblComputerPunkte(ZERO) Then
        If StichBesitzer = Spieler Then
            
            AgentSpeak myText(25) & vbCr & myText(24), True
            Write2Log myText(25) & vbCr & myText(24)
            
            setSpielerRndPunkte CInt(lblSpielerRundenPunkte(ZERO)) + ONE
            Geber = Computer
        Else
            If Rnd(ONE) > 0.9 And Playermodus = singleplayer Then
                AgentAnim Explain
                AgentSpeak myText(205)
            End If
            
            AgentSpeak myText(25) & vbCr & Gegner.SpielerName & myText(27), True
            Write2Log myText(25) & vbCr & Gegner.SpielerName & myText(27)
            
            setComputerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + ONE
            Geber = Spieler
        End If
    Else
        If ErsterStichSpieler Then
        
            AgentSpeak myText(28) & vbCr & Gegner.SpielerName & myText(27) & myText(29), True
            Write2Log myText(28) & vbCr & Gegner.SpielerName & myText(27) & myText(29)
            
            setComputerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + 2
        Else
            AgentSpeak Gegner.SpielerName & myText(27), True
            Write2Log Gegner.SpielerName & myText(27)
            
            setComputerRndPunkte CInt(lblComputerRundenPunkte(ZERO)) + ONE
        End If
        Geber = Spieler
    End If
    SetSpielerInDB AktuellerSpieler, False
    Write2DB Spiel, Me.lblSpielerPunkte(ZERO), Me.lblComputerPunkte(ZERO), Now, (Geber = Computer)
    
    Write2Log AktuellerSpieler.SpielerName & gstrDblDot & Me.lblSpielerPunkte(ZERO) & vbTab & _
        Gegner.SpielerName & gstrDblDot & Me.lblComputerPunkte(ZERO)
    Write2Log cstrLinie
    
    If Me.lblComputerRundenPunkte(ZERO) >= 7 And Me.lblComputerRundenPunkte(ZERO) - Me.lblSpielerRundenPunkte(ZERO) >= 2 Then
        menStatistik_Click
        PlaySound ComputerGewinntRunde
        
        AgentAnim Sad
        If (Playermodus = singleplayer) Or (Playermodus = multiplayer And AktuellerSpieler.SpielOption = Liga) Then
            Write2DB Runde, Me.lblSpielerRundenPunkte(ZERO), Me.lblComputerRundenPunkte(ZERO), Now
            frmStatistik.StatistikAktualisieren
            frmStatistik.SetPlayerLevel
        End If
        
        If Not Test Then
            Antw = AgentQuestion(Gegner.SpielerName & myText(30) & vbCr & myText(31), Gegner.SpielerName)
        Else
            Antw = vbYes
        End If
        
        If Antw = vbYes Then
            Geber = Spieler
            Init
        Else
            LaufendesSpiel = False
            #If Not Tiny Then
                If Playermodus = multiplayer Then
                    frmChat.SendNetworkMessage MsgSendSpielAbbruch, gstrNullstr, ZERO
                    If ServerConnected And Not ServerEventsForm Is Nothing Then
                        ServerEventsForm.SendMsg2Server Msg_PlayerLost
                    End If
                End If
            #End If
            If Test Then End
        End If
    ElseIf Me.lblSpielerRundenPunkte(ZERO) >= 7 And Me.lblSpielerRundenPunkte(ZERO) - Me.lblComputerRundenPunkte(ZERO) >= 2 Then
        menStatistik_Click
        PlaySound SpielerGewinntRunde
        AgentAnim Congratulate
        
        If (Playermodus = singleplayer) Or (Playermodus = multiplayer And AktuellerSpieler.SpielOption = Liga) Then
            Write2DB Runde, Me.lblSpielerRundenPunkte(ZERO), Me.lblComputerRundenPunkte(ZERO), Now
            frmStatistik.StatistikAktualisieren
            frmStatistik.SetPlayerLevel
        End If
        
        If Not Test Then
            Antw = AgentQuestion(myText(32) & vbCr & myText(31), Gegner.SpielerName)
        Else
            Antw = vbYes
        End If
        
        If Antw = vbYes Then
            Geber = Computer
            Init
        Else
            LaufendesSpiel = False
            Me.cmdStichEnde.Visible = False
            #If Not Tiny Then
                If Playermodus = multiplayer Then
                    frmChat.SendNetworkMessage MsgSendSpielAbbruch, gstrNullstr, ZERO
                    Unload Me
                    If ServerConnected And Not ServerEventsForm Is Nothing Then
                        ServerEventsForm.SendMsg2Server Msg_PlayerWon
                    End If
                End If
            #End If
        End If
        If Test Then End
    Else
        frmStatistik.StatistikAktualisieren
        If Geber = Spieler Then
            AgentAnim Writes
        Else
            AgentAnim Confused
        End If
        Aktualisieren
        Init True
    End If
Exit Sub
ERRHand:
If ErrorBox("checkSieg", Err) Then Resume Next
End Sub


Sub sendMixedGame()
'Sendet das gemsichte Spiel an den Gegenespieler

Dim str As String
Dim i As Integer

On Error Resume Next

#If Not Tiny Then
    For i = 1 To KartenAnzahl
        str = str & i & gstrDblDot
        str = str & CStr(gemischtesSpiel.Karte(i).Bild) & gstrDblDot
        str = str & CStr(gemischtesSpiel.Karte(i).Color) & gstrDblDot
        str = str & CStr(gemischtesSpiel.Karte(i).BildName) & gstrDblDot
        str = str & CStr(gemischtesSpiel.Karte(i).Caption) & gstrDblDot
    Next
    frmChat.SendNetworkMessage MsgSendMixedGame, str, ZERO
    UpdateChat SystemMsg, myText(33), frmChat
#End If
End Sub

Sub sendGeber(Geber As Integer)
'Sndet den Geber zum Gegenspieler
On Error Resume Next

#If Not Tiny Then
    frmChat.SendNetworkMessage MsgSendGeber, CStr(Geber), ZERO
    UpdateChat SystemMsg, myText(34) & " (" _
        & IIf(-Geber = Spieler, AktuellerSpieler.SpielerName, Gegner.SpielerName) & ")", frmChat
#End If
End Sub

Private Sub picKarte_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'mouseover bei den Spielrkarten
'bewegt die Karte ein Stückchen nach oben
    picKarte(Index).Move picKarte(Index).Left, (Me.Height * picKarte_Top) - (Me.Height * 0.03)
    KarteGehoben = True
End Sub



Private Sub picKarteStich_Click(Index As Integer)
    If Me.cmdStichEnde.Visible And Me.cmdStichEnde.Visible Then
        cmdStichEnde_Click
    End If
End Sub

Private Sub picKarteStich_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'Moseover beim Stich
'zeigt den Stichindikator an
'geht im Tutorial weiter

On Error Resume Next

    Me.picKarteStich(Index).ToolTipText = IIf(StichBesitzer = Spieler, myText(67), Gegner.SpielerName & myText(68))
    If ShowIndikator Then ShowStichIndikator True
    If Tutorial And TutorialStepCompleted(8) And Not TutorialStepCompleted(9) Then
        ShowTutorial 9
    End If
End Sub

Private Sub ShareWareTimer_Timer()
'Überprüft ob die DemoZeit vorbei ist
'Fordert zur Registrierunug auf oder beendet das Spiel
Static minutes As Long
On Error Resume Next

minutes = minutes + 1
If minutes >= DemoTime Then
    AgentSpeak myText(65), True
    If Not IsRegisteredFree And StartAnz < MaxStartsWReg * 2 Then
        AgentOffer
    Else
        AgentSpeak myText(66)
        RegisterApp Me
    End If
    DoSleep 2000
    Unload Me
End If
End Sub

Private Sub AgentOffer()
'Agent offerriert Registrierung
On Error Resume Next
    AgentSpeak myText(59), True
    AgentSpeak myText(60) & myText(61), True
    DoSleep 8888
    If AgentQuestion(myText(62), myText(63)) = vbYes Then
        getWebRegKey
    Else
        AgentSpeak myText(64)
    End If
End Sub

Private Sub StichEndeTimer_Timer()
On Error Resume Next

    Me.lblTime(ZERO) = val(Me.lblTime(ZERO)) - 1 & SecundeEinheit
    Me.lblTime(ONE) = val(Me.lblTime(ONE)) - 1 & SecundeEinheit
    Me.lblTime(ONE).Refresh
    Me.lblTime(ZERO).Refresh
    PlaySound TimeTick, -2500 + ((7 - val(Me.lblTime(ZERO))) * 333)
    If val(Me.lblTime(ONE)) <= ZERO Then
        ZeitUeberschreitung = True
        checkSieg
    ElseIf val(Me.lblTime(ONE)) = 3 Then
        Me.lblTime(ONE).ForeColor = vbRed
    End If
End Sub

Private Sub WriteMsg(str As String, Optional delay)
'Schreibt eine Message auf zwei Labels mitten in die Hauptform
'optional Zeit nach der die Message verschwindet

Dim i As Long
Const LeftP As Long = 120
Const Breite As Long = 5205

On Error Resume Next

'Me.lblMsg(zero).left = LeftP
'Me.lblMsg(one).left = LeftP - 30
If frmMain_Loaded Then
    Me.lblMsg(ZERO).Visible = Not (str = vbNullString)
    Me.lblMsg(ONE).Visible = Not (str = vbNullString)
    Me.lblMsg(ZERO).Enabled = Not (str = vbNullString)
    Me.lblMsg(ONE).Enabled = Not (str = vbNullString)
    
    Me.lblMsg(ZERO) = str
    Me.lblMsg(ONE) = str
    
    Me.lblMsg(ZERO).Refresh
    Me.lblMsg(ONE).Refresh
    
    'For i = Zero To Breite
    '    Me.lblMsg(zero).Width = i
    '    Me.lblMsg(one).Width = i
    'Next
    'DoEvents
    'Me.lblMsg(zero).Refresh
    'Me.lblMsg(one).Refresh
    
    If Not IsMissing(delay) Then Sleep delay
    
    'For i = Zero To Breite
    '    Me.lblMsg(zero).Width = Width - i
    '    Me.lblMsg(one).Width = Width - i
    '    Me.lblMsg(one).left = LeftP + i - 30
    '    Me.lblMsg(zero).left = LeftP + i
    'Next
    
    'Me.lblMsg(zero).Visible = False
    'Me.lblMsg(one).Visible = False
End If

End Sub

Private Sub Timer1_Timer()
'berprüft ob sich das chat fenster bewegt hat und dockt die hauptform an
#If Not Tiny Then
    Static X As Long
    Static Y As Long
    If frmChat_Loaded Then
        If frmChat.Top <> Y Or frmChat.Left <> X Then
            Y = frmMain.Top
            X = frmMain.Left
            Dock2Chat
        End If
    End If
#End If
End Sub

Private Sub setSpielerPunkte(Pt As Integer)
    Me.lblSpielerPunkte(ZERO).Caption = CStr(Pt)
    Me.lblSpielerPunkte(ONE).Caption = CStr(Pt)
End Sub
Private Sub setComputerPunkte(Pt As Integer)
    Me.lblComputerPunkte(ZERO).Caption = CStr(Pt)
    Me.lblComputerPunkte(ONE).Caption = CStr(Pt)
End Sub

Public Sub setSpielerRndPunkte(Pt As Integer)
'Setzt die Punkte des Spielrs hoch
    
    'Anzeige
    Me.lblSpielerRundenPunkte(ZERO).Caption = CStr(Pt)
    Me.lblSpielerRundenPunkte(ONE).Caption = CStr(Pt)
    
    boolPlayerWon = True
    
    'Highscore
    If Playermodus = singleplayer Then
        AktuellerSpieler.Points = AktuellerSpieler.Points + AktuellerSpieler.SpielerLevel + 1
    ElseIf Playermodus = multiplayer And AktuellerSpieler.SpielOption = Liga Then
        AktuellerSpieler.Points = AktuellerSpieler.Points + Gegner.SpielerLevel + 1
    End If
    
    setHighscore AktuellerSpieler.Points
    
End Sub

Private Sub setComputerRndPunkte(Pt As Integer)
    Me.lblComputerRundenPunkte(ZERO).Caption = CStr(Pt)
    Me.lblComputerRundenPunkte(ONE).Caption = CStr(Pt)
    boolPlayerWon = False
End Sub

