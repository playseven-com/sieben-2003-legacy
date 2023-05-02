VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStatistik 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800000&
   BorderStyle     =   0  'Kein
   Caption         =   "Statistik"
   ClientHeight    =   6645
   ClientLeft      =   9660
   ClientTop       =   735
   ClientWidth     =   2955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   2955
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLevel 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   90
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   40
      Top             =   2910
      Width           =   420
   End
   Begin VB.TextBox txtGelegteKarten 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3540
      TabIndex        =   36
      Top             =   780
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   3180
      Top             =   5220
   End
   Begin VB.ListBox lstVerlauf 
      BackColor       =   &H00404040&
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
      Height          =   630
      IntegralHeight  =   0   'False
      ItemData        =   "frmStatistik.frx":0000
      Left            =   60
      List            =   "frmStatistik.frx":0002
      TabIndex        =   35
      Top             =   5940
      Width           =   2835
   End
   Begin VB.TextBox txtRundenVerl 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   31
      Top             =   5160
      Width           =   615
   End
   Begin VB.TextBox txtRundenGew 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1680
      TabIndex        =   29
      Top             =   5520
      Width           =   615
   End
   Begin VB.TextBox txtRundenGes 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   27
      Top             =   4800
      Width           =   615
   End
   Begin VB.TextBox txtSpieleVerl 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   1680
      TabIndex        =   24
      Top             =   3960
      Width           =   615
   End
   Begin VB.TextBox txtSpieleGew 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   285
      Left            =   1680
      TabIndex        =   21
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtSpieleGes 
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1680
      TabIndex        =   19
      Top             =   3600
      Width           =   615
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   14
      Left            =   390
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   720
      Width           =   585
   End
   Begin VB.TextBox txtKartenImStapel 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3540
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   300
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   10
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C000C0&
      Height          =   375
      Index           =   7
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   900
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   11
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1590
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   12
      Left            =   1110
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1590
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   13
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1590
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   9
      Left            =   720
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   585
   End
   Begin VB.TextBox txtKarteGef 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   8
      Left            =   1500
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   2280
      Width           =   585
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3180
      Top             =   5700
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
            Picture         =   "frmStatistik.frx":0004
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":12F7
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":16A5
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":3A35
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":5F0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":8356
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStatistik.frx":AA06
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblName 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
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
      Height          =   255
      Left            =   630
      TabIndex        =   41
      Top             =   2850
      Width           =   2295
   End
   Begin VB.Label lblExit 
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
      Left            =   2550
      TabIndex        =   38
      ToolTipText     =   "Exit"
      Top             =   0
      Width           =   255
   End
   Begin VB.Label lblLevel 
      BackStyle       =   0  'Transparent
      Caption         =   "Level"
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
      Height          =   255
      Left            =   630
      TabIndex        =   37
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   2940
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   2940
      Y1              =   5880
      Y2              =   5880
   End
   Begin VB.Label lblRdVerlPrz 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2310
      TabIndex        =   34
      Top             =   5160
      Width           =   585
   End
   Begin VB.Label lblRdpGewPrz 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2310
      TabIndex        =   33
      Top             =   5520
      Width           =   585
   End
   Begin VB.Label lblRundenVerl 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Runden verloren"
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
      Height          =   285
      Left            =   60
      TabIndex        =   32
      Top             =   5160
      Width           =   1695
   End
   Begin VB.Label lblRundenGew 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Runden gewonnen"
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
      Height          =   285
      Left            =   60
      TabIndex        =   30
      Top             =   5520
      Width           =   1785
   End
   Begin VB.Label lblRundenGes 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Runden gesamt"
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
      Height          =   285
      Left            =   60
      TabIndex        =   28
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Label lblSpVerlPrz 
      BackColor       =   &H00000000&
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2310
      TabIndex        =   26
      Top             =   3960
      Width           =   585
   End
   Begin VB.Label lblSpieleVerl 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Spiele verloren"
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
      Height          =   285
      Left            =   60
      TabIndex        =   25
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label lblSpGewPrz 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   2310
      TabIndex        =   23
      Top             =   4320
      Width           =   585
   End
   Begin VB.Label lblSpieleGew 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Spiele gewonnen"
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
      Height          =   285
      Left            =   60
      TabIndex        =   22
      Top             =   4320
      Width           =   1695
   End
   Begin VB.Label lblSpieleGes 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Spiele gesamt"
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
      Height          =   285
      Left            =   60
      TabIndex        =   20
      Top             =   3600
      Width           =   1695
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   60
      X2              =   2940
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Karten gefallen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2940
      TabIndex        =   18
      Top             =   780
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblSieben 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "7"
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
      Left            =   1110
      TabIndex        =   17
      Top             =   660
      Width           =   585
   End
   Begin VB.Label lblZehn 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Zehn"
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
      Left            =   1830
      TabIndex        =   16
      Top             =   480
      Width           =   585
   End
   Begin VB.Label lblAs 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "As"
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
      Left            =   390
      TabIndex        =   15
      Top             =   480
      Width           =   585
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Karten im Haufen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   2940
      TabIndex        =   14
      Top             =   300
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblKoenig 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "König"
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
      Left            =   300
      TabIndex        =   13
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label lblDame 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Dame"
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
      Left            =   1110
      TabIndex        =   12
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label lblBube 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Bube"
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
      Left            =   1920
      TabIndex        =   11
      Top             =   1350
      Width           =   585
   End
   Begin VB.Label lblAcht 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "8"
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
      Left            =   1500
      TabIndex        =   10
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label lblNeun 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "9"
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
      Left            =   720
      TabIndex        =   9
      Top             =   2040
      Width           =   585
   End
   Begin VB.Label Label3 
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
      Height          =   345
      Left            =   2640
      TabIndex        =   39
      ToolTipText     =   "Exit"
      Top             =   60
      Width           =   195
   End
   Begin VB.Label lblCaption 
      Appearance      =   0  '2D
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
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
      Index           =   0
      Left            =   0
      TabIndex        =   42
      Top             =   0
      Width           =   8685
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Undurchsichtig
      Height          =   645
      Left            =   0
      Top             =   2820
      Width           =   2955
   End
End
Attribute VB_Name = "frmStatistik"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
Private myText() As String
'Private oldHeight As Long
'Private cFormResizer As New clFormResizer

Private Sub Form_Activate()
    Me.StatistikAktualisieren
    ShowPlayerLevel
End Sub

Private Sub Form_Load()
'initialisierungen
    
'    cFormResizer.Initialize Me
'    cFormResizer.AutoResize = False
    
'    oldHeight = Me.Height
    
    'form andocken
    Dock2Main
    
    'Hintergrund setzen
    SetBackGround Me
    'form Rund machen
    makeRoundEdges Me
    
    frmStatistik_Loaded = True
    Me.Timer1.Enabled = True
    
    'Texte laden
    LoadObjectText Me.Name, myText()
    Me.lblKoenig = myText(ZERO)
    Me.lblDame = myText(ONE)
    Me.lblBube = myText(2)
    Me.lblAs = myText(3)
    Me.lblZehn = myText(4)
    Me.lblRundenGes = myText(5)
    Me.lblRundenGew = myText(6)
    Me.lblRundenVerl = myText(7)
    Me.lblSpieleGes = myText(8)
    Me.lblSpieleGew = myText(6)
    Me.lblSpieleVerl = myText(7)
    
    
End Sub
'Public Sub Maximize()
'    Me.Move Screen.Width - Me.Width, ZERO, Me.Width, Screen.Height
'End Sub
'Public Sub Normalize()
'    Me.Height = oldHeight
'    Dock2Main
'End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.ForeColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    SetWindowPos frmStatistik.hWnd, HWND_NOTOPMOST, frmStatistik.Left, frmStatistik.Top, frmStatistik.Width, frmStatistik.Height, 3
    AnimWindow Me, AW_HIDE + AW_BLEND
    If AktuellerSpieler.SpielerLevel < 3 Then
        frmMain.menStatistik.Enabled = True
    ElseIf AktuellerSpieler.SpielerLevel >= 3 And Not LaufendesSpiel Then
        frmMain.menStatistik.Enabled = True
    Else
        'nix
    End If
    frmStatistik_Loaded = False
End Sub

Public Sub Dock2Main()
On Error Resume Next
    Me.Move frmMain.Left + frmMain.Width - 40, frmMain.Top
End Sub

Private Sub lblCaption_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub lblExit_Click()
    Unload Me
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblExit.FontSize = Me.lblExit.FontSize - 2
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblExit.ForeColor = vbRed
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblExit.FontSize = Me.lblExit.FontSize + 2
End Sub


Private Sub Timer1_Timer()
On Error Resume Next
Static X As Long
Static Y As Long


'Überprüfung ob Position des hauptfensters sich geämdert hat
If (frmMain.Top <> Y Or frmMain.Left <> X) Then
    Y = frmMain.Top
    X = frmMain.Left
    'wenn ja, an das fenster andocken
    If frmMain.WindowState = ZERO Then
        Dock2Main
    ElseIf frmMain.WindowState = vbMinimized Then
        Me.WindowState = vbMinimized
    Else
        If Me.WindowState <> ZERO Then Me.WindowState = ZERO
    End If
End If

End Sub

Public Sub StatistikAktualisieren()
Dim PunkteInf As PunkteInfo
Const cstrShortLinie As String = "---"
On Error GoTo ERRHand

'Statistik berechnen
    PunkteInf = ReadFromDB(Spiel)
    Me.txtSpieleGes = PunkteInf.GesamtSpiele
    Me.txtSpieleGew = PunkteInf.GewonneneSpiele
    Me.txtSpieleVerl = PunkteInf.GesamtSpiele - PunkteInf.GewonneneSpiele
    If PunkteInf.GesamtSpiele > ZERO Then
        Me.lblSpGewPrz.Caption = Format$(CStr(PunkteInf.GewonneneSpiele / PunkteInf.GesamtSpiele), PercFormat)
        Me.lblSpVerlPrz = Format$(CStr(1 - PunkteInf.GewonneneSpiele / PunkteInf.GesamtSpiele), PercFormat)
    Else
        Me.lblSpGewPrz.Caption = cstrShortLinie
        Me.lblSpVerlPrz.Caption = cstrShortLinie
    End If
    
    PunkteInf = ReadFromDB(Runde)
    Me.txtRundenGes = PunkteInf.GesamtSpiele
    Me.txtRundenGew = PunkteInf.GewonneneSpiele
    Me.txtRundenVerl = PunkteInf.GesamtSpiele - PunkteInf.GewonneneSpiele
    If PunkteInf.GesamtSpiele > ZERO Then
        Me.lblRdpGewPrz.Caption = Format$(CStr(PunkteInf.GewonneneSpiele / PunkteInf.GesamtSpiele), PercFormat)
        Me.lblRdVerlPrz = Format$(CStr(1 - PunkteInf.GewonneneSpiele / PunkteInf.GesamtSpiele), PercFormat)
    Else
        Me.lblRdpGewPrz.Caption = cstrShortLinie
        Me.lblRdVerlPrz.Caption = cstrShortLinie
    End If
    
Exit Sub
ERRHand:
If ErrorBox("StatistikAktualisieren", Err) Then Resume Next
End Sub

Public Sub SetPlayerLevel()
Dim WinPerc As Single, setNewLevel As Boolean
Dim str As String

On Error GoTo ERRHand

'Gewinnquote errechnen
WinPerc = val(Replace(Me.lblRdpGewPrz, gstrDot, gstrKomma))
Select Case WinPerc
    Case Is > 50
        'level UP
        If AktuellerSpieler.SpielerLevel < 5 And boolPlayerWon Then
            AktuellerSpieler.SpielerLevel = AktuellerSpieler.SpielerLevel + ONE
            setNewLevel = True
        End If
    Case Is <= 33
        'Level Down
        If AktuellerSpieler.SpielerLevel > 0 And Not boolPlayerWon Then
            AktuellerSpieler.SpielerLevel = AktuellerSpieler.SpielerLevel - ONE
            setNewLevel = True
        End If
End Select

If setNewLevel Then
    If WinPerc > 50 Then
        PlaySound KissSound
        AgentSpeak myText(10), True
    Else
        AgentSpeak myText(22), True
        PlaySound LevelDown
    End If
    SetSpielerInDB AktuellerSpieler, False
    AgentSpeak myText(9) & gstrSpace & strPlayerLevel(AktuellerSpieler.SpielerLevel), True
    
    Select Case AktuellerSpieler.SpielerLevel
        Case ZERO
            If Playermodus = singleplayer Then
                str = myText(11) & vbCr & myText(12)
                'If Not Schmulen Then frmMain.menSchmulen = True
            Else
                str = myText(13) & vbCr & myText(14)
            End If
        Case ONE
            If Playermodus = singleplayer Then
                str = myText(15) & vbCr & myText(16)
            Else
                'to do ?
            End If
        Case 2
            If Playermodus = singleplayer Then
                str = myText(17) & vbCr & myText(18)
            Else
                'to do ?
            End If
        Case 3
            str = myText(19)
        Case 4
            str = myText(20)
        Case 5
            str = myText(21)
    End Select
    
    AgentSpeak str, True
    'in hauptform aktualisieren
    If frmMain_Loaded Then frmMain.SetCaption
End If

ShowPlayerLevel

Exit Sub
ERRHand:
If ErrorBox("SetPlayerLevel", Err) Then Resume Next
End Sub

Private Sub ShowPlayerLevel()
'Anzeigen für Playerlevel aktualisieren
On Error Resume Next

    WriteLblLevel Me.lblLevel, AktuellerSpieler.SpielerLevel
    Me.picLevel.Picture = Me.ImageList1.ListImages(AktuellerSpieler.SpielerLevel + 1).Picture
    Me.lblName = AktuellerSpieler.SpielerName
    Me.Refresh
End Sub
