VERSION 5.00
Begin VB.Form frmLic 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Lizenzeingabe"
   ClientHeight    =   4110
   ClientLeft      =   30
   ClientTop       =   270
   ClientWidth     =   4830
   Icon            =   "frmLic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Begin VB.CommandButton cmdBUY 
      Caption         =   "buy online"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2880
      Style           =   1  'Grafisch
      TabIndex        =   12
      Top             =   2130
      Width           =   645
   End
   Begin VB.CommandButton cmdClipBKto 
      Caption         =   "Kontodaten in Zwischenablage"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   3540
      Style           =   1  'Grafisch
      TabIndex        =   11
      Top             =   2130
      Width           =   1305
   End
   Begin VB.CommandButton cmdSendRegister 
      Caption         =   "Lizenz anfordern"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1950
      Style           =   1  'Grafisch
      TabIndex        =   9
      Top             =   2130
      Width           =   945
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   288
      Left            =   1470
      TabIndex        =   8
      Top             =   3690
      Width           =   1800
   End
   Begin VB.Frame fraLic 
      BackColor       =   &H00000000&
      Height          =   615
      Left            =   690
      TabIndex        =   3
      Top             =   2970
      Width           =   3450
      Begin VB.TextBox txtLic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   2580
         MaxLength       =   4
         TabIndex        =   7
         Top             =   192
         Width           =   800
      End
      Begin VB.TextBox txtLic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   1755
         MaxLength       =   4
         TabIndex        =   6
         Top             =   192
         Width           =   800
      End
      Begin VB.TextBox txtLic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   900
         MaxLength       =   4
         TabIndex        =   5
         Top             =   192
         Width           =   800
      End
      Begin VB.TextBox txtLic 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   96
         MaxLength       =   4
         TabIndex        =   4
         Top             =   192
         Width           =   800
      End
   End
   Begin VB.Label labInfo 
      Alignment       =   1  'Rechts
      BackStyle       =   0  'Transparent
      Caption         =   "ProgID:"
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
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   10
      Top             =   2250
      Width           =   660
   End
   Begin VB.Label labInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Sie erhalten dann den Lizenz-Code, der in diesen vier Feldern eingetragen werden kann."
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
      Height          =   435
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   2580
      Width           =   4650
   End
   Begin VB.Label labInfo 
      BackStyle       =   0  'Transparent
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
      Height          =   1995
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   150
      Width           =   4650
   End
   Begin VB.Label labID 
      Alignment       =   2  'Zentriert
      BorderStyle     =   1  'Fest Einfach
      Height          =   285
      Left            =   870
      TabIndex        =   1
      Top             =   2190
      Width           =   1050
   End
End
Attribute VB_Name = "frmLic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public ChangesOk As Boolean
Public LicenseCode As String

Private Sub cmdBUY_Click()
    Go2URL "https://orders.digitalcandle.com/order_php/order.php?productid=000972"
End Sub

Private Sub cmdClipBKto_Click()
    Clipboard.SetText "ProgID: " & Me.labID & vbCrLf & KontoVerbindung
End Sub

Private Sub cmdSendRegister_Click()
Dim body As String
body = ModText(18) & mailCrLf & _
    ModText(19) & gstrSpace & App.ProductName & ModText(20) & Me.labID & mailCrLf & ModText(21)
    
    SendMail "mw@playseven.com", ModText(22) & AppInfo, body
    
End Sub

Private Sub Form_Load()
  Dim s As String
  '
  '---- ID zeigen
  labID = MyLic.ID
  '---- Lizenz zeigen
  s = MyLic.UserLicenseString
  txtLic(0) = Mid(s, 1, 4)
  txtLic(1) = Mid(s, 6, 4)
  txtLic(2) = Mid(s, 11, 4)
  txtLic(3) = Mid(s, 16, 4)
  SetBackGround Me
  Me.labInfo(0).Caption = ModText(17) & vbCr & vbCr & KontoVerbindung
  Me.labInfo(1).Caption = ModText(25)
  Me.cmdSendRegister.Caption = ModText(26)
  Me.cmdClipBKto.Caption = ModText(27)
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub


Private Sub txtLic_GotFocus(Index As Integer)
'Text markieren
  With txtLic(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub btnOK_Click()
  LicenseCode = UCase(Right("0000" & txtLic(0).Text, 4) & _
                      Right("0000" & txtLic(1).Text, 4) & _
                      Right("0000" & txtLic(2).Text, 4) & _
                      Right("0000" & txtLic(3).Text, 4))
  SaveSetting AppExeName, cstrDefault, cstrRegID, LicenseCode
  
  Write2Reg KEY_LICCODE, Encrypt(LicenseCode, True)
  Unload Me
End Sub


Private Sub txtLic_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(txtLic(Index)) >= 4 Then
    If Index = 3 Then
        Me.btnOK.SetFocus
    Else
        txtLic(Index + 1).SetFocus
    End If
End If
End Sub
