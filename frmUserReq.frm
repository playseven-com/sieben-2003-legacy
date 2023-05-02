VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUserReq 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   2055
   ClientLeft      =   7155
   ClientTop       =   2265
   ClientWidth     =   4680
   Icon            =   "frmUserReq.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   660
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Max             =   60
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   2730
      TabIndex        =   3
      Top             =   1590
      Width           =   1365
   End
   Begin VB.Timer timServAntw 
      Interval        =   434
      Left            =   4170
      Top             =   1560
   End
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1350
      TabIndex        =   2
      Top             =   90
      Width           =   3195
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "connecting to:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   3
      Height          =   1995
      Left            =   30
      Top             =   30
      Width           =   4605
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   120
      TabIndex        =   0
      Top             =   870
      Width           =   4425
   End
End
Attribute VB_Name = "frmUserReq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public GlobalID As String

Private Sub cmdCancel_Click()
    ServerEventsForm.SendMsg2Server msg_gamestart_cancel, GlobalID
    Unload Me
End Sub

Private Sub Form_Load()
    SetBackGround Me
    SetzFensterMittig Me
    Me.timServAntw.Enabled = True
    Me.lblMsg = "User wird kontaktiert"
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub timServAntw_Timer()
Static tim As Integer
    If Me.ProgressBar1.Value < MaxTime Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
        tim = 0
    ElseIf tim >= 1 Then
       tim = tim + 1
       If tim >= 4 Then
           Me.timServAntw.Enabled = False
           Unload Me
       End If
    ElseIf Me.ProgressBar1.Value = MaxTime Then
        Me.lblMsg = "User hat nicht geantwortet"
        ServerEventsForm.SendMsg2Server msg_gamestart_cancel, GlobalID
        tim = 1
    End If
'DoEvents

End Sub
