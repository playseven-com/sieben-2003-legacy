VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmUserAnswer 
   BorderStyle     =   3  'Fester Dialog
   ClientHeight    =   1980
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   4605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmUserAnswer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmdYes 
      Caption         =   "&Ja"
      Default         =   -1  'True
      Height          =   345
      Left            =   420
      TabIndex        =   4
      Top             =   1530
      Width           =   1365
   End
   Begin VB.Timer timServAntw 
      Interval        =   434
      Left            =   4140
      Top             =   1530
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "&Nein"
      Height          =   345
      Left            =   2670
      TabIndex        =   1
      Top             =   1530
      Width           =   1365
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   225
      Left            =   90
      TabIndex        =   0
      Top             =   1170
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   397
      _Version        =   393216
      Appearance      =   1
      Max             =   60
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Der Spieler möchte mit Ihnen einen Spielraum betreten. Stimmen Sie zu ?"
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
      Height          =   585
      Left            =   90
      TabIndex        =   3
      Top             =   600
      Width           =   4425
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00800000&
      BorderWidth     =   3
      Height          =   1995
      Left            =   0
      Top             =   0
      Width           =   4605
   End
   Begin VB.Label lblUserInfo 
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   4455
   End
End
Attribute VB_Name = "frmUserAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Public Antwort As Long

Private Sub cmdNo_Click()
    Me.Antwort = vbNo
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Me.Antwort = vbYes
    Unload Me
End Sub

Private Sub Form_Load()
    SetBackGround Me
    SetzFensterMittig Me
    Me.timServAntw.Enabled = True
    Me.lblUserInfo = Gegner.SpielerName
    Me.Show 1, ServerEventsForm
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub


Private Sub timServAntw_Timer()
    If Me.ProgressBar1.Value < MaxTime Then
        Me.ProgressBar1.Value = Me.ProgressBar1.Value + 1
    ElseIf Me.ProgressBar1.Value = MaxTime Then
        Me.Antwort = vbCancel
        Unload Me
    End If
'DoEvents

End Sub
