VERSION 5.00
Begin VB.Form frmTip 
   Caption         =   "Tips und Tricks"
   ClientHeight    =   3285
   ClientLeft      =   2370
   ClientTop       =   2400
   ClientWidth     =   5415
   Icon            =   "frmTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5415
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Tips beim Starten anzeigen"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   2970
      Width           =   2415
   End
   Begin VB.CommandButton cmdNextTip 
      Caption         =   "&Nächster Tip"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2715
      Left            =   120
      Picture         =   "frmTip.frx":08CA
      ScaleHeight     =   2655
      ScaleWidth      =   3675
      TabIndex        =   1
      Top             =   120
      Width           =   3735
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Wußten Sie schon.."
         Height          =   255
         Left            =   540
         TabIndex        =   5
         Top             =   180
         Width           =   2655
      End
      Begin VB.Label lblTipText 
         BackColor       =   &H00FFFFFF&
         Height          =   1635
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmTip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

' Tip-Datenbank im Speicher.
Dim Tips() As String

' Index in der Tip-Auflistung, die momentan angezeigt wird.
Dim CurrentTip As Long


Private Sub DoNextTip()

    ' Einen Tip willkürlich auswählen.
    CurrentTip = Int((UBound(Tips) * Rnd) + 1)
    
    ' Oder die Tips der Reihenfolge nach durchgehen.

'    CurrentTip = CurrentTip + 1
'    If Tips.count < CurrentTip Then
'        CurrentTip = 1
'    End If
    
    ' Tip anzeigen.
    frmTip.DisplayCurrentTip
    
End Sub



Private Sub chkLoadTipsAtStartup_Click()
    ' Speichern, ob dieses Formular beim Start angezeigt werden soll oder nicht
    ShowTipAtStartup = IIf(chkLoadTipsAtStartup.Value = 1, True, False)
    SaveSetting AppExeName, cstrOptions, cstrShowTips, ShowTipAtStartup
End Sub

Private Sub cmdNextTip_Click()
    DoNextTip
End Sub

Private Sub cmdOK_Click()
    Me.Hide
End Sub

Private Sub Form_Load()
On Error Resume Next
    chkLoadTipsAtStartup.Caption = ModText(12)
    Me.cmdNextTip.Caption = ModText(23)
    Me.Label1.Caption = ModText(24)
    ' Kontrollkästchen festlegen. Hierdurch wird der Wert in die Registrierung geschrieben
    Me.chkLoadTipsAtStartup.Value = vbChecked
    
    ' Randomisieren beginnen
    Randomize
    
    ' Tips einlesen und einen Tip willkürlich anzeigen.
    LoadObjectText "TIPOFDAY", Tips
    ' Tips willkürlich anzeigen.
    DoNextTip

    
End Sub

Public Sub DisplayCurrentTip()
On Error Resume Next
lblTipText.Caption = Tips(CurrentTip)

End Sub
