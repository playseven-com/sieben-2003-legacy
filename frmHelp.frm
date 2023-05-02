VERSION 5.00
Begin VB.Form frmHelp 
   BorderStyle     =   4  'Festes Werkzeugfenster
   Caption         =   "Regeln"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox Picture2 
      Height          =   1155
      Left            =   1260
      ScaleHeight     =   1095
      ScaleWidth      =   885
      TabIndex        =   1
      Top             =   1500
      Width           =   945
   End
   Begin VB.PictureBox Picture1 
      Height          =   1125
      Left            =   180
      ScaleHeight     =   1065
      ScaleWidth      =   825
      TabIndex        =   0
      Top             =   1500
      Width           =   885
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    SetBackGround Me
End Sub
