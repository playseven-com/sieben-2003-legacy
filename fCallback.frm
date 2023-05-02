VERSION 5.00
Begin VB.Form fCallback 
   BorderStyle     =   0  'Kein
   Caption         =   "Form1"
   ClientHeight    =   975
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1545
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   975
   ScaleWidth      =   1545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Visible         =   0   'False
End
Attribute VB_Name = "fCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements DirectXEvent8

Private Sub DirectXEvent8_DXCallback(ByVal eventid As Long)
On Error Resume Next
  Dim i As Integer

    'Find what sound we are being notified about
    For i = 1 To UBound(Sounds)
        If Sounds(i).Notification = eventid Then
            Exit For
        End If
    Next i

    Sounds(i).Playing = False

End Sub


