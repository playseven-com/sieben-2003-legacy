Attribute VB_Name = "modLog"
Option Explicit
Option Compare Text

Public Sub Write2Log(str As String)
Dim f As Long
Dim path As String
Static NoLogWhileError As Boolean


On Error GoTo ERRHand
If NoLogWhileError Then Exit Sub

path = App.path & "\LOGS"
If Not DirExists(path) Then
    MkDir path
End If

f = FreeFile
Open path & gstrDirSep & Replace(Date, gstrDot, vbNullString) & ".txt" For Append As #f
Print #f, Now & " - " & str
Close #f

Exit Sub
ERRHand:
NoLogWhileError = True
'If ErrorBox("Write2log", Err) Then Resume Next
End Sub

