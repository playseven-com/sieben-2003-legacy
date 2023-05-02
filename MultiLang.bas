Attribute VB_Name = "MultiLang"
Option Explicit
Option Compare Text

Public strAvailableLangs() As String
Public LangID As Integer

Private Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public StandardLanguage As String
Private DBLang As dao.Database


Public Sub InitLangDb()


On Error GoTo ERRHand

StandardLanguage = GetSetting(AppExeName, cstrOptions, cstrLanguage)
If StandardLanguage = gstrNullstr Then
    LangID = GetUserDefaultLangID
    Select Case LangID
'        Case &H407, &H807, &HC07, &H1007, &H1407
'            StandardLanguage = "deutsch"
'        Case &H415
'            StandardLanguage = "polski"
        Case Else
            StandardLanguage = "english"
    End Select
Else
    Select Case StandardLanguage
        Case "deutsch"
            LangID = &H407
        Case "polski"
            LangID = &H415
        Case "english"
            LangID = &H409
    End Select
End If

Set DBLang = WS.OpenDatabase(App.path & cstrLangDBPath, False)
GetAvailableLangs
Exit Sub
ERRHand:
If ErrorBox("InitLangDb", Err) Then Resume Next
End Sub

Private Sub GetAvailableLangs()
Dim TABdef As dao.TableDef
Dim i As Integer

On Error Resume Next

Set TABdef = DBLang.TableDefs("frmMain")

For i = 0 To TABdef.Fields.count - 1
    ReDim Preserve strAvailableLangs(0 To i)
    strAvailableLangs(i) = TABdef.Fields(i).Name
Next

End Sub

Public Sub LoadObjectText(ObjectName As String, txt() As String)
Dim rs As dao.Recordset
Dim i As Integer

On Error GoTo ERRHand

Set rs = DBLang.OpenRecordset("Select count(*) as ANZAHL from " & ObjectName)
If rs!Anzahl > ZERO Then
    ReDim Preserve txt(0 To rs!Anzahl - 1)
Else
    MsgBox "No entries found in '" & DBLang.Name & "' for '" & ObjectName & "'" & vbCr & _
        "Apllication will abort !", vbCritical
    CloseAll
    End
End If

Set rs = DBLang.OpenRecordset(ObjectName)
While Not rs.EOF
    txt(rs!StringID) = vbNullString & rs(StandardLanguage)
    rs.MoveNext
    i = i + 1
Wend
rs.Close

Exit Sub
ERRHand:
    If ErrorBox("'MultiLang:LoadObjectText'" & ObjectName, Err) Then Resume Next
End Sub

Public Sub CloseLangDb()
If Not DBLang Is Nothing Then
    DBLang.Close
    Set DBLang = Nothing
End If
End Sub
