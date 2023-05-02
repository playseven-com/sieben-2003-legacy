Attribute VB_Name = "DBStats"
Option Explicit
Option Compare Text

Public WS As dao.Workspace
Private DB As dao.Database
Private Const ps = "5C435A252C7367707C727C7B6A5D73404A7A"
'Private Const cstrTabbFILst = "FIList"
Private Const cstrTabbRegInfo = "RegInfo"

Public Enum SpielModus
    Spiel = 0
    Runde = 1
End Enum

Public Enum TableList
    FriendS
    Ignore
End Enum

Public Enum Operation
    adder
    Deleter
End Enum

Public Type PunkteInfo
    GesamtSpiele As Long
    GewonneneSpiele As Long
End Type

Public Enum FriendIgnore
    Undefined
    FriendS
    Ignore
End Enum

Public Sub DBInit()
On Error GoTo ERRHand
Dim pass As String
' Microsoft Jet Workspace-Objekt erstellen.
Set WS = CreateWorkspace("", "Admin", "", dbUseJet)

' Database-Objekt aus gespeicherter Microsoft Jet-Datenbank
' nicht exklusiv öffnen 4 Test
pass = Encrypt(ps, False)
Set DB = WS.OpenDatabase(App.path & "\Source\stat.mdb", Not Test, False, "" & pass)
'debug.Print DB.Name

On Error Resume Next
DB.Execute "ALTER TABLE Spiele ADD COLUMN won YesNo;"
           
Exit Sub
ERRHand:
If ErrorBox("DBInit", Err) Then Resume Next
End Sub

Public Sub DBClose()

If Not DB Is Nothing Then
    DB.Close
    WS.Close
    Set DB = Nothing
    Set WS = Nothing
End If

End Sub


Public Sub Write2DB(mode As SpielModus, PunkteSpieler As Integer, PunkteComputer As Integer, ZeitStempel As Date, Optional SpielerWon)

Dim rs As dao.Recordset
Dim tabb As String

On Error GoTo ERRHand

If Playermodus = multiplayer And mode = Runde Then
    If AktuellerSpieler.SpielOption = Liga Then
        tabb = "MultiPlayer"
        Set rs = DB.OpenRecordset(tabb, dbOpenTable)
        rs.AddNew
        
        rs!Spieler1ID = AktuellerSpieler.GlobalID
        rs!Spieler2ID = Gegner.GlobalID
        
        rs!SpielerName2 = Gegner.SpielerName
        
        rs!PunkteSpieler1 = PunkteSpieler
        rs!PunkteSpieler2 = PunkteComputer
        
        rs!LevelSpieler1 = AktuellerSpieler.SpielerLevel
        rs!LevelSpieler2 = Gegner.SpielerLevel
        
        rs!Datum = ZeitStempel
        
        rs.Update
        rs.Close

    End If
End If

Select Case mode
    Case Spiel
        tabb = "Spiele"
    Case Runde
        tabb = "Runden"
End Select

Set rs = DB.OpenRecordset(tabb, dbOpenTable)
rs.AddNew

rs!Spieler = AktuellerSpieler.SpielerName
rs!Datum = ZeitStempel
rs!PunkteSpieler = PunkteSpieler
rs!PunkteComputer = PunkteComputer
rs!Level = AktuellerSpieler.SpielerLevel
If Not IsMissing(SpielerWon) Then rs!Won = SpielerWon

rs.Update
rs.Close
Set rs = Nothing

Exit Sub
ERRHand:
If ErrorBox("Write2DB", Err) Then Resume Next

End Sub

Function ReadFromDB(mode As SpielModus) As PunkteInfo
Dim rs As dao.Recordset
Dim tabb As String
Dim SQLMsg As String

On Error GoTo ERRHand

Select Case mode
    Case Spiel
        tabb = "Spiele"
    Case Runde
        tabb = "Runden"
    Case multiplayer
        'mussen wir noch machen
End Select

SQLMsg = "SELECT count(*) as Anzahl from " & tabb & " WHERE Spieler like '" & AktuellerSpieler.SpielerName & "'"

Set rs = DB.OpenRecordset(SQLMsg)
ReadFromDB.GesamtSpiele = rs!Anzahl
rs.Close

If mode = Spiel Then
    SQLMsg = SQLMsg & " AND (PunkteSpieler > PunkteComputer OR WON = true)"
Else
    SQLMsg = SQLMsg & " AND PunkteSpieler > PunkteComputer"
End If
Set rs = DB.OpenRecordset(SQLMsg)
ReadFromDB.GewonneneSpiele = rs!Anzahl
rs.Close
Set rs = Nothing

Exit Function
ERRHand:
If ErrorBox("ReadFromDB", Err) Then Resume Next

End Function


Public Function GetSpielerFromDB(Spieler() As SpielerInfo) As Boolean
Dim rs As dao.Recordset
Dim SQLMsg As String
Dim i As Integer

On Error GoTo ERRHand

SQLMsg = "Select * from Spieler"
Set rs = DB.OpenRecordset(SQLMsg)
i = 0
While Not rs.EOF
    i = i + 1
    ReDim Preserve Spieler(1 To i)
    Spieler(i).GlobalID = gstrNullstr & rs!GlobalID
    Spieler(i).SpielerName = rs!SpielerName
    Spieler(i).SpielerLevel = rs!SpielerLevel
    If Spieler(i).SpielerLevel < 0 Then Spieler(i).SpielerLevel = 0
    If Spieler(i).SpielerLevel > 6 Then Spieler(i).SpielerLevel = 6
    Spieler(i).Points = val(gstrNullstr & rs!punkte)
    Spieler(i).AvatarFileName = vbNullString & rs!AvatarPath
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
If i > ZERO Then GetSpielerFromDB = True

Exit Function
ERRHand:
If ErrorBox("GetSpielerFromDB", Err) Then Resume Next

End Function

Public Function SetSpielerInDB(Spieler As SpielerInfo, Neu As Boolean) As Boolean
Dim rs As dao.Recordset
Dim SQLMsg As String

On Error GoTo ERRHand

If Spieler.SpielerName = gstrNullstr Then
    SetSpielerInDB = False
    Exit Function
End If

If DB Is Nothing Then
    MsgBox "Database is not initialized!" & vbCr & _
        "Please restart the game and report this problem to betatester@playseven.com", vbCritical, "Error"
    SetSpielerInDB = False
    Exit Function
End If

SQLMsg = "Select * from Spieler where " & IIf(Spieler.GlobalID <> vbNullString, "GlobalID like '" & Spieler.GlobalID & "'", "Spielername like '" & Spieler.SpielerName & "'")
Set rs = DB.OpenRecordset(SQLMsg)

If Neu And Not rs.EOF Then
    MsgBox ModText(14) & vbCr & ModText(15)
    SetSpielerInDB = False
    GoTo RAUS
End If

If rs.EOF And Neu Then
    rs.AddNew
Else
    rs.Edit
End If

rs!GlobalID = IIf(Spieler.GlobalID = gstrNullstr, GetGUID, Spieler.GlobalID)
rs!SpielerName = Spieler.SpielerName
rs!SpielerLevel = Spieler.SpielerLevel
rs!punkte = Spieler.Points
rs!AvatarPath = vbNullString & Spieler.AvatarFileName
rs.Update
SetSpielerInDB = True

RAUS:
rs.Close
Set rs = Nothing

Exit Function
ERRHand:
If ErrorBox("SetSpielerInDB", Err) Then Resume Next

End Function

Public Function AlterList(Guid As String, Name As String, Table As TableList, op As Operation) As Boolean
Dim rs As dao.Recordset
Dim tabb As String

On Error Resume Next

If Guid = vbNullString Then Exit Function

Set rs = DB.OpenRecordset("Select * from FIList where SpielerGuid like '" & Guid & "'")

If op = adder Then

    If rs.EOF Then
        rs.AddNew
    Else
        rs.Edit
    End If
    rs!SpielerGuid = Guid
    rs!SpielerName = Name
    rs!FriendEnemy = IIf(Table = TableList.FriendS, True, False)
    rs.Update
        
Else
    If Not rs.EOF Then rs.Delete
End If


If Err.Number = 3022 Then
    AgentSpeak "Spieler schon in Liste vorhanden", True
ElseIf Err.Number = 0 Then
    AlterList = True
Else
    If ErrorBox("DBStats:AlterList", Err) Then Resume Next
End If

rs.Close
Set rs = Nothing
End Function




Public Function IsInList(Guid As String) As FriendIgnore
Dim rs As dao.Recordset
Dim tabb As String
On Error GoTo ERRHand

Set rs = DB.OpenRecordset("Select * from FIList where SpielerGuid like '" & Guid & "'")
If rs.EOF Then
    IsInList = Undefined
Else
    If rs!FriendEnemy Then
        IsInList = FriendIgnore.FriendS
    Else
        IsInList = FriendIgnore.Ignore
    End If
End If

rs.Close
Set rs = Nothing
Exit Function
ERRHand:
If ErrorBox("DBStats:IsInList", Err) Then Resume Next

End Function

Public Function GetMPGewinnQuote(Player As SpielerInfo) As Single
Dim rs As dao.Recordset
Dim count As Long
Dim SQLMsg As String

SQLMsg = "SELECT Count(Multiplayer.Spieler1ID) AS [Anzahl] From Multiplayer WHERE (((Multiplayer.Spieler1ID) like '" & AktuellerSpieler.GlobalID & "') AND ((Multiplayer.Spieler2ID) like '" & Player.GlobalID & "'))"
Set rs = DB.OpenRecordset(SQLMsg)
count = rs!Anzahl
If count > 0 Then
    SQLMsg = SQLMsg & " AND (Multiplayer.PunkteSpieler1 > Multiplayer.PunkteSpieler2)"
    Set rs = DB.OpenRecordset(SQLMsg)
Else
    GetMPGewinnQuote = -1
    Exit Function
End If

GetMPGewinnQuote = rs!Anzahl / count

rs.Close
Set rs = Nothing

End Function

Public Function Write2Reg(Key As String, KeyValue As String)
Dim rs As dao.Recordset
On Error Resume Next

Set rs = DB.OpenRecordset("Select * from " & cstrTabbRegInfo & " WHERE key like '" & Key & "'")
If rs.EOF Then
    rs.AddNew
Else
    rs.Edit
End If
rs!Key = Key
rs!KeyString = KeyValue
rs.Update
rs.Close
Set rs = Nothing

End Function

Public Function GetFromReg(Key As String) As String
Dim rs As dao.Recordset
On Error Resume Next

Set rs = DB.OpenRecordset("Select KeyString from " & cstrTabbRegInfo & " WHERE key like '" & Key & "'")
If Not rs.EOF Then
    GetFromReg = vbNullString & rs!KeyString
End If
rs.Close
Set rs = Nothing

End Function

