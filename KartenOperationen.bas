Attribute VB_Name = "KartenOperationen"
Option Explicit
Option Compare Text

Public Kartenspiel As Kartenspiel
Public gemischtesSpiel As Kartenspiel
Public NullKarte As Karte
Public lastCard As Karte

Public Geber As Integer
Public KartenAnzahl As Integer

Public RueckSeitePic As IPictureDisp
Public Type Karte
    Bild As KartenBilder
    Color As Farbe
    Caption As String
    BildName As String
    Position As Integer
    Pic As IPictureDisp
End Type

Type Kartenspiel
    Karte() As Karte
End Type

Public Enum Farbe
    Kreuz = 1
    Karo = 2
    Herz = 3
    Pik = 4
End Enum

Public Enum KartenBilder
    Zwei = 2
    Drei = 3
    Vier = 4
    Fünf = 5
    Sechs = 6
    sieben = 7
    Acht = 8
    Neun = 9
    Zehn = 10
    Bube = 11
    Dame = 12
    Koenig = 13
    Ass = 14
End Enum

Public HaufenSpieler() As Karte
Public HaufenComputer() As Karte

Public HandkartenSpieler(1 To 4) As Karte
Public HandkartenComputer(1 To 4) As Karte

Public Stich() As Karte
Public StichTrumpf As KartenBilder

Public StichAktion  As Integer
Public StichBesitzer As Integer
Public SiebenGefunden As Integer
Public GemischtesSpielAktPosition As Integer
Public GemischtesSpielEnde As Integer

'Public ZuEndeSpielen As Boolean
Public ErsterStichSpieler As Boolean
Public ErsterStichComputer As Boolean
Public Bluff As Boolean

Sub KartenSpiel_Init()
Dim Bild As Integer, Farbe As Integer, i As Integer
'initialisiert das kartenSpiel wie fabrikneu (sortiert)

Set RueckSeitePic = LoadPicture(App.path & cstrSubPathDeck & cstrCardBackName)

ReDim Preserve Kartenspiel.Karte(1 To KartenAnzahl)
ReDim Preserve gemischtesSpiel.Karte(1 To KartenAnzahl)

i = 1
For Bild = IIf(KartenAnzahl = 52, 2, 7) To 14
    For Farbe = 1 To 4
        Kartenspiel.Karte(i).Bild = Bild
        Select Case Bild
            Case 2
                Kartenspiel.Karte(i).Caption = ModText(29)
            Case 3
                Kartenspiel.Karte(i).Caption = ModText(30)
            Case 4
                Kartenspiel.Karte(i).Caption = ModText(31)
            Case 5
                Kartenspiel.Karte(i).Caption = ModText(32)
            Case 6
                Kartenspiel.Karte(i).Caption = ModText(33)
            Case 7
                Kartenspiel.Karte(i).Caption = ModText(34)
            Case 8
                Kartenspiel.Karte(i).Caption = ModText(35)
            Case 9
                Kartenspiel.Karte(i).Caption = ModText(36)
            Case 10
                Kartenspiel.Karte(i).Caption = ModText(37)
            Case 11
                Kartenspiel.Karte(i).Caption = ModText(38)
            Case 12
                Kartenspiel.Karte(i).Caption = ModText(39)
            Case 13
                Kartenspiel.Karte(i).Caption = ModText(40)
            Case 14
                Kartenspiel.Karte(i).Caption = ModText(41)
        End Select
        Kartenspiel.Karte(i).Color = Farbe
        Select Case Farbe
            Case 4
                Kartenspiel.Karte(i).Caption = Kartenspiel.Karte(i).Caption & gstrSpace & ModText(42)
            Case 3
                Kartenspiel.Karte(i).Caption = Kartenspiel.Karte(i).Caption & gstrSpace & ModText(43)
            Case 2
                Kartenspiel.Karte(i).Caption = Kartenspiel.Karte(i).Caption & gstrSpace & ModText(44)
            Case 1
                Kartenspiel.Karte(i).Caption = Kartenspiel.Karte(i).Caption & gstrSpace & ModText(42)
        End Select
        Kartenspiel.Karte(i).BildName = cstrSubPathDeck & i + IIf(KartenAnzahl = 52, 4, 24) & ".JPG"
        'Debug.Print i, Kartenspiel.Karte(i).Caption, Kartenspiel.Karte(i).Bild & gstrSpace & Kartenspiel.Karte(i).Color, Kartenspiel.Karte(i).BildPfad
        
        Set Kartenspiel.Karte(i).Pic = LoadPicture(App.path & Kartenspiel.Karte(i).BildName)

        i = i + 1
    Next
Next
End Sub

Function KartenMischen(Karten As Kartenspiel) As Kartenspiel

Dim Position() As Integer
Dim PositionBesetzt() As Boolean
Dim gemischteKarten As Kartenspiel
Dim retKartenspiel As Kartenspiel
Dim i As Integer
On Error GoTo ERRHand


ReDim Position(1 To KartenAnzahl)
ReDim PositionBesetzt(1 To KartenAnzahl)

ReDim retKartenspiel.Karte(1 To KartenAnzahl)
'If Not Test Then
Randomize Timer
'zufällig neue Positionen generieren
For i = 1 To KartenAnzahl
    Do
        Position(i) = Int(KartenAnzahl * Rnd + 1)
    Loop While PositionBesetzt(Position(i))
    PositionBesetzt(Position(i)) = True
Next

'neue KartenPositionen dem ungemischten Spiel zuordnen
For i = LBound(Karten.Karte) To UBound(Karten.Karte)
    retKartenspiel.Karte(i) = Karten.Karte(Position(i))
    'Debug.Print i, retKartenspiel.Karte(i).Caption
Next
KartenMischen = retKartenspiel

Exit Function
ERRHand:
If ErrorBox("KartenMischen", Err) Then Resume Next
End Function


Function KartenAbheben(Karten As Kartenspiel, ByVal pos As Integer, SiebenGefunden As Integer) As Kartenspiel
Dim i As Integer
Dim retKarten As Kartenspiel

On Error GoTo ERRHand

ReDim retKarten.Karte(1 To UBound(Karten.Karte))
For i = 1 To UBound(Karten.Karte)
    If i > pos Then
        retKarten.Karte(i - pos) = Karten.Karte(i)
    Else
        retKarten.Karte(UBound(Karten.Karte) - pos + i) = Karten.Karte(i)
    End If
'    Debug.Print i, retKarten.Karte(i).Caption
Next
'For i = 1 To UBound(Karten.Karte)
''    Debug.Print i, retKarten.Karte(i).Caption
'Next

'gefunden Sieben aus dem gemischten haufen entferenen
ReDim Preserve gemischtesSpiel.Karte(1 To KartenAnzahl - SiebenGefunden)

KartenAbheben = retKarten

Exit Function
ERRHand:
If ErrorBox("KartenAbheben", Err) Then Resume Next
End Function


Public Function GetKartenSpielFromString(sChat As String)
'erstellt ein Kartenspielobjekt aus dem übers Netzwerk übertragenem String
Dim TStr As String
Dim pos As Long, i As Long, ii As Long

On Error GoTo ERRHand

Debug.Print sChat
i = 0
ii = 0
pos = InStr(1, sChat, gstrDblDot, vbBinaryCompare)

Do While pos > 0
    i = i + 1
    TStr = Mid$(sChat, 1, pos - 1)
    
    sChat = Mid$(sChat, pos + 1)
    pos = InStr(1, sChat, gstrDblDot, vbTextCompare)
    Select Case i
        Case 1
            ii = ii + 1
            ReDim Preserve gemischtesSpiel.Karte(1 To ii)
            If ii <> TStr Then Stop
        Case 2
            gemischtesSpiel.Karte(ii).Bild = TStr
        Case 3
            gemischtesSpiel.Karte(ii).Color = TStr
        Case 4
            gemischtesSpiel.Karte(ii).BildName = TStr
        Case 5
            gemischtesSpiel.Karte(ii).Caption = TStr
            i = 0
    End Select
Loop
If ii < 32 Then
    MsgBox "Zu wenige Karten erhalten.", vbCritical, "Fehler"
    Write2Log "Zu wenige Karten erhalten: " & sChat
End If

Exit Function
ERRHand:
If ErrorBox("GetKartenSpielFromString: " & sChat, Err) Then Resume Next

End Function



