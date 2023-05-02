Attribute VB_Name = "AI"
Option Explicit

Private myText() As String

Type StechInfo
    Karte As Karte
    GesamtStiche As Integer
    AnzahlSieben As Integer
End Type

Type LuschenInfo
    Karte() As Karte
    GesamtLuschen As Integer
End Type

Public KI_Erklaerung As String

Public Function KI_SpielKarte() As Integer
On Error GoTo ERRHand
KI_Erklaerung = vbNullString
If StichAktion = 0 Then
'spieler kommt
    KI_SpielKarte = SuchGuteKarte(HandkartenSpieler)
    If KI_SpielKarte = ZERO Then
        KI_SpielKarte = ComputerLegtSicher(HandkartenSpieler, StichTrumpf).Position
    End If
ElseIf StichAktion Mod 2 = 1 Then
    'Spieler in der Hinterhand
    KI_SpielKarte = ComputerAntwort(HandkartenSpieler, 5, CInt(frmMain.lblComputerPunkte(ZERO).Caption), CInt(frmMain.lblSpielerPunkte(ZERO).Caption), ErsterStichSpieler)
Else
    'Spieler ist wieder dran und verlängert oder beendet Stich
    KI_SpielKarte = ComputerVerlaengert(HandkartenSpieler, Spieler, CInt(frmMain.lblSpielerPunkte(ZERO).Caption), CInt(frmMain.lblComputerPunkte(ZERO).Caption), ErsterStichComputer)
End If
Exit Function
ERRHand:
If ErrorBox("KI_SpielKarte", Err) Then Resume Next

End Function

Public Sub Init_KI()
    LoadObjectText "AI", myText
End Sub


Public Sub ComputerAktion()
Dim pos As Integer
Dim i As Integer

Select Case AktuellerSpieler.SpielerLevel
    Case 0
        'Anfänger Level
        'Keine Angriffslogik
        frmMain.picKarteComp_Wirf SuchZufallsKarte(HandkartenComputer).Position - 1
    Case 1
        'Normal Level
        'Defensiv Spiel
        frmMain.picKarteComp_Wirf ComputerLegtSicher(HandkartenComputer, StichTrumpf).Position - 1
    Case Else
        'Fortgeschrittener Level
        'Aggresives Spiel
        pos = SuchGuteKarte(HandkartenComputer)
        If pos > ZERO Then
            frmMain.picKarteComp_Wirf pos - 1
        Else
            frmMain.picKarteComp_Wirf ComputerLegtSicher(HandkartenComputer, StichTrumpf).Position - 1
        End If
End Select
If Test Then frmMain.ComputerSpieltFuerSpieler
End Sub

Private Function stichMoeglich(Karten() As Karte, Trumpf As KartenBilder) As StechInfo
Dim ret As StechInfo
Dim i As Integer
Dim found As Boolean
i = 0
For i = LBound(Karten) To UBound(Karten)
    If Karten(i).Bild = Trumpf Then
        ret.Karte.Position = i
        ret.Karte.Bild = Trumpf
        found = True
        ret.GesamtStiche = ret.GesamtStiche + 1
    End If
Next

'wenn kein anderer trumpf zu finden war nach 7 suchen
For i = LBound(Karten) To UBound(Karten)
    If Karten(i).Bild = 7 Then
        If Not found Then
            ret.Karte.Position = i
            ret.Karte.Bild = 7
        End If
        ret.AnzahlSieben = ret.AnzahlSieben + 1
        ret.GesamtStiche = ret.GesamtStiche + 1
    End If
Next

stichMoeglich = ret

End Function

Private Function ZeigLuschen(Karten() As Karte, Trumpf As KartenBilder) As LuschenInfo
Dim X As Integer
Dim ret As LuschenInfo
Dim i As Integer

For i = LBound(Karten) To UBound(Karten)
    If Karten(i).Bild <> Ass And Karten(i).Bild <> 10 And Karten(i).Bild <> 7 And _
        Karten(i).Bild > ZERO And Karten(i).Bild <> Trumpf Then
        
        ret.GesamtLuschen = ret.GesamtLuschen + 1
        X = X + 1
        ReDim Preserve ret.Karte(1 To X)
        ret.Karte(X).Bild = Karten(i).Bild
        ret.Karte(X).Position = i
    End If
Next
If ret.GesamtLuschen > 1 Then ret.Karte = SortiereKarten(ret.Karte)
ZeigLuschen = ret

End Function
Public Function ComputerVerlaengert(handkarten() As Karte, SpielerAmZug As Integer, SpielerPunkte As Integer, GegnerPunkte As Integer, GegnerErsterStich As Boolean) As Integer

Dim StechKarte As StechInfo
Dim StichPunkte As Integer

StichPunkte = GetPunkteImStich(Stich)
StechKarte = stichMoeglich(handkarten, StichTrumpf)

'wenn punkte drin sind und ist eh unser
If StichPunkte >= 1 And StichBesitzer = SpielerAmZug Then
    'Stich nehmen
    ComputerVerlaengert = -1
    KI_Erklaerung = IIf(StichPunkte = 1, myText(0), myText(1) & gstrSpace & StichPunkte & myText(2))
    
'wenn wir 4 punkte haben und eine Sieben in der Hinterhand behalten können -> Stechen, auf 16er spielen
ElseIf SpielerPunkte = 4 And StechKarte.GesamtStiche > 1 And StechKarte.AnzahlSieben > 0 And GegnerErsterStich Then
    ComputerVerlaengert = StechKarte.Karte.Position
    KI_Erklaerung = KI_Erklaerung & vbCr & myText(3)

'wenn wir mehr als 2 punkte haben und drei Sieben in der Hinterhand -> Stechen, auf 16er spielen
ElseIf SpielerPunkte >= 3 And StechKarte.AnzahlSieben > 2 And GegnerErsterStich And StichBesitzer <> SpielerAmZug Then
    ComputerVerlaengert = StechKarte.Karte.Position
    KI_Erklaerung = KI_Erklaerung & vbCr & myText(3)

'wenn wir schon gewonenn haben -> auf 16er spielen
ElseIf SpielerPunkte > 4 And StechKarte.GesamtStiche > 0 And StichBesitzer = -SpielerAmZug Then
    ComputerVerlaengert = StechKarte.Karte.Position
    KI_Erklaerung = KI_Erklaerung & vbCr & myText(3)

'abgeben wenn wir auf 4 und letzten spielen können
ElseIf SpielerPunkte = 4 And StechKarte.AnzahlSieben = 1 And StichBesitzer = -SpielerAmZug Then
    'stich abgeben
    ComputerVerlaengert = -1
    KI_Erklaerung = myText(4) & vbCr & myText(5)
    
'abgeben wenn wir auf 4 und letzten spielen können
ElseIf frmStatistik.txtKartenImStapel = ZERO And StichPunkte + GegnerPunkte <= 4 And StechKarte.AnzahlSieben = 1 Then
    ComputerVerlaengert = -1
    KI_Erklaerung = myText(5)

Else
    'wenn wir stechen können
    If StechKarte.Karte.Position > ZERO Then
        'wenn verlängerungskarte keine 7 ist
        If StechKarte.Karte.Bild <> 7 Then
            'stechen
            ComputerVerlaengert = StechKarte.Karte.Position
            KI_Erklaerung = myText(6)
        Else
            'wenn Punkte im Stich sind
            If StichPunkte >= 1 Then
                ComputerVerlaengert = StechKarte.Karte.Position
                KI_Erklaerung = myText(7)
            Else
                ComputerVerlaengert = -1
                KI_Erklaerung = myText(8)
            End If
        End If
    Else
        'wenn wir nicht stechen können --> Stich beenden
        ComputerVerlaengert = -1
        KI_Erklaerung = myText(9)
    End If
End If
End Function

Private Function GetPunkteImStich(Stich() As Karte) As Integer
Dim i As Integer
For i = LBound(Stich) To UBound(Stich)
    If Stich(i).Bild = Ass Or Stich(i).Bild = Zehn Then
        GetPunkteImStich = GetPunkteImStich + 1
    End If
Next

End Function

Private Function SuchGuteKarte(Karten() As Karte) As Integer
Dim i As Integer
Dim Paare(1 To 14) As Karte
Dim Anzahl As Integer

'Paare suchen
For i = LBound(Karten) To UBound(Karten)
    If Karten(i).Bild <> ZERO Then
        Paare(Karten(i).Bild).Bild = Paare(Karten(i).Bild).Bild + 1
        Paare(Karten(i).Bild).Position = i
    End If
Next

'günstigstes Paar ermitteln
'das von dem am meisten gefallen ist
For i = 1 To 14
    If Paare(i).Bild > 1 And i <> 7 Then
        If Anzahl < Paare(i).Bild + frmStatistik.txtKarteGef(i) Then
            SuchGuteKarte = Paare(i).Position
            Anzahl = Paare(i).Bild + frmStatistik.txtKarteGef(i)
            KI_Erklaerung = myText(10) & gstrSpace & Paare(i).Bild & myText(11) & _
                IIf(frmStatistik.txtKarteGef(i) > ZERO, myText(12) & gstrSpace & frmStatistik.txtKarteGef(i) & myText(13), vbNullString)
        End If
    End If
Next

'wenn wir >1 Siebenen haben dann mit Dicken versuchen zu kommen
If SuchGuteKarte = ZERO And Paare(sieben).Bild > 1 Then
    If Paare(Zehn).Bild > Paare(Ass).Bild Then
        SuchGuteKarte = Paare(Zehn).Position
    Else
        SuchGuteKarte = Paare(Ass).Position
    End If
    KI_Erklaerung = myText(14) & gstrSpace & Paare(sieben).Bild & myText(15)
End If

'Letzte Karten schmeissen, wenn keine Siebenen drin sind
If SuchGuteKarte = ZERO And (frmStatistik.txtKarteGef(7) + Paare(sieben).Bild = 4) Then
    For i = 1 To 13
        If Paare(i).Bild > ZERO And i <> sieben Then
            If Anzahl < Paare(i).Bild + frmStatistik.txtKarteGef(i) Then
                Anzahl = Paare(i).Bild + frmStatistik.txtKarteGef(i)
                If Anzahl = 4 Then
                    SuchGuteKarte = Paare(i).Position
                    KI_Erklaerung = myText(16) & gstrSpace & frmStatistik.txtKarteGef(i) & myText(17)
                End If
            End If
        End If
    Next
End If


'wenn kein paar gefunden günstigste Angriffskarte (die am meisten gefallen ist) ohne Punkte
If SuchGuteKarte = ZERO Then
    For i = 1 To 13
        If Paare(i).Bild > ZERO And i <> sieben And i <> 10 And i <> Ass Then
            If Anzahl < Paare(i).Bild + frmStatistik.txtKarteGef(i) Then
                Anzahl = Paare(i).Bild + frmStatistik.txtKarteGef(i)
                If Anzahl > 1 Then
                    SuchGuteKarte = Paare(i).Position
                    KI_Erklaerung = myText(16) & gstrSpace & frmStatistik.txtKarteGef(i) & myText(17)
                End If
            End If
        End If
    Next
End If

'wenn kein paar gefunden und sonst nix schauen ob wir eine Karte haben wie die, die unten liegt.
If SuchGuteKarte = ZERO And frmStatistik.txtGelegteKarten < 8 Then
    For i = LBound(Karten) To UBound(Karten)
        If Karten(i).Bild = lastCard.Bild Then
            SuchGuteKarte = i
            KI_Erklaerung = myText(18) & gstrSpace & lastCard.Caption & myText(19)
            Exit For
        End If
    Next
End If

'wenn keine gute Karte gefunden günstigste Angriffskarte inkl. Punkte
If SuchGuteKarte = ZERO Then
    For i = 1 To 13
        If Paare(i).Bild > ZERO And i <> sieben Then
            If Anzahl < Paare(i).Bild + frmStatistik.txtKarteGef(i) Then
                Anzahl = Paare(i).Bild + frmStatistik.txtKarteGef(i)
                If Anzahl > 1 Then
                    SuchGuteKarte = Paare(i).Position
                    KI_Erklaerung = myText(16) & gstrSpace & frmStatistik.txtKarteGef(i) & myText(17)
                End If
            End If
        End If
    Next
End If

End Function

Private Function SortiereKarten(Karten() As Karte) As Karte()
'sortiert Karten nach Häufigkeit wie sie gefallen sind
Dim temp As Karte
Dim i As Integer, j As Integer
For i = UBound(Karten) To 1 Step -1
    For j = 1 To i - 1
        If frmStatistik.txtKarteGef(Karten(j + 1).Bild) > frmStatistik.txtKarteGef(Karten(j).Bild) Then
            temp = Karten(j)
            Karten(j) = Karten(j + 1)
            Karten(j + 1) = temp
        End If
    Next
Next
SortiereKarten = Karten
End Function


Private Function PaarBluff(PGegner As Integer, PMeine As Integer) As Boolean
Dim erg As Integer, t As Integer
On Error Resume Next

'wenn Spieler führt Wahrscheinlichkeit für Bluff steigern, da ein kartenlauf beim Spieler angenommen wird
t = (PGegner - PMeine) + 1

erg = Int(t * (Rnd))

If erg > 1 Then
    PaarBluff = True
Else
    PaarBluff = False
End If
'Debug.Print "Bluffindex:" & erg & " Bluff = " & PaarBluff

End Function

Private Function ComputerLegtSicher(handkarten() As Karte, StichTrumpf As KartenBilder) As Karte
Dim LuscheSuche As LuschenInfo
Dim StechKarte As StechInfo
Dim pos As Integer
Dim TKarten() As Karte
Dim i As Integer


'beste Lusche suchen
LuscheSuche = ZeigLuschen(handkarten, StichTrumpf)

'Versuchen eine Lusche abzuschmeißen
If LuscheSuche.GesamtLuschen > ZERO Then
    ComputerLegtSicher = LuscheSuche.Karte(ONE)
    If frmStatistik.txtKarteGef(LuscheSuche.Karte(ONE).Bild) > ZERO Then KI_Erklaerung = myText(20)
Else
    'wenn keine Lusche abgeschmissen werden kann,
    pos = 0
    For i = UBound(handkarten) To LBound(handkarten) Step -1
        If handkarten(i).Bild > ZERO And handkarten(i).Bild <> sieben Then
            pos = pos + ONE
            ReDim Preserve TKarten(1 To pos)
            TKarten(pos) = handkarten(i)
            TKarten(pos).Position = i
        End If
    Next
    If pos > 0 Then
        'Handkarten danach sortieren wie häufig sie gefallen sind
        TKarten = SortiereKarten(TKarten)
        'häufigste Karte zurückgeben
        ComputerLegtSicher = TKarten(1)
    
    'If TKarten(1).Position > 0 Then
        KI_Erklaerung = myText(21)
    Else
        Debug.Print "Und es passiert doch"
        Write2Log "Und es passiert doch"
        StechKarte = stichMoeglich(handkarten, StichTrumpf)
        If StechKarte.Karte.Bild > ZERO Then
            ComputerLegtSicher = StechKarte.Karte
            KI_Erklaerung = myText(22)
        Else
            KI_Erklaerung = myText(23)
        End If
    End If
End If

End Function

Public Function SuchZufallsKarte(HandkartenComputer() As Karte) As Karte
Dim pos As Integer
Dim TKarten() As Karte
Dim i As Integer

pos = 0
For i = UBound(HandkartenComputer) To 1 Step -1
    If HandkartenComputer(i).Bild > ZERO Then
        pos = pos + 1
        ReDim Preserve TKarten(1 To pos)
        TKarten(pos) = HandkartenComputer(i)
        TKarten(pos).Position = i
    End If
Next

pos = Int((UBound(TKarten)) * Rnd + 1)
SuchZufallsKarte = TKarten(pos)

End Function
Public Function ComputerAntwort(handkarten() As Karte, Level As Integer, GegnerPunkte As Integer, MeinePunkte As Integer, MeinErsterStich As Boolean) As Integer
Dim StechKarte As StechInfo
Dim DefensivKarte As Karte

DefensivKarte = ComputerLegtSicher(handkarten, StichTrumpf)
Select Case Level
    Case ZERO
        'Anfänger Level
        'Keine Angriffslogik
        ComputerAntwort = SuchZufallsKarte(handkarten).Position
    Case 1
        'Normal Level
        'Defensiv Spiel
        ComputerAntwort = DefensivKarte.Position
    Case Else
        'Fortgeschrittener level
        'Aggresives Spiel
    
        StechKarte = stichMoeglich(handkarten, StichTrumpf)
        If StechKarte.Karte.Bild > ZERO Then
            
            'wenn ein 16er droht, dann Lusche sticht, und letzte Sieben behalten zum Schluss
            If GegnerPunkte > 4 And MeinErsterStich Then
                KI_Erklaerung = KI_Erklaerung & vbCrLf & myText(24)
                If StechKarte.Karte.Bild <> sieben Then
                    ComputerAntwort = StechKarte.Karte.Position
                    KI_Erklaerung = KI_Erklaerung & vbCrLf & myText(25)
                ElseIf StechKarte.AnzahlSieben = 1 And frmStatistik.txtGelegteKarten < 31 Then
                    ComputerAntwort = DefensivKarte.Position
                    KI_Erklaerung = KI_Erklaerung & vbCrLf & myText(26)
                Else
                    ComputerAntwort = StechKarte.Karte.Position
                    KI_Erklaerung = KI_Erklaerung & vbCrLf & myText(27) & gstrSpace & StechKarte.AnzahlSieben & myText(28)
                End If
                
            'wenn Spieler mit dem Stich mehr als 4 hätte, dann muss Comp stechen
            ElseIf GegnerPunkte + GetPunkteImStich(Stich) > 4 Then
                ComputerAntwort = StechKarte.Karte.Position
                KI_Erklaerung = myText(29)

            'wenn Spieler mit dem Stich mehr als 4 hätte, dann muss Comp stechen
            ElseIf (DefensivKarte.Bild = Ass Or DefensivKarte.Bild = Zehn) And GegnerPunkte + GetPunkteImStich(Stich) >= 4 Then
                ComputerAntwort = StechKarte.Karte.Position
                KI_Erklaerung = myText(29)

            'wenn Spieler mit dem Stich mehr als 4 hätte, dann muss Comp stechen
            '#############
            'to do ... hier könnte noch eine Überprüfung stattfinden, ob der Stich damit wirklich verloren wäre
            '#############
            ElseIf (StechKarte.Karte.Bild = Ass Or StechKarte.Karte.Bild = Zehn) And GegnerPunkte + GetPunkteImStich(Stich) >= 3 And StechKarte.GesamtStiche = 1 Then
                ComputerAntwort = DefensivKarte.Position
                KI_Erklaerung = myText(30)
            
            'wenn Spieler schon 4 Punkte hat und wir Ihm ansonsten ein Punkt reinschmeißen würden ---> stechen
            ElseIf GegnerPunkte = 4 And (DefensivKarte.Bild = Ass Or DefensivKarte.Bild = 10) Then
                ComputerAntwort = StechKarte.Karte.Position
                KI_Erklaerung = myText(31)
                
            'wenn Computer schon 4 Punkte und noch eine Sieben hat, dann auf 4 und letzen spielen
            ElseIf MeinePunkte = 4 And StechKarte.AnzahlSieben = 1 Then
                If frmStatistik.txtGelegteKarten < 31 Then
                    ComputerAntwort = DefensivKarte.Position
                Else
                    ComputerAntwort = StechKarte.Karte.Position
                End If
                KI_Erklaerung = KI_Erklaerung & vbCr & myText(4)
            
            'bei 2 Punkten im Stich --> stechen
            ElseIf GetPunkteImStich(Stich) > 1 Then
                    ComputerAntwort = StechKarte.Karte.Position
                    KI_Erklaerung = myText(34) & gstrSpace & GetPunkteImStich(Stich) & myText(35)
                

            'wenn trumpf punkt ist, nicht ungedeckt einstechen
            ElseIf (StechKarte.Karte.Bild = Ass Or StechKarte.Karte.Bild = Zehn) And StechKarte.GesamtStiche = 1 Then
                ComputerAntwort = DefensivKarte.Position
                KI_Erklaerung = myText(37)
          
            
            ' Zufallsbluff bei Paaren
            ElseIf StechKarte.Karte.Bild <> sieben Or StechKarte.GesamtStiche > 2 Then
                'Paare rauslocken mit Zufallsbluff
                'If StechKarte.GesamtStiche = 1 Then
                If (frmStatistik.txtKarteGef(StechKarte.Karte.Bild) <= 2) And Not (StichTrumpf = Ass Or StichTrumpf = 10) And Not (DefensivKarte.Bild = Ass Or DefensivKarte.Bild = Zehn) Then
                    Bluff = PaarBluff(GegnerPunkte, MeinePunkte)
                    If Bluff Then
                        ComputerAntwort = DefensivKarte.Position
                        KI_Erklaerung = KI_Erklaerung & myText(32)
                    Else
                        ComputerAntwort = StechKarte.Karte.Position
                        KI_Erklaerung = myText(33)
                    End If
                Else
                    
                    ComputerAntwort = StechKarte.Karte.Position
                    Bluff = False
                    ' Debug.Print "kein Bluff!"
                    KI_Erklaerung = myText(33)
                End If
            Else
                ComputerAntwort = DefensivKarte.Position
                KI_Erklaerung = KI_Erklaerung 'Es gibt kein Grund zu stechen
            End If
        Else
            ComputerAntwort = DefensivKarte.Position
            KI_Erklaerung = myText(36)
        End If
    End Select
End Function


