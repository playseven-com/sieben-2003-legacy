VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clFormResizer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'** FormResizer-Klasse
Option Explicit

'** die zu behandelnde Form
Private WithEvents TheForm As Form
Attribute TheForm.VB_VarHelpID = -1
'** automatisches Resizen?
Public AutoResize As Boolean

'** Ausnahme-Flags (öffentlich)
Public Enum Exceptions
    exNoMove = 1
    exNoScale = 2
    exScaleKeepAspect = 4
    exNoHScaleRightAlign = 8
    exNoVScaleBottomAlign = 16
    exNoScaleFont = 32
End Enum

'** Sicherungstyp für Controls
Private Type ControlType
    Index   As Integer
    Left    As Single
    Top     As Single
    Width   As Single
    Height  As Single
    FontSize As Single
    Flags   As Exceptions
End Type
Private Ctrls() As ControlType

'** Original-Fenstergröße
Private FoScaleWidth As Single
Private FoScaleHeight As Single
Private FoWidth As Single
Private FoHeight As Single
'** Mindestgröße beibehalten?
Private KeepMinSize As Boolean

Private Sub Class_Initialize()
'** Klassen-Initialisierung
    ReDim Ctrls(0)
    AutoResize = True
End Sub

Public Sub Initialize(Fo As Form, _
    Optional ByVal SetMinSize As Boolean)
'** Initialisieren: Elementezustände sichern
'**  SetMinSize = True: Aktuelle Größe ist Mindestgröße!
    '** Form übernehmen
    Set TheForm = Fo
    
    '** Startgröße sichern
    FoWidth = Fo.Width
    FoHeight = Fo.Height
    FoScaleWidth = Fo.ScaleWidth
    FoScaleHeight = Fo.ScaleHeight
    
    '** Mindestgröße beachten?
    KeepMinSize = SetMinSize
    
    '** Controlgrößen sichern
    Dim Anz As Long
    Anz = Fo.Controls.count - 1
    ReDim Ctrls(Anz)
    
    '** für die Prüfung, ob Resizing möglich
    On Error Resume Next
    Dim WIDummy As Long
    
    Dim i As Long, Index As Long
    For i = 0 To Anz
        With Fo.Controls(i)
            '** prüfen, ob Resize möglich
            Err.Clear
            WIDummy = .Width
            If Err.Number = 0 Then
                '** ja: skalierbares Element eintragen
                Ctrls(Index).Index = i
                Ctrls(Index).Left = .Left
                Ctrls(Index).Top = .Top
                Ctrls(Index).Width = .Width
                Ctrls(Index).Height = .Height
                Ctrls(Index).FontSize = .FontSize
                Index = Index + 1
            End If
        End With
    Next
    
    '** Arraygröße anpassen
    ReDim Preserve Ctrls(Index - 1)
End Sub

Function SetExceptionCtrl(TheControl As Control, ByVal ExceptionFlags As Exceptions) As Boolean
'** Ausnahme definieren
'** Ergebnis True, wenn OK
    Dim Index As Long
    Index = GetCtlIndex(TheControl)
    
    '** bei Fehlschlag abbrechen!
    If Index = -1 Then Exit Function
    
    '** Typ einsetzen
    Ctrls(Index).Flags = ExceptionFlags
    
    '** Erfolg melden
    SetExceptionCtrl = True
End Function

Function ResizeSingleCtrl(TheControl As Control) As Boolean
'** einzelnes Steuerelement wurde manuell verändert
'** Ergebnis True, wenn OK
    Dim Index As Long
    Index = GetCtlIndex(TheControl)
    
    '** bei Fehlschlag abbrechen!
    If Index = -1 Then Exit Function
    
    '** Skalierung berechnen
    Dim FaktorX As Double
    Dim FaktorY As Double
    FaktorX = TheForm.ScaleWidth / FoScaleWidth
    FaktorY = TheForm.ScaleHeight / FoScaleHeight
    
    '** neue Größe eintragen
    With TheControl
        Ctrls(Index).Left = .Left / FaktorX
        Ctrls(Index).Top = .Top / FaktorY
        Ctrls(Index).Width = .Width / FaktorX
        Ctrls(Index).Height = .Height / FaktorY
    End With
    
    ResizeSingleCtrl = True
End Function


Private Function GetCtlIndex(TheControl As Control) As Long
'** Index des Controls ermitteln
'** Ergebnis: Index oder -1 bei Fehlschlag
Dim i As Long
    GetCtlIndex = -1
    
    For i = 0 To UBound(Ctrls)
        If TheForm.Controls(Ctrls(i).Index) Is TheControl Then
            '** gefunden!
            GetCtlIndex = i
            Exit For
        End If
    Next
End Function

Private Sub TheForm_Resize()
'** ggf. automatisch Resizen
    If AutoResize Then
        ResizeForm
    End If
End Sub

Sub ResizeForm()
'** Form resizen => Elemente anpassen
    If TheForm Is Nothing Then Exit Sub
    '** Abbruch bei Minimierung
    If TheForm.WindowState = vbMinimized Then Exit Sub
    
    If KeepMinSize Then
        '** Mindestgröße erhalten
        If TheForm.Width < FoWidth Then
            '** Mindestbreite einstellen!
            TheForm.Width = FoWidth
            Exit Sub
        End If
        If TheForm.Height < FoHeight Then
            '** Mindesthöhe einstellen!
            TheForm.Height = FoHeight
            Exit Sub
        End If
    End If
    
    '** Skalierungen berechnen
    Dim FaktorX As Double
    Dim FaktorY As Double
    Dim Faktor As Double
    FaktorX = TheForm.ScaleWidth / FoScaleWidth
    FaktorY = TheForm.ScaleHeight / FoScaleHeight
    
    '** Elemente skalieren
    Dim i As Long
    Dim NewX As Single, NewY As Single
    Dim NewW As Single, NewH As Single
    '** Fehler ignorieren (z.B. bei Combos)
    On Error Resume Next
    For i = 0 To UBound(Ctrls)
        '** pro Element:
        With Ctrls(i)
            '** angepasste Position/Größe berechnen
            NewX = .Left * FaktorX
            NewY = .Top * FaktorY
            NewW = .Width * FaktorX
            NewH = .Height * FaktorY
            
            If (.Flags And exNoMove) Then
                '** nicht verschieben
                NewX = .Left
                NewY = .Top
            Else
                If .Flags And exNoHScaleRightAlign Then
                    '** rechts-orientiert
                    NewX = TheForm.ScaleWidth - _
                        (FoScaleWidth - .Left)
                    NewW = .Width
                End If
                If .Flags And exNoVScaleBottomAlign Then
                    '** unten-orientiert
                    NewY = TheForm.ScaleHeight - _
                        (FoScaleHeight - .Top)
                    NewH = .Height
                End If
            End If
            
            If .Flags And exNoScale Then
                '** nicht skalieren!
                NewW = .Width
                NewH = .Height
            Else
                '** ggf. skalieren/einpassen
                If .Flags And exScaleKeepAspect Then
                    '** resizen, Seitenverhältnis beibehalten
                    '** Minimalwert ermitteln
                    Faktor = FaktorX
                    If Faktor > FaktorY Then Faktor = FaktorY
                    NewW = .Width * Faktor
                    NewH = .Height * Faktor
                    If Faktor >= 1 Then
                        '** nur wenn größer als Original!
                        NewX = .Left + _
                            .Width * (FaktorX - Faktor) / 2
                        NewY = .Top + _
                            .Height * (FaktorY - Faktor) / 2
                    End If
                End If
            End If
            
            '** Element setzen
            TheForm.Controls(.Index).Move NewX, NewY, NewW, NewH
            
            If Not .Flags And exNoScaleFont Then
                TheForm.Controls(.Index).FontSize = .FontSize * FaktorX
            End If
        End With
    Next
End Sub

