Attribute VB_Name = "SoundNorm"
Option Explicit

Private Const SND_FILENAME = &H20000    'name is a file name
Private Const SND_ASYNC = &H1           'play asynchronously
Private Const SND_NOSTOP = &H10         'Es wird kein aktuell spielender Sound gestoppt um diesen abzuspielen
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
    ByVal uFlags As Long) As Long

Dim SoundPath(14) As String

Public Enum SoundName
    KissSound = 1
    Karte_legen = 2
    SMPlayer = 3
    ComputerNimmt = 4
    SpielerNimmt = 5
    ComputerGewinntRunde = 6
    SpielerGewinntRunde = 7
    SMPlayerChoose = 8
    ENDE = 9
    SpielerNimmtPunkte = 10
    ComputerNimmtPunkte = 11
    spielerFehler = 12
    TimeTick = 13
    LevelDown = 14
End Enum

Public Sub DXSoundInit()
    SoundPath(1) = App.path & cstrSubPathAudio & "LevelUpgrade.wav"
    SoundPath(2) = App.path & cstrSubPathAudio & "Karte_legen.wav"
    SoundPath(3) = App.path & cstrSubPathAudio & "SPIRIT.WAV"
    SoundPath(4) = App.path & cstrSubPathAudio & "RESORAVE.WAV"
    SoundPath(5) = App.path & cstrSubPathAudio & "GLOCKE01.WAV"
    SoundPath(6) = App.path & cstrSubPathAudio & "ComputerWinRound.wav"
    SoundPath(7) = App.path & cstrSubPathAudio & "SpielerWinRound.wav"
    SoundPath(8) = App.path & cstrSubPathAudio & "Button10.wav"
    SoundPath(9) = App.path & cstrSubPathAudio & "ENDE.WAV"
    SoundPath(10) = App.path & cstrSubPathAudio & "GLOCKE02.WAV"
    SoundPath(11) = App.path & cstrSubPathAudio & "Poing1.WAV"
    SoundPath(12) = App.path & cstrSubPathAudio & "crowdohh.wav"
    SoundPath(13) = App.path & cstrSubPathAudio & "tick.wav"
    SoundPath(14) = App.path & cstrSubPathAudio & "Button2.wav"
End Sub
    

Public Sub Playsound(Sound As Integer, Optional Volume As Long)
Dim ret As Integer
On Error GoTo ERRHand
If AudioOn Then
    ret = sndPlaySound(SoundPath(Sound), SND_ASYNC + SND_NOSTOP + SND_FILENAME)
End If

Exit Sub
ERRHand:
If ErrorBox("PlaySound", Err) Then Resume Next

'    If ret = 0 Then
'        MsgBox "Die Wavedatei " & SoundPath(Sound) & " konnte nicht gefunden werden."
'    End If
End Sub
