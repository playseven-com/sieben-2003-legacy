Attribute VB_Name = "DXSound"
Option Explicit

Private Type DSOUND
    DSBuffer             As DirectSoundSecondaryBuffer8
    Notification         As Long
    ChannelVolume        As Long
    Playing              As Boolean
    path                 As String
End Type

Public Sounds() As DSOUND
Private SW As DirectSound8Wrapper

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

Sub DXSoundInit()
On Error Resume Next
AudioOn = CBool(GetSetting(AppExeName, cstrOptions, "Audio", True))
Set SW = New DirectSound8Wrapper
If SW Is Nothing Then
    MsgBox "DirectSound could not be intilalized", vbCritical, "please install DirectX 8"
Else
    LoadSounds
End If
On Error GoTo 0
End Sub
  
Public Sub TerminateSound()
On Error GoTo ERRHand
    Set SW = Nothing
    'Erase the full sound array
    Erase Sounds
Exit Sub
ERRHand:
If ErrorBox("TerminateSound", Err) Then Resume Next
End Sub


Sub LoadSounds()

On Error GoTo ERRHand
    SW.CreateSoundBuffer App.path & cstrSubPathAudio & "LevelUpgrade.wav", False, True, False, True
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "Karte_legen.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 2, -1800
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "SPIRIT.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 3, -1300
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "RESORAVE.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 4, -1300
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "GLOCKE01.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 5, -1300
    End If
    SW.CreateSoundBuffer App.path & cstrSubPathAudio & "ComputerWinRound.wav", False, True, False, True
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "SpielerWinRound.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 7, -1300
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "button10.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 8, -1300
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "ENDE.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 9, -1600
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "GLOCKE02.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 10, -1600
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "Poing1.WAV", False, True, False, True) <> -1 Then
        SW.SetVolume 11, -1600
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "crowdohh.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 12, -1600
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "tick.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 13, -1000
    End If
    If SW.CreateSoundBuffer(App.path & cstrSubPathAudio & "button2.wav", False, True, False, True) <> -1 Then
        SW.SetVolume 14, -1600
    End If

Exit Sub
ERRHand:
    If ErrorBox("LoadSounds", Err) Then Resume Next
End Sub

Sub PlaySound(Sound As SoundName, Optional Volume)
On Error Resume Next
    If Sound > ZERO And AudioOn And Not SW Is Nothing Then
        If Not IsMissing(Volume) Then
            SW.SetVolume Sound, Volume
'            Debug.Print Volume
        End If
        SW.PlaySound Sound
    End If

End Sub

Sub SetVolume(val As Long)
    SW.SetVolume 5, val
End Sub
