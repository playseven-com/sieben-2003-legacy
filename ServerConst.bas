Attribute VB_Name = "ServerConst"
Option Explicit
Option Compare Text

#If Not Tiny Then
    Public Enum ServerMsgTypes
        Msg_NoOtherPlayers
        Msg_EnumPlayers
        Msg_EnumPlayersResponse
        Msg_PlayerWon
        Msg_PlayerLost
        Msg_PlayerInfo
        Msg_PlayerInfo_OK
        Msg_PlayerInfo_NOK
        Msg_StartGame
        Msg_GameStarted_OK
        Msg_GameStarted_NOK
        msg_gamestart_cancel
        Msg_PopUpMsg
        Msg_SystemMsg
        Msg_PlayerStatus
        Msg_PlayerStatusRe
        Msg_UpdateGame 'to do on server
        Msg_Chat
        Msg_SetGrandMasterSuperior 'to do
        Msg_SetHighScore 'to do
        Msg_NoAnswer 'to do on server
    End Enum
#End If

Public Type SpielerInfo
    SpielerName As String
    GlobalID As String
    SpielerLevel As Integer
    Points As Long
    IP_Adress As String
    ClientID As String
    RegID As String
    Status As PlayerStatus
    SpielOption As SpielOptionen
    AvatarFileName As String
End Type


Public Enum SpielOptionen
    Liga = 0
    Freundschaft = 1
End Enum

Public Enum PlayerStatus
    PlayingMP
    PlayingSP
    Idle
End Enum

Public strPlayerStatus(0 To 2) As String
Public strSpielOption(0 To 1) As String
Public strPlayerLevel(0 To 6) As String

Public Function getSpielOption(str As String) As Integer
Dim i As Integer
For i = LBound(strSpielOption) To UBound(strSpielOption)
    If str = strSpielOption(i) Then Exit For
Next
getSpielOption = i
End Function

Public Function getPlayerStatus(str As String) As Integer
Dim i As Integer
For i = LBound(strPlayerStatus) To UBound(strPlayerStatus)
    If str = strPlayerStatus(i) Then Exit For
Next
getPlayerStatus = i
End Function

Public Function getPlayerLevel(str As String) As Integer
Dim i As Integer
For i = LBound(strPlayerLevel) To UBound(strPlayerLevel)
    If str = strPlayerLevel(i) Then Exit For
Next
getPlayerLevel = i
End Function

