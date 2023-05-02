Attribute VB_Name = "modDplay"
Option Explicit
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
'  Copyright (C) 1999-2001 Microsoft Corporation.  All Rights Reserved.
'
'  File:       modDplay.bas
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Enum vbDplayChatMsgType
    MsgChat
    MsgWhisper
    MsgSendMixedGame
    MsgSendGeber
    MsgSendAbheben
    MsgSendCard
    MsgSendStichEnde
    MsgSendSpielAbbruch
    MsgSendMeisterfehler
    MsgSendZeitueberschreitung
End Enum

Public Const StartGame = "/7?"
Public Const StartGameOK = "/7OK"
Public Const StartGameNOK = "/7NOK"

'Constants
Public dx As DirectX8
Public dpp As DirectPlay8Peer
Public dpc As DirectPlay8Client
Public dvServer As DirectPlayVoiceServer8
Public dvClient As DirectPlayVoiceClient8

Public glMyID As Long

Public Const glDefaultPort As Long = 80
Public Const gstrServerName = "21235E36305C2D3C654133663F05753079783906"

Public AppGuid As String
Public AppGuidP2P As String
Public Const AppGuidEncrypted = "A449370C0C31111360377E2A1D1B230309224705710B0A382C1123400A48761270520C3F269C"
Public Const AppGuidP2PEncrypted = "A41B360E0C31111360377E2A1D1B230309224705710B0A382C1123400A48761270520E353D9C"

Public SpielErhalten As Boolean
Public MakeHost As Boolean
Public ServerConnected As Boolean

'App specific variables
Public gsUserName As String
'Our connection form and message pump
'for p2p playerConnection
Public DPlayEventsForm As DPlayConnect
'for c/s GameserverConnection
Public ServerEventsForm As ServerConnect

Public frm As frmUserReq

Public Type HostFound
    AppDesc As DPN_APPLICATION_DESC
    AddressHost As String
    AddressDevice As String
    TimeLastFound As Long
End Type

'Declares for closing the form without waiting
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const WM_CLOSE = &H10

'Public P2PHostFromServer As HostFound

Public Sub InitDPlay(caller As String)
On Error GoTo ERRHand
Dim Done_OK As Boolean
    'Create our DX/DirectPlay objects
    Set dx = New DirectX8
    
'    Init_DPP caller
'    Exit Sub
    
    Set dpc = dx.DirectPlayClientCreate
    Set ServerEventsForm = New ServerConnect
    Done_OK = ServerEventsForm.StartClientConnectWizard(dx, dpc, AppGuid)
    If Not Done_OK Then
        CleanUpServer
        Init_DPP caller
    End If
Exit Sub
ERRHand:
If ErrorBox("Init_DPlay", Err) Then Resume Next
End Sub
Public Sub Init_DPP(caller As String)

On Error GoTo ERRHand
Set dpp = dx.DirectPlayPeerCreate

Set DPlayEventsForm = New DPlayConnect
'Start the connection form (it will either create or join a session)
If Not DPlayEventsForm.StartConnectWizard(dx, dpp, AppGuidP2P, 4, frmChat) Then
    CleanUpP2P
    Exit Sub
Else
    AnimWindow frmChat, AW_ACTIVATE + IIf(caller = "frmMain", AW_SLIDE + AW_VER_NEGATIVE, AW_BLEND)
End If
Debug.Print "DPP Init"

Exit Sub
ERRHand:
If ErrorBox("Init_DPP", Err) Then Resume Next
End Sub


Public Sub CleanUp()
On Error GoTo ERRHand
CleanUpP2P
CleanUpServer
    
Set dx = Nothing
Exit Sub
ERRHand:
If ErrorBox("CleanUp", Err) Then Resume Next
End Sub
Public Sub CleanUpP2P()
    CleanUpVoice
            
    If Not (dpp Is Nothing) Then dpp.UnRegisterMessageHandler
    
    'Close down our session
    If Not (dpp Is Nothing) Then
        dpp.Close
        DoSleep 500
    End If
    
    If Not (DPlayEventsForm Is Nothing) Then
        'Get rid of our message pump
        DPlayEventsForm.GoUnload
        
        Set DPlayEventsForm = Nothing
    End If
    

    Set dpp = Nothing
    Debug.Print "CleanedUpP2P"
   
End Sub
Public Sub CleanUpServer()

On Error GoTo ERRHand
    If Not (dpc Is Nothing) Then dpc.UnRegisterMessageHandler
    If Not (dpc Is Nothing) Then dpc.Close
    'dpc.CancelAsyncOperation
    DoSleep 500
    Set dpc = Nothing

    If Not ServerEventsForm Is Nothing Then
        ServerEventsForm.GoUnload
        
        Set ServerEventsForm = Nothing
    End If
    ServerConnected = False
    Debug.Print "CleanedUpServer"
Exit Sub
ERRHand:
MsgBox Err.Number & vbCr & Err.Description, vbCritical
Resume Next

End Sub

Public Sub CleanUpVoice()

On Error Resume Next
'Disconnect and destroy the client
If Not (dvClient Is Nothing) Then
    dvClient.UnRegisterMessageHandler
    dvClient.Disconnect DVFLAGS_SYNC
    Set dvClient = Nothing
End If
'Stop and Destroy the server
If Not (dvServer Is Nothing) Then
    dvServer.UnRegisterMessageHandler
    dvServer.StopSession 0
    Set dvServer = Nothing
End If

Debug.Print "CleanedUpVoice"

End Sub

Public Sub CloseForm(oForm As Form)
    'Anytime we need to close a form from within a DirectPlay callback
    'we need to use this function.  The reason is that DirectPlay uses multiple
    'threads to spawn all of it's messages back to the application.  However
    'it cannot close down until all of it's threads have returned.
    'If we attempt to simply call Unload Me in the callback, we will run into
    'a deadlock instance, since the callback will be running on the DirectPlay
    'thread waiting for the unload to finish, and the unload will be waiting
    'for the DirectPlay thread to finish.
    
    'PostMessage puts the message on the queue for our form and returns immediately
    'allowing the thread to finish
    PostMessage oForm.hWnd, WM_CLOSE, 0, 0
End Sub



