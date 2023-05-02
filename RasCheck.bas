Attribute VB_Name = "RasCheck"


' Feststellen, ob eine RAS-Verbindung existiert
Option Explicit
Option Compare Text

'Const RAS95_MaxEntryName = 256
'Const RAS_MaxDeviceType = 16
'Const RAS95_MaxDeviceName = 128
'Const RASCS_DONE = &H2000&
'
'Private Type RASCONN95
'  ' dwsize erhält den Wert 412
'  dwSize As Long
'  hRasConn As Long
'  szEntryName(RAS95_MaxEntryName) As Byte
'  szDeviceType(RAS_MaxDeviceType) As Byte
'  szDeviceName(RAS95_MaxDeviceName) As Byte
'End Type
'
'Private Type RASCONNSTATUS95
'  ' dwsize erhält den Wert 160
'  dwSize As Long
'  RasConnState As Long
'  dwError As Long
'  szDeviceType(RAS_MaxDeviceType) As Byte
'  szDeviceName(RAS95_MaxDeviceName) As Byte
'End Type
'
'Private Declare Function RasEnumConnections Lib "RasApi32.dll" Alias "RasEnumConnectionsA" (lprasconn As Any, lpcb As Long, lpcConnections As Long) As Long
'Private Declare Function RasGetConnectStatus Lib "RasApi32.dll" Alias "RasGetConnectStatusA" (ByVal hRasConn As Long, lpRASCONNSTATUS As Any) As Long

'###### zweite möglichkeit
Private Declare Function InternetGetConnectedState _
        Lib "wininet.dll" (ByRef lpSFlags As Long, _
        ByVal dwReserved As Long) As Long


'Private Const INTERNET_CONNECTION_MODEM As Long = &H1
'Private Const INTERNET_CONNECTION_LAN As Long = &H2
'Private Const INTERNET_CONNECTION_PROXY As Long = &H4
'Private Const INTERNET_CONNECTION_MODEM_BUSY As Long = &H8
'Private Const INTERNET_CONNECTION_OFFLINE As Long = &H20
'Private Const INTERNET_CONNECTION_CONFIGURED As Long = &H40
'Private Const INTERNET_RAS_INSTALLED As Long = &H10


'Public Function BestehtVerbindung2() As Boolean
'    Dim TRasCon(255) As RASCONN95
'    Dim n As Long
'    Dim nStatus  As Long
'    Dim nRetval  As Long
'    Dim Tstatus As RASCONNSTATUS95
'    TRasCon(0).dwSize = 412
'    n = 256 * TRasCon(0).dwSize
'    nRetval = RasEnumConnections(TRasCon(0), n, nStatus)
'
'    If nRetval <> 0 Then
'        MsgBox "Es trat ein Fehler auf!", vbCritical, "RAS Verbindungscheck"
'        Exit Function
'    End If
'    Tstatus.dwSize = 160
'    nRetval = RasGetConnectStatus(TRasCon(0).hRasConn, Tstatus)
'    If Tstatus.RasConnState = RASCS_DONE Then
'       BestehtVerbindung2 = True
'    Else
'       BestehtVerbindung2 = False
'    End If
'End Function

Public Function BestehtVerbindung() As Boolean
Dim icFlags As Long
    
BestehtVerbindung = InternetGetConnectedState(icFlags, 0&)

'Select Case icFlags
'    Case icFlags And INTERNET_CONNECTION_LAN, icFlags And INTERNET_CONNECTION_MODEM, _
'            icFlags And INTERNET_CONNECTION_PROXY, icFlags And INTERNET_CONNECTION_MODEM_BUSY
'        BestehtVerbindung = True
'    Case icFlags And INTERNET_RAS_INSTALLED
'        BestehtVerbindung = BestehtVerbindung2()
'    Case Else
'        BestehtVerbindung = False
'End Select

End Function

