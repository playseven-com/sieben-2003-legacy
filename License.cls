VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "License"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
'
' Lizenzierungstool
'
'zur Festellung der VolumeSerialNumber
Private Declare Function GetVolumeInformation _
  Lib "kernel32" Alias "GetVolumeInformationA" _
  (ByVal lpRootPathName As String, _
   ByVal lpVolumeNameBuffer As String, _
   ByVal nVolumeNameSize As Long, _
   ByRef lpVolumeSerialNumber As Long, _
   ByRef lpMaximumComponentLength As Long, _
   ByRef lpFileSystemFlags As Long, _
   ByVal lpFileSystemNameBuffer As String, _
   ByVal nFileSystemNameSize As Long) As Long
   
'---- Verschlüsselungswerte
Private CRYPT_KEY As String
Private ID_KEY As String
'---- Seriennummer
Private LocalSerial As String
'---- Originallizenz
Public MyLicense As String


'==== PUBLIC ====================

Public Sub Init()
'LocalSerial und Lizenz ermitteln
  Dim APP_NAME As String
  Dim s As String
  '
  '---- Seriennummer von Laufwerk C: holen
  LocalSerial = HDserialString
  '---- falls keine HD-Seriennummer verfügbar ist,
  If (Len(LocalSerial) = 0) Then
    '---- Ersatz-Seriennummer aus der INI holen
    LocalSerial = GetFromReg(KEY_LICCODE2)
    '---- falls auch noch keine Ersatz-Seriennummer vorhanden ist,
    If (Len(LocalSerial) = 0) Then
      '---- Ersatz-Seriennummer generieren
      LocalSerial = Long2String(MakeRandomSerial)
      '---- und in der INI speichern
      Write2Reg KEY_LICCODE2, Encrypt(String2Hex(LocalSerial), True)
    Else
      '---- in Byte-String wandeln
      LocalSerial = Long2String(Hex2Long(LocalSerial))
    End If
  End If
  '---- Produktstring generieren
  '---- evtl. mit "_" o.ä. auffüllen
  APP_NAME = Left(App.ProductName & String(8, "_"), 8)
  '---- Lizenzcode errechnen
  s = LocalSerial & LocalSerial
  s = XORstring(s, APP_NAME)
  s = XORstring(s, CRYPT_KEY)
  '---- in lesbare HexDigits wandeln
  MyLicense = String2Hex(s)
End Sub

Public Property Get ValidLicenseString() As String
'Originallizenz
  '---- formatieren
  ValidLicenseString = FormatLicense(MyLicense)
End Property

Public Property Get UserLicenseString() As String
'eingegebene Lizenz
  Dim s As String
  '
  '---- Lizenz holen
  s = GetSetting(AppExeName, cstrDefault, cstrRegID, vbNullString)
  s = Encrypt(GetFromReg(KEY_LICCODE), False)
  '---- formatieren

  UserLicenseString = FormatLicense(s)
End Property

Public Property Get ID() As String
  ID = String2Hex(XORstring(LocalSerial, ID_KEY))
End Property

Public Function License_OK(Key As String) As Boolean
'Auf gültige Registrierung checken
'liefert TRUE, wenn Lizenz in Ordnung ist,
'andernfalls FALSE
  '---- vorhandene Lizenz mit
  '---- korrekter Lizenz vergleichen
  License_OK = (Key = MyLicense)
End Function

Public Function GenerateLicense(ByVal AppName As String, _
                                ByVal UserId As String) As String
  Dim APP_NAME As String
  Dim s As String
  '
  APP_NAME = Left(AppName & String(8, "_"), 8)
  '---- UserID in Byte-String wandeln
  s = Long2String(Hex2Long(UserId))
  '---- Seriennummer extrahieren
  s = XORstring(s, ID_KEY)
  '---- Lizenzcode errechnen
  s = s & s
  s = XORstring(s, APP_NAME)
  s = XORstring(s, CRYPT_KEY)
  '---- in lesbare HexDigits wandeln
  GenerateLicense = String2Hex(s)
End Function

'==== STUFF ====================

Private Function GetVolumeSerial() As Long
  Dim VolumeNameBuffer As String
  Dim VolumeNameSize As Long
  Dim VolumeSerialNumber As Long
  Dim MaximumComponentLength As Long
  Dim FileSystemFlags As Long
  Dim FileSystemNameBuffer As String
  Dim FileSystemNameSize As Long
  Dim i As Long
  '
  GetVolumeInformation _
    "c:\", _
    VolumeNameBuffer, _
    VolumeNameSize, _
    VolumeSerialNumber, _
    MaximumComponentLength, _
    FileSystemFlags, _
    FileSystemNameBuffer, _
    FileSystemNameSize
  'Seriennummer zurückliefern
  GetVolumeSerial = VolumeSerialNumber
End Function

Private Function MakeRandomSerial() As Long
'generiert eine ZufallsSeriennummer
  MakeRandomSerial = CLng(Rnd() * &H7FFFFFFF)
End Function

Private Function Long2Hex(ByVal l As Long) As String
'wandelt eine 4-Byte-Zahl in einen 8-Zeichen-HexDigitString
  Long2Hex = UCase(Right(String(7, 48) & Hex(l), 8))
End Function

Private Function Hex2Long(ByVal s As String) As Long
'wandelt HexDigit-String zu Long
  Hex2Long = val("&h" & s)
End Function

Private Function Long2String(ByVal l As Long) As String
'wandelt eine 4-Byte-Zahl in einen 4-Zeichen-String
  Long2String = Chr(((l And &HFF000000) \ &H1000000) And &HFF)
  Long2String = Long2String & Chr((l And &HFF0000) \ &H10000)
  Long2String = Long2String & Chr((l And &HFF00&) \ &H100)
  Long2String = Long2String & Chr(l And &HFF&)
End Function

Private Function String2Hex(ByVal s As String) As String
'wandelt einen String in einen HexDigitString
  Dim i As Long, le As Long
  '
  le = Len(s)
  String2Hex = Space(2 * le)
  For i = 0 To le - 1
    Mid(String2Hex, 1 + i * 2 _
      ) = Right("0" & Hex(Asc(Mid(s, i + 1, 1))), 2)
  Next
End Function

Private Function XORstring(ByVal s1 As String, _
                           ByVal s2 As String) As String
'verknüpft zwei Strings gleicher(!) Länge
  Dim i As Long, le As Long
  '
  le = Len(s1)
  XORstring = Space(le)
  For i = 1 To le
    Mid(XORstring, i, 1) = Chr(Asc(Mid(s1, i, 1)) Xor _
                               Asc(Mid(s2, i, 1)))
  Next
End Function

Private Function HDserialString() As String
'liefert die Seriennummer von Laufwerk C: als Byte-String
'bzw. Leerstring, falls keine Seriennummer verfügbar ist
  Dim i As Long
  '
  i = GetVolumeSerial
  '---- in Byte-String wandeln
  HDserialString = IIf(i = 0, "", Long2String(i))
End Function

Private Function FormatLicense(ByVal s As String) As String
'Lizenzstring formatieren
  s = Right(String(16, "0") & s, 16)
  FormatLicense = Mid(s, 1, 4) & " " & _
                  Mid(s, 5, 4) & " " & _
                  Mid(s, 9, 4) & " " & _
                  Mid(s, 13, 4)
End Function

'==== CLASS ====================

Private Sub Class_Initialize()
  Dim i As Long
  Dim s As String
  '
  'Verschlüsselungswerte mit Hilfe
  'einer Sinus-Funktion generieren
  '---- CRYPT_KEY (184,129,112,195,111,48,226,170)
  CRYPT_KEY = ""
  For i = 1 To 8
    s = Mid(str(Abs(Sin(i * 77)) + 13), 4, 3)
    CRYPT_KEY = CRYPT_KEY & Chr(val(s) And 255)
  Next
  '---- ID_KEY (201,113,125,98)
  ID_KEY = ""
  For i = 1 To 4
    s = Mid(str(Abs(Sin(i * 10))), 4, 3)
    ID_KEY = ID_KEY & Chr(val(s) And 255)
  Next
End Sub

