Attribute VB_Name = "global"
Option Explicit
Option Compare Text

Public Const gstrNullstr As String = ""
Public Const gstrSpace As String = " "
Public Const gstrDblDot As String = ":"
Public Const gstrDot As String = "."
Public Const gstrKomma As String = ","
Public Const gstrMinus As String = "-"
Public Const gstrDirSep As String = "\"
Public Const ZERO As Integer = 0
Public Const ONE As Integer = 1

Public Const mailCrLf As String = "%0D%0A"

Public Const AppExeName As String = "Sieben"

Public Const cstrXXX As String = "XXXXXXX"
Public Const cstrLinie As String = "-------------------"
Public Const cstrSystemLbl As String = " | --<SYSTEM>-- : "
Public Const cstrGegnerStandardName As String = "Computer"
Public Const gstrTimeInterval As String = "7 s"

#If Not Tiny Then
    Public Const cstrshowSystemMessage As String = "showSystemMessage"
    Public Const cstrMakeLog As String = "makeLog"
    Public Const cstrNoHigherLevel As String = "NoHigherLevel"
    Public Const cstrSpielOption As String = "GameOption"
#End If

'Pfade
Public Const cstrSubPathBackGround As String = "\Source\BackGrounds\"
Public Const cstrSubPathAvatars As String = "\Source\Avatars\"
Public Const cstrSubPathAudio As String = "\Source\Audio\"
Public Const cstrSubPathDeck As String = "\Source\Decks\Standard\"
'Dateien
Public Const cstrCardBackName As String = "Rueckseite.jpg"
Public Const cstrLangDBPath As String = "\Source\lang.mdb"

'Registriereinträge
Public Const cstrLastBackGround As String = "lastBackGround"
Public Const cstrUseAgent As String = "useAgent"
Public Const cstrAgentGivesTips As String = "AgentGivesTips"
Public Const cstrShowIndikator As String = "show Indikator"
Public Const cstrAgentTalkChat As String = "AgentTalkChat"
Public Const cstrShowTips As String = "Show Tips at Startup"
Public Const cstrStarts As String = "Starts"
Public Const cstrHighscore As String = "Highscore"
Public Const cstrLastPlayer As String = "Last Player"
Public Const cstrTransparency As String = "Transparency"
Public Const cstrOptions As String = "Options"
Public Const cstrDefault As String = "Defaults"
Public Const cstrAudio As String = "Audio"
Public Const cstrClientID As String = "CID"
Public Const cstrRegID As String = "RID"
Public Const cstrLanguage = "LastLanguage"

Public Const PercFormat As String = "0.00%"
Public Const MaxStartsWReg As Integer = 7
Public Const MaxTime As Integer = 60

Public Const Computer As Integer = -1
Public Const Spieler As Integer = 1

'Farben
Public Const ROT = vbRed
Public Const hellROT = &HAA&
Public Const GRUEN = vbGreen
Public Const hellGRUEN = &HAA00&
Public Const BLAU = vbBlue
Public Const hellBLAU = &HAA0000
Public Const SCHWARZ = vbBlack

'spielintern
Public LaufendesSpiel As Boolean
Public ZeitUeberschreitung As Boolean
Public ZeitUeberschreitungGegner As Boolean
Public MeisterFehler As Boolean
Public MeisterFehlerGegner As Boolean

Public boolPlayerWon As Boolean

Public Test As Boolean
'Public BOT As Boolean
Public ShowTipAtStartup As Boolean
Public ShowIndikator As Boolean
Public useAgent As Boolean
Public AudioOn As Boolean
Public AgentGivesTips As Boolean
Public frmChat_Loaded As Boolean
Public frmMain_Loaded As Boolean
Public frmStatistik_Loaded As Boolean

Public bool_isRegistered As Boolean

Public Enum PlayerMode
    singleplayer = 0
    multiplayer = 1
End Enum

Public User As String
Public CompName As String

Public glHighScore As Long

Public Playermodus As PlayerMode
Public AktuellerSpieler As SpielerInfo
Public Gegner As SpielerInfo

Public AppInfo As String
Public StartAnz As Long
Public WindowsVersion As Long
'Public strPlayerlevel(0 To 6) As String

'#############################################
'##############----API----####################
'#############################################

'apis für SysTrayIcon
Public Type NOTIFYICONDATA
    cbSize As Long
    hWnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    sTip As String * 64
End Type

Public sysIcon As NOTIFYICONDATA

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const NIF_DOALL = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONUP = &H205

Public Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

'Api für Form verschieben
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Declare Sub ReleaseCapture Lib "user32" ()
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

'Api für 'Form immmer im Vordergrund'
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_FRAMECHANGED = &H20  'Das Fenster wird an seiner neuen Position vollkommen neu gezeichnet
Public Const SWP_HIDEWINDOW = &H80  'Das Fenster wird versteckt
Public Const SWP_NOACTIVATE = &H10  'Das Fenster wird nicht Akteviert
Public Const SWP_NOCOPYBITS = &H100  'Der inhalt der Form wird nicht mitkopiert
Public Const SWP_NOMOVE = &H2  'Die Position des Fensters wird nicht verändert
Public Const SWP_NOSIZE = &H1  'Die größe de Fensters wird nicht verändert
Public Const SWP_NOREDRAW = &H8  'Zeichent das Fenster nicht neu
Public Const SWP_NOZORDER = &H4  'Ignoriert die einstellungen der Z-Order
Public Const SWP_SHOWWINDOW = &H40  'Zeigt das Fenster an


'Api für Browser und Mail starten
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
      (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, _
      ByVal lpParameters As String, ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long
      
Public Const SW_HIDE = 0
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 5
Public Const SW_MINIMIZE = 6
Public Const SW_SHOWMINNOACTIVE = 7
Public Const SW_SHOWNA = 8
Public Const SW_RESTORE = 9
Public Const SW_SHOWDEFAULT = 10


'Api für Erzeugeung einer eindeutigen Spielerkennung und anderer GUIDs
Private Type GUIDByteArray
  GuidData(20) As Byte
End Type
Private Declare Function CoCreateGuid Lib "ole32.dll" (pguid As Any) As Long
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpstrClsId As Long, ByVal cbMax As Long) As Long

'Konstanten für Listviewelement per Sendmessage api
Private Const LVM_FIRST = &H1000
Private Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Private Const LVSCW_AUTOSIZE = -1
Private Const LVSCW_AUTOSIZE_USEHEADER = -2


'Sonstige Api's
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'###################

'Api um Titelleiste wegzublenden, bei beibehaltung des Icons und der Titelleiste
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_STYLE = (-16)
Private Const WS_SYSMENU = &H80000


'---------
Public ModText() As String

Public Function HideTitleBar(frm As Form)
Dim llngStyle As Long

On Error Resume Next

    llngStyle = GetWindowLong(frm.hWnd, GWL_STYLE) Or WS_SYSMENU
    SetWindowLong frm.hWnd, GWL_STYLE, llngStyle

End Function

Public Function GetGUID() As String
Dim uGUID As GUIDByteArray, sGuid As String, GuidLen As Long

  sGuid = Space$(38)
  CoCreateGuid uGUID
  GuidLen = StringFromGUID2(uGUID, StrPtr(sGuid), 39)
  If GuidLen = ZERO Then GetGUID = "00000000": Exit Function
  sGuid = Mid$(sGuid, 2, GuidLen - 3)
  'GetGUID = Left$(sGUID, 8) & Mid$(sGUID, 10, 4) & Mid$(sGUID, 15, 4) & Mid$(sGUID, 20, 4) & Right$(sGUID, 12)
  ' Ab VB6 einfacher:
  GetGUID = Replace(sGuid, gstrMinus, vbNullString)
End Function

Public Function ShowAvatar(frm As Form, path As String) As Boolean
'Lädt die AvatarPics in die Controls des übergeben Fensters
On Error Resume Next
If FileExists(path) Then
    With frm
        If Right(path, 4) <> ".gif" Then
            .aniGifAvatar.Visible = False
            .ImgAvatar.Picture = LoadPicture(path)
            .ImgAvatar.Visible = True
        Else
            .ImgAvatar.Visible = False
            .aniGifAvatar.FileName = path
            .aniGifAvatar.Visible = True
            .aniGifAvatar.ZOrder
        End If
    End With
End If

End Function

'Zeigt den fehler an und schickt Mail mit Fehlermeldung
Public Function ErrorBox(Proc As String, ByRef error As ErrObject) As Boolean
Dim Antw As Integer
Dim str As String
    
    str = "An Error occured in Procedure: " & Proc & vbCr & _
        " Errorcode: " & Err.Number & gstrSpace & Err.Description & vbCr & _
        " LastDllError: " & Err.LastDllError & vbCr & _
        " Source : " & Err.Source
    Antw = MsgBox(str & vbCr & "Would you like to continue this Applikation" _
        , vbYesNo + vbCritical, "Fehler !")
    
    SendMail "betatester@playseven.com", "Error in " & AppInfo, Replace(str, vbCr, mailCrLf)
    If Antw = vbYes Then
        ErrorBox = True
        Err.Clear
    Else
        CloseAll
        End
    End If
End Function

'Überprüft ob ein Feld mit Wreten gefüllt ist
Public Function ArrayIsFilled(a() As String) As Boolean
Dim X As Long

On Error GoTo ERRHand
X = UBound(a)
ArrayIsFilled = True

Exit Function

ERRHand:
ArrayIsFilled = False
End Function

'Ermittelt den User und ComputerNamen
Sub getUserAndComp(User As String, CompName As String)
On Error GoTo ERRHand
Dim errval As Long, i As Long
    User = String(25, 0): i = Len(User)
    errval = GetUserName(ByVal User, i)
    User = Left$(User, i - 1)
    CompName = Space(16): i = Len(CompName)
    errval = GetComputerName(ByVal CompName, i)
    CompName = Left$(CompName, i)
Exit Sub
ERRHand:
If ErrorBox("getUserAndComp", Err) Then Resume Next
End Sub


Public Sub SetzFensterMittig(ByRef frm As Form)
On Error Resume Next
frm.Move (Screen.Width - frm.Width) \ 2, (Screen.Height - frm.Height) \ 2

End Sub


Private Sub InitBase()
'initialisiert die Grundkomponeneten
GetWindowsVersion
SetAppInfo
getUserAndComp User, CompName
DBInit
InitLangDb
DXSoundInit
End Sub

Public Sub SetAppInfo()
    AppInfo = vbCrLf & App.EXEName & " Version " & App.Major & gstrDot & App.Minor & gstrDot & App.Revision & " /" & StartAnz & " /" & AktuellerSpieler.SpielerLevel & " /" & WindowsVersion
End Sub

Public Sub WriteLblLevel(lbl As Label, ii As Integer)
    lbl.Caption = strPlayerLevel(ii)
End Sub

Public Function getLevelFromStr(str As String) As Integer
Dim i As Integer
For i = LBound(strPlayerLevel) To UBound(strPlayerLevel)
    If strPlayerLevel(i) = str Then
        getLevelFromStr = i
        Exit For
    End If
Next
End Function

Public Sub CloseAll()

On Error GoTo ERRHand
    #If Not Tiny Then
        TerminateSound
        CleanUp
    #End If
    Shell_NotifyIcon NIM_DELETE, sysIcon
    CloseLangDb
    DBClose
Exit Sub
ERRHand:
If ErrorBox("CloseAll", Err) Then Resume Next
End Sub

Public Sub DropDown(Combo As ImageCombo, Optional ByVal ShowHide As Boolean = True)
  Const CB_SHOWDROPDOWN = &H14F
  Const CB_SETITEMHEIGHT = &H153
  SendMessage Combo.hWnd, CB_SHOWDROPDOWN, ShowHide, 0
End Sub

Public Function isBitSet(i As Long, BitPos As Byte) As Boolean
On Error GoTo ERRHand
    isBitSet = (i And (2 ^ (BitPos - 1)))
Exit Function
ERRHand:
If ErrorBox("isBitSet " & i & ":" & BitPos, Err) Then Resume Next
End Function

Private Sub getStartAnz()
    StartAnz = val(GetSetting(AppExeName, cstrOptions, cstrStarts))
    StartAnz = val(GetFromReg(cstrStarts))
    StartAnz = StartAnz + 1
    SaveSetting AppExeName, cstrOptions, cstrStarts, StartAnz
    Write2Reg cstrStarts, CStr(StartAnz)
End Sub

Sub Main()
    
On Error GoTo ERRHand:
    If App.PrevInstance Then
        MsgBox App.EXEName & " is already started.... "
        Exit Sub
    End If
    
    InitBase
    
    getStartAnz
    
    glHighScore = GetSetting(AppExeName, cstrDefault, cstrHighscore, 1000)
    
    useAgent = GetSetting(AppExeName, cstrOptions, cstrUseAgent, True)
    
    SaveSetting AppExeName, cstrDefault, "Version", App.Major & gstrDot & App.Minor & gstrDot & App.Revision
    'initialisierung der DBs, Sound etc.
    
    #If Not Tiny Then
        AppGuid = Encrypt(AppGuidEncrypted, False)
        AppGuidP2P = Encrypt(AppGuidP2PEncrypted, False)
    #End If
    
    'Sprachstrings laden
    LoadObjectText "AllModuls", ModText()
    Init_KI
    setStrPlayerLevel
    
    AnimWindow frmSplash, AW_ACTIVATE + AW_BLEND
    
Exit Sub
ERRHand:
    If ErrorBox("Main", Err) Then Resume Next
End Sub

Public Sub AnimWindow(frm As Form, effect As AW_FLAGS)
Dim ANIMED As Boolean
On Error GoTo ERRHand

    LinkWndProc frm
    ANIMED = AnimateWnd(frm.hWnd, 250, effect)
    DetachWndProc frm
    
    If Not ANIMED Then
        If isBitSet(effect, 17) Then
            frm.Hide
        Else
            frm.Show
            RefreshAllCtls frm
        End If
    End If
    
Exit Sub
ERRHand:
If ErrorBox("AnimWindow", Err) Then Resume Next

End Sub

Public Function RefreshAllCtls(frm As Form)

Dim ctl As Control
On Error Resume Next
    'frm.Refresh
    For Each ctl In frm.Controls
        If Not TypeOf ctl Is Menu And Not TypeOf ctl Is Timer And Not TypeOf ctl Is Agent And Not TypeOf ctl Is ImageList Then
            ctl.Refresh
       End If
    Next
    'frm.Refresh

Exit Function
ERRHand:
If ErrorBox("RefreshAllCtls", Err) Then Resume Next
End Function

'Public Function HasContainer(ctl As Control) As Boolean
'Dim a As String
'On Error Resume Next
'a = ctl.Container
'Debug.Print a
'If Err.Number = ZERO Then
'    HasContainer = True
'Else
'    HasContainer = False
'End If
'End Function

Public Sub MakeMsg(str As String)
    frmStatistik.lstVerlauf.AddItem str
    frmStatistik.lstVerlauf.Selected(frmStatistik.lstVerlauf.ListCount - 1) = True
    Write2Log str
End Sub




Public Sub SendMail(sTo As String, sSubject As String, sBody As String)
    ShellExecute 0, "open", "mailto:" & sTo & " ?Subject=" & sSubject & " &Body=" & sBody, vbNullString, vbNullString, SW_SHOW
    DoEvents
End Sub

Private Sub GetWindowsVersion()
    WindowsVersion = GetVersion()
    WindowsVersion = WindowsVersion And &HFF&
End Sub

Public Function DirExists(path As String) As Boolean
  On Error Resume Next
    DirExists = CBool(GetAttr(path) And vbDirectory)
  On Error GoTo 0
End Function

Public Function FileExists(File As String) As Boolean
   On Error Resume Next
   FileExists = (Len(Dir$(File, vbNormal)) <> 0)
End Function
Public Sub setHighscore(Points As Long)
    If Points > val(glHighScore) Then
        glHighScore = Points
        If frmMain_Loaded Then frmMain.lblHighScore = Points
        SaveSetting AppExeName, cstrDefault, cstrHighscore, Points
    End If
End Sub


Public Sub ColumnAutoSize(ByRef ListView As ListView, _
                 Optional ByVal Column As Long = -1, _
                 Optional ByVal UseHeader As Boolean)
  Dim lngFlag As Long
 
  If (UseHeader) Then
    lngFlag = LVSCW_AUTOSIZE_USEHEADER
  Else
    lngFlag = LVSCW_AUTOSIZE
  End If
 
  If (Column <> -1) Then
    SendMessage ListView.hWnd, LVM_SETCOLUMNWIDTH, Column - 1, ByVal lngFlag
  Else
    With ListView
      For Column = 0 To .ColumnHeaders.count - 1
        SendMessage .hWnd, LVM_SETCOLUMNWIDTH, Column, ByVal lngFlag
      Next
    End With
  End If
End Sub

Public Sub GoHome()
Go2URL "www.playseven.com"
End Sub

Public Sub Go2URL(Url As String)
Dim rc As Integer
If Url <> vbNullString Then
    rc = ShellExecute(1, "open", Url, vbNullString, vbNullString, SW_SHOWNORMAL)
End If
End Sub
    
Public Sub MoveME(frm As Form)
Dim lResult As Long

Call ReleaseCapture
lResult = SendMessage(frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
End Sub



Public Sub setStrPlayerLevel()
Dim i As Integer

For i = ZERO To 5
    strPlayerLevel(i) = ModText(6 + i)
Next
strPlayerLevel(6) = ModText(11) & " Superior"

End Sub

'Public Function IsForm_Loaded(f As Form) As Boolean
'Dim Frmm As Form
'For Each Frmm In Forms
''Debug.Print Frmm.Name
'    If Frmm Is f Then
'        IsForm_Loaded = True
'    End If
'Next
'End Function

Public Sub DoSleep(Optional ByVal lMilliSec As Long = 0)
    'The DoSleep function allows other threads to have a time slice
    'and still keeps the main VB thread alive (since DPlay callbacks
    'run on separate threads outside of VB).
Dim ii As Long

    For ii = 1 To lMilliSec \ 10
        Sleep 10
        DoEvents
    Next

End Sub
