VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSplash 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fest Einfach
   ClientHeight    =   6885
   ClientLeft      =   225
   ClientTop       =   1380
   ClientWidth     =   8640
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   FillColor       =   &H00004000&
   FillStyle       =   0  'Ausgefüllt
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSplash.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   459
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   576
   Tag             =   "101"
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   88
      Left            =   690
      Top             =   6390
   End
   Begin VB.PictureBox scrollPic 
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2055
      Left            =   2520
      ScaleHeight     =   133
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   309
      TabIndex        =   11
      Top             =   2670
      Width           =   4695
   End
   Begin MSComctlLib.ImageCombo cmbSpieler 
      Height          =   420
      Left            =   3030
      TabIndex        =   10
      Top             =   720
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   741
      _Version        =   393216
      ForeColor       =   16777215
      BackColor       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Text            =   "ImageCombo1"
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'Kein
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   4470
      Picture         =   "frmSplash.frx":08CA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   9
      ToolTipText     =   "www.playseven.com"
      Top             =   2100
      Width           =   480
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   6270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483647
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   8421376
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":0DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":20EF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":249D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":482D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":6D04
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":914E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSplash.frx":B7FE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin GIF89LibCtl.Gif89a aniGifAvatar 
      Height          =   1155
      Left            =   870
      OleObjectBlob   =   "frmSplash.frx":DCF1
      TabIndex        =   12
      Top             =   510
      Width           =   1155
   End
   Begin VB.Shape Ellipse1 
      BorderColor     =   &H00FFFFFF&
      Height          =   11685
      Left            =   2400
      Shape           =   2  'Oval
      Top             =   1620
      Visible         =   0   'False
      Width           =   11685
   End
   Begin VB.Image ImgAvatar 
      BorderStyle     =   1  'Fest Einfach
      Height          =   1200
      Left            =   840
      Stretch         =   -1  'True
      ToolTipText     =   "www.avatarus.de"
      Top             =   480
      Width           =   1200
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   915
      Left            =   1980
      Shape           =   4  'Gerundetes Rechteck
      Top             =   6030
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Shape Nase 
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   2865
      Left            =   300
      Shape           =   4  'Gerundetes Rechteck
      Top             =   60
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Shape EckOben 
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   1875
      Left            =   390
      Shape           =   4  'Gerundetes Rechteck
      Top             =   360
      Visible         =   0   'False
      Width           =   8295
   End
   Begin VB.Shape Ellipse2 
      BorderColor     =   &H00FFFFFF&
      Height          =   11685
      Left            =   4860
      Shape           =   2  'Oval
      Top             =   1440
      Visible         =   0   'False
      Width           =   11685
   End
   Begin VB.Shape Balken 
      FillColor       =   &H00808080&
      FillStyle       =   7  'Diagonalkreuz
      Height          =   3645
      Left            =   1620
      Shape           =   2  'Oval
      Top             =   1950
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label lblLevel 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   3060
      TabIndex        =   5
      Top             =   1230
      Width           =   120
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2310
      TabIndex        =   4
      Top             =   1290
      Width           =   645
   End
   Begin VB.Label lblPlayer 
      BackStyle       =   0  'Transparent
      Caption         =   "Player:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   2310
      TabIndex        =   3
      Top             =   900
      Width           =   645
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   5
      Height          =   705
      Left            =   3420
      Shape           =   2  'Oval
      Top             =   5460
      Width           =   660
   End
   Begin VB.Label lblExit 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   675
      Left            =   3540
      TabIndex        =   2
      ToolTipText     =   "Exit"
      Top             =   5460
      Width           =   405
   End
   Begin VB.Label lblMultiPlayer 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplayer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   2580
      TabIndex        =   1
      Top             =   6270
      Width           =   2310
   End
   Begin VB.Label lblSinglePlayer 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Singleplayer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3300
      TabIndex        =   0
      Top             =   4770
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Single Player"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   3300
      TabIndex        =   8
      Top             =   4830
      Width           =   2685
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Multiplayer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   2610
      TabIndex        =   7
      Top             =   6330
      Width           =   2310
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   675
      Left            =   3630
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   5490
      Width           =   405
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      BorderWidth     =   5
      Height          =   705
      Left            =   3495
      Shape           =   2  'Oval
      Top             =   5490
      Width           =   630
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private Spieler() As SpielerInfo
Private myText() As String
Private ShowText() As String
Private TextHeighter As Integer
Private ShowIt As Integer
Private curline As Integer


Private Sub cmbSpieler_Click()
'Spieler asuwählen
    If Me.cmbSpieler.SelectedItem.Index = Me.cmbSpieler.ComboItems.count Then
        SpielerEingabe
    ElseIf Me.cmbSpieler.SelectedItem.Index = Me.cmbSpieler.ComboItems.count - 1 Then
        SpielerEingabe
    Else
        AktuellerSpieler = Spieler(Me.cmbSpieler.SelectedItem.Index)
        WriteLblLevel Me.lblLevel, AktuellerSpieler.SpielerLevel
        ShowAvatar Me, App.path & cstrSubPathAvatars & AktuellerSpieler.AvatarFileName
    End If
    Me.Picture2.SetFocus
End Sub


Private Sub cmbSpieler_GotFocus()
    DropDown Me.cmbSpieler
End Sub



Private Sub Form_Activate()
    RefreshAllCtls Me
End Sub

Private Sub Form_Load()
    
On Error GoTo ERRHand
    If InStr(1, Command$, "/test", vbTextCompare) Then Test = True
'    If InStr(1, Command$, "/bot", vbTextCompare) Then BOT = True
    'Test = True
    LoadObjectText Me.Name, myText()
    
    HideTitleBar Me
    'rundes fenster
    MakeFormRound
        
    Me.cmbSpieler.ComboItems.Clear
    Set Me.cmbSpieler.ImageList = Me.ImageList1
    
    SetBackGround Me
    #If Tiny Then
        Me.lblMultiPlayer.Visible = False
        Me.Label3.Visible = False
        Me.lblSinglePlayer.Move (Me.Width \ 2 \ Screen.TwipsPerPixelX) - Me.lblSinglePlayer.Width \ 2
        Me.Label4.Move (Me.Width \ 2 \ Screen.TwipsPerPixelX) - Me.lblSinglePlayer.Width \ 2
    #End If
    SetzFensterMittig frmSplash
    getSpieler

    GetText
    Me.Visible = True
    Me.ZOrder
    


Exit Sub
ERRHand:
If ErrorBox("frmSplash:Load", Err) Then Resume Next
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub



Private Sub MakeFormRound()
Dim dots(0 To 5) As POINTAPI
Dim i As Integer

Dim RegionA As Long
Dim RegionB As Long
Dim tmp As Long

'MakeTransparent Me

On Error GoTo ERRHand
    'Fuss machen
    RegionA = CreateEllipticRgn(Me.Ellipse1.Left, Me.Ellipse1.Top, Me.Ellipse1.Left + Me.Ellipse1.Width, Me.Ellipse1.Height)
    RegionB = CreateEllipticRgn(Me.Ellipse2.Left, Me.Ellipse2.Top, Me.Ellipse2.Left + Me.Ellipse2.Width, Me.Ellipse2.Height)
    CombineRgn RegionA, RegionA, RegionB, RGN_XOR
    
    RegionB = CreateRoundRectRgn(Me.EckOben.Left, Me.EckOben.Top, Me.EckOben.Left + Me.EckOben.Width, Me.EckOben.Height, Me.EckOben.Height \ 3, Me.EckOben.Height \ 3)
    CombineRgn RegionA, RegionA, RegionB, RGN_OR

    RegionB = CreateEllipticRgn(Me.Balken.Left, Me.Balken.Top, Me.Balken.Left + Me.Balken.Width, Me.Balken.Top + Me.Balken.Height)
'    RegionB = CreateRoundRectRgn(Me.Balken.Left, Me.Balken.Top, Me.Balken.Left + Me.Balken.Width, Me.Balken.Top + Me.Balken.Height, Me.Balken.Height / 6, Me.Balken.Height / 6)
    CombineRgn RegionA, RegionA, RegionB, RGN_OR
    
'    dots(0).X = Me.Line1(0).X1
'    dots(0).Y = Me.Line1(0).Y1
'
'    For i = 1 To 5
'        dots(i).X = Me.Line1(i - 1).X2
'        dots(i).Y = Me.Line1(i - 1).Y2
'    Next
    

'    RegionB = CreatePolygonRgn(dots(0), 6, 1)
    RegionB = CreateRoundRectRgn(Me.Nase.Left, Me.Nase.Top, Me.Nase.Left + Me.Nase.Width, Me.Nase.Height, Me.Nase.Height \ 6, Me.Nase.Height \ 6)
    CombineRgn RegionA, RegionA, RegionB, RGN_OR
    
    RegionB = CreateRoundRectRgn(Me.Shape3.Left, Me.Shape3.Top, Me.Shape3.Left + Me.Shape3.Width, Me.Shape3.Top + Me.Shape3.Height, Me.Shape3.Height \ 6, Me.Shape3.Height \ 6)
    CombineRgn RegionA, RegionA, RegionB, RGN_OR
    
    tmp = SetWindowRgn(Me.hWnd, RegionA, True)
Exit Sub
ERRHand:
If ErrorBox("MakeFormRound", Err) Then Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
'Farbänderung bei MouseOver
    Me.lblSinglePlayer.ForeColor = vbWhite
    Me.lblMultiPlayer.ForeColor = vbWhite
    Me.lblExit.ForeColor = vbWhite
    Me.Shape1.BorderColor = vbWhite
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Me.Timer1.Enabled = False
    AnimWindow Me, AW_HIDE + AW_BLEND
End Sub


Private Sub lblExit_Click()
    Unload Me
    CloseAll
    End
End Sub

Private Sub lblExit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.Font.Size = Me.lblExit.Font.Size - 3
End Sub

Private Sub lblExit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Me.lblExit.ForeColor <> ROT Then
        PlaySound ENDE
        Me.lblExit.ForeColor = ROT
        Me.Shape1.BorderColor = ROT
    End If
    Me.lblMultiPlayer.ForeColor = vbWhite
    Me.lblSinglePlayer.ForeColor = vbWhite
End Sub

Private Sub lblExit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblExit.Font.Size = Me.lblExit.Font.Size + 3
End Sub

Private Sub lblMultiPlayer_Click()
    
#If Not Tiny Then

    If Me.cmbSpieler.SelectedItem.Index = Me.cmbSpieler.ComboItems.count Then
        SpielerEingabe
        Exit Sub
    End If
    
    If Not SpielerAuswahl_OK() Then Exit Sub
    
    AktuellerSpieler = Spieler(Me.cmbSpieler.SelectedItem.Index)
    SetAppInfo
    If Not IsRegistered Then
        RegisterApp Me
        Exit Sub
    End If
    
    If Not BestehtVerbindung Then
        MsgBox myText(6), vbInformation, myText(7)
        Exit Sub
    End If
    
'    If AktuellerSpieler.SpielerLevel < ZERO Then
'        MsgBox myText(0) & vbCr & myText(1), vbInformation, myText(2)
'    Else

    PlaySound SMPlayer
    Playermodus = multiplayer
    SaveSetting AppExeName, cstrDefault, cstrLastPlayer, AktuellerSpieler.SpielerName
    Me.Hide
    Unload Me
    InitDPlay "frmSplash"


'    End If
#End If
End Sub

Private Sub lblMultiPlayer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblMultiPlayer.Font.Size = Me.lblMultiPlayer.Font.Size - 2
End Sub

Private Sub lblMultiPlayer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Farbänderung + Sound bei MoseOver
    If Me.lblMultiPlayer.ForeColor <> ROT Then
        PlaySound SMPlayerChoose
        Me.lblMultiPlayer.ForeColor = ROT
    End If
    Me.lblSinglePlayer.ForeColor = vbWhite
    Me.lblExit.ForeColor = vbWhite
End Sub

Private Sub lblMultiPlayer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblMultiPlayer.Font.Size = Me.lblMultiPlayer.Font.Size + 2
End Sub

Private Sub lblSinglePlayer_Click()
    If Me.cmbSpieler.SelectedItem.Index = Me.cmbSpieler.ComboItems.count - 1 Then
        SpielerEingabe
        Exit Sub
    End If
    If Not SpielerAuswahl_OK() Then Exit Sub
    
    AktuellerSpieler = Spieler(Me.cmbSpieler.SelectedItem.Index)
    
    SetAppInfo
    If Not IsRegisteredFree Then chkRegister
    
    
    PlaySound SMPlayer
    Playermodus = singleplayer
    Unload Me
    
    'Spieler speichern
    SaveSetting AppExeName, cstrDefault, cstrLastPlayer, AktuellerSpieler.SpielerName
    Gegner.SpielerName = cstrGegnerStandardName
    
    ' Feststellen, ob das Dialogfeld beim Start angezeigt werden soll
    ShowTipAtStartup = CBool(GetSetting(AppExeName, cstrOptions, cstrShowTips, 1))
    If ShowTipAtStartup Then
        frmTip.Show
    End If
    
    'spielfenster zeigen
    AnimWindow frmMain, AW_ACTIVATE + AW_BLEND
    frmMain.Refresh
    'spiel intialisieren
    frmMain.Init
    
End Sub

Private Sub lblSinglePlayer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

Me.lblSinglePlayer.Font.Size = Me.lblSinglePlayer.Font.Size - 2
End Sub

Private Sub lblSinglePlayer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

'Farbänderung + Sound bei MoseOver
    If Me.lblSinglePlayer.ForeColor <> ROT Then
        PlaySound SMPlayerChoose
        Me.lblSinglePlayer.ForeColor = ROT
    End If
    Me.lblMultiPlayer.ForeColor = vbWhite
    Me.lblExit.ForeColor = vbWhite
End Sub


Public Sub getSpieler()
Dim lS As String
Dim i As Long, ii As Long
On Error GoTo ERRHand

'AktuellenSpieler ermitteln und anzeigen

ii = -1
Me.cmbSpieler.ComboItems.Clear
lS = GetSetting(AppExeName, cstrDefault, cstrLastPlayer)
If GetSpielerFromDB(Spieler) Then
    For i = LBound(Spieler) To UBound(Spieler)
        Me.cmbSpieler.ComboItems.Add , , Spieler(i).SpielerName & " - (Level " & Spieler(i).SpielerLevel & ")", Spieler(i).SpielerLevel + 1
        If lS = Spieler(i).SpielerName Then
            ii = i
            WriteLblLevel Me.lblLevel, Spieler(i).SpielerLevel
            ShowAvatar Me, App.path & cstrSubPathAvatars & Spieler(i).AvatarFileName
        End If
    Next
    Me.cmbSpieler.ComboItems.Add , , "-----------------"
    Me.cmbSpieler.ComboItems.Add , , myText(3)
Else
    SpielerEingabe
    Exit Sub
End If

If ii >= ZERO Then
    Me.cmbSpieler.ComboItems(ii).Selected = True
Else
    Me.cmbSpieler.ComboItems(Me.cmbSpieler.ComboItems.count - 2).Selected = True
End If

Me.cmbSpieler.SelStart = 0
Me.cmbSpieler.SelLength = Len(Me.cmbSpieler.Text)

Exit Sub
ERRHand:
If ErrorBox("getSpieler", Err) Then Resume Next

End Sub


Private Sub SpielerEingabe()
    frmGetPlayer.Show 1
End Sub

Private Function SpielerAuswahl_OK() As Boolean
'überprüft ob ein gültige Auswahl in Combobox gewählt ist
    If Me.cmbSpieler.SelectedItem.Index = -1 Or Me.cmbSpieler.SelectedItem.Index = Me.cmbSpieler.ComboItems.count - 1 Then
        MsgBox myText(4), vbInformation
        SpielerAuswahl_OK = False
    Else
        SpielerAuswahl_OK = True
    End If
End Function

Private Sub lblSinglePlayer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lblSinglePlayer.Font.Size = Me.lblSinglePlayer.Font.Size + 2
End Sub

Private Sub Picture2_Click()
    GoHome
End Sub



Private Sub Timer1_Timer()
Dim ret As Integer
On Error Resume Next

If (ShowIt% = TextHeighter) Then    'play with this value for desired effect
    scrollPic.CurrentX = 0
    scrollPic.CurrentY = scrollPic.ScaleHeight - TextHeighter  'play with this also
    scrollPic.Print ShowText(curline)
    curline = curline + 1
    If curline > UBound(ShowText) Then curline = 0
    ShowIt% = 0
Else
    ret = BitBlt(scrollPic.hdc, 0, 0, scrollPic.ScaleWidth, scrollPic.ScaleHeight - 1, scrollPic.hdc, 0, 1, SRCCOPY)
    ShowIt% = ShowIt% + 1
End If
End Sub

Private Sub GetText()

Dim i As Long
Dim s() As String
Dim pos As Long, border As Long
Const StoryFileName As String = "\Story.txt"

On Error Resume Next

'FileNumber = FreeFile
border = Me.scrollPic.ScaleWidth \ Me.scrollPic.TextWidth("e")
TextHeighter = scrollPic.TextHeight("T") + 2

'Open App.path & StoryFileName For Input As #FileNumber
'If Err.Number = 53 Then
'    AgentSpeak App.path & StoryFileName & myText(5), True
'    Exit Sub
'End If

ReDim Preserve ShowText(0)
LoadObjectText "Story", s
Do While i <= UBound(s)
'    Line Input #FileNumber, s
    Do While Len(s(i)) > border
        pos = InStrRev(Left$(s(i), border), gstrSpace, border, vbTextCompare)
        If pos = ZERO Then pos = border
        ReDim Preserve ShowText(UBound(ShowText) + 1)
        ShowText(UBound(ShowText)) = Left$(s(i), pos)
        s(i) = Mid$(s(i), pos + 1)
    Loop
    ReDim Preserve ShowText(UBound(ShowText) + 2)
    ShowText(UBound(ShowText) - 1) = s(i)
    ShowText(UBound(ShowText)) = vbNullString
    i = i + 1
Loop
'Close #FileNumber   ' Close file.
Timer1.Enabled = True
Exit Sub
ERRHand:
If ErrorBox("gettext", Err) Then Resume Next

End Sub


