VERSION 5.00
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Begin VB.Form frmGetPlayer 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'Kein
   ClientHeight    =   3570
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5715
   Icon            =   "frmGetPlayer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   238
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows-Standard
   Begin GIF89LibCtl.Gif89a aniGifAvatar 
      Height          =   1125
      Left            =   4290
      OleObjectBlob   =   "frmGetPlayer.frx":08CA
      TabIndex        =   23
      ToolTipText     =   "www.avatarus.de"
      Top             =   480
      Width           =   1125
   End
   Begin VB.TextBox txtSpielerName 
      BackColor       =   &H00404040&
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
      Height          =   405
      Left            =   300
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   90
      Width           =   3735
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   5
      Left            =   300
      Picture         =   "frmGetPlayer.frx":090C
      Style           =   1  'Grafisch
      TabIndex        =   5
      Top             =   3100
      Width           =   405
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   4
      Left            =   300
      Picture         =   "frmGetPlayer.frx":2FAF
      Style           =   1  'Grafisch
      TabIndex        =   4
      Top             =   2610
      Width           =   405
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   2
      Left            =   300
      Picture         =   "frmGetPlayer.frx":53E9
      Style           =   1  'Grafisch
      TabIndex        =   3
      Top             =   1600
      Width           =   405
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   1
      Left            =   300
      Picture         =   "frmGetPlayer.frx":7769
      Style           =   1  'Grafisch
      TabIndex        =   2
      Top             =   1100
      Width           =   405
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   0
      Left            =   300
      Picture         =   "frmGetPlayer.frx":7B07
      Style           =   1  'Grafisch
      TabIndex        =   1
      Top             =   600
      Value           =   -1  'True
      Width           =   405
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00800000&
      ForeColor       =   &H000000FF&
      Height          =   405
      Index           =   3
      Left            =   300
      Picture         =   "frmGetPlayer.frx":8DEA
      Style           =   1  'Grafisch
      TabIndex        =   0
      Top             =   2100
      Width           =   405
   End
   Begin VB.Label lblMSG 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   945
      Left            =   4290
      TabIndex        =   22
      ToolTipText     =   "www.avatarus.de"
      Top             =   570
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fest Einfach
      Caption         =   "Avatar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4260
      TabIndex        =   21
      Top             =   210
      Width           =   1200
   End
   Begin VB.Image ImgAvatar 
      BorderStyle     =   1  'Fest Einfach
      Height          =   1200
      Left            =   4260
      Stretch         =   -1  'True
      ToolTipText     =   "www.avatarus.de"
      Top             =   450
      Width           =   1200
   End
   Begin VB.Label lblPlus 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4905
      TabIndex        =   19
      Top             =   1620
      Width           =   285
   End
   Begin VB.Label lblMinus 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   4485
      TabIndex        =   17
      Top             =   1620
      Width           =   285
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   5
      Left            =   810
      TabIndex        =   16
      Top             =   3120
      Width           =   3165
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   4
      Left            =   780
      TabIndex        =   15
      Top             =   2580
      Width           =   3165
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   3
      Left            =   810
      TabIndex        =   14
      Top             =   2100
      Width           =   3165
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   2
      Left            =   810
      TabIndex        =   13
      Top             =   1620
      Width           =   3165
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Index           =   1
      Left            =   810
      TabIndex        =   12
      Top             =   1125
      Width           =   3165
   End
   Begin VB.Label lblOption 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   525
      Index           =   0
      Left            =   810
      TabIndex        =   11
      Top             =   615
      Width           =   3165
   End
   Begin VB.Label lbl_Cancel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   810
      Left            =   4620
      TabIndex        =   8
      ToolTipText     =   "Exit"
      Top             =   2790
      Width           =   495
   End
   Begin VB.Label lbl_OK 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   585
      Left            =   4545
      TabIndex        =   7
      Top             =   2160
      Width           =   675
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   810
      Left            =   4680
      TabIndex        =   9
      ToolTipText     =   "Exit"
      Top             =   2850
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   585
      Left            =   4620
      TabIndex        =   10
      Top             =   2220
      Width           =   675
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4530
      TabIndex        =   18
      Top             =   1680
      Width           =   285
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackColor       =   &H8000000B&
      BackStyle       =   0  'Transparent
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4950
      TabIndex        =   20
      Top             =   1680
      Width           =   285
   End
End
Attribute VB_Name = "frmGetPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Private PicIndex As Long
Private ImageList() As String

Private Sub Form_Activate()
    Me.txtSpielerName.SelLength = Len(Me.txtSpielerName)
    Me.txtSpielerName.SetFocus
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = ZERO To 5
    Me.lblOption(i).Caption = strPlayerLevel(i)
Next
Me.txtSpielerName = "New Player"
Me.BackColor = 0
SetzFensterMittig Me
SetBackGround Me
makeRoundEdges Me
ImageList = GetAvatars
End Sub

Private Function GetAvatars() As String()
Dim ret() As String, pfad1 As String, Name1 As String
Dim i As Long

On Error GoTo ERRHand
pfad1 = App.path & cstrSubPathAvatars
Name1 = Dir$(pfad1)   ' Ersten Eintrag abrufen.

i = -1

Do While Name1 <> gstrNullstr   ' Schleife beginnen.
    
    ' Aktuelles und übergeordnetes Verzeichnis ignorieren.
    If Name1 <> gstrDot And Name1 <> ".." Then
        i = i + 1
        ReDim Preserve ret(0 To i)
        ret(i) = Name1
    End If
    Name1 = Dir$   'Nächsten Eintrag abrufen.
Loop
If i > -1 Then
    GetAvatars = ret
    Me.ImgAvatar.BorderStyle = 1
    ShowAvatar Me, pfad1 & ret(0)
Else
    Me.lblMsg.Caption = "No Avatars installed.." & vbCr & "click here 2 get some"
End If
Exit Function
ERRHand:
If ErrorBox("GetAvatars", Err) Then Resume Next
End Function


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = vbWhite
    Me.lbl_OK.ForeColor = vbWhite
    Me.lblMinus.ForeColor = vbWhite
    Me.lblPlus.ForeColor = vbWhite
End Sub

Private Sub ImgAvatar_Click()
    lblMSG_Click
End Sub


Private Sub lbl_Cancel_Click()
    Unload Me
    frmSplash.getSpieler
End Sub

Private Sub lbl_Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize - 3
End Sub

Private Sub lbl_Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = ROT
    Me.lbl_OK.ForeColor = vbWhite
    Me.lblMinus.ForeColor = vbWhite
    Me.lblPlus.ForeColor = vbWhite
End Sub

Private Sub lbl_Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize + 3
End Sub

Private Sub lbl_OK_Click()
Dim NeuSpieler As SpielerInfo
Dim i As Integer
NeuSpieler.SpielerName = CutApostroph(Me.txtSpielerName)
NeuSpieler.AvatarFileName = ImageList(PicIndex)

For i = Me.Option1.LBound To Me.Option1.ubound
    If Me.Option1(i) Then Exit For
Next
NeuSpieler.SpielerLevel = i

    If SetSpielerInDB(NeuSpieler, True) Then
        SaveSetting AppExeName, cstrDefault, cstrLastPlayer, NeuSpieler.SpielerName
        frmSplash.getSpieler
        Unload Me
    End If
End Sub

Private Function CutApostroph(str As String) As String
'apostroph eleminieren wg. schreiben in DB
Dim i As Integer
Dim ret As String, b As String
For i = 1 To Len(str)
    b = Mid$(str, i, 1)
    If b <> "'" Then ret = ret & b
Next
CutApostroph = ret
End Function

Private Sub lbl_OK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_OK.FontSize = Me.lbl_OK.FontSize - 3
End Sub

Private Sub lbl_OK_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_OK.ForeColor = ROT
    Me.lbl_Cancel.ForeColor = vbWhite
    Me.lblMinus.ForeColor = vbWhite
    Me.lblPlus.ForeColor = vbWhite
End Sub

Private Sub lbl_OK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    Me.lbl_OK.FontSize = Me.lbl_OK.FontSize + 3
End Sub

Private Sub lblMinus_Click()
'voriges Avatar anzeigen
On Error Resume Next
PicIndex = PicIndex - 1
If PicIndex < LBound(ImageList) Then
    PicIndex = UBound(ImageList)
Else
    ShowAvatar Me, App.path & cstrSubPathAvatars & ImageList(PicIndex)
End If

End Sub

Private Sub lblMSG_Click()
    Go2URL Me.ImgAvatar.ToolTipText
End Sub

Private Sub lblOption_Click(Index As Integer)
'Level als ausgewählt markieren
    Me.Option1(Index).Value = True
End Sub

Private Sub lblPlus_Click()
'Das nächste Avatarbild anzeigen
On Error Resume Next
PicIndex = PicIndex + 1
If PicIndex > UBound(ImageList) Then
    PicIndex = LBound(ImageList)
Else
    ShowAvatar Me, App.path & cstrSubPathAvatars & ImageList(PicIndex)
End If
End Sub


Private Sub lblMinus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinus.Font.Size = 16.25
End Sub

Private Sub lblMinus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinus.ForeColor = ROT
    Me.lblPlus.ForeColor = vbWhite
End Sub

Private Sub lblMinus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblMinus.Font.Size = 18.25
End Sub


Private Sub lblPlus_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Me.lblPlus.Font.Size = 16.25
End Sub

Private Sub lblPlus_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblPlus.ForeColor = ROT
    Me.lblMinus.ForeColor = vbWhite
End Sub

Private Sub lblPlus_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblPlus.Font.Size = 18.25
End Sub


Private Sub Option1_Click(Index As Integer)
'zugehöriges Labels auch markieren
Dim i As Integer
    For i = Me.lblOption.LBound To Me.lblOption.ubound
        Me.lblOption(i).ForeColor = vbWhite
    Next
    Me.lblOption(Index).ForeColor = vbRed
    
End Sub

Private Sub txtSpielerName_KeyDown(KeyCode As Integer, Shift As Integer)
'Wenn ENTER Dann Eingabe beenden
    If KeyCode = 13 Then
        Me.lbl_OK.FontSize = Me.lbl_OK.FontSize - 3
        lbl_OK_Click
        Me.lbl_OK.FontSize = Me.lbl_OK.FontSize + 3
    End If
End Sub
