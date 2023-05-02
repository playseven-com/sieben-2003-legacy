VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'Kein
   ClientHeight    =   2415
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5010
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1666.876
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   4704.65
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'Fenstermitte
   Visible         =   0   'False
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Height          =   420
      Left            =   90
      ScaleHeight     =   360
      ScaleWidth      =   4800
      TabIndex        =   5
      Top             =   660
      Width           =   4860
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   420
      Left            =   90
      Picture         =   "frmAbout.frx":08CA
      ScaleHeight     =   360
      ScaleWidth      =   4800
      TabIndex        =   3
      Top             =   780
      Visible         =   0   'False
      Width           =   4860
   End
   Begin VB.Label lbl_Cancel 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   4470
      TabIndex        =   6
      ToolTipText     =   "Exit"
      Top             =   1740
      Width           =   285
   End
   Begin VB.Label lblURL 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "www.playseven.com"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   810
      TabIndex        =   4
      Top             =   1980
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Zentriert
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2002 Milosz Weckowski mw@playseven.com"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   570
      TabIndex        =   2
      Top             =   1380
      Width           =   3015
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Height          =   240
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
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
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   420
      Width           =   3885
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Zentriert
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   495
      Left            =   4530
      TabIndex        =   7
      ToolTipText     =   "Exit"
      Top             =   1800
      Width           =   285
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Compare Text

Dim XPos, YPos As Integer 'Current X and Y positions of the "laser"
Dim Color As Long 'The color that the "laser" is currentlly drawing

Dim vLeft As Boolean, hLeft As Boolean

Private Enum LaserDrawModes
    LaserCorner
    PrinterScan
    WierdDraw
    WierdDrawSlow
End Enum


Private Sub Form_Activate()
    
    LaserDraw Picture1, Me.Picture2, Me.ScaleWidth, Me.ScaleHeight, vbRed, WierdDrawSlow

End Sub

Private Sub Form_Load()
    SetBackGround Me
    makeRoundEdges Me
    
    Me.Caption = "Info: " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MoveME Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.ForeColor = SCHWARZ
End Sub

Private Sub lbl_Cancel_Click()
    Unload Me
End Sub

Private Sub lbl_Cancel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize - 3
End Sub

Private Sub lbl_Cancel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.ForeColor = ROT
End Sub

Private Sub lbl_Cancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lbl_Cancel.FontSize = Me.lbl_Cancel.FontSize + 3
End Sub


Private Sub lblURL_Click()
GoHome
End Sub

Private Sub lblURL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.Font.Size = Me.lblURL.Font.Size - 2
End Sub

Private Sub lblURL_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.ForeColor = ROT
End Sub

Private Sub lblURL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.lblURL.Font.Size = Me.lblURL.Font.Size + 2
End Sub



'LaserDraw
'**** Description ***********
'Copies a picture from one picture box (or form) to another, with an animated "laser" effect
'**** Usage ***************
'LaserDraw PictureToDraw, DrawSurface, LaserOriginX, LaserOriginY, BackColor
'**** Inputs ***************
'0 PictureToDraw - Picturebox containing the picture to be copied
'0 DrawSurface - Picturebox or Form which the picture should be copied to
'0 LaserOriginX - Optional; The x coordinate of where the laser should come from.
'   Default is the width of the PictureToDraw picture box
'0 LaserOriginY - Optional; The y coordinate of where the laser should come from.
'   Default is the height of the PictureToDraw picture box
'0 BackColor - Optional; The background color of the DrawSurface
'   Default is the current background color of DrawSurface
'0 LaserDrawMode - Optional; The style of the laser draw
'   LaserCorner - Original mode, draws the picture, one line at a time, as if from a laser in a corner
'   PrinterScan - Draws the picture as if a printer were going along and drawing each dot
'   WierdDraw - Wierd draw mode, similar to PrinterScan. Try it for yourself :-)
'   Default is LaserCorner
'**** Outputs *****************
'None

Private Sub LaserDraw(PictureToDraw As PictureBox, DrawSurface As PictureBox, Optional LaserOriginX = -1, Optional LaserOriginY = -1, Optional BackColor As ColorConstants = -1, Optional LaserDrawMode As LaserDrawModes = LaserCorner)
    'Set up the DrawSurface picture box
        DrawSurface.ScaleMode = vbPixels 'Set the scale mode of the "canvas" to pixels
        If BackColor <> -1 Then 'Background color specified
            DrawSurface.BackColor = BackColor 'Set the background color of the "canvas" to the desired background color
        End If
    'Set up the PictureToDraw picture box
        PictureToDraw.ScaleMode = vbPixels 'Set the scale mode of the picturebox containing the picture to be drawn to pixels
        PictureToDraw.AutoRedraw = True 'Set the autoredraw property of the picturebox containing the picture to be drawn to true
        PictureToDraw.Visible = False 'Hide the picturebox containing the picture to be drawn
    'Set up the X and Y coordinates of the "laser"
        If LaserOriginX = -1 Then 'No X coordinate of the "laser" is specified
            LaserOriginX = PictureToDraw.ScaleWidth 'Set it to the width of the picturebox containing the picture to be drawn
        End If
        If LaserOriginY = -1 Then 'No Y coordinate of the "laser" is specified
            LaserOriginY = PictureToDraw.ScaleHeight 'Set it to the height of the picturebox containing the picture to be drawn
        End If
    'Start the "Laser" effect
        For XPos = ZERO To PictureToDraw.ScaleWidth 'Move the "laser" horizantally along the "canvas"
            DoEvents 'Allow input to be prosessed
            For YPos = ZERO To PictureToDraw.ScaleHeight 'Move the "laser" verticlly along the "canvas"
                Color = PictureToDraw.Point(XPos, YPos) 'Determine the color of the pixel to be drawn
                If LaserDrawMode = LaserCorner Then 'Normal Drawing
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, LaserOriginY), Color 'Draw a line from the origin coordinates to the coordinates of the pixel to be drawn
                ElseIf LaserDrawMode = PrinterScan Then '"Printer Scanning" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                    DrawSurface.Line (XPos + 1, YPos - 1)-(LaserOriginX, YPos - 1), BackColor 'Erase the last position of the "laser"
                    DoEvents 'Alow input to be prosessed
                ElseIf LaserDrawMode = WierdDrawSlow Then '"Weird Draw Slow" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                    DoEvents 'Alow input to be prosessed
                Else '"Wierd Draw" mode
                    DrawSurface.Line (XPos, YPos)-(LaserOriginX, YPos), Color 'Draw a straight line from the pixel to LaserOrginX
                End If
            Next
        Next
        DrawSurface.Picture = PictureToDraw.Picture
End Sub




