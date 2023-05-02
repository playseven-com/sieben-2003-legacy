Attribute VB_Name = "Transitions"
Option Explicit

Public Const WHITE_BRUSH = 0
Public Const BLACK_BRUSH = 4
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long

'Public Enum SideUD_Enum
'    sUp = 1
'    sDown = 2
'End Enum
'Public Enum SideLR_Enum
'    sLeft = 1
'    sRight = 2
'End Enum
'Public Enum Side_all
'    aUp = 1
'    aDown = 2
'    aLeft = 4
'    aRight = 8
'End Enum
'Public Enum Side_HV
'    VerticalSide = 1
'    HorizontalSide = 2
'End Enum
'Public Enum PushModeEnum
'    Pushing = 1
'    Hiding = 2
'    Moving = 3
'End Enum


'Public Const MS_DELAY = 1
'Public mblnRunning As Boolean, Ended As Boolean
'Public mlngTimer As Long
'Public lngSpeed As Long
'
'Private Type POINTAPI
'        x As Long
'        y As Long
'End Type
'
'Type SAFEARRAYBOUND
'    cElements As Long
'    lLbound As Long
'End Type
'
'Type SAFEARRAY2D
'    cDims As Integer
'    fFeatures As Integer
'    cbElements As Long
'    cLocks As Long
'    pvData As Long
'    Bounds(0 To 1) As SAFEARRAYBOUND
'End Type
'
'Type BITMAP
'    bmType As Long
'    bmWidth As Long
'    bmHeight As Long
'    bmWidthBytes As Long
'    bmPlanes As Integer
'    bmBitsPixel As Integer
'    bmBits As Long
'End Type
'
'Public Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
'Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
'Public Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long

'The user defined type 'SAFEARRAY2D' is used by Visual Basic for internal management multiple dimension arrays.
'The user defined type 'BITMAP' will keep some information about our picture.
'VarPtrArray' returns the memory address of an array.
'CopyMemory' copies blocks in memory from one position to another (extremely fast).
'GetObjectAPI' returns information about our bitmap, that will be written into our user defined type 'BITMAP'.




'Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
'Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
'Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
'Public Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
'
'Private Pic1() As Byte, Pic2() As Byte, Pic3() As Byte 'Our Memory
'Private SA1 As SAFEARRAY2D, SA2 As SAFEARRAY2D, SA3 As SAFEARRAY2D   'our Array Dimension
'Private Bmp1 As BITMAP, Bmp2 As BITMAP, Bmp3 As BITMAP 'Bitmap info
'Dim int_i As Long, int_j As Long

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As _
        Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As _
        Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal _
        xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) _
        As Long

Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As _
        Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As _
        Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal _
        xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As _
        Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As _
        Long

Const SRCCOPY = &HCC0020

Public Sub SetBackGround(ByRef frm As Form)
Dim i As Long, ii As Long, X As Long
Dim picName() As String, OLDVal As String
On Error GoTo ERRHand

    If frm.Name = "frmMain" Then
        
        For i = frmMain.menBackStandard.LBound To frmMain.menBackStandard.ubound
            'Debug.Print frmMain.menBackStandard(i).Caption
            If frmMain.menBackStandard(i).Checked Then Exit For
        Next
        'Bildnamen in Array schreiben
        ii = 0
        ReDim Preserve picName(ZERO To ii)
        picName(ii) = Dir(App.path & cstrSubPathBackGround & frmMain.menBackStandard(i).Caption & "\*.jpg")
        
        Do While picName(ii) <> gstrNullstr
            ii = ii + 1
            ReDim Preserve picName(ZERO To ii)
            picName(ii) = Dir
            'Debug.Print picName(ii)
        Loop
        'Index aus Array zufällig auswählen
        If ii > 1 Then
            Do
                X = Int((ii) * Rnd)
            Loop While frm.Tag = App.path & cstrSubPathBackGround & frmMain.menBackStandard(i).Caption & gstrDirSep & picName(X)
        End If
        
        OLDVal = frm.Tag
        frm.Tag = App.path & cstrSubPathBackGround & frmMain.menBackStandard(i).Caption & gstrDirSep & picName(X)
        frm.PictureDummy = LoadPicture(frm.Tag)
        
        If OLDVal = frm.Tag And UBound(picName) > 1 Then
            'nix machen
        Else
            If frm.WindowState = vbMaximized Then
                frmMain.PaintPicture frmMain.PictureDummy.Picture, ZERO, ZERO, frmMain.Width, frmMain.Height
                'RenderPicture frmMain.PictureDummy.Picture, frmMain.Picture.Handle, ZERO, ZERO, frmMain.Width, frmMain.Height, , , , , , vbTwips, vbTwips
                'SavePicture frmMain.Picture, "c:\temp\testpic.jpg"
                'RefreshAllCtls frm
            Else
                ChangePic frmMain, frmMain.PictureDummy, Int((6) * Rnd) + 1
            End If
            'Debug.Print frm.Tag
        End If
    Else
        Dim path As String
        path = App.path & cstrSubPathBackGround & frm.Name & "Scr2.jpg"
        If FileExists(path) Then
            frm.Tag = path
        Else
            frm.Tag = App.path & cstrSubPathBackGround & "Standard\Background2.jpg"
        End If
        frm.Picture = LoadPicture(frm.Tag)
    End If
    'Debug.Print frm.Tag

Exit Sub
ERRHand:
If ErrorBox("SetBackGround", Err) Then Resume Next

End Sub

Public Sub WaitTick(time_to_wait As Long)
Dim ctim As Long
Dim tim As Long

tim = GetTickCount&()
  Do
    DoEvents 'mach die anderen Dinge .....
  Loop Until GetTickCount&() > time_to_wait + tim&
End Sub



Public Sub ChangePic(ByRef Dest As Form, ByRef Source As PictureBox, effect As Integer)

Dim pixelwidth As Long
Dim pixelheight As Long
Dim u As Integer, chk As Integer, a As Integer, b As Integer, u1 As Integer, u2 As Integer, b0 As Integer, b1 As Integer
Dim chky As Integer, chkx As Integer, i As Integer, st As Integer, to1 As Integer, to2 As Integer
Dim stepsize As Integer, steptime As Long
Dim ScreenTX As Long, ScreenTY As Long

On Error GoTo ERRHand

Debug.Print "Effect =" & effect

ScreenTX = Screen.TwipsPerPixelX
ScreenTY = Screen.TwipsPerPixelY
pixelwidth = Dest.ScaleWidth \ ScreenTX
pixelheight = Dest.ScaleHeight \ ScreenTY


Select Case effect
    Case 1
        stepsize = 20
        steptime = 40
    Case 2, Is >= 4
        stepsize = 10
        steptime = 1
    Case 3
        stepsize = 20
        steptime = 20
End Select

With Dest
      
      Select Case effect
               
        Case 1 'vertikale Streifen, erst nach oben, dann nach unten
               chk% = (pixelheight \ stepsize)
               For a% = ZERO To stepsize Step 2
                 
                 u% = BitBlt(.hdc, 0, chk% * a%, pixelwidth, _
                             chk%, Source.hdc, 0, chk% * a%, _
                             SRCCOPY)
                 u% = DoEvents()
                 Call WaitTick(steptime)
               Next
               For a% = stepsize - 1 To 1 Step -2
                 
                 u% = BitBlt(.hdc, 0, chk% * a%, pixelwidth, chk%, _
                              Source.hdc, 0, chk% * a%, SRCCOPY)
                 u% = DoEvents()
                 Call WaitTick(steptime)
               Next a%
               
        Case 2, Is >= 4 'Zufallsblöcke
               Dim blks%()
               ReDim blks%(1, stepsize ^ 2)
               For a% = ZERO To stepsize - 1
                 For b% = ZERO To stepsize - 1
                   blks%(0, a% + b% * stepsize) = a%
                   blks%(1, a% + b% * stepsize) = b%
                 Next b%
               Next a%
               
               'mixen
               For a% = 1 To stepsize * 10
                 u1% = Int(Rnd(1) * (stepsize ^ 2))
                 u2% = Int(Rnd(1) * (stepsize ^ 2))
                 b0% = blks%(0, u1%): b1% = blks%(1, u1%)
                 blks%(0, u1%) = blks%(0, u2%)
                 blks%(1, u1%) = blks%(1, u2%)
                 blks%(0, u2%) = b0%
                 blks%(1, u2%) = b1%
               Next a%
               chkx% = (pixelwidth \ stepsize)
               chky% = (pixelheight \ stepsize)
               
               'Blöcke blitten
               For a% = ZERO To (stepsize ^ 2) - 1
                 
                 u% = BitBlt(.hdc, blks%(0, a%) * chkx%, _
                             blks%(1, a%) * chky%, chkx% + 1, _
                             chky% + 1, Source.hdc, blks%(0, a%) _
                             * chkx%, blks%(1, a%) * chky%, SRCCOPY)
                 
                 Call WaitTick(steptime)
               Next a%
               
        Case 3 'Scrollen von rechts oder links
               chk% = (pixelwidth \ stepsize)
               i% = Int(Rnd(1) * 2) 'Zufällige Richtung
               If i% < 1 Then
                 st% = 1: to1% = 0: to2% = stepsize
               Else
                 to1% = stepsize
                 to2% = 0
                 st% = -1
               End If
               
               For a% = to1% To to2% Step st%
                 
                 u% = BitBlt(.hdc, chk% * a%, 0, chk%, _
                             pixelheight, Source.hdc, chk% * a%, _
                             0, SRCCOPY)
                  
                 Call WaitTick(steptime)
               Next a%
'            Case 4
'                RandomLines Dest, Source, Int(Rnd + 1)
'            Case 5
'                Slide Dest, Dest, Source, Int(Rnd + 1)
'            Case 6
'                Wipe Dest, Source, Int(Rnd * 3 + 1)
        End Select
    End With
    Dest.Picture = Source.Picture
Exit Sub
ERRHand:
If ErrorBox("ChangePic", Err) Then Resume Next
End Sub

'Private Sub GetRGB(R As Long, G As Long, B As Long, Flag As Long)
'    Select Case Flag
'    Case 1 '1
'        R = Pic1(int_i + 2, int_j)
'        G = Pic1(int_i + 1, int_j)
'        B = Pic1(int_i, int_j)
'    Case 2 '2
'        R = Pic2(int_i + 2, int_j)
'        G = Pic2(int_i + 1, int_j)
'        B = Pic2(int_i, int_j)
'    Case 3 '3
'        R = Pic3(int_i + 2, int_j)
'        G = Pic3(int_i + 1, int_j)
'        B = Pic3(int_i, int_j)
'    End Select
'End Sub
'Private Sub CheckRGB(R As Long, G As Long, B As Long)
'        If R > 255 Then R = 255
'        If R < Zero Then R = 0
'        If G > 255 Then G = 255
'        If G < Zero Then G = 0
'        If B > 255 Then B = 255
'        If B < Zero Then B = 0
'End Sub
'
'Private Sub RandomLines(DestPic As Form, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional RefreshRate As Long = 0)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim X_Arr() As Long, Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim RRate As Long, Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = VerticalSide Then
'            Xleng = pxWidth
'        Else
'            Xleng = pxHeight
'        End If
'        ReDim X_Arr(Xleng)
'        'Create Table
'        For i = 1 To Xleng
'            X_Arr(i) = i
'        Next
'        'Mixing table!
'        For j = 1 To 3
'            For i = 1 To Xleng
'                r1 = CInt(Rnd * Xleng)
'                t = X_Arr(r1)
'                X_Arr(r1) = X_Arr(i)
'                X_Arr(i) = t
'            Next
'        Next
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    For RRate = Zero To RefreshRate
'                        If Cntr >= Xleng Then
'                            'we want to stop
'                            mblnRunning = False
'                            'Set new picture, you can use bitblt too.
'                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                            Exit Sub
'                        End If
'                        If Side = VerticalSide Then
'                            BitBlt DestPic.hdc, X_Arr(Cntr), 0, 1, pxHeight, NewPic.hdc, X_Arr(Cntr), 0, SRCCOPY
'                        Else
'                            BitBlt DestPic.hdc, 0, X_Arr(Cntr), pxWidth, 1, NewPic.hdc, 0, X_Arr(Cntr), SRCCOPY
'                        End If
'                        Cntr = Cntr + 1
'                    Next
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Slide(DestPic As Form, PrevPic As Form, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1)
''Not Completed : Left and Right
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim RRate As Long, Cntr As Long
'        Dim Xleng As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, PrevPic.hdc, 0, 0, SRCCOPY
'        If Side > 2 Then
'            Xleng = pxWidth \ 2
'        Else
'            Xleng = pxHeight \ 2
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    If Side = aUp Then
'                        'Prev Picture go up
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, Cntr, SRCCOPY
'                        'New pic go down
'                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - (2 * Cntr), SRCCOPY
'                    ElseIf Side = aDown Then
'                        'Prev pic go up
'                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, 0, SRCCOPY
'                        'New pic come down
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, Cntr, SRCCOPY
'                    ElseIf Side = aLeft Then
'                    ElseIf Side = aRight Then
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop loop and then restart another loop!
'                        mblnRunning = False
'                    End If
'                End If
'            DoEvents
'            Loop
'            mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr < Zero Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    If Side = aUp Then
'                        'Prev
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, PrevPic.hdc, 0, Cntr, SRCCOPY
'                        'New
'                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, NewPic.hdc, 0, 0, SRCCOPY
'                    ElseIf Side = aDown Then
'                        'Prev pic go up
'                        BitBlt DestPic.hdc, 0, Cntr, pxWidth, pxHeight - Cntr, PrevPic.hdc, 0, 0, SRCCOPY
'                        'New pic come down
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight - Cntr, NewPic.hdc, 0, Cntr, SRCCOPY
'                    ElseIf Side = aLeft Then
'                    ElseIf Side = aRight Then
'                    End If
'                    Cntr = Cntr - Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'
'Public Function IsReady() As Boolean
'    IsReady = Not mblnRunning
'End Function
'Public Sub Stretching(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Optional Side As SideLR_Enum = sLeft, Optional Step_all As Long = 1, Optional RefreshRate As Long = 0, Optional PushMode As PushModeEnum)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim RRate As Long, Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        Xleng = pxWidth
'        SetStretchBltMode DestPic.hdc, 4 'This is ColorOnColor(3)
'                                         'HalfTone (4) mode is better but slower and need to call SetBrushOrgEx
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    For RRate = Zero To RefreshRate
'                        If Cntr >= Xleng Then
'                            'we want to stop
'                            mblnRunning = False
'                            'Set new picture, you can use bitblt too.
'                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                            Exit Sub
'                        End If
'                        Select Case Side
'                        Case sLeft
'                            StretchBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            If PushMode = 1 Then
'                                'Push
'                                StretchBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            ElseIf PushMode = 3 Then
'                                'Move
'                                BitBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, SRCCOPY
'                            End If
'                        Case sRight
'                            StretchBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            If PushMode = 1 Then
'                                'Push
'                                StretchBlt DestPic.hdc, 0, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            ElseIf PushMode = 3 Then
'                                'Move
'                                BitBlt DestPic.hdc, 0, 0, pxWidth - Cntr, pxHeight, PrevPic.hdc, Cntr, 0, SRCCOPY
'                            End If
'                        End Select
'                        Cntr = Cntr + Step_all
'                    Next
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'Public Sub Wipe(DestPic As Form, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side < aLeft Then
'            Xleng = pxHeight
'        Else
'            Xleng = pxWidth
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    Select Case Side
'                    Case aUp
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, SRCCOPY
'                    Case aDown
'                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - Cntr, SRCCOPY
'                    Case aLeft
'                        BitBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                    Case aRight
'                        BitBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, pxWidth - Cntr, 0, SRCCOPY
'                    End Select
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Wipe_In(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional Steps As Long = 1)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = VerticalSide Then
'            Xleng = pxHeight / 2
'        Else
'            Xleng = pxWidth / 2
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    If Side = VerticalSide Then
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, SRCCOPY
'                        BitBlt DestPic.hdc, 0, pxHeight - Cntr, pxWidth, Cntr, NewPic.hdc, 0, pxHeight - Cntr, SRCCOPY
'                    ElseIf Side = HorizontalSide Then
'                        BitBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        BitBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, pxWidth - Cntr, 0, SRCCOPY
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'Public Sub Wipe_Out(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = VerticalSide, Optional Steps As Long = 1)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = VerticalSide Then
'            Xleng = pxHeight / 2
'        Else
'            Xleng = pxWidth / 2
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    If Side = VerticalSide Then
'                        BitBlt DestPic.hdc, 0, Xleng - Cntr, pxWidth, Cntr, NewPic.hdc, 0, Xleng - Cntr, SRCCOPY
'                        BitBlt DestPic.hdc, 0, Xleng, pxWidth, Cntr, NewPic.hdc, 0, Xleng, SRCCOPY
'                    ElseIf Side = HorizontalSide Then
'                        BitBlt DestPic.hdc, Xleng - Cntr, 0, Cntr, pxHeight, NewPic.hdc, Xleng - Cntr, 0, SRCCOPY
'                        BitBlt DestPic.hdc, Xleng, 0, Cntr, pxHeight, NewPic.hdc, Xleng, 0, SRCCOPY
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Bars_Draw(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Steps As Long = 1, Optional BarSize As Long = 10, Optional FirstBar_RightToLeft As Boolean = True)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long, OthXLeng As Long
'        Dim tBars As Long, bltside As Boolean
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = HorizontalSide Then
'            Xleng = pxWidth
'            OthXLeng = pxHeight
'        Else
'            Xleng = pxHeight
'            OthXLeng = pxWidth
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    bltside = FirstBar_RightToLeft
'                    If Side = VerticalSide Then
'                        For tBars = Zero To OthXLeng Step BarSize
'                            If bltside Then
'                                BitBlt DestPic.hdc, tBars, 0, BarSize, Cntr, NewPic.hdc, tBars, 0, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, tBars, pxHeight - Cntr, BarSize, Cntr, NewPic.hdc, tBars, pxHeight - Cntr, SRCCOPY
'                            End If
'                            bltside = Not bltside
'                        Next
'                    Else
'                        For tBars = Zero To OthXLeng Step BarSize
'                            If bltside Then
'                                BitBlt DestPic.hdc, 0, tBars, Cntr, BarSize, NewPic.hdc, 0, tBars, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, pxWidth - Cntr, tBars, Cntr, BarSize, NewPic.hdc, pxWidth - Cntr, tBars, SRCCOPY
'                            End If
'                            bltside = Not bltside
'                        Next
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Bars_Move(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Steps As Long = 1, Optional BarSize As Long = 10, Optional FirstBar_RightToLeft As Boolean = True)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long, OthXLeng As Long
'        Dim tBars As Long, bltside As Boolean
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = HorizontalSide Then
'            Xleng = pxWidth
'            OthXLeng = pxHeight
'        Else
'            Xleng = pxHeight
'            OthXLeng = pxWidth
'        End If
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    bltside = FirstBar_RightToLeft
'                    If Side = VerticalSide Then
'                        For tBars = Zero To OthXLeng Step BarSize
'                            If bltside Then
'                                BitBlt DestPic.hdc, tBars, 0, BarSize, Cntr, NewPic.hdc, tBars, pxHeight - Cntr, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, tBars, pxHeight - Cntr, BarSize, Cntr, NewPic.hdc, tBars, 0, SRCCOPY
'                            End If
'                            bltside = Not bltside
'                        Next
'                    Else
'                        For tBars = Zero To OthXLeng Step BarSize
'                            If bltside Then
'                                BitBlt DestPic.hdc, 0, tBars, Cntr, BarSize, NewPic.hdc, pxWidth - Cntr, tBars, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, pxWidth - Cntr, tBars, Cntr, BarSize, NewPic.hdc, 0, tBars, SRCCOPY
'                            End If
'                            bltside = Not bltside
'                        Next
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Bars_Wipe(DestPic As PictureBox, NewPic As PictureBox, Optional Side As Side_all = aUp, Optional Steps As Long = 1, Optional BarSize As Long = 10)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim tBars As Long
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= BarSize Then
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    If Side < aLeft Then
'                        For tBars = Zero To pxHeight Step BarSize
'                            If Side = aUp Then
'                                BitBlt DestPic.hdc, 0, tBars, pxWidth, Cntr, NewPic.hdc, 0, tBars, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, 0, tBars + BarSize - Cntr, pxWidth, Cntr, NewPic.hdc, 0, tBars + BarSize - Cntr, SRCCOPY
'                            End If
'                        Next
'                    Else
'                        For tBars = Zero To pxWidth Step BarSize
'                            If Side = aLeft Then
'                                BitBlt DestPic.hdc, tBars, 0, Cntr, pxHeight, NewPic.hdc, tBars, 0, SRCCOPY
'                            Else
'                                BitBlt DestPic.hdc, tBars + BarSize - Cntr, 0, Cntr, pxHeight, NewPic.hdc, tBars + BarSize - Cntr, 0, SRCCOPY
'                            End If
'                        Next
'                    End If
'                    Cntr = Cntr + Steps
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub Stretching_Wipe_In(DestPic As PictureBox, PrevPic As PictureBox, NewPic As PictureBox, Optional Side As Side_HV = HorizontalSide, Optional Step_all As Long = 1, Optional RefreshRate As Long = 0, Optional PushMode As PushModeEnum)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Long, i As Long, j As Long, t As Long
'        Dim RRate As Long, Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        If Side = HorizontalSide Then
'            Xleng = pxWidth \ 2
'        Else
'            Xleng = pxHeight \ 2
'        End If
'        SetStretchBltMode DestPic.hdc, 4
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    For RRate = Zero To RefreshRate
'                        If Cntr >= Xleng Then
'                            'we want to stop
'                            mblnRunning = False
'                            'Set new picture, you can use bitblt too.
'                            BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                            Exit Sub
'                        End If
'                        If Side = HorizontalSide Then
'                            StretchBlt DestPic.hdc, 0, 0, Cntr, pxHeight, NewPic.hdc, 0, 0, Xleng, pxHeight, SRCCOPY
'                            StretchBlt DestPic.hdc, pxWidth - Cntr, 0, Cntr, pxHeight, NewPic.hdc, Xleng, 0, Xleng, pxHeight, SRCCOPY
'                            If PushMode = Pushing Then
'                                StretchBlt DestPic.hdc, Cntr, 0, pxWidth - Cntr - Cntr, pxHeight, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            End If
'                        Else
'                            StretchBlt DestPic.hdc, 0, 0, pxWidth, Cntr, NewPic.hdc, 0, 0, pxWidth, Xleng, SRCCOPY
'                            StretchBlt DestPic.hdc, 0, pxHeight - Cntr - 1, pxWidth, Cntr, NewPic.hdc, 0, Xleng, pxWidth, Xleng, SRCCOPY
'                            If PushMode = Pushing Then
'                                StretchBlt DestPic.hdc, 0, Cntr, pxWidth, pxWidth - Cntr - Cntr, PrevPic.hdc, 0, 0, pxWidth, pxHeight, SRCCOPY
'                            End If
'                        End If
'                        Cntr = Cntr + Step_all
'                    Next
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If
'    Ended = True
'End Sub
'
'Public Sub MaskEffect(DestPic As PictureBox, NewPic As PictureBox, MaskIndex As Integer, FormHdc As Long, Optional Steps As Long = 10)
'    If IsReady Then
'        Ended = False
'        Dim pxWidth As Long, pxHeight As Long
'        Dim ScreenTX As Long, ScreenTY As Long
'        Dim Xleng As Long
'        Dim r1 As Double, i As Long, j As Long, t As Long
'        Dim Cntr As Long
'
'        ScreenTX = Screen.TwipsPerPixelX
'        ScreenTY = Screen.TwipsPerPixelY
'        pxWidth = DestPic.ScaleWidth \ ScreenTX
'        pxHeight = DestPic.ScaleHeight \ ScreenTY
'
'        Dim T1_hdc As Long, T2_hdc As Long
'        Dim T1_bmp As Long, T2_bmp As Long
'        Dim RetPnt As POINTAPI
'
'        T1_hdc = CreateCompatibleDC(DestPic.hdc)
'        T1_bmp = CreateCompatibleBitmap(DestPic.hdc, pxWidth + 2, pxHeight + 2)
'        SelectObject T1_hdc, T1_bmp
'        'Clear Pic
'        For i = -1 To pxWidth
'        MoveToEx T1_hdc, i, -1, RetPnt
'        LineTo T1_hdc, i, pxHeight
'        Next
'
'        T2_hdc = CreateCompatibleDC(DestPic.hdc)
'        T2_bmp = CreateCompatibleBitmap(DestPic.hdc, pxWidth + 2, pxHeight + 2)
'        SelectObject T2_hdc, T2_bmp
'
'        SelectObject T1_hdc, GetStockObject(6) 'White pen
'        Dim MaxDPR As Long
'        Select Case MaskIndex
'        Case 1
'            Xleng = (2 * pxWidth) + (2 * pxHeight)
'        Case 2
'            Xleng = CLng(Sqr((pxWidth / 2) ^ 2 + (pxHeight / 2) ^ 2))
'        Case 3
'            Xleng = pxWidth + pxHeight
'        Case 4
'            Xleng = pxWidth
'        Case 5
'            Xleng = pxWidth
'        Case 6
'            Xleng = pxWidth
'        End Select
'        Dim Now_side As Integer, Cntr2 As Long, DPR As Long
'        mblnRunning = True
'            'Loop starts here
'            Do While mblnRunning
'                If mlngTimer + lngSpeed <= GetTickCount() Then
'                    'BitBlting
'                    If Cntr >= Xleng Or Now_side = -1 Then
'                        'We must Delete temporary hDC
'                        DeleteDC T1_hdc
'                        DeleteDC T2_hdc
'                        DeleteObject T1_bmp
'                        DeleteObject T2_bmp
'                        'we want to stop
'                        mblnRunning = False
'                        'Set new picture, you can use bitblt too.
'                        BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCCOPY
'                        Exit Sub
'                    End If
'                    Select Case MaskIndex
'                        Case 1
'                        For DPR = 1 To Steps
'                            'Radial Wipe
'                            MoveToEx T1_hdc, pxWidth / 2, pxHeight / 2, RetPnt
'                            Select Case Now_side
'                            Case 0
'                                LineTo T1_hdc, Cntr2, -1
'                                If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 1
'                            Case 1
'                                LineTo T1_hdc, pxWidth, Cntr2
'                                If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = 2
'                            Case 2
'                                LineTo T1_hdc, pxWidth - Cntr2, pxHeight
'                                If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 3
'                            Case 3
'                                LineTo T1_hdc, -1, pxHeight - Cntr2
'                                If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = -1
'                            End Select
'                            Cntr2 = Cntr2 + 1
'                        Next
'                            '*****************************************
'                        Case 2
'                            ' Circle Wipe
'                            Cntr = Cntr - 1
'                            For DPR = 1 To Steps
'                            Ellipse T1_hdc, pxWidth / 2 - Cntr, pxHeight / 2 - Cntr, pxWidth / 2 + Cntr, pxHeight / 2 + Cntr
'                            Cntr = Cntr + 1
'                            Next
'                        Case 3
'                            'Side Radial Wipe
'                            For DPR = 1 To Steps
'                                MoveToEx T1_hdc, 0, 0, RetPnt
'                                If Now_side = Zero Then
'                                    If Cntr2 > pxWidth Then Cntr2 = 0: Now_side = 1
'                                    LineTo T1_hdc, Cntr2, pxHeight
'                                ElseIf Now_side = 1 Then
'                                    If Cntr2 > pxHeight Then Cntr2 = 0: Now_side = -1
'                                    LineTo T1_hdc, pxWidth, pxHeight - Cntr2
'                                End If
'                                Cntr2 = Cntr2 + 1
'                            Next
'                        Case 4
'                            ' Triangles Wipe
'                            For DPR = 1 To Steps
'                                If Now_side = Zero Then
'                                    Cntr2 = Cntr2 + 1
'                                    If Cntr2 = pxWidth Then Now_side = -1
'                                    t = ((Cntr2 / pxWidth) * pxHeight) + 1
'                                    MoveToEx T1_hdc, Cntr2, 0, RetPnt
'                                    LineTo T1_hdc, Cntr2, t
'                                    MoveToEx T1_hdc, pxWidth - Cntr2, pxHeight, RetPnt
'                                    LineTo T1_hdc, pxWidth - Cntr2, pxHeight - t
'                                End If
'                            Next
'                        Case 5
'                            For DPR = 1 To Steps
'                                If Now_side = Zero Then
'                                    If Cntr2 = Xleng Then Now_side = -1
'                                    t = (Cntr2 / pxWidth) * pxHeight
'                                    MoveToEx T1_hdc, pxWidth - Cntr2, -1, RetPnt
'                                    LineTo T1_hdc, -1, pxHeight - t
'
'                                    MoveToEx T1_hdc, pxWidth, t, RetPnt
'                                    LineTo T1_hdc, Cntr2, pxHeight
'                                    Cntr2 = Cntr2 + 1
'                                End If
'                            Next
'                        Case 6
'                            For DPR = 1 To Steps
'                                If Now_side = Zero Then
'                                    If Cntr2 = Xleng Then Now_side = -1
'                                    MoveToEx T1_hdc, 0, 0, RetPnt
'                                    LineTo T1_hdc, pxWidth - Cntr2, pxHeight
'                                    MoveToEx T1_hdc, pxWidth, pxHeight, RetPnt
'                                    LineTo T1_hdc, Cntr2, -1
'                                    Cntr2 = Cntr2 + 1
'                                End If
'                            Next
'                    End Select
'                    BitBlt T2_hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, NOTSRCCOPY
'                    BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T2_hdc, 0, 0, SRCAND
'                    BitBlt T1_hdc, 0, 0, pxWidth, pxHeight, NewPic.hdc, 0, 0, SRCAND
'                    BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, SRCPAINT
'                    'BitBlt DestPic.hdc, 0, 0, pxWidth, pxHeight, T1_hdc, 0, 0, SRCCOPY
'
'                    Cntr = Cntr + 1
'                    'Refresh Picture
'                    DestPic.Refresh
'                    'Refresh Timer
'                    mlngTimer = GetTickCount()  'Reset the timer variable
'                End If
'            DoEvents
'            Loop
'        mblnRunning = False
'    End If 'If IsReady
'    Ended = True
'End Sub
'Public Sub SwapPictures(Picture1 As PictureBox, Picture2 As PictureBox, Picture3 As PictureBox)
'Dim UB As Long, UB2 As Long
'
'        GetObjectAPI Picture1.Picture, Len(Bmp1), Bmp1
'        GetObjectAPI Picture2.Picture, Len(Bmp2), Bmp2
'        GetObjectAPI Picture3.Picture, Len(Bmp3), Bmp3
'
'        With SA1
'            .cbElements = 1
'            .cDims = 2
'            .Bounds(0).lLbound = 0
'            .Bounds(0).cElements = Bmp1.bmHeight
'            .Bounds(1).lLbound = 0
'            .Bounds(1).cElements = Bmp1.bmWidthBytes
'            .pvData = Bmp1.bmBits
'        End With
'        With SA2
'            .cbElements = 1
'            .cDims = 2
'            .Bounds(0).lLbound = 0
'            .Bounds(0).cElements = Bmp2.bmHeight
'            .Bounds(1).lLbound = 0
'            .Bounds(1).cElements = Bmp2.bmWidthBytes
'            .pvData = Bmp2.bmBits
'        End With
'        With SA3
'            .cbElements = 1
'            .cDims = 2
'            .Bounds(0).lLbound = 0
'            .Bounds(0).cElements = Bmp3.bmHeight
'            .Bounds(1).lLbound = 0
'            .Bounds(1).cElements = Bmp3.bmWidthBytes
'            .pvData = Bmp3.bmBits
'        End With
'
'        CopyMemory ByVal VarPtrArray(Pic1), VarPtr(SA1), 4
'        CopyMemory ByVal VarPtrArray(Pic2), VarPtr(SA2), 4
'        CopyMemory ByVal VarPtrArray(Pic3), VarPtr(SA3), 4
'        UB = UBound(Pic1, 1) + 1
'        UB2 = UBound(Pic1, 2) + 1
'
'        CopyMemory Pic1(0, 0), Pic2(0, 0), UB * UB2
'        CopyMemory Pic2(0, 0), Pic3(0, 0), UB * UB2
'        CopyMemory Pic3(0, 0), Pic1(0, 0), UB * UB2
'        Set Picture1.Picture = Picture1.Picture
'        Set Picture2.Picture = Picture2.Picture
'        Set Picture3.Picture = Picture3.Picture
'        CopyMemory ByVal VarPtrArray(Pic1), 0&, 4
'        CopyMemory ByVal VarPtrArray(Pic2), 0&, 4
'        CopyMemory ByVal VarPtrArray(Pic3), 0&, 4
'End Sub
'
'Public Function PictureToBytes(Picture As StdPicture, Optional Key As String = "Bild") As Byte()
'
'  With New PropertyBag
'    .WriteProperty Key, Picture
'    PictureToBytes = .Contents
'  End With
'End Function
'
'Public Function BytesToPicture(Bytes() As Byte, Optional Key As String = "Bild") As StdPicture
'
'  With New PropertyBag
'    .Contents = Bytes
'    Set BytesToPicture = .ReadProperty(Key)
'  End With
'End Function


