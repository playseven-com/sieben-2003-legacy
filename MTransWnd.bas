Attribute VB_Name = "MTranslucentWnd"
' *************************************************************************
'  Copyright ©2000 Karl E. Peterson
'  All Rights Reserved, http://www.mvps.org/vb
' *************************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code, non-compiled, without prior written consent.
' *************************************************************************
Option Explicit

' BOOL SetLayeredWindowAttributes(
'   HWND hwnd,       // handle to the layered window
'   COLORREF crKey,  // specifies the color key
'   BYTE bAlpha,     // value for the blend function
'   DWORD dwFlags    // action
' );
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Long, ByVal dwFlags As Long) As Long
Private Const LWA_COLORKEY = &H1&
Private Const LWA_ALPHA = &H2&

' Style setting APIs
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000

' Win32 APIs to determine OS information.
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Type OSVERSIONINFO
   dwOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128
End Type
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Used to determine parentage.
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long

'#############################################
'Api's für runde und transparente  fenster
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

Public Const RGN_OR = 2
Public Const RGN_XOR = 3

Public Type POINTAPI
   X As Long
   Y As Long
End Type

Public Sub makeRoundEdges(frm As Form)
Dim RegionA As Long
Dim tmp As Long

On Error Resume Next

Const EckeWeite = 15
frm.ScaleMode = vbPixels
RegionA = CreateRoundRectRgn(ZERO, ZERO, frm.ScaleWidth, frm.ScaleHeight, EckeWeite, EckeWeite)
Debug.Print frm.ScaleHeight
tmp = SetWindowRgn(frm.hWnd, RegionA, True)
frm.ScaleMode = vbTwips
frm.Refresh
End Sub

'Public Function ClearWindowTranslucency(ByVal hWnd As Long) As Boolean
'   Dim nStyle As Long
'   If IsWin2000 Then
'      ' Only work with top-level.
'      hWnd = GetTopLevel(hWnd)
'      ' Set translucency to fully
'      ' opaque (255).
'      Call SetLayeredWindowAttributes(hWnd, 0, 255&, LWA_ALPHA)
'      ' Clear exstyle bit.
'      nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) And Not WS_EX_LAYERED
'      ClearWindowTranslucency = CBool(SetWindowLong(hWnd, GWL_EXSTYLE, nStyle))
'   End If
'End Function
'
'Public Function SetWindowTranslucency(ByVal hWnd As Long, ByVal Alpha As Byte) As Boolean
'    Dim nStyle As Long
''   If IsWin2000 Then
'    ' Only work with top-level.
'    hWnd = GetTopLevel(hWnd)
'    ' Set exstyle bit.
'    nStyle = GetWindowLong(hWnd, GWL_EXSTYLE) Or WS_EX_LAYERED
'    If SetWindowLong(hWnd, GWL_EXSTYLE, nStyle) Then
'       ' Set window translucency to
'       ' requested Alpha value.
'       SetWindowTranslucency = CBool(SetLayeredWindowAttributes(hWnd, 0, CLng(Alpha), LWA_ALPHA))
'    End If
''   End If
'End Function

Public Function IsWin2000() As Boolean
   Dim os As OSVERSIONINFO
   ' Layered windows are only available in
   ' Windows 2000. This function shouldn't
   ' be called often, so check on demand.
   os.dwOSVersionInfoSize = Len(os)
   Call GetVersionEx(os)
   If os.dwPlatformId = VER_PLATFORM_WIN32_NT Then
      IsWin2000 = (os.dwMajorVersion >= 5)
   End If
End Function

'Private Function GetTopLevel(ByVal hChild As Long) As Long
'   Dim hWnd As Long
'
'   ' Read parent chain up to highest visible.
'   hWnd = hChild
'   Do While IsWindowVisible(GetParent(hWnd))
'      hWnd = GetParent(hChild)
'      hChild = hWnd
'   Loop
'   GetTopLevel = hWnd
'End Function


'Public Sub MakeTransparent(TransForm As Form)
'   Dim ErrorTest As Double
'   On Error Resume Next
'   Dim Regn As Long
'   Dim TmpRegn As Long
'   Dim TmpControl As Control
'   Dim LinePoints(4) As POINTAPI
'
'   'Weil die API mit Pixeln arbeitet, die Maßeinheit auf Pixel setzen
'   TransForm.ScaleMode = 3
'   'Die Form darf keinen Rand haben, deshalb wird der Rand erstmal geprüft
'   If TransForm.BorderStyle <> 0 Then MsgBox "Change the borderstyle to 0!", vbCritical, "ACK!": End
'   'Macht alles unsichtbar
'   Regn = CreateRectRgn(0, 0, 0, 0)
'   'Für jedes Steuerelement der Form
'   For Each TmpControl In TransForm
'      'Wenn das Steuerelement eine Linie ist
'      If TypeOf TmpControl Is Line Then
'         If Abs((TmpControl.Y1 - TmpControl.Y2) / (TmpControl.X1 - TmpControl.X2)) > 1 Then
'            LinePoints(0).X = TmpControl.X1 - 1
'            LinePoints(0).Y = TmpControl.Y1
'            LinePoints(1).X = TmpControl.X2 - 1
'            LinePoints(1).Y = TmpControl.Y2
'            LinePoints(2).X = TmpControl.X2 + 1
'            LinePoints(2).Y = TmpControl.Y2
'            LinePoints(3).X = TmpControl.X1 + 1
'            LinePoints(3).Y = TmpControl.Y1
'         Else
'            LinePoints(0).X = TmpControl.X1
'            LinePoints(0).Y = TmpControl.Y1 - 1
'            LinePoints(1).X = TmpControl.X2
'            LinePoints(1).Y = TmpControl.Y2 - 1
'            LinePoints(2).X = TmpControl.X2
'            LinePoints(2).Y = TmpControl.Y2 + 1
'            LinePoints(3).X = TmpControl.X1
'            LinePoints(3).Y = TmpControl.Y1 + 1
'         End If
'         TmpRegn = CreatePolygonRgn(LinePoints(0), 4, 1)
'      'Wenn das Steuerelement eine Form (Shape) ist
'      ElseIf TypeOf TmpControl Is Shape Then
'         'Typ der Form
'         If TmpControl.Shape = 0 Then
'            'Es ist ein Rechteck
'            TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
'         ElseIf TmpControl.Shape = 1 Then
'            'Es ist ein Quadrat
'            If TmpControl.Width < TmpControl.Height Then
'               TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width)
'            Else
'               TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height)
'            End If
'         ElseIf TmpControl.Shape = 2 Then
'            'Es ist eine Ellipse
'            TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
'         ElseIf TmpControl.Shape = 3 Then
'            'Es ist ein Kreis
'            If TmpControl.Width < TmpControl.Height Then
'               TmpRegn = CreateEllipticRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 0.5)
'            Else
'               TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 0.5, TmpControl.Top + TmpControl.Height + 0.5)
'            End If
'         ElseIf TmpControl.Shape = 4 Then
'            'Es ist ein Rechteck mit abgerundeten Ecken
'            If TmpControl.Width > TmpControl.Height Then
'               TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
'            Else
'               TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Width / 4, TmpControl.Width / 4)
'            End If
'         ElseIf TmpControl.Shape = 5 Then
'            'Es ist ein Quadrat mit abgerundeten Ecken
'            If TmpControl.Width > TmpControl.Height Then
'               TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2, TmpControl.Top, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height + 1, TmpControl.Top + TmpControl.Height + 1, TmpControl.Height / 4, TmpControl.Height / 4)
'            Else
'               TmpRegn = CreateRoundRectRgn(TmpControl.Left, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2, TmpControl.Left + TmpControl.Width + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width + 1, TmpControl.Width / 4, TmpControl.Width / 4)
'            End If
'         End If
'         If TmpControl.BackStyle = 0 Then
'            'Kombiniert die Regionen im Speicher und erstellt eine neue
'            CombineRgn Regn, Regn, TmpRegn, RGN_XOR
'
'            If TmpControl.Shape = 0 Then
'               'Rechteck
'               TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + TmpControl.Height - 1)
'            ElseIf TmpControl.Shape = 1 Then
'               'Quadrat
'               If TmpControl.Width < TmpControl.Height Then
'                  TmpRegn = CreateRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 1)
'               Else
'                  TmpRegn = CreateRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 1, TmpControl.Top + TmpControl.Height - 1)
'               End If
'            ElseIf TmpControl.Shape = 2 Then
'               'Ellipse
'               TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
'            ElseIf TmpControl.Shape = 3 Then
'               'Kreis
'               If TmpControl.Width < TmpControl.Height Then
'                  TmpRegn = CreateEllipticRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width - 0.5, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width - 0.5)
'               Else
'                  TmpRegn = CreateEllipticRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height - 0.5, TmpControl.Top + TmpControl.Height - 0.5)
'               End If
'            ElseIf TmpControl.Shape = 4 Then
'               'Rechteck mit abgerundeten Ecken
'               If TmpControl.Width > TmpControl.Height Then
'                  TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
'               Else
'                  TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height, TmpControl.Width / 4, TmpControl.Width / 4)
'               End If
'            ElseIf TmpControl.Shape = 5 Then
'               'Quadrat mit abgerundeten Ecken
'               If TmpControl.Width > TmpControl.Height Then
'                  TmpRegn = CreateRoundRectRgn(TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + 1, TmpControl.Top + 1, TmpControl.Left + (TmpControl.Width - TmpControl.Height) / 2 + TmpControl.Height, TmpControl.Top + TmpControl.Height, TmpControl.Height / 4, TmpControl.Height / 4)
'               Else
'                  TmpRegn = CreateRoundRectRgn(TmpControl.Left + 1, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + 1, TmpControl.Left + TmpControl.Width, TmpControl.Top + (TmpControl.Height - TmpControl.Width) / 2 + TmpControl.Width, TmpControl.Width / 4, TmpControl.Width / 4)
'               End If
'            End If
'         End If
'      Else
'         'Eine Rechteckige Region erstellen
'         TmpRegn = CreateRectRgn(TmpControl.Left, TmpControl.Top, TmpControl.Left + TmpControl.Width, TmpControl.Top + TmpControl.Height)
'      End If
'      'Prüft ob das Steuerelement überhaupt eine Breite hat (die größer als 0 ist)
'      ErrorTest = 0
'      ErrorTest = TmpControl.Width
'      If ErrorTest <> 0 Or TypeOf TmpControl Is Line Then
'         'Die Regionen kombinieren
'         CombineRgn Regn, Regn, TmpRegn, RGN_XOR
'      End If
'   Next TmpControl
'   'Die Regionen erstellen
'   SetWindowRgn TransForm.hwnd, Regn, True
'End Sub
