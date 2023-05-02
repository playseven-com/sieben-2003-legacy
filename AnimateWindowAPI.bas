Attribute VB_Name = "AnimateWindowAPI"
''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright (c) by Christoph Haltiner, 1999 - 2000
' ChristophH@BasicPro.de
'
' Hilfsmodul für die  AnimateWindow API Funktion
''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit

' AnimateWindow API Funktion
Public Declare Function AnimateWindow Lib "user32" (ByVal hwnd As Long, ByVal dwTime As Long, ByVal dwFlags As Long) As Boolean

' AW_* Flags
Public Enum AW_FLAGS
    AW_SLIDE = &H40000
    AW_ACTIVATE = &H20000
    AW_BLEND = &H80000
    AW_CENTER = &H10
    AW_HIDE = &H10000
    AW_HOR_POSITIVE = &H1
    AW_HOR_NEGATIVE = &H2
    AW_VER_POSITIVE = &H4
    AW_VER_NEGATIVE = &H8
End Enum

' benötigte Window Message
Public Const WM_PRINTCLIENT = &H318

' restliche API/GDI Funktionen
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetVersion Lib "kernel32" () As Long
Public Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

' restliche API Konstanten und Typen
Public Const GWL_WNDPROC = (-4)
Public Const SRCCOPY = &HCC0020

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

' Variable für die Adresse der ursprünglichen Window Procedure
Dim WndProc As Long

' Richtet die neue Window Procedure ein
Public Sub LinkWndProc(frmWindow As Form)
On Error GoTo errHand
    WndProc = SetWindowLong(frmWindow.hwnd, GWL_WNDPROC, AddressOf WindowProc)

Exit Sub
errHand:
If ErrorBox("LinkWndProc", Err) Then Resume Next
End Sub

' Entfernt die neue Window Procedure
Public Sub DetachWndProc(frmWindow As Form)
On Error GoTo errHand
    Dim lResult As Long
    lResult = SetWindowLong(frmWindow.hwnd, GWL_WNDPROC, WndProc)
    
Exit Sub
errHand:
If ErrorBox("DetachWndProc", Err) Then Resume Next

End Sub

' Window Procedure
Public Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error GoTo errHand
' Client Area zeichnen
If uMsg = WM_PRINTCLIENT Then
    Dim lResult As Long
    Dim newhBrush As Long
    Dim oldhBrush As Long
    
    ' Bitmap aus der Res. Datei auf den DC wParam ausgeben
    Dim frm As Form
    For Each frm In Forms
        If frm.hwnd = hwnd Then
            ' Pinsel mit Bitmap Muster erzeugen...
            newhBrush = CreatePatternBrush(LoadPicture(frm.Tag))
            ' ...und zuweisen.
            oldhBrush = SelectObject(wParam, newhBrush)
            ' Rechteck malen
            lResult = Rectangle(wParam, 0, 0, frm.Width, frm.Height)
            ' Alter Pinsel wieder zuweisen
            lResult = SelectObject(wParam, oldhBrush)
        End If
    Next
    
End If

' ursprüngliche WindowProc aufrufen
WindowProc = CallWindowProc(WndProc, hwnd, uMsg, wParam, lParam)
Exit Function
errHand:
If ErrorBox("WindowProc", Err) Then Resume Next
End Function

Public Function AnimateWnd(hwnd As Long, dwTime As Long, dwFlags As AW_FLAGS) As Boolean
On Error GoTo errHand
    ' Windows Version überprüfen
    
    If WindowsVersion < 5 Then
        'Beep
        Exit Function
    End If
    ' AnimateWindow aufrufen
    AnimateWnd = AnimateWindow(hwnd, dwTime, dwFlags)

Exit Function
errHand:
If ErrorBox("AnimateWnd", Err) Then Resume Next

End Function
