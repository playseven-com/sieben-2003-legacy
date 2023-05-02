Attribute VB_Name = "modRenderPic"
Option Explicit

' *** Den Artikel zu diesem Modul finden Sie unter http://www.aboutvb.de/khw/artikel/khwrenderpic.htm ***

Public Sub RenderPicture(Picture As StdPicture, ByVal hdc As Long, ByVal XDest As Long, ByVal YDest As Long, Optional ByVal WidthDest As Long, Optional ByVal HeightDest As Long, Optional ByVal xSrc As Long, Optional ByVal ySrc As Long, Optional ByVal WidthSrc As Long, Optional ByVal HeightSrc As Long, Optional Container As Object, Optional ByVal DestScaleMode As Integer = vbPixels, Optional ByVal SrcScaleMode As Integer = vbTwips)
    Dim nXDest As Long
    Dim nYDest As Long
    Dim nWidthDest As Long
    Dim nHeightDest As Long
    Dim nXSrc As Double
    Dim nYSrc As Double
    Dim nWidthSrc As Double
    Dim nHeightSrc As Double
    Dim nContainer As Object
    
    If nContainer Is Nothing Then
        Set nContainer = Screen.ActiveForm
    End If
    With Picture
        nXDest = XDest
        If WidthDest = 0 Then
            nWidthDest = nContainer.ScaleX(.Width, vbHimetric, vbPixels)
        Else
            nWidthDest = WidthDest
        End If
        If HeightDest = 0 Then
            nHeightDest = nContainer.ScaleY(.Height, vbHimetric, vbPixels)
        Else
            nHeightDest = HeightDest
        End If
        nYDest = YDest + nHeightDest
        nHeightDest = nHeightDest * -1
        With nContainer
            If DestScaleMode <> vbPixels Then
                nXDest = .ScaleX(nXDest, DestScaleMode, vbPixels)
                nYDest = .ScaleY(nYDest, DestScaleMode, vbPixels)
                nWidthDest = .ScaleX(nWidthDest, DestScaleMode, vbPixels)
                nHeightDest = .ScaleY(nHeightDest, DestScaleMode, vbPixels)
            End If
            nYDest = nYDest - 1
            nXSrc = .ScaleX(xSrc, SrcScaleMode, vbHimetric)
            nYSrc = .ScaleY(ySrc, SrcScaleMode, vbHimetric) + .ScaleY(1, vbPixels, vbHimetric)
        End With
        If WidthSrc = 0 Then
            nWidthSrc = .Width
        Else
            nWidthSrc = nContainer.ScaleX(WidthSrc, SrcScaleMode, vbHimetric)
        End If
        If HeightSrc = 0 Then
            nHeightSrc = .Height
        Else
            nHeightSrc = nContainer.ScaleY(HeightSrc, SrcScaleMode, vbHimetric)
        End If
        .Render CLng(hdc), CLng(nXDest), CLng(nYDest), CLng(nWidthDest), CLng(nHeightDest), nXSrc, nYSrc, nWidthSrc, nHeightSrc, 0&
    End With
End Sub

