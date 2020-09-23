Attribute VB_Name = "modGlobalUtils"


'dessine la bordure
Public Sub UtilDrawShapeStyle(ByVal lngHdc As Long, _
                              ByVal X1 As Long, _
                              ByVal Y1 As Long, _
                              ByVal X2 As Long, _
                              ByVal Y2 As Long, _
                              ByVal radius As Long)

    RoundRect lngHdc, X1, Y1, X2, Y2, radius, radius
End Sub

'remplissage arr plan
Public Sub UtilDrawBackground(ByVal lngHdc As Long, _
                              ByVal colorStart As Long, _
                              ByVal colorEnd As Long, _
                              ByVal lngLeft As Long, _
                              ByVal lngTop As Long, _
                              ByVal lngWidth As Long, _
                              ByVal lngHeight As Long, _
                              Optional ByVal horizontal As Long = 0)


    Dim tR As RECT

    With tR
        .Left = lngLeft
        .Top = lngTop
        .Right = lngWidth    'lngLeft + lngWidth
        .Bottom = lngHeight    'lngTop + lngHeight
        ' gradient fill vertical:
    End With    'tR
    GradientFillRect lngHdc, tR, colorStart, colorEnd, IIf(horizontal = 0, GRADIENT_FILL_RECT_H, GRADIENT_FILL_RECT_V)

End Sub


Private Sub GradientFillRect(ByVal lHDC As Long, _
                             tR As RECT, _
                             ByVal oStartColor As OLE_COLOR, _
                             ByVal oEndColor As OLE_COLOR, _
                             ByVal eDir As GradientFillRectType)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR As GRADIENT_RECT
    Dim hBrush As Long
    Dim lStartColor As Long
    Dim lEndColor As Long

    'Dim lR As Long
    ' Use GradientFill:
    If Not (HasGradientAndTransparency) Then
        lStartColor = TranslateColor(oStartColor)
        lEndColor = TranslateColor(oEndColor)
        setTriVertexColor tTV(0), lStartColor
        tTV(0).x = tR.Left
        tTV(0).y = tR.Top
        setTriVertexColor tTV(1), lEndColor
        tTV(1).x = tR.Right
        tTV(1).y = tR.Bottom
        tGR.UpperLeft = 0
        tGR.LowerRight = 1
        GradientFill lHDC, tTV(0), 2, tGR, 1, eDir
    Else
        ' Fill with solid brush:
        hBrush = CreateSolidBrush(TranslateColor(oEndColor))
        FillRect lHDC, tR, hBrush
        DeleteObject hBrush
    End If

End Sub


Private Function TranslateColor(ByVal oClr As OLE_COLOR, _
                                Optional hPal As Long = 0) As Long

' Convert Automation color to Windows color
'--------- Drawing

    If OleTranslateColor(oClr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function

Private Sub setTriVertexColor(tTV As TRIVERTEX, _
                              ByVal lColor As Long)


    Dim lRed As Long
    Dim lGreen As Long
    Dim lBlue As Long

    lRed = (lColor And &HFF&) * &H100&
    lGreen = (lColor And &HFF00&)
    lBlue = (lColor And &HFF0000) \ &H100&
    With tTV
        setTriVertexColorComponent .Red, lRed
        setTriVertexColorComponent .Green, lGreen
        setTriVertexColorComponent .Blue, lBlue
    End With    'tTV

End Sub

Private Sub setTriVertexColorComponent(ByRef iColor As Integer, _
                                       ByVal lComponent As Long)

    If (lComponent And &H8000&) = &H8000& Then
        iColor = (lComponent And &H7F00&)
        iColor = iColor Or &H8000
    Else
        iColor = lComponent
    End If

End Sub



Public Property Get dBlendColor(ByVal oColorFrom As OLE_COLOR, _
                                ByVal oColorTo As OLE_COLOR, _
                                Optional ByVal alpha As Long = 128) As Long

    Dim lSrcR As Long
    Dim lSrcG As Long
    Dim lSrcB As Long
    Dim lDstR As Long
    Dim lDstG As Long
    Dim lDstB As Long
    Dim lCFrom As Long
    Dim lCTo As Long
    lCFrom = TranslateColor(oColorFrom)
    lCTo = TranslateColor(oColorTo)
    lSrcR = lCFrom And &HFF
    lSrcG = (lCFrom And &HFF00&) \ &H100&
    lSrcB = (lCFrom And &HFF0000) \ &H10000
    lDstR = lCTo And &HFF
    lDstG = (lCTo And &HFF00&) \ &H100&
    lDstB = (lCTo And &HFF0000) \ &H10000
    dBlendColor = RGB(((lSrcR * alpha) / 255) + ((lDstR * (255 - alpha)) / 255), ((lSrcG * alpha) / 255) + ((lDstG * (255 - alpha)) / 255), ((lSrcB * alpha) / 255) + ((lDstB * (255 - alpha)) / 255))

End Property



