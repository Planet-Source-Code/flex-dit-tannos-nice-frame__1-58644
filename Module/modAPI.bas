Attribute VB_Name = "modAPI"
'===Types=============================================================================================================

Public Type POINTAPI
    X As Long
    Y As Long
End Type

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type


Public Type TRIVERTEX
    X As Long
    Y As Long
    Red As Integer
    Green As Integer
    Blue As Integer
    alpha As Integer
End Type
Public Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type

Public Enum GradientFillRectType
    GRADIENT_FILL_RECT_H = 0
    GRADIENT_FILL_RECT_V = 1
End Enum


'=CONSTANTES de texte==================================================================================================

Public Const DT_RIGHT = &H2
Public Const DT_LEFT = &H0
Public Const DT_CENTER = &H1
Public Const DT_CALCRECT = &H400
Public Const DT_TOP = &H0
Public Const DT_BOTTOM = &H8
Public Const DT_VCENTER = &H4
Public Const DT_SINGLELINE = &H20
Public Const DT_END_ELLIPSIS = &H8000&






'=API POUR LE DESSIN==================================================================================================

Private Declare Function Rectangle Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function RoundRect Lib "gdi32" (ByVal hDc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hDc As Long, pVertex As TRIVERTEX, ByVal dwNumVertex As Long, pMesh As GRADIENT_RECT, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function GradientFill Lib "msimg32" (ByVal hDc As Long, pVertex As Any, ByVal dwNumVertex As Long, pMesh As Any, ByVal dwNumMesh As Long, ByVal dwMode As Long) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, _
                                               ByVal nWidth As Long, _
                                               ByVal crColor As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Public Declare Function GetCurrentObject Lib "gdi32" (ByVal hDc As Long, ByVal uObjectType As Long) As Long
Public Declare Function PaintRgn Lib "gdi32" (ByVal hDc As Long, ByVal hRgn As Long) As Long

Public Declare Function MoveToEx Lib "gdi32" (ByVal hDc As Long, _
                                              ByVal X As Long, _
                                              ByVal Y As Long, _
                                              lpPoint As POINTAPI) As Long
Public Declare Function LineTo Lib "gdi32" (ByVal hDc As Long, _
                                            ByVal X As Long, _
                                            ByVal Y As Long) As Long

Public Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, _
                                                              ByVal HPALETTE As Long, _
                                                              pccolorref As Long) As Long



Public Declare Function FillRect Lib "user32" (ByVal hDc As Long, _
                                               lpRect As RECT, _
                                               ByVal hBrush As Long) As Long


Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long

Public Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long

Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal X As Long, ByVal Y As Long) As Long


Public Declare Function GetGDIObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

