VERSION 5.00
Begin VB.UserControl MyNiceFrame 
   Alignable       =   -1  'True
   ClientHeight    =   780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2505
   ControlContainer=   -1  'True
   DrawWidth       =   56
   EditAtDesignTime=   -1  'True
   ScaleHeight     =   52
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   167
   ToolboxBitmap   =   "MyNiceFrame.ctx":0000
End
Attribute VB_Name = "MyNiceFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'===Declarations enumérations et types======================================================================================================


Private Declare Function GetIconInfo Lib "user32" (ByVal hIcon As Long, piconinfo As ICONINFO) As Long
Private Declare Function CopyImage Lib "user32" (ByVal Handle As Long, ByVal imageType As Long, ByVal newWidth As Long, ByVal newHeight As Long, ByVal lFlags As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetMapMode Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetMapMode Lib "gdi32" (ByVal hDc As Long, ByVal nMapMode As Long) As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDc As Long, ByVal HPALETTE As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetBkColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long
Private Declare Function GetTextColor Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hDc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObj As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObj As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Const SRCAND = &H8800C6
Const SRCCOPY = &HCC0020
Const SRCERASE = &H440328
Const SRCINVERT = &H660046
Const SRCPAINT = &HEE0086

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type ICONINFO
    fIcon As Long
    xHotSpot As Long
    yHotSpot As Long
    hbmMask As Long
    hbmColor As Long
End Type

Public Enum Theme
    xThemeDarkBlue
    xThemeMoney
    xThemeMediaPlayer
    xThemeMediaPlayer2
    xThemeGreen
    xThemeMetallic
    xThemeOffice2003
    xThemeOrange
    xThemeTurquoise
    xThemeGray
    xThemeDarkBlue2
    xThemeLightBlue
End Enum

'Public Enum PictureAlign
'    xAlignLeftEdge                      '-->Left edge
'    xAlignRightEdge                     '-->Right Edge
'    xAlignLeftOfCaption                 '-->Left of the caption
'    xAlignRightOfCaption                '-->Right of the caption
'End Enum

Public Enum TextAlign
    xAlignLefttop
    xAlignLeftMiddle
    xAlignLeftBottom
    xAlignRightTop
    xAlignRightMiddle
    xAlignRightBottom
    xAlignCenterTop
    xAlignCenterMiddle
    xAlignCenterBottom
End Enum

Public Enum FillStyle
    HorizontalFading
    VerticalFading
End Enum

Public Enum HeaderFooterStyleSize
    Small
    Medium
    Large
End Enum

Public Enum TraceBorderStyle
    SOLID = 0                   '  _______
    DASH = 1                    '  -------
    DOT = 2                     '  .......
    DASHDOT = 3                 '  _._._._
    DASHDOTDOT = 4              '  _.._.._
    NONE = 5
End Enum

Public Enum ShapeStyle
    Squared
    Rounded
End Enum

Public Enum HeaderBarStates
    eBarcollapsed
    eBarExpanded
End Enum




'===Constantes=========================================================================================================

'valeurs par defaut des propriétés

'#Header
Private Const m_def_sHeaderText As String = "Header"      'texte du header
Private Const m_def_eHeaderTextAlign As Integer = xAlignLefttop      'alignement du caption
Private Const m_def_oHeaderTextColor As Long = vbWhite     'couleur texte header
Private Const m_def_eHeaderSize As Integer = HeaderFooterStyleSize.Medium      'taille
Private Const m_def_bHeaderVisible As Boolean = True        'visibilité
Private Const m_def_oHeaderBackColor As Long = &HFFFFFF          'couleur utilisée pour l'arrière plan
Private Const m_def_oHeaderFadeColor As Long = &HD96C00          'couleur utilisée pour le gradient
Private Const m_def_eHeaderFillStyle As Integer = VerticalFading        'type du gradient

'#Footer
Private Const m_def_sFooterText As String = "Footer"      'texte du Footer
Private Const m_def_eFooterTextAlign As Integer = xAlignLefttop      'alignement du caption
Private Const m_def_oFooterTextColor As Long = vbWhite     'couleur texte Footer
Private Const m_def_eFooterSize As Integer = HeaderFooterStyleSize.Medium      'taille
Private Const m_def_bFooterVisible As Boolean = True        'visibilité
Private Const m_def_oFooterBackColor As Long = &HD96C00       'couleur utilisée pour l'arrière plan
Private Const m_def_oFooterFadeColor As Long = &HFFFFFF       'couleur utilisée pour le gradient
Private Const m_def_eFooterFillStyle As Integer = VerticalFading        'type du gradient


'#Container
Private Const m_def_oContainerBackColor As Long = vbButtonFace       'couleur utilisée pour l'arrière plan
Private Const m_def_oContainerFadeColor As Long = vbButtonFace       'couleur utilisée pour le gradient
Private Const m_def_eContainerFillStyle As Integer = VerticalFading        'type du gradient
Private Const m_def_oContainerForeColor As Long = vbButtonText       'couleur utilisée pour le texte
Private Const m_def_oContainerBorderColor As Long = vbBlack       'couleur utilisée pour la bordure
Private Const m_def_eContainerBorderStyle As Integer = SOLID     'style de bordure
Private Const m_def_eContainerShapeStyle As Long = Squared       'forme du controle
Private Const m_def_iContainerCornerRadius As Integer = 20   'angle utilise pour les angles arrondis

'#divers
Private Const m_def_iTheme As Integer = xThemeDarkBlue      'theme par defaut
Private Const m_def_UseCustomColors As Boolean = False      'couleurs par defaut=faux
Private Const m_def_CollapseButtonColor As Long = vbRed
Private ValFooterTextAlign As Long
Private ValHeaderTextAlign As Long

Private bInitializing As Boolean   'pour eviter de redessiner sur initialize



'===Propriétés========================================================================================================

'#Header
Private m_sHeaderText As String  'texte du header
Private m_fHeaderTextFont As Font  'police du header
Private m_eHeaderTextAlign As Integer       'alignement du texte
Private m_oHeaderTextColor As OLE_COLOR       'couleur texte header
Private m_iHeaderSize As Integer        'taille
Private m_bHeaderVisible As Boolean          'visibilité
Private m_oHeaderBackColor As OLE_COLOR         'couleur utilisée pour l'arrière plan
Private m_oHeaderFadeColor As OLE_COLOR         'couleur utilisée pour le gradient
Private m_eHeaderFillStyle As Integer          'type du gradient
Private m_HeaderPicture As StdPicture           'image du header
Private m_HeaderPictureSize As Long             'taille de l'image du header


'#Footer
Private m_sFooterText As String  'texte du Footer
Private m_fFooterTextFont As Font  'police du Footer
Private m_eFooterTextAlign As Integer       'alignement du texte
Private m_oFooterTextColor As OLE_COLOR       'couleur texte Footer
Private m_iFooterSize As Integer        'taille
Private m_bFooterVisible As Boolean          'visibilité
Private m_oFooterBackColor As OLE_COLOR         'couleur utilisée pour l'arrière plan
Private m_oFooterFadeColor As OLE_COLOR         'couleur utilisée pour le gradient
Private m_eFooterFillStyle As Integer          'type du gradient


'#Container
Private m_fContainerTextFont As Font        'police du container
Private m_oContainerBackColor As OLE_COLOR         'couleur utilisée pour l'arrière plan
Private m_oContainerFadeColor As OLE_COLOR         'couleur utilisée pour le gradient
Private m_eContainerFillStyle As Integer          'type du gradient
Private m_oContainerForeColor As OLE_COLOR         'couleur utilisée pour le texte
Private m_oContainerBorderColor As OLE_COLOR         'couleur utilisée pour la bordure
Private m_eContainerBorderStyle As TraceBorderStyle       'style de bordure
Private m_eContainerShapeStyle As ShapeStyle          'forme du controle
Private m_iContainerCornerRadius As Integer     'angle utilise pour les angles arrondis
Private m_ContainerPicture As StdPicture        'image du container
Private m_ContainerPictureSize As Long          'taille de l'image du container


'#Divers
Private m_UseCustomColors As Boolean        'couleurs par defaut
Private m_enmTheme As Theme     'theme
Private m_bEnabled As Boolean       'enable
Private m_CanExpand As Boolean          'expandable
Private State As HeaderBarStates          'état du controle
Private CollapseOffset As Long          'hauteur expandable/collapsable
Private m_CollapseButtonColor As OLE_COLOR      'couleur du bouton

'#Divers intermédiares
'#Header
Private m_lColorHeaderColorOne As OLE_COLOR     '1ère couleur du gradient
Private m_lColorHeaderColorTwo As OLE_COLOR     '2ème couleur du gradient
Private m_lColorHeaderForeColor As OLE_COLOR        'recupère la couleur du texte
Private m_hRegion As Long                   'region du header
Private m_lColorCollapseButtonColor As OLE_COLOR
'#Footer
Private m_lColorFooterColorOne As OLE_COLOR     '1ère couleur du gradient
Private m_lColorFooterColorTwo As OLE_COLOR     '2ème couleur du gradient
Private m_lColorFooterForeColor As OLE_COLOR        'recupère la couleur du texte
'#Container
Private m_lColorContainerColorOne As OLE_COLOR     '1ère couleur du gradient
Private m_lColorContainerColorTwo As OLE_COLOR     '2ème couleur du gradient
Private m_lColorContainerColorBorder As OLE_COLOR     'recupere la couleur de la bordure


'=Evenements Public =======================================================================================================

Public Event Click()
Attribute Click.VB_UserMemId = -600
Public Event DblClick()
Attribute DblClick.VB_UserMemId = -601
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_UserMemId = -602
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_UserMemId = -603
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_UserMemId = -604
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_UserMemId = -605
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_UserMemId = -606
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_UserMemId = -607


'=Debut du code =======================================================================================================
'
'#Header
'HeaderText
Public Property Get HeaderText() As String
Attribute HeaderText.VB_Description = "Renvoie ou définit le texte affiché dans la barre d'entête"
Attribute HeaderText.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderText = m_sHeaderText
End Property

Public Property Let HeaderText(ByVal sHeaderText As String)
    m_sHeaderText = sHeaderText
    Call UserControl.PropertyChanged("HeaderText")
    Call DrawControl

End Property

'HeaderTextFont
Public Property Get HeaderTextFont() As Font
Attribute HeaderTextFont.VB_Description = "Police utilisée pour l'affichage dans l'entête"
Attribute HeaderTextFont.VB_ProcData.VB_Invoke_Property = ";Police"
    Set HeaderTextFont = m_fHeaderTextFont
End Property

Public Property Set HeaderTextFont(objHeaderTextFont As Font)
    Set m_fHeaderTextFont = objHeaderTextFont
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderTextFont")
End Property

'HeaderTextAlign
Public Property Get HeaderTextAlign() As TextAlign
Attribute HeaderTextAlign.VB_Description = "Renvoie ou définit l'alignement du texte dans la zone concernée"
Attribute HeaderTextAlign.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderTextAlign = m_eHeaderTextAlign
End Property
Public Property Let HeaderTextAlign(ByVal eHeaderTextAlign As TextAlign)

    Select Case eHeaderTextAlign
    Case xAlignLefttop
        ValHeaderTextAlign = DT_LEFT Or DT_TOP Or DT_SINGLELINE
    Case xAlignLeftMiddle
        ValHeaderTextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignLeftBottom
        ValHeaderTextAlign = DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE
    Case xAlignRightTop
        ValHeaderTextAlign = DT_RIGHT Or DT_TOP Or DT_SINGLELINE
    Case xAlignRightMiddle
        ValHeaderTextAlign = DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignRightBottom
        ValHeaderTextAlign = DT_RIGHT Or DT_BOTTOM Or DT_SINGLELINE
    Case xAlignCenterTop
        ValHeaderTextAlign = DT_CENTER Or DT_TOP Or DT_SINGLELINE
    Case xAlignCenterMiddle
        ValHeaderTextAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignCenterBottom
        ValHeaderTextAlign = DT_CENTER Or DT_BOTTOM Or DT_SINGLELINE
    End Select
    m_eHeaderTextAlign = eHeaderTextAlign
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderTextAlign")

End Property

'HeaderTextColor
Public Property Get HeaderTextColor() As OLE_COLOR
Attribute HeaderTextColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderTextColor = m_oHeaderTextColor
End Property
Public Property Let HeaderTextColor(ByVal eHeaderTextColor As OLE_COLOR)
    m_oHeaderTextColor = eHeaderTextColor
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderTextColor")
End Property

'HeaderSize
Public Property Get HeaderSize() As HeaderFooterStyleSize
Attribute HeaderSize.VB_Description = "Renvoie ou définit la taille de l'entête"
Attribute HeaderSize.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderSize = m_iHeaderSize
End Property
Public Property Let HeaderSize(ByVal eHeaderSize As HeaderFooterStyleSize)
    m_iHeaderSize = eHeaderSize
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderSize")
End Property

'HeaderVisible
Public Property Get HeaderVisible() As Boolean
Attribute HeaderVisible.VB_Description = "Renvoie ou définit une valeur qui détermine si l'objet est visible ou masqué"
Attribute HeaderVisible.VB_ProcData.VB_Invoke_Property = ";Comportement"
    HeaderVisible = m_bHeaderVisible
End Property
Public Property Let HeaderVisible(ByVal bHeaderVisible As Boolean)
    m_bHeaderVisible = bHeaderVisible
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderVisible")
End Property

'HeaderBackColor
Public Property Get HeaderBackColor() As OLE_COLOR
Attribute HeaderBackColor.VB_Description = "Renvoie ou définit la couleur utilisée pour l'entête"
Attribute HeaderBackColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderBackColor = m_oHeaderBackColor
End Property
Public Property Let HeaderBackColor(ByVal objHeaderBackColor As OLE_COLOR)
    m_oHeaderBackColor = objHeaderBackColor
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderBackColor")
End Property

'HeaderFadeColor
Public Property Get HeaderFadeColor() As OLE_COLOR
Attribute HeaderFadeColor.VB_Description = "Renvoie ou définit la couleur utilisée pour le dégradé de l'entête"
Attribute HeaderFadeColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderFadeColor = m_oHeaderFadeColor
End Property
Public Property Let HeaderFadeColor(ByVal objHeaderFadeColor As OLE_COLOR)
    m_oHeaderFadeColor = objHeaderFadeColor
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderFadeColor")
End Property

'HeaderFillStyle
Public Property Get HeaderFillStyle() As FillStyle
Attribute HeaderFillStyle.VB_Description = "Renvoie ou définit le type de dégradé utilisé pour l'entête"
Attribute HeaderFillStyle.VB_ProcData.VB_Invoke_Property = ";Apparence"
    HeaderFillStyle = m_eHeaderFillStyle
End Property
Public Property Let HeaderFillStyle(ByVal eHeaderFillStyle As FillStyle)
    m_eHeaderFillStyle = eHeaderFillStyle
    Call DrawControl
    Call UserControl.PropertyChanged("HeaderFillStyle")
End Property

'HeaderPicture
Public Property Set HeaderPicture(NewIcon As StdPicture)
    Set m_HeaderPicture = NewIcon
    PropertyChanged "HeaderPicture"
    Call DrawControl
End Property
Public Property Get HeaderPicture() As StdPicture
    Set HeaderPicture = m_HeaderPicture
End Property


Private Sub DrawPicture(ByRef tP As RECT, sPic As StdPicture, Optional newSize As Long)
    Dim BMInf As BITMAP
    Dim ICInf As ICONINFO
    Dim dRect As RECT
    Dim BMtR As RECT
    Dim TransImage As Long
    Dim PicW As Long
    Dim PicH As Long

PicW = newSize: PicH = newSize
    '-- on recupere les dimensions des images
    If Not sPic Is Nothing Then
        Call GetObjectAPI(sPic.Handle, Len(BMInf), BMInf)
        If BMInf.bmBits = 0 Then
            Call GetIconInfo(sPic.Handle, ICInf)
            If ICInf.hbmColor <> 0 Then    '--il s'agit d'une icone
                Call GetObjectAPI(ICInf.hbmColor, Len(BMInf), BMInf)
                DeleteObject ICInf.hbmColor
                If ICInf.hbmMask <> 0 Then
                    DeleteObject ICInf.hbmMask
                End If
            End If
        End If
    End If

    dRect = tP

    If (sPic.Type = vbPicTypeIcon) Then
        '--cas d'une icone
        '--on dessine avec la taille passée en paramètre
        DrawIconEx UserControl.hDc, dRect.Left, dRect.Top, sPic.Handle, PicW, PicH, 0, 0, &H3
    Else
        '--cas d'un bitmap
        '--on dessine l'image de toute sa taille
        TransImage = CopyImage(sPic.Handle, 0, PicW, PicH, ByVal 0&)
        DrawTransparentBitmap UserControl.hDc, dRect, TransImage, BMtR, PicW, PicH
    End If


End Sub


Private Sub DrawTransparentBitmap(lHDCdest As Long, destRect As RECT, _
                                  lBMPsource As Long, bmpRect As RECT, _
                                  ByVal bmpSizeX As Long, _
                                  ByVal bmpSizeY As Long)
    Const DSna = &H220326
    
    Dim lMask2Use As Long
    Dim bmpMask As Long, bmpMemory As Long, bmpColor As Long
    Dim bmpObjectOld As Long, bmpMemoryOld As Long, bmpColorOld As Long
    Dim lBackDC As Long, hWndDc As Long, lHDCsrc As Long, lMaskDC As Long, lHDCcolor As Long
    Dim bmpPointSizeX As Long, bmpPointSizeY As Long, srcX As Long, srcY As Long

    Dim hPalOld As Long, hPalMem As Long

    hWndDc = GetDC(0&): If hWndDc = 0 Then Exit Sub
    lHDCsrc = CreateCompatibleDC(hWndDc)
    SelectObject lHDCsrc, lBMPsource

    srcX = bmpSizeX
    srcY = bmpSizeY

    bmpRect.Right = srcX
    bmpRect.Bottom = srcY

    bmpPointSizeX = bmpSizeX
    bmpPointSizeY = bmpSizeY

    lMask2Use = ConvertColor(GetPixel(lHDCsrc, 0, 0))

    '
    lMaskDC = CreateCompatibleDC(hWndDc)
    lBackDC = CreateCompatibleDC(hWndDc)
    lHDCcolor = CreateCompatibleDC(hWndDc)

    bmpColor = CreateCompatibleBitmap(hWndDc, srcX, srcY)
    bmpMemory = CreateCompatibleBitmap(hWndDc, bmpPointSizeX, bmpPointSizeY)
    bmpMask = CreateBitmap(srcX, srcY, 1&, 1&, ByVal 0&)

    bmpColorOld = SelectObject(lHDCcolor, bmpColor)
    bmpMemoryOld = SelectObject(lBackDC, bmpMemory)
    bmpObjectOld = SelectObject(lMaskDC, bmpMask)

    ReleaseDC 0&, hWndDc

    '
    SetMapMode lBackDC, GetMapMode(lHDCdest)
    hPalMem = SelectPalette(lBackDC, 0, True)
    RealizePalette lBackDC

    BitBlt lBackDC, 0&, 0&, bmpPointSizeX, bmpPointSizeY, lHDCdest, destRect.Left, destRect.Top, vbSrcCopy

    hPalOld = SelectPalette(lHDCcolor, 0, True)
    RealizePalette lHDCcolor
    SetBkColor lHDCcolor, GetBkColor(lHDCsrc)
    SetTextColor lHDCcolor, GetTextColor(lHDCsrc)

    BitBlt lHDCcolor, 0&, 0&, srcX, srcY, lHDCsrc, bmpRect.Left, bmpRect.Top, vbSrcCopy

    SetBkColor lHDCcolor, lMask2Use
    SetTextColor lHDCcolor, vbWhite

    BitBlt lMaskDC, 0&, 0&, srcX, srcY, lHDCcolor, 0&, 0&, vbSrcCopy
    
    SetTextColor lHDCcolor, vbBlack
    SetBkColor lHDCcolor, vbWhite
    BitBlt lHDCcolor, 0, 0, srcX, srcY, lMaskDC, 0, 0, DSna

    StretchBlt lBackDC, 0, 0, bmpSizeX, bmpSizeY, lMaskDC, 0&, 0&, srcX, srcY, vbSrcAnd

    StretchBlt lBackDC, 0&, 0&, bmpSizeX, bmpSizeY, lHDCcolor, 0, 0, srcX, srcY, vbSrcPaint

    BitBlt lHDCdest, destRect.Left, destRect.Top, bmpPointSizeX, bmpPointSizeY, lBackDC, 0&, 0&, vbSrcCopy

    '--efface les bitmaps en mémoires et les DC
    DeleteObject SelectObject(lHDCcolor, bmpColorOld)
    DeleteObject SelectObject(lMaskDC, bmpObjectOld)
    DeleteObject SelectObject(lBackDC, bmpMemoryOld)
    DeleteDC lBackDC
    DeleteDC lMaskDC
    DeleteDC lHDCcolor
    DeleteDC lHDCsrc
End Sub


Private Function ConvertColor(tColor As Long) As Long

' Converts VB color constants to real color values

    If tColor < 0 Then
        ConvertColor = GetSysColor(tColor And &HFF&)
    Else
        ConvertColor = tColor
    End If
End Function



'HeaderPictureSize
Public Property Get HeaderPictureSize() As Integer
    HeaderPictureSize = m_HeaderPictureSize
End Property

Public Property Let HeaderPictureSize(ByVal NewIconSize As Integer)
    m_HeaderPictureSize = NewIconSize
    PropertyChanged "HeaderPictureSize"
    Call DrawControl
End Property

'
'
'#Footer
'FooterText
Public Property Get FooterText() As String
Attribute FooterText.VB_Description = "Renvoie ou définit le texte affiché dans la barre pied de frame"
Attribute FooterText.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterText = m_sFooterText
End Property

Public Property Let FooterText(ByVal sFooterText As String)
    m_sFooterText = sFooterText
    Call DrawControl
    Call UserControl.PropertyChanged("FooterText")
End Property

'FooterTextFont
Public Property Get FooterTextFont() As Font
Attribute FooterTextFont.VB_Description = "Police utilisée pour l'affichage"
Attribute FooterTextFont.VB_ProcData.VB_Invoke_Property = ";Police"
    Set FooterTextFont = m_fFooterTextFont
End Property

Public Property Set FooterTextFont(objFooterTextFont As Font)
    Set m_fFooterTextFont = objFooterTextFont
    Call DrawControl
    Call UserControl.PropertyChanged("FooterTextFont")
End Property

'FooterTextAlign
Public Property Get FooterTextAlign() As TextAlign
Attribute FooterTextAlign.VB_Description = "Renvoie ou définit l'alignement du texte dans la zone concernée"
Attribute FooterTextAlign.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterTextAlign = m_eFooterTextAlign
End Property
Public Property Let FooterTextAlign(ByVal eFooterTextAlign As TextAlign)
    
    Select Case eFooterTextAlign
    Case xAlignLefttop
        ValFooterTextAlign = DT_LEFT Or DT_TOP Or DT_SINGLELINE
    Case xAlignLeftMiddle
        ValFooterTextAlign = DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignLeftBottom
        ValFooterTextAlign = DT_LEFT Or DT_BOTTOM Or DT_SINGLELINE
    Case xAlignRightTop
        ValFooterTextAlign = DT_RIGHT Or DT_TOP Or DT_SINGLELINE
    Case xAlignRightMiddle
        ValFooterTextAlign = DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignRightBottom
        ValFooterTextAlign = DT_RIGHT Or DT_BOTTOM Or DT_SINGLELINE
    Case xAlignCenterTop
        ValFooterTextAlign = DT_CENTER Or DT_TOP Or DT_SINGLELINE
    Case xAlignCenterMiddle
        ValFooterTextAlign = DT_CENTER Or DT_VCENTER Or DT_SINGLELINE
    Case xAlignCenterBottom
        ValFooterTextAlign = DT_CENTER Or DT_BOTTOM Or DT_SINGLELINE
    End Select
    m_eFooterTextAlign = eFooterTextAlign
    Call DrawControl
    Call UserControl.PropertyChanged("FooterTextAlign")
End Property

'FooterTextColor
Public Property Get FooterTextColor() As OLE_COLOR
Attribute FooterTextColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterTextColor = m_oFooterTextColor
End Property
Public Property Let FooterTextColor(ByVal eFooterTextColor As OLE_COLOR)
    m_oFooterTextColor = eFooterTextColor
    Call DrawControl
    Call UserControl.PropertyChanged("FooterTextColor")
End Property

'FooterSize
Public Property Get FooterSize() As HeaderFooterStyleSize
Attribute FooterSize.VB_Description = "Renvoie ou définit la taille du pied"
Attribute FooterSize.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterSize = m_iFooterSize
End Property
Public Property Let FooterSize(ByVal eFooterSize As HeaderFooterStyleSize)
    m_iFooterSize = eFooterSize
    Call DrawControl
    Call UserControl.PropertyChanged("FooterSize")
End Property

'FooterVisible
Public Property Get FooterVisible() As Boolean
Attribute FooterVisible.VB_Description = "Renvoie ou définit une valeur qui détermine si l'objet est visible ou masqué"
Attribute FooterVisible.VB_ProcData.VB_Invoke_Property = ";Comportement"
    FooterVisible = m_bFooterVisible
End Property
Public Property Let FooterVisible(ByVal bFooterVisible As Boolean)
    m_bFooterVisible = bFooterVisible
    Call DrawControl
    Call UserControl.PropertyChanged("FooterVisible")
End Property

'FooterBackColor
Public Property Get FooterBackColor() As OLE_COLOR
Attribute FooterBackColor.VB_Description = "Renvoie ou définit la couleur utilisée pour le pied"
Attribute FooterBackColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterBackColor = m_oFooterBackColor
End Property
Public Property Let FooterBackColor(ByVal objFooterBackColor As OLE_COLOR)
    m_oFooterBackColor = objFooterBackColor
    Call DrawControl
    Call UserControl.PropertyChanged("FooterBackColor")
End Property

'FooterFadeColor
Public Property Get FooterFadeColor() As OLE_COLOR
Attribute FooterFadeColor.VB_Description = "Renvoie ou définit la couleur utilisée pour le dégradé du pied"
Attribute FooterFadeColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterFadeColor = m_oFooterFadeColor
End Property
Public Property Let FooterFadeColor(ByVal objFooterFadeColor As OLE_COLOR)
    m_oFooterFadeColor = objFooterFadeColor
    Call DrawControl
    Call UserControl.PropertyChanged("FooterFadeColor")
End Property

'FooterFillStyle
Public Property Get FooterFillStyle() As FillStyle
Attribute FooterFillStyle.VB_Description = "Renvoie ou définit le type de dégradé utilisé pour le pied"
Attribute FooterFillStyle.VB_ProcData.VB_Invoke_Property = ";Apparence"
    FooterFillStyle = m_eFooterFillStyle
End Property
Public Property Let FooterFillStyle(ByVal eFooterFillStyle As FillStyle)
    m_eFooterFillStyle = eFooterFillStyle
    Call DrawControl
    Call UserControl.PropertyChanged("FooterFillStyle")
End Property
'
'
'#Container
'ContainerTextFont
Public Property Get ContainerTextFont() As Font
Attribute ContainerTextFont.VB_Description = "Police utilisée pour l'affichage de l'arrière plan"
Attribute ContainerTextFont.VB_ProcData.VB_Invoke_Property = ";Police"
    Set ContainerTextFont = m_fContainerTextFont
End Property

Public Property Set ContainerTextFont(objContainerTextFont As Font)
    Set m_fContainerTextFont = objContainerTextFont
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerTextFont")
End Property

'BackColor
Public Property Get ContainerBackColor() As OLE_COLOR
Attribute ContainerBackColor.VB_Description = "Renvoie ou définit la couleur d'arrière plan de l'objet"
Attribute ContainerBackColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerBackColor = m_oContainerBackColor
End Property
Public Property Let ContainerBackColor(ByVal objContainerBackColor As OLE_COLOR)
    m_oContainerBackColor = objContainerBackColor
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerBackColor")
End Property

'FadeColor
Public Property Get ContainerFadeColor() As OLE_COLOR
Attribute ContainerFadeColor.VB_Description = "Renvoie ou définit la couleur utilisée pour le dégradé de l'arrière plan "
Attribute ContainerFadeColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerFadeColor = m_oContainerFadeColor
End Property
Public Property Let ContainerFadeColor(ByVal objContainerFadeColor As OLE_COLOR)
    m_oContainerFadeColor = objContainerFadeColor
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerFadeColor")
End Property

'FillStyle
Public Property Get ContainerFillStyle() As FillStyle
Attribute ContainerFillStyle.VB_Description = "Renvoie ou définit le type de dégradé utilisé pour l'arrière plan"
Attribute ContainerFillStyle.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerFillStyle = m_eContainerFillStyle
End Property
Public Property Let ContainerFillStyle(ByVal eContainerFillStyle As FillStyle)
    m_eContainerFillStyle = eContainerFillStyle
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerFillStyle")
End Property

'ForeColor
Public Property Get ContainerForeColor() As OLE_COLOR
Attribute ContainerForeColor.VB_Description = "Renvoie ou définit la couleur utilisée pour l'affichage du texte ou des graphiques de l'arrière plan"
Attribute ContainerForeColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerForeColor = m_oContainerForeColor
End Property
Public Property Let ContainerForeColor(ByVal objContainerForeColor As OLE_COLOR)
    m_oContainerForeColor = objContainerForeColor
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerForeColor")
End Property

'BorderColor
Public Property Get ContainerBorderColor() As OLE_COLOR
Attribute ContainerBorderColor.VB_Description = "Renvoie ou définit la couleur de la bordure"
Attribute ContainerBorderColor.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerBorderColor = m_oContainerBorderColor
End Property
Public Property Let ContainerBorderColor(ByVal objContainerBorderColor As OLE_COLOR)
    m_oContainerBorderColor = objContainerBorderColor
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerBorderColor")
End Property

'BorderStyle
Public Property Get ContainerBorderStyle() As TraceBorderStyle
Attribute ContainerBorderStyle.VB_Description = "Renvoie ou définit le style de bordure"
Attribute ContainerBorderStyle.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerBorderStyle = m_eContainerBorderStyle
End Property
Property Let ContainerBorderStyle(ByVal eBorderStyle As TraceBorderStyle)
    m_eContainerBorderStyle = eBorderStyle
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerBorderStyle")
End Property

'ShapeStyle
Public Property Get ContainerShapeStyle() As ShapeStyle
Attribute ContainerShapeStyle.VB_Description = "Renvoie ou définit la forme de l'objet"
Attribute ContainerShapeStyle.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerShapeStyle = m_eContainerShapeStyle
End Property
Property Let ContainerShapeStyle(ByVal eShapeStyle As ShapeStyle)
    m_eContainerShapeStyle = eShapeStyle
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerShapeStyle")
End Property

'CornerRadius
Public Property Get ContainerCornerRadius() As Integer
Attribute ContainerCornerRadius.VB_Description = "Renvoie ou définit la valeur de l'angle utilisée"
Attribute ContainerCornerRadius.VB_ProcData.VB_Invoke_Property = ";Apparence"
    ContainerCornerRadius = m_iContainerCornerRadius
End Property
Property Let ContainerCornerRadius(ByVal iContainerCornerRadius As Integer)
    m_iContainerCornerRadius = iContainerCornerRadius
    Call DrawControl
    Call UserControl.PropertyChanged("ContainerCornerRadius")
End Property

'ContainerPicture
Public Property Set ContainerPicture(NewIcon As StdPicture)
    Set m_ContainerPicture = NewIcon
    PropertyChanged "ContainerPicture"
    Call DrawControl
End Property
Public Property Get ContainerPicture() As StdPicture
    Set ContainerPicture = m_ContainerPicture
End Property
'ContainerPictureSize
Public Property Get ContainerPictureSize() As Integer
    ContainerPictureSize = m_ContainerPictureSize
End Property

Public Property Let ContainerPictureSize(ByVal NewIconSize As Integer)
    m_ContainerPictureSize = NewIconSize
    PropertyChanged "ContainerPictureSize"
    Call DrawControl
End Property
'
'
'#Divers
'UseCustonColors
Public Property Get UseCustomColors() As Boolean
Attribute UseCustomColors.VB_Description = "Renvoie ou définit si on utilises des couleurs personnalisées pour l'affichage de l'objet"
Attribute UseCustomColors.VB_ProcData.VB_Invoke_Property = ";Comportement"
    UseCustomColors = m_UseCustomColors
End Property

Public Property Let UseCustomColors(ByVal New_UseCustomColors As Boolean)
    m_UseCustomColors = New_UseCustomColors
    PropertyChanged "UseCustomColors"
    GetGradientColors
    Call DrawControl
End Property

'Appearence
Public Property Get Appearence() As Theme
Attribute Appearence.VB_Description = "Renvoie ou définit l'apparence de l'objet"
Attribute Appearence.VB_ProcData.VB_Invoke_Property = ";Apparence"
    Appearence = m_enmTheme
End Property

Public Property Let Appearence(enmNewTheme As Theme)
    m_enmTheme = enmNewTheme
    Call UserControl.PropertyChanged("Appearence")
    GetGradientColors
    Call DrawControl
End Property

'Enable
Public Property Get Enabled() As Boolean
    Enabled = m_bEnabled
End Property
Public Property Let Enabled(ByVal vNewValue As Boolean)
    On Error Resume Next
    m_bEnabled = vNewValue
    Dim Ctl As Control
    For Each Ctl In UserControl.ContainedControls
        Ctl.Enabled = m_bEnabled
    Next
End Property
'State
Public Property Get Expandable() As Boolean
Attribute Expandable.VB_Description = "Renvoie / défint l'état du contrôle"
Attribute Expandable.VB_ProcData.VB_Invoke_Property = ";Comportement"
    Expandable = m_CanExpand
End Property

Public Property Let Expandable(ByVal New_State As Boolean)
    m_CanExpand = New_State
    State = eBarcollapsed
    PropertyChanged "Expandable"
    Call DrawControl
End Property

'CollapseButtonColor
Public Property Get CollapseButtonColor() As OLE_COLOR
    CollapseButtonColor = m_CollapseButtonColor
End Property
Public Property Let CollapseButtonColor(ByVal objCollapseButtonColor As OLE_COLOR)
    m_CollapseButtonColor = objCollapseButtonColor
    Call UserControl.PropertyChanged("CollapseButtonColor")
    Call DrawControl
End Property
'
'
'Prodédures diverses et utiles
'definir le rectangle de chaque partie
Private Sub GetItemClientRect(tR As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long)
    SetRect tR, X1, Y1, X2, Y2
End Sub
'definir la zone région de la  fenêtre
Private Sub GetItemWindowReg(ByVal Region)
    SetWindowRgn UserControl.hwnd, Region, True
End Sub

'fonction qui donne la taille des bars: header et Footer
Private Function ReadSizeBar(bar As String) As Long
    Select Case bar
    Case "Header"
        Select Case m_iHeaderSize
        Case Small
            ReadSizeBar = 18
        Case Medium
            ReadSizeBar = 25
        Case Large
            ReadSizeBar = 40
        End Select

    Case "Footer"
        Select Case m_iFooterSize
        Case Small
            ReadSizeBar = 12
        Case Medium
            ReadSizeBar = 20
        Case Large
            ReadSizeBar = 34
        End Select
    End Select
End Function

Private Sub DrawControl()
    Dim tR As RECT
    Dim BorderStyle As Long
    Dim lWinRgn As Long
    Dim hPen As Long
    Dim hPenOld As Long
    Dim lRadius As Long
    Dim lHeaderSize As Long
    Dim lFooterSize As Long

    With UserControl
        .Cls
        'recupere le type de bordure selectionnée
        BorderStyle = m_eContainerBorderStyle
        'Taille du header
        lHeaderSize = ReadSizeBar("Header")
        'taille du Footer
        lFooterSize = ReadSizeBar("Footer")
        If m_eContainerShapeStyle = Rounded Then
            'angle pour rectangle arrondi
            lRadius = m_iContainerCornerRadius
        Else
            lRadius = 0
        End If

        'création la zone région de la  fenêtre - carré ou arrondi selon l'angle
        lWinRgn = CreateRoundRectRgn(0, 0, .ScaleWidth, .ScaleHeight, lRadius, lRadius)
        GetItemWindowReg lWinRgn


        If m_UseCustomColors Then
            '#définir la région du container
            GetItemClientRect tR, 0, 1, .ScaleWidth, ScaleHeight
            UtilDrawBackground .hDc, m_oContainerBackColor, m_oContainerFadeColor, tR.Left, tR.Top, tR.Right - tR.Left, tR.Bottom - tR.Top, m_eContainerFillStyle

            '#definit la region du Footer
            If m_bFooterVisible Then
                GetItemClientRect tR, 0, .ScaleHeight - lFooterSize, .ScaleWidth - 1, ScaleHeight
                UtilDrawBackground .hDc, m_oFooterBackColor, m_oFooterFadeColor, tR.Left, tR.Top, tR.Right, tR.Bottom, m_eFooterFillStyle
                DrawCaption tR, m_sFooterText, m_oFooterTextColor, m_fFooterTextFont, ValFooterTextAlign
            End If

            '#definit la region du header
            If m_bHeaderVisible Then
                GetItemClientRect tR, 0, 1, .ScaleWidth - 2, lHeaderSize
                UtilDrawBackground .hDc, m_oHeaderBackColor, m_oHeaderFadeColor, 0, 0, tR.Right, tR.Bottom, m_eHeaderFillStyle
                'texte
                'décalage en fonction de l'icone du header
                If Not m_HeaderPicture Is Nothing Then
                    If HeaderPicture <> 0 Then
                        tR.Left = tR.Left + 3 + m_HeaderPictureSize
                        DrawCaption tR, m_sHeaderText, m_oHeaderTextColor, m_fHeaderTextFont, ValHeaderTextAlign
                    End If
                End If
            End If


            'recupere la couleur de la bordure
            UserControl.ForeColor = m_oContainerBorderColor

            'création du stylo (personnalisé)
            hPen = CreatePen(BorderStyle, 1, m_oContainerBorderColor)
            hPenOld = SelectObject(.hDc, hPen)

        Else

            '#définir la région du container
            GetItemClientRect tR, 0, 1, .ScaleWidth, .ScaleHeight
            UtilDrawBackground .hDc, m_lColorContainerColorOne, m_lColorContainerColorTwo, tR.Left, tR.Top, tR.Right, tR.Bottom, m_eContainerFillStyle


            '#definit la region du Footer
            If m_bFooterVisible Then
                GetItemClientRect tR, 0, .ScaleHeight - lFooterSize, .ScaleWidth - 1, ScaleHeight
                UtilDrawBackground .hDc, m_lColorFooterColorOne, m_lColorFooterColorTwo, tR.Left, tR.Top, tR.Right, tR.Bottom, m_eFooterFillStyle
                DrawCaption tR, m_sFooterText, m_lColorFooterForeColor, m_fFooterTextFont, ValFooterTextAlign
            End If

            '#definit la region du header
            If m_bHeaderVisible Then
                GetItemClientRect tR, 0, 1, .ScaleWidth - 1, lHeaderSize
                UtilDrawBackground .hDc, m_lColorHeaderColorOne, m_lColorHeaderColorTwo, 0, 0, tR.Right, tR.Bottom, m_eHeaderFillStyle
                'texte
                'décalage en fonction de l'icone du header
                If Not m_HeaderPicture Is Nothing Then
                    If HeaderPicture <> 0 Then
                        tR.Left = tR.Left + 3 + m_HeaderPictureSize
                        DrawCaption tR, m_sHeaderText, m_lColorHeaderForeColor, m_fHeaderTextFont, ValHeaderTextAlign
                    End If
                End If
            End If

            'recupere la couleur de la bordure
            UserControl.ForeColor = m_lColorContainerColorBorder

            'création du stylo (theme)
            hPen = CreatePen(BorderStyle, 1, m_lColorContainerColorBorder)
            hPenOld = SelectObject(.hDc, hPen)
        End If

        'on trace la bordure du controle carré ou arrondi selon l'angle
        UtilDrawShapeStyle .hDc, 0, 0, .ScaleWidth - 1, .ScaleHeight - 1, lRadius
        'on detruit les objets
        SelectObject .hDc, hPenOld
        DeleteObject hPen
    End With



    '#dessin de l'image
    'Header
    If m_bHeaderVisible Then
        If Not m_HeaderPicture Is Nothing Then
            GetItemClientRect tR, 4, (lHeaderSize - m_HeaderPictureSize) / 2, m_HeaderPictureSize, lHeaderSize
            DrawPicture tR, m_HeaderPicture, m_HeaderPictureSize
        End If
        '#dessin du chevron
        If m_CanExpand Then Call DrawChevron
    End If
    'Container
    If Not m_ContainerPicture Is Nothing Then
        Dim W As Long, H As Long
        W = UserControl.ScaleWidth: H = UserControl.ScaleHeight
        If m_bFooterVisible Then H = H - lFooterSize
        If m_bHeaderVisible Then H = H - lHeaderSize
        GetItemClientRect tR, 4, UserControl.ScaleHeight - H, W, H
        DrawPicture tR, m_ContainerPicture, m_ContainerPictureSize
    End If


End Sub

Private Sub DrawCaption(ByRef rct As RECT, _
                        ByVal sCaption As String, _
                        ByVal lTextColor As OLE_COLOR, _
                        ByVal oTextFont As Font, _
                        ByVal Alignment As Long)

    Dim AlignmentCushion As Integer
    Dim oColor As OLE_COLOR           '--couleur du texte du caption
    Dim bFnt As StdFont               '--police


    AlignmentCushion = 3

    '--on recupere les parametres Font et ForeColor
    Set bFnt = Font
    Set Font = oTextFont
    oColor = UserControl.ForeColor
    UserControl.ForeColor = lTextColor

    '--le decalage (texte)
    With rct
        .Left = .Left + AlignmentCushion
        .Top = .Top + AlignmentCushion - 1
        .Right = .Right - 3
        .Bottom = .Bottom - 1
    End With


    '--décalage en fonction du bouton
    If m_CanExpand Then
        rct.Right = rct.Right - 20
    End If

    'dessin du texte
    DrawText UserControl.hDc, sCaption, Len(sCaption), rct, Alignment

    '--on restaure les parametres Font et ForeColor
    UserControl.ForeColor = oColor
    Set Font = bFnt
End Sub

'Gestion des différents thèmes
Private Sub GetGradientColors()
    If Not m_UseCustomColors Then
        Select Case m_enmTheme
        Case xThemeDarkBlue
            m_lColorHeaderColorOne = RGB(137, 170, 224)
            m_lColorHeaderColorTwo = RGB(7, 33, 100)
            m_lColorHeaderForeColor = RGB(215, 230, 251)
            m_lColorCollapseButtonColor = RGB(172, 191, 227)
            '
            m_lColorFooterColorOne = RGB(9, 42, 127)
            m_lColorFooterColorTwo = RGB(102, 145, 215)
            m_lColorFooterForeColor = RGB(215, 230, 251)
            '
            m_lColorContainerColorOne = RGB(142, 179, 231)
            m_lColorContainerColorTwo = RGB(142, 179, 231)
            m_lColorContainerColorBorder = RGB(1, 45, 150)

        Case xThemeMediaPlayer
            m_lColorHeaderColorOne = RGB(216, 228, 248)
            m_lColorHeaderColorTwo = RGB(49, 106, 197)
            m_lColorHeaderForeColor = RGB(215, 230, 251)
            m_lColorCollapseButtonColor = RGB(65, 105, 225)
            '
            m_lColorFooterColorOne = RGB(49, 106, 197)
            m_lColorFooterColorTwo = RGB(211, 229, 250)
            m_lColorFooterForeColor = RGB(215, 230, 251)
            '
            m_lColorContainerColorOne = RGB(243, 255, 255)
            m_lColorContainerColorTwo = RGB(142, 179, 231)
            m_lColorContainerColorBorder = RGB(192, 211, 240)

        Case xThemeGreen
            m_lColorHeaderColorOne = RGB(228, 235, 200)
            m_lColorHeaderColorTwo = RGB(175, 194, 142)
            m_lColorHeaderForeColor = RGB(100, 144, 88)
            m_lColorCollapseButtonColor = RGB(143, 188, 139)

            m_lColorFooterColorOne = RGB(165, 182, 121)
            m_lColorFooterColorTwo = RGB(233, 244, 207)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(233, 244, 207)
            m_lColorContainerColorTwo = RGB(233, 244, 207)
            m_lColorContainerColorBorder = RGB(100, 144, 88)

        Case xThemeLightBlue
            m_lColorHeaderColorOne = RGB(0, 45, 150)
            m_lColorHeaderColorTwo = RGB(89, 135, 214)
            m_lColorHeaderForeColor = RGB(172, 191, 227)
            m_lColorCollapseButtonColor = RGB(172, 191, 227)

            m_lColorFooterColorOne = RGB(255, 255, 255)
            m_lColorFooterColorTwo = RGB(201, 211, 243)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(246, 247, 248)
            m_lColorContainerColorTwo = RGB(246, 247, 248)
            m_lColorContainerColorBorder = RGB(124, 124, 148)

        Case xThemeMediaPlayer2
            m_lColorHeaderColorOne = RGB(255, 255, 255)
            m_lColorHeaderColorTwo = RGB(201, 211, 243)
            m_lColorHeaderForeColor = RGB(0, 41, 99)
            m_lColorCollapseButtonColor = RGB(172, 191, 227)

            m_lColorFooterColorOne = RGB(255, 255, 255)
            m_lColorFooterColorTwo = RGB(184, 205, 236)
            m_lColorFooterForeColor = RGB(0, 41, 99)

            m_lColorContainerColorOne = RGB(255, 255, 255)
            m_lColorContainerColorTwo = RGB(201, 211, 243)
            m_lColorContainerColorBorder = RGB(184, 205, 236)

        Case xThemeMetallic
            m_lColorHeaderColorOne = RGB(219, 220, 232)
            m_lColorHeaderColorTwo = RGB(149, 147, 177)
            m_lColorHeaderForeColor = RGB(119, 118, 151)
            m_lColorCollapseButtonColor = RGB(106, 90, 205)

            m_lColorFooterColorOne = RGB(149, 147, 177)
            m_lColorFooterColorTwo = RGB(207, 223, 239)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(232, 232, 232)
            m_lColorContainerColorTwo = RGB(232, 232, 232)
            m_lColorContainerColorBorder = RGB(119, 118, 151)

        Case xThemeOrange
            m_lColorHeaderColorOne = RGB(255, 122, 0)
            m_lColorHeaderColorTwo = RGB(130, 0, 0)
            m_lColorHeaderForeColor = RGB(255, 222, 173)
            m_lColorCollapseButtonColor = RGB(246, 172, 84)

            m_lColorFooterColorOne = RGB(180, 99, 1)
            m_lColorFooterColorTwo = RGB(130, 0, 0)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(255, 222, 173)
            m_lColorContainerColorTwo = RGB(224, 180, 97)
            m_lColorContainerColorBorder = RGB(139, 0, 0)

        Case xThemeTurquoise
            m_lColorHeaderColorOne = RGB(72, 209, 204)
            m_lColorHeaderColorTwo = RGB(43, 103, 109)
            m_lColorHeaderForeColor = RGB(233, 250, 248)
            m_lColorCollapseButtonColor = RGB(193, 240, 234)

            m_lColorFooterColorOne = RGB(0, 139, 139)
            m_lColorFooterColorTwo = RGB(0, 128, 128)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(224, 255, 255)
            m_lColorContainerColorTwo = RGB(224, 255, 255)
            m_lColorContainerColorBorder = RGB(65, 131, 111)

        Case xThemeGray
            m_lColorHeaderColorOne = RGB(192, 192, 192)
            m_lColorHeaderColorTwo = RGB(51, 51, 51)
            m_lColorHeaderForeColor = RGB(235, 235, 235)
            m_lColorCollapseButtonColor = RGB(220, 220, 220)

            m_lColorFooterColorOne = RGB(128, 128, 128)
            m_lColorFooterColorTwo = RGB(211, 211, 211)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(235, 235, 235)
            m_lColorContainerColorTwo = RGB(235, 235, 235)
            m_lColorContainerColorBorder = RGB(51, 51, 51)

        Case xThemeDarkBlue2
            m_lColorHeaderColorOne = RGB(81, 128, 208)
            m_lColorHeaderColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
            m_lColorHeaderForeColor = vbRed
            m_lColorCollapseButtonColor = RGB(191, 0, 0)

            m_lColorFooterColorOne = RGB(81, 128, 208)
            m_lColorFooterColorTwo = dBlendColor(RGB(11, 63, 153), vbBlack, 230)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(142, 179, 231)
            m_lColorContainerColorTwo = RGB(142, 179, 231)
            m_lColorContainerColorBorder = RGB(0, 45, 150)

        Case xThemeMoney
            m_lColorHeaderColorOne = RGB(160, 160, 160)
            m_lColorHeaderColorTwo = dBlendColor(RGB(90, 90, 90), vbBlack, 230)
            m_lColorHeaderForeColor = vbWhite
            m_lColorCollapseButtonColor = RGB(220, 220, 220)

            m_lColorFooterColorOne = RGB(169, 169, 169)
            m_lColorFooterColorTwo = RGB(105, 105, 105)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(112, 112, 112)
            m_lColorContainerColorTwo = RGB(112, 112, 112)
            m_lColorContainerColorBorder = RGB(68, 68, 68)

        Case xThemeOffice2003
            m_lColorHeaderColorOne = RGB(209, 227, 251)
            m_lColorHeaderColorTwo = RGB(106, 140, 203)
            m_lColorHeaderForeColor = RGB(110, 109, 143)
            m_lColorCollapseButtonColor = RGB(45, 80, 153)

            m_lColorFooterColorOne = RGB(176, 196, 222)
            m_lColorFooterColorTwo = RGB(100, 149, 237)
            m_lColorFooterForeColor = RGB(215, 230, 251)

            m_lColorContainerColorOne = RGB(255, 255, 255)
            m_lColorContainerColorTwo = RGB(255, 255, 255)
            m_lColorContainerColorBorder = RGB(0, 0, 128)
        End Select

    Else
        m_lColorHeaderColorOne = m_oHeaderBackColor
        m_lColorHeaderColorTwo = m_oHeaderFadeColor
        m_lColorHeaderForeColor = m_oHeaderTextColor
        m_lColorCollapseButtonColor = m_CollapseButtonColor

        m_lColorFooterColorOne = m_oFooterBackColor
        m_lColorFooterColorTwo = m_oFooterFadeColor
        m_lColorFooterForeColor = m_oFooterTextColor

        m_lColorContainerColorOne = m_oContainerBackColor
        m_lColorContainerColorOne = m_oContainerFadeColor
        m_lColorContainerColorBorder = m_oContainerBorderColor
    End If
End Sub
Private Sub DrawChevron()
    Dim tWorkR As RECT
    Dim hPen As Long
    Dim lHDC As Long
    Dim hPenOld As Long
    Dim tPoint As POINTAPI

    lHDC = UserControl.hDc

    '--Rectangle représentant le header
    GetItemClientRect tWorkR, 0, 1, UserControl.ScaleWidth - 1, ReadSizeBar("Header")

    tWorkR.Left = tWorkR.Right - 22
    tWorkR.Top = tWorkR.Top + (tWorkR.Bottom - tWorkR.Top - ReadSizeBar("Header")) \ 2 + 1
    tWorkR.Right = tWorkR.Left + 18
    tWorkR.Bottom = tWorkR.Top + 16

    If m_UseCustomColors Then
        hPen = CreatePen(0, 1, m_CollapseButtonColor)
        hPenOld = SelectObject(lHDC, hPen)
    Else
        hPen = CreatePen(0, 1, m_lColorCollapseButtonColor)
        hPenOld = SelectObject(lHDC, hPen)
    End If
    MoveToEx lHDC, tWorkR.Left + 1, tWorkR.Top + tWorkR.Bottom - 4, tPoint
    LineTo lHDC, tWorkR.Left + 1, tWorkR.Top + 1
    LineTo lHDC, tWorkR.Right - 2, tWorkR.Top + 1
    LineTo lHDC, tWorkR.Right - 2, tWorkR.Bottom - 2
    LineTo lHDC, tWorkR.Left + 1, tWorkR.Top + tWorkR.Bottom - 4

    '--on définit la région du chevron sensible à la souris
    m_hRegion = CreateRectRgn(tWorkR.Left + 1, tWorkR.Top + 1, tWorkR.Right - 2, tWorkR.Top + tWorkR.Bottom - 3)

    If (State = eBarExpanded) Then
        MoveToEx lHDC, tWorkR.Left + 5, tWorkR.Top + 7, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 4
        LineTo lHDC, tWorkR.Left + 12, tWorkR.Top + 8
        MoveToEx lHDC, tWorkR.Left + 6, tWorkR.Top + 7, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 5
        LineTo lHDC, tWorkR.Left + 11, tWorkR.Top + 8

        MoveToEx lHDC, tWorkR.Left + 5, tWorkR.Top + 11, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 8
        LineTo lHDC, tWorkR.Left + 12, tWorkR.Top + 12
        MoveToEx lHDC, tWorkR.Left + 6, tWorkR.Top + 11, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 9
        LineTo lHDC, tWorkR.Left + 11, tWorkR.Top + 12
    Else
        MoveToEx lHDC, tWorkR.Left + 5, tWorkR.Top + 4, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 7
        LineTo lHDC, tWorkR.Left + 12, tWorkR.Top + 3
        MoveToEx lHDC, tWorkR.Left + 6, tWorkR.Top + 4, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 6
        LineTo lHDC, tWorkR.Left + 11, tWorkR.Top + 3

        MoveToEx lHDC, tWorkR.Left + 5, tWorkR.Top + 8, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 11
        LineTo lHDC, tWorkR.Left + 12, tWorkR.Top + 7
        MoveToEx lHDC, tWorkR.Left + 6, tWorkR.Top + 8, tPoint
        LineTo lHDC, tWorkR.Left + 8, tWorkR.Top + 10
        LineTo lHDC, tWorkR.Left + 11, tWorkR.Top + 7
    End If

    SelectObject UserControl.hDc, hPenOld
    DeleteObject hPen
End Sub

' expande/collapse le control
Private Sub fExpandBar(ByVal iDir As Long)
    Dim lStart As Long
    Dim lTarget As Long
    Dim lPos As Long

    If (iDir > 0) Then

        lStart = ReadSizeBar("Header") * Screen.TwipsPerPixelY
        lTarget = CollapseOffset    'UserControl.Height
    Else
        lTarget = ReadSizeBar("Header") * Screen.TwipsPerPixelY
        lStart = CollapseOffset    'UserControl.Height
    End If

    lPos = lStart

    Do While Not (lPos = lTarget)
        lPos = lPos + iDir
        UserControl.Height = lPos
        '        DoEvents
    Loop

    If (iDir > 0) Then
        State = eBarcollapsed
    Else
        State = eBarExpanded
    End If
    UserControl.Refresh
End Sub

'
'
'===Usercontrol======================================================================================================

Private Sub UserControl_InitProperties()
    bInitializing = True

    Set m_fHeaderTextFont = Ambient.Font
    Set m_fFooterTextFont = Ambient.Font
    Set m_fContainerTextFont = Ambient.Font

    m_UseCustomColors = m_def_UseCustomColors
    m_CanExpand = False

    m_bHeaderVisible = m_def_bHeaderVisible
    m_bFooterVisible = m_def_bFooterVisible
    m_eHeaderFillStyle = m_def_eHeaderFillStyle
    m_eFooterFillStyle = m_def_eFooterFillStyle

    m_eContainerFillStyle = m_def_eContainerFillStyle

    m_oHeaderTextColor = m_def_oHeaderTextColor
    m_oFooterTextColor = m_def_oFooterTextColor
    m_oContainerForeColor = m_def_oContainerForeColor

    m_eHeaderTextAlign = m_def_eHeaderTextAlign
    m_eFooterTextAlign = m_def_eFooterTextAlign

    m_oHeaderBackColor = m_def_oHeaderBackColor
    m_oFooterBackColor = m_def_oFooterBackColor
    m_oContainerBackColor = m_def_oContainerBackColor

    m_oHeaderFadeColor = m_def_oHeaderFadeColor
    m_oFooterFadeColor = m_def_oFooterFadeColor
    m_oContainerFadeColor = m_def_oContainerFadeColor

    m_oContainerBorderColor = m_def_oContainerBorderColor
    m_eContainerBorderStyle = m_def_eContainerBorderStyle
    m_eContainerShapeStyle = m_def_eContainerShapeStyle
    m_iContainerCornerRadius = m_def_iContainerCornerRadius

    m_bEnabled = True
    m_enmTheme = m_def_iTheme
    m_iHeaderSize = m_def_eHeaderSize
    m_iFooterSize = m_def_eFooterSize

    m_CollapseButtonColor = m_def_CollapseButtonColor

    m_sHeaderText = Ambient.DisplayName

    m_HeaderPictureSize = 16
    Set m_HeaderPicture = LoadPicture()

    m_ContainerPictureSize = 32
    Set m_ContainerPicture = LoadPicture()

    bInitializing = False
    GetGradientColors
    DrawControl
End Sub


Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Set m_fHeaderTextFont = PropBag.ReadProperty("HeaderTextFont", Ambient.Font)
    Set m_fFooterTextFont = PropBag.ReadProperty("footerTextFont", Ambient.Font)
    Set m_fContainerTextFont = PropBag.ReadProperty("ContainerTextFont", Ambient.Font)

    m_UseCustomColors = PropBag.ReadProperty("UseCustomColors", m_def_UseCustomColors)
    m_CanExpand = PropBag.ReadProperty("Expandable", False)

    m_bHeaderVisible = PropBag.ReadProperty("HeaderVisible", m_def_bHeaderVisible)
    m_bFooterVisible = PropBag.ReadProperty("FooterVisible", m_def_bFooterVisible)

    m_sHeaderText = PropBag.ReadProperty("HeaderText", m_def_sHeaderText)
    m_sFooterText = PropBag.ReadProperty("FooterText", m_def_sFooterText)
    m_eHeaderTextAlign = PropBag.ReadProperty("HeaderTextAlign", m_def_eHeaderTextAlign)
    m_eFooterTextAlign = PropBag.ReadProperty("FooterTextAlign", m_def_eFooterTextAlign)

    m_eHeaderFillStyle = PropBag.ReadProperty("HeaderFillStyle", m_def_eHeaderFillStyle)
    m_eFooterFillStyle = PropBag.ReadProperty("FooterFillStyle", m_def_eFooterFillStyle)
    m_eContainerFillStyle = PropBag.ReadProperty("ContainerFillStyle", m_def_eContainerFillStyle)

    m_oHeaderTextColor = PropBag.ReadProperty("HeaderTextColor", m_def_oHeaderTextColor)
    m_oFooterTextColor = PropBag.ReadProperty("FooterTextColor", m_def_oFooterTextColor)
    m_oContainerForeColor = PropBag.ReadProperty("ContainerForeColor", m_def_oContainerForeColor)

    m_oHeaderBackColor = PropBag.ReadProperty("HeaderBackColor", m_def_oHeaderBackColor)
    m_oFooterBackColor = PropBag.ReadProperty("FooterBackColor", m_def_oFooterBackColor)
    m_oContainerBackColor = PropBag.ReadProperty("ContainerBackColor", m_def_oContainerBackColor)

    m_oHeaderFadeColor = PropBag.ReadProperty("HeaderFadeColor", m_def_oHeaderFadeColor)
    m_oFooterFadeColor = PropBag.ReadProperty("FooterFadeColor", m_def_oFooterFadeColor)
    m_oContainerFadeColor = PropBag.ReadProperty("ContainerFadeColor", m_def_oContainerFadeColor)

    m_oContainerBorderColor = PropBag.ReadProperty("ContainerBorderColor", m_def_oContainerBorderColor)

    m_eContainerBorderStyle = PropBag.ReadProperty("ContainerBorderStyle", m_def_eContainerBorderStyle)

    m_eContainerShapeStyle = PropBag.ReadProperty("ContainerShapeStyle", m_def_eContainerShapeStyle)
    m_iContainerCornerRadius = PropBag.ReadProperty("ContainerCornerRadius", m_def_iContainerCornerRadius)

    m_bEnabled = PropBag.ReadProperty("Enabled", True)
    m_enmTheme = PropBag.ReadProperty("Appearence", m_def_iTheme)

    m_iHeaderSize = PropBag.ReadProperty("HeaderSize", m_def_eHeaderSize)
    m_iFooterSize = PropBag.ReadProperty("FooterSize", m_def_eFooterSize)

    m_sHeaderText = PropBag.ReadProperty("HeaderText", m_def_sHeaderText)
    m_sFooterText = PropBag.ReadProperty("FooterText", m_def_sFooterText)

    m_CollapseButtonColor = PropBag.ReadProperty("CollapseButtonColor", m_def_CollapseButtonColor)

    Set m_HeaderPicture = PropBag.ReadProperty("HeaderPicture", Nothing)
    m_HeaderPictureSize = PropBag.ReadProperty("HeaderPictureSize", 16)

    Set m_ContainerPicture = PropBag.ReadProperty("ContainerPicture", Nothing)
    m_ContainerPictureSize = PropBag.ReadProperty("ContainerPictureSize", 32)
    
    GetGradientColors
    CollapseOffset = UserControl.Height
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("HeaderTextFont", m_fHeaderTextFont, Ambient.Font)
    Call PropBag.WriteProperty("footerTextFont", m_fFooterTextFont, Ambient.Font)
    Call PropBag.WriteProperty("ContainerTextFont", m_fContainerTextFont, Ambient.Font)

    Call PropBag.WriteProperty("UseCustomColors", m_UseCustomColors, m_def_UseCustomColors)
    Call PropBag.WriteProperty("Expandable", m_CanExpand, False)

    Call PropBag.WriteProperty("HeaderVisible", m_bHeaderVisible, m_def_bHeaderVisible)
    Call PropBag.WriteProperty("FooterVisible", m_bFooterVisible, m_def_bFooterVisible)
    Call PropBag.WriteProperty("HeaderText", m_sHeaderText, m_def_sHeaderText)
    Call PropBag.WriteProperty("FooterText", m_sFooterText, m_def_sFooterText)

    Call PropBag.WriteProperty("HeaderTextAlign", m_eHeaderTextAlign, m_def_eHeaderTextAlign)
    Call PropBag.WriteProperty("FooterTextAlign", m_eFooterTextAlign, m_def_eFooterTextAlign)

    Call PropBag.WriteProperty("HeaderFillStyle", m_eHeaderFillStyle, m_def_eHeaderFillStyle)
    Call PropBag.WriteProperty("FooterFillStyle", m_eFooterFillStyle, m_def_eFooterFillStyle)
    Call PropBag.WriteProperty("ContainerFillStyle", m_eContainerFillStyle, m_def_eContainerFillStyle)

    Call PropBag.WriteProperty("HeaderTextColor", m_oHeaderTextColor, m_def_oHeaderTextColor)
    Call PropBag.WriteProperty("FooterTextColor", m_oFooterTextColor, m_def_oFooterTextColor)
    Call PropBag.WriteProperty("ContainerForeColor", m_oContainerForeColor, m_def_oContainerForeColor)

    Call PropBag.WriteProperty("HeaderBackColor", m_oHeaderBackColor, m_def_oHeaderBackColor)
    Call PropBag.WriteProperty("FooterBackColor", m_oFooterBackColor, m_def_oFooterBackColor)
    Call PropBag.WriteProperty("ContainerBackColor", m_oContainerBackColor, m_def_oContainerBackColor)

    Call PropBag.WriteProperty("HeaderFadeColor", m_oHeaderFadeColor, m_def_oHeaderFadeColor)
    Call PropBag.WriteProperty("FooterFadeColor", m_oFooterFadeColor, m_def_oFooterFadeColor)
    Call PropBag.WriteProperty("ContainerFadeColor", m_oContainerFadeColor, m_def_oContainerFadeColor)

    Call PropBag.WriteProperty("ContainerBorderColor", m_oContainerBorderColor, m_def_oContainerBorderColor)
    Call PropBag.WriteProperty("ContainerBorderStyle", m_eContainerBorderStyle, m_def_eContainerBorderStyle)

    Call PropBag.WriteProperty("ContainerShapeStyle", m_eContainerShapeStyle, m_def_eContainerShapeStyle)
    Call PropBag.WriteProperty("ContainerCornerRadius", m_iContainerCornerRadius, m_def_iContainerCornerRadius)

    Call PropBag.WriteProperty("Appearence", m_enmTheme, m_def_iTheme)
    Call PropBag.WriteProperty("Enabled", m_bEnabled, True)

    Call PropBag.WriteProperty("HeaderSize", m_iHeaderSize, m_def_eHeaderSize)
    Call PropBag.WriteProperty("FooterSize", m_iFooterSize, m_def_eFooterSize)

    Call PropBag.WriteProperty("HeaderText", m_sHeaderText, m_def_sHeaderText)
    Call PropBag.WriteProperty("FooterText", m_sFooterText)

    Call PropBag.WriteProperty("CollapseButtonColor", m_CollapseButtonColor, m_def_CollapseButtonColor)

    Call PropBag.WriteProperty("HeaderPicture", m_HeaderPicture)
    Call PropBag.WriteProperty("HeaderPictureSize", m_HeaderPictureSize, 16)
        Call PropBag.WriteProperty("ContainerPicture", m_ContainerPicture)
    Call PropBag.WriteProperty("ContainerPictureSize", m_ContainerPictureSize, 32)
End Sub


Private Sub UserControl_Paint()
    DrawControl
End Sub

'
'
'===Evenements du controle======================================================================================================
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'RaiseEvent MouseDown(Button, Shift, x, y)

    Dim pt As POINTAPI
    If GetCursorPos(pt) Then
        If ScreenToClient(UserControl.hwnd, pt) Then
            If PtInRegion(m_hRegion, pt.X, pt.Y) Then
                'si le controle est expandable
                If (m_CanExpand) Then

                    If (State = eBarcollapsed) Then
                        'on collapse le header
                        fExpandBar -1
                    Else
                        ' on expand le header
                        fExpandBar 1
                    End If
                    DrawChevron
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub


