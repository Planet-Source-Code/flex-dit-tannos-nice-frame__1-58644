VERSION 5.00
Object = "*\A..\prjMaFrame2.vbp"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8715
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraHeader 
      Height          =   2655
      Left            =   4560
      TabIndex        =   6
      Top             =   240
      Width           =   4095
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1560
         TabIndex        =   49
         Top             =   1155
         Width           =   375
      End
      Begin VB.Frame Frame3 
         Caption         =   "Header Size"
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   1860
         Width           =   3855
         Begin VB.OptionButton optHeaderSize 
            Caption         =   "Medium"
            Height          =   375
            Index           =   1
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optHeaderSize 
            Caption         =   "Small"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optHeaderSize 
            Caption         =   "Large"
            Height          =   375
            Index           =   2
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.CheckBox chkHeaderVisible 
         Appearance      =   0  'Flat
         Caption         =   "Header Visible"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   40
         ToolTipText     =   $"test.frx":0000
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.CheckBox chkExpandable 
         Appearance      =   0  'Flat
         Caption         =   "Expandable"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   39
         ToolTipText     =   "Select to make the grid editable."
         Top             =   680
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Fill Style"
         Height          =   735
         Left            =   1560
         TabIndex        =   36
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton optFeaderFillStyle 
            Caption         =   "Horizontal"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFeaderFillStyle 
            Caption         =   "Vertical"
            Height          =   375
            Index           =   1
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2640
         TabIndex        =   33
         Top             =   1155
         Width           =   1335
      End
      Begin VB.ComboBox cboHeaderTextAlign 
         Height          =   315
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Picture Size"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1170
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "Text"
         Height          =   195
         Left            =   2160
         TabIndex        =   32
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         Caption         =   "Header"
         ForeColor       =   &H80000016&
         Height          =   240
         Index           =   4
         Left            =   0
         TabIndex        =   17
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label2 
         Caption         =   "Text Align"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   975
      End
   End
   Begin VB.Frame fraContainer 
      Height          =   2655
      Left            =   4560
      TabIndex        =   5
      Top             =   240
      Width           =   4095
      Begin VB.ComboBox cboBorderStyle 
         Height          =   315
         Left            =   2520
         TabIndex        =   30
         Top             =   1560
         Width           =   1455
      End
      Begin VB.ComboBox cboTheme 
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   2280
         Width           =   3855
      End
      Begin VB.Frame Frame6 
         Caption         =   "Shape Style"
         Height          =   735
         Left            =   120
         TabIndex        =   24
         Top             =   1200
         Width           =   2295
         Begin VB.OptionButton optContainerShapeStyle 
            Caption         =   "Rounded"
            Height          =   375
            Index           =   1
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optContainerShapeStyle 
            Caption         =   "Squared"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Fill Style"
         Height          =   735
         Left            =   1680
         TabIndex        =   21
         Top             =   360
         Width           =   2295
         Begin VB.OptionButton optContainerFillStyle 
            Caption         =   "Horizontal"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
         Begin VB.OptionButton optContainerFillStyle 
            Caption         =   "Vertical"
            Height          =   375
            Index           =   1
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   240
            Value           =   -1  'True
            Width           =   855
         End
      End
      Begin VB.CheckBox chkCustomColours 
         Appearance      =   0  'Flat
         Caption         =   "Custom Colours"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         ToolTipText     =   "Toggles a custom colour set for the grid."
         Top             =   840
         Width           =   1815
      End
      Begin VB.CheckBox chkEnabled 
         Appearance      =   0  'Flat
         Caption         =   "E&nabled"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         ToolTipText     =   "Enable or disable the grid."
         Top             =   480
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Border Style"
         Height          =   255
         Left            =   2520
         TabIndex        =   31
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Theme"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   1080
         TabIndex        =   28
         Top             =   480
         Width           =   45
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         Caption         =   "Container"
         ForeColor       =   &H80000016&
         Height          =   240
         Index           =   0
         Left            =   0
         TabIndex        =   18
         Top             =   120
         Width           =   4065
      End
   End
   Begin VB.Frame fraFooter 
      Height          =   2655
      Left            =   4560
      TabIndex        =   9
      Top             =   240
      Width           =   4095
      Begin VB.Frame Frame1 
         Caption         =   "Footer Size"
         Height          =   735
         Left            =   120
         TabIndex        =   41
         Top             =   1860
         Width           =   3855
         Begin VB.OptionButton optFooterSize 
            Caption         =   "Large"
            Height          =   375
            Index           =   2
            Left            =   2640
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFooterSize 
            Caption         =   "Small"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFooterSize 
            Caption         =   "Medium"
            Height          =   375
            Index           =   1
            Left            =   1320
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1560
         TabIndex        =   34
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CheckBox chkFooterVisible 
         Appearance      =   0  'Flat
         Caption         =   "Footer Visible"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         ToolTipText     =   $"test.frx":008C
         Top             =   600
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.ComboBox cboFooterTextAlign 
         Height          =   315
         Left            =   1560
         TabIndex        =   13
         Top             =   1560
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Caption         =   "Fill Style"
         Height          =   735
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   2415
         Begin VB.OptionButton optFooterFillStyle 
            Caption         =   "Horizontal"
            Height          =   375
            Index           =   0
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optFooterFillStyle 
            Caption         =   "Vertical"
            Height          =   375
            Index           =   1
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.Label Label8 
         Caption         =   "Text"
         Height          =   195
         Left            =   120
         TabIndex        =   35
         Top             =   1200
         Width           =   315
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         Caption         =   "Footer"
         ForeColor       =   &H80000016&
         Height          =   240
         Index           =   3
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   4065
      End
      Begin VB.Label Label4 
         Caption         =   "Text Align"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
   End
   Begin prjMaFrame2.MyNiceFrame MyNiceFrame1 
      Height          =   2895
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   5106
      BeginProperty HeaderTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty footerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty ContainerTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Expandable      =   -1  'True
      HeaderText      =   "Cool ma frame ?"
      FooterText      =   ""
      ContainerBackColor=   12632319
      ContainerFadeColor=   16744703
      Appearence      =   7
      HeaderText      =   "Cool ma frame ?"
      FooterText      =   ""
      HeaderPicture   =   "test.frx":0118
      HeaderPictureSize=   24
      ContainerPicture=   "test.frx":0531
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000D&
      Caption         =   "Quitter"
      Height          =   495
      Left            =   7680
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Header"
      Height          =   495
      Left            =   4560
      TabIndex        =   2
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Footer"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Container"
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label12 
      Caption         =   "Alignement de l'image dans le container"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   54
      Top             =   2160
      Width           =   3495
   End
   Begin VB.Label Label11 
      Caption         =   "Plus de possibilités de ""Gradient"" (=angle, direction et plus de couleurs)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   53
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Plus de ""Shape"" pour le contour"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   52
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "A venir:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   720
      Width           =   675
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboBorderStyle_Click()
Me.MyNiceFrame1.ContainerBorderStyle = Me.cboBorderStyle.ListIndex
End Sub

Private Sub cboheaderTextAlign_Click()
Me.MyNiceFrame1.headerTextAlign = Me.cboHeaderTextAlign.ListIndex
End Sub

Private Sub cboFooterTextAlign_Click()
Me.MyNiceFrame1.FooterTextAlign = Me.cboFooterTextAlign.ListIndex
End Sub

Private Sub cboTheme_Click()
Me.MyNiceFrame1.Appearence = Me.cboTheme.ListIndex
End Sub

Private Sub chkCustomColours_Click()
Me.MyNiceFrame1.UseCustomColors = (Me.chkCustomColours.Value = Checked)
End Sub

Private Sub chkEnabled_Click()
Me.MyNiceFrame1.Enabled = (Me.chkEnabled.Value = Checked)
End Sub

Private Sub chkExpandable_Click()
Me.MyNiceFrame1.Expandable = (Me.chkExpandable.Value = Checked)
End Sub

Private Sub chkFooterVisible_Click()
Me.MyNiceFrame1.FooterVisible = (Me.chkFooterVisible.Value = Checked)
End Sub

Private Sub chkHeaderVisible_Click()
Me.MyNiceFrame1.HeaderVisible = (Me.chkHeaderVisible.Value = Checked)
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Me.fraHeader.ZOrder
End Sub

Private Sub Command4_Click()
Me.fraContainer.ZOrder
End Sub

Private Sub Command5_Click()
Me.fraFooter.ZOrder
End Sub

Private Sub Form_Load()
    With Me.cboHeaderTextAlign
        .AddItem "xAlignLefttop"
        .AddItem "xAlignLeftMiddle"
        .AddItem "xAlignLeftBottom"
        .AddItem "xAlignRightTop"
        .AddItem "xAlignRightMiddle"
        .AddItem "xAlignRightBottom"
        .AddItem "xAlignCenterTop"
        .AddItem "xAlignCenterMiddle"
        .AddItem "xAlignCenterBottom"
        .ListIndex = 0
    End With
    With Me.cboFooterTextAlign
        .AddItem "xAlignLefttop"
        .AddItem "xAlignLeftMiddle"
        .AddItem "xAlignLeftBottom"
        .AddItem "xAlignRightTop"
        .AddItem "xAlignRightMiddle"
        .AddItem "xAlignRightBottom"
        .AddItem "xAlignCenterTop"
        .AddItem "xAlignCenterMiddle"
        .AddItem "xAlignCenterBottom"
        .ListIndex = 0
    End With
    With Me.cboBorderStyle
        .AddItem "SOLID"
        .AddItem "DASH"
        .AddItem "DOT"
        .AddItem "DASHDOT"
        .AddItem "DASHDOTDOT"
        .AddItem "NONE"
        .ListIndex = 0
    End With
    With Me.cboTheme
        .AddItem "xThemeDarkBlue"
        .AddItem "xThemeMoney"
        .AddItem "xThemeMediaPlayer"
        .AddItem "xThemeMediaPlayer2"
        .AddItem "xThemeGreen"
        .AddItem "xThemeMetallic"
        .AddItem "xThemeOffice2003"
        .AddItem "xThemeOrange"
        .AddItem "xThemeTurquoise"
        .AddItem "xThemeGray"
        .AddItem "xThemeDarkBlue2"
        .AddItem "xThemeLightBlue"
        .ListIndex = 0
    End With

End Sub

Private Sub optContainerFillStyle_Click(Index As Integer)
Me.MyNiceFrame1.ContainerFillStyle = Index
End Sub

Private Sub optContainerShapeStyle_Click(Index As Integer)
Me.MyNiceFrame1.ContainerShapeStyle = Index
End Sub

Private Sub optFeaderFillStyle_Click(Index As Integer)
Me.MyNiceFrame1.HeaderFillStyle = Index
End Sub

Private Sub optFooterFillStyle_Click(Index As Integer)
Me.MyNiceFrame1.FooterFillStyle = Index
End Sub

Private Sub optFooterSize_Click(Index As Integer)
Me.MyNiceFrame1.FooterSize = Index
End Sub

Private Sub optHeaderSize_Click(Index As Integer)
Me.MyNiceFrame1.HeaderSize = Index
End Sub

Private Sub Text1_Change()
On Error Resume Next
Me.MyNiceFrame1.HeaderPictureSize = Me.Text1.Text
End Sub

Private Sub Text2_Change()
Me.MyNiceFrame1.HeaderText = Me.Text2.Text
End Sub

Private Sub Text3_Change()
Me.MyNiceFrame1.FooterText = Me.Text3.Text
End Sub
