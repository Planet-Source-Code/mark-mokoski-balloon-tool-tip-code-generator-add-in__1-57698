VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolTips 
   Caption         =   "Tool Tip Code Builder"
   ClientHeight    =   7515
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   9135
   Icon            =   "ToolTip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   9135
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   49
      Top             =   7020
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   873
      ShowTips        =   0   'False
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5186
            Picture         =   "ToolTip.frx":0442
            Text            =   "Visit PSC for code updates"
            TextSave        =   "Visit PSC for code updates"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5186
            Picture         =   "ToolTip.frx":0894
            Text            =   "Hide Tool Tip Builder Add In"
            TextSave        =   "Hide Tool Tip Builder Add In"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   1
            Bevel           =   2
            Object.Width           =   5186
            Picture         =   "ToolTip.frx":116E
            Text            =   "Close Tool Tip Builder Add In"
            TextSave        =   "Close Tool Tip Builder Add In"
         EndProperty
      EndProperty
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frameCodeGen 
      Caption         =   "CodeGen Output"
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   120
      TabIndex        =   25
      Top             =   3600
      Width           =   8895
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   4880
         Picture         =   "ToolTip.frx":15C0
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   48
         Top             =   2400
         Width           =   495
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   4880
         Picture         =   "ToolTip.frx":1A02
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   47
         Top             =   1440
         Width           =   495
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   4880
         Picture         =   "ToolTip.frx":1E44
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         Top             =   840
         Width           =   495
      End
      Begin VB.PictureBox picArrow 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   4880
         Picture         =   "ToolTip.frx":2286
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   45
         Top             =   240
         Width           =   495
      End
      Begin VB.CommandButton cmdTipCode 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy"
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
         Left            =   5400
         MouseIcon       =   "ToolTip.frx":26C8
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2400
         Width           =   615
      End
      Begin VB.CommandButton cmdTipDIM 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy"
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
         Left            =   5400
         MouseIcon       =   "ToolTip.frx":29D2
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   1440
         Width           =   615
      End
      Begin VB.CommandButton cmdAPIsubcall 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy"
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
         Left            =   5400
         MouseIcon       =   "ToolTip.frx":2CDC
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   840
         Width           =   615
      End
      Begin VB.CommandButton cmdAPIDeclare 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copy"
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
         Left            =   5400
         MouseIcon       =   "ToolTip.frx":2FE6
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   240
         Width           =   615
      End
      Begin RichTextLib.RichTextBox CodeText 
         Height          =   1215
         Left            =   120
         TabIndex        =   36
         Top             =   2040
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   2143
         _Version        =   393217
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   32767
         TextRTF         =   $"ToolTip.frx":32F0
      End
      Begin RichTextLib.RichTextBox DeclareText 
         Height          =   495
         Left            =   120
         TabIndex        =   35
         Top             =   1440
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"ToolTip.frx":3372
      End
      Begin RichTextLib.RichTextBox CallSubText 
         Height          =   495
         Left            =   120
         TabIndex        =   34
         Top             =   840
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"ToolTip.frx":33F4
      End
      Begin RichTextLib.RichTextBox APItext 
         Height          =   495
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   873
         _Version        =   393217
         ReadOnly        =   -1  'True
         TextRTF         =   $"ToolTip.frx":3476
      End
      Begin VB.Frame frameCustom 
         Caption         =   "Customize Your Tool Tip"
         ForeColor       =   &H00000080&
         Height          =   3015
         Left            =   6120
         TabIndex        =   26
         Top             =   240
         Width           =   2655
         Begin VB.TextBox ParentNameText 
            Height          =   285
            Left            =   120
            TabIndex        =   30
            Top             =   2520
            Width           =   2400
         End
         Begin VB.TextBox TipNameText 
            Height          =   285
            Left            =   120
            TabIndex        =   29
            Top             =   1800
            Width           =   2400
         End
         Begin VB.CommandButton cmdCustomize 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Customize !"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   1080
            MouseIcon       =   "ToolTip.frx":34F8
            MousePointer    =   99  'Custom
            Picture         =   "ToolTip.frx":3802
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   240
            Width           =   1335
         End
         Begin VB.PictureBox picCustomize 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   360
            MouseIcon       =   "ToolTip.frx":40CC
            MousePointer    =   99  'Custom
            Picture         =   "ToolTip.frx":43D6
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   27
            Top             =   480
            Width           =   495
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Enter your Control Names Below"
            ForeColor       =   &H00000080&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label labelControlName 
            Alignment       =   2  'Center
            Caption         =   " Your Parent Control Name"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   2280
            Width           =   2415
         End
         Begin VB.Label labelTipName 
            Alignment       =   2  'Center
            Caption         =   "Your Tool Tip Name"
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1560
            Width           =   2415
         End
      End
   End
   Begin VB.Frame frameStyle 
      Caption         =   "Tool Tip Style"
      ForeColor       =   &H00000080&
      Height          =   1695
      Left            =   6480
      TabIndex        =   17
      Top             =   0
      Width           =   2535
      Begin VB.PictureBox picCenter 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         MouseIcon       =   "ToolTip.frx":4818
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":4B22
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   43
         Top             =   1200
         Width           =   495
      End
      Begin VB.PictureBox picStyle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   120
         MouseIcon       =   "ToolTip.frx":4E2C
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":5136
         ScaleHeight     =   495
         ScaleWidth      =   615
         TabIndex        =   42
         Top             =   640
         Width           =   615
      End
      Begin VB.PictureBox picStyle 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   120
         MouseIcon       =   "ToolTip.frx":5440
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":574A
         ScaleHeight     =   375
         ScaleWidth      =   495
         TabIndex        =   41
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton optionStyle 
         Caption         =   "Balloon Tip"
         Height          =   195
         Index           =   1
         Left            =   840
         MouseIcon       =   "ToolTip.frx":5A54
         MousePointer    =   99  'Custom
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optionStyle 
         Caption         =   "Rectangular Tip"
         Height          =   195
         Index           =   0
         Left            =   840
         MouseIcon       =   "ToolTip.frx":5D5E
         MousePointer    =   99  'Custom
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   360
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.CheckBox checkCenter 
         Caption         =   "Center Tool Tip"
         Height          =   195
         Left            =   840
         MouseIcon       =   "ToolTip.frx":6068
         MousePointer    =   99  'Custom
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Frame frameToolTipTitle 
      Caption         =   "Tool Tip Title"
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   2880
      Width           =   6015
      Begin VB.TextBox textTitle 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   1920
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   140
         Width           =   3975
      End
      Begin VB.Label labelTitleinfo 
         Caption         =   "(Title needed for Icons)"
         Height          =   200
         Left            =   120
         TabIndex        =   24
         Top             =   200
         Width           =   1695
      End
   End
   Begin VB.Frame frameToolTipText 
      Caption         =   "Tool Tip Text"
      ForeColor       =   &H00000080&
      Height          =   1575
      Left            =   240
      TabIndex        =   11
      Top             =   1320
      Width           =   6015
      Begin VB.TextBox BackColorRGB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   22
         Text            =   "Text6"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox ForeColorRGB 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   21
         Text            =   "Text5"
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdBackColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Background Color"
         Height          =   375
         Left            =   120
         MouseIcon       =   "ToolTip.frx":6372
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdForeColor 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Text Color"
         Height          =   375
         Left            =   120
         MouseIcon       =   "ToolTip.frx":667C
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox textTip 
         BackColor       =   &H80000018&
         Height          =   1245
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "ToolTip.frx":6986
         Top             =   240
         Width           =   3975
      End
   End
   Begin VB.Frame frameIcon 
      Caption         =   "Select Tip Icon"
      ForeColor       =   &H00000080&
      Height          =   1815
      Left            =   6480
      TabIndex        =   3
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton optionIcon 
         Caption         =   "None"
         Height          =   255
         Index           =   0
         Left            =   120
         MouseIcon       =   "ToolTip.frx":6997
         MousePointer    =   99  'Custom
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   600
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton optionIcon 
         Caption         =   "Warning"
         Height          =   255
         Index           =   1
         Left            =   1320
         MouseIcon       =   "ToolTip.frx":6CA1
         MousePointer    =   99  'Custom
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   600
         Width           =   975
      End
      Begin VB.PictureBox iconWarning 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         MouseIcon       =   "ToolTip.frx":6FAB
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":72B5
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   120
         Width           =   495
      End
      Begin VB.OptionButton optionIcon 
         Caption         =   "Information"
         Height          =   255
         Index           =   2
         Left            =   120
         MouseIcon       =   "ToolTip.frx":76F7
         MousePointer    =   99  'Custom
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.PictureBox iconInfo 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   480
         MouseIcon       =   "ToolTip.frx":7A01
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":7D0B
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
      Begin VB.OptionButton optionIcon 
         Caption         =   "ERROR"
         Height          =   255
         Index           =   3
         Left            =   1320
         MouseIcon       =   "ToolTip.frx":814D
         MousePointer    =   99  'Custom
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1440
         Width           =   1095
      End
      Begin VB.PictureBox iconError 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   1560
         MouseIcon       =   "ToolTip.frx":8457
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":8761
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   960
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Frame frameTipBuild 
      Caption         =   "Change Test Tool Tip Parameters"
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   6255
      Begin VB.CommandButton cmdCodeGen 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Click To Generate Code"
         Height          =   950
         Left            =   4240
         MouseIcon       =   "ToolTip.frx":8BA3
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":8EAD
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   240
         Width           =   1870
      End
      Begin VB.CommandButton cmdTest 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test Your Tool Tip Here!"
         CausesValidation=   0   'False
         Height          =   950
         Left            =   2160
         MouseIcon       =   "ToolTip.frx":92EF
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":95F9
         Style           =   1  'Graphical
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   240
         Width           =   1870
      End
      Begin VB.CommandButton cmdApply 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apply Changes"
         Height          =   950
         Left            =   120
         MouseIcon       =   "ToolTip.frx":9D3B
         MousePointer    =   99  'Custom
         Picture         =   "ToolTip.frx":A045
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   240
         Width           =   1870
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Close Add In"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmToolTips"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    '**************************************************************
    '
    '   Custom Tool Tip Demo
    '
    '   Mark Mokoski
    '   16-NOV-2004
    '
    '   See clsToolTips Class Module for details
    '
    '**************************************************************

    Option Explicit

    'Make new tool tip objects for this project

    Dim cmdCustomizetip              As New clsTooltips
    Dim cmdApplyTip                  As New clsTooltips
    Dim cmdAPIDeclareTip             As New clsTooltips
    Dim cmdTestTip                   As New clsTooltips
    Dim cmdCodeGenTip                As New clsTooltips
    Dim cmdAPIsubcallTip             As New clsTooltips
    Dim cmdTipDIMTip                 As New clsTooltips
    Dim cmdTipCodeTip                As New clsTooltips
    Dim textTipTip                   As New clsTooltips
    Dim textTitleTip                 As New clsTooltips
    Dim ForeColorRGBtip              As New clsTooltips
    Dim BackColorRGBtip              As New clsTooltips
    Dim picCustomizeTip              As New clsTooltips
    Dim picStyleTip(1)               As New clsTooltips
    Dim picCenterTip                 As New clsTooltips
    Dim checkCenterTip               As New clsTooltips
    Dim optionStyleTip(1)            As New clsTooltips
    
    'Public Var's used in this and other modules
    Public TipText                   As String
    Public TipTitleText              As String
    Public TipCentered               As Boolean
    Public TipStyle                  As toolStyleEnum
    Public TipIcon                   As toolIconType
    Public TipForeColor              As Long
    Public TipBackColor              As Long
    Public VBInstance                As VBIDE.VBE
    Public Connect                   As Connect

    

Private Sub checkCenter_Click()

    TipCentered = CBool(checkCenter.Value)

        Select Case CBool(checkCenter.Value)
            Case False
                optionStyleTip(0).CreateTip optionStyle(0), "Tip looks like this."
                optionStyleTip(0).Centered = False
                picStyleTip(0).CreateTip picStyle(0), "Tip looks like this."
                picStyleTip(0).Centered = False
                optionStyleTip(1).CreateBalloon optionStyle(1), "Tip looks like this."
                optionStyleTip(1).Centered = False
                picStyleTip(1).CreateBalloon picStyle(1), "Tip looks like this."
                picStyleTip(1).Centered = False
                picStyle(0).MouseIcon = LoadResPicture(101, 2)
                picStyle(1).MouseIcon = LoadResPicture(101, 2)
                optionStyle(0).MouseIcon = LoadResPicture(101, 2)
                optionStyle(1).MouseIcon = LoadResPicture(101, 2)
                
            Case True
                optionStyleTip(0).CreateTip optionStyle(0), "Tip looks like this."
                optionStyleTip(0).Centered = True
                picStyleTip(0).CreateTip picStyle(0), "Tip looks like this."
                picStyleTip(0).Centered = True
                optionStyleTip(1).CreateBalloon optionStyle(1), "Tip looks like this."
                optionStyleTip(1).Centered = True
                picStyleTip(1).CreateBalloon picStyle(1), "Tip looks like this."
                picStyleTip(1).Centered = True
                picStyle(0).MouseIcon = LoadResPicture(102, 2)
                picStyle(1).MouseIcon = LoadResPicture(102, 2)
                optionStyle(0).MouseIcon = LoadResPicture(102, 2)
                optionStyle(1).MouseIcon = LoadResPicture(102, 2)

        End Select

End Sub

Private Sub cmdCustomize_Click()

    Call GenCode
    
End Sub

Private Sub cmdTipCode_Click()

    Clipboard.Clear
    CodeText.SelStart = (InStr(1, CodeText.Text, vbCrLf)) + 1
    CodeText.SelLength = ((Len(CodeText.Text) + 1) - CodeText.SelStart)
    Clipboard.SetText CodeText.SelText
    CodeText.SelStart = 0
    
End Sub

Private Sub cmdApply_Click()

        If textTip.Text = "" And textTitle.Text = "" Then
            MsgBox "Tool Tip Text ERROR" & vbCrLf & vbCrLf & _
            "Tool Tip Text is Blank" & vbCrLf & _
            "For proper Tool Tip operation, Tip Text and/or a Tip Title is needed", vbCritical, "Tool Tip ERROR"
            Exit Sub
        Else
           
            'Change Tool Tip Text and other properties

                If textTip.Text = "" Then
                    cmdTestTip.TipText = " "
                Else
                    cmdTestTip.TipText = textTip.Text
                End If

                Select Case TipStyle
                    Case styleStandard
                        cmdTest.MouseIcon = LoadResPicture(103, 2)

                    Case styleBalloon

                        If TipCentered = True Then
                            cmdTest.MouseIcon = LoadResPicture(103, 2)
                        Else
                            cmdTest.MouseIcon = LoadResPicture(104, 2)
                        End If
                        
                End Select

            cmdTestTip.Title = TipTitleText
            cmdTestTip.Style = TipStyle
            cmdTestTip.Centered = TipCentered
            cmdTestTip.Icon = TipIcon
            cmdTestTip.ForeColor = TipForeColor
            cmdTestTip.BackColor = TipBackColor
        End If

End Sub

Private Sub cmdAPIDeclare_Click()

    Clipboard.Clear
    APItext.SelStart = (InStr(1, APItext.Text, vbCrLf)) + 1
    APItext.SelLength = ((Len(APItext.Text) + 1) - APItext.SelStart)
    Clipboard.SetText APItext.SelText & vbCrLf & vbCrLf
    APItext.SelStart = 0
    
End Sub

Private Sub cmdTest_Click()

    cmdApply.SetFocus
    
End Sub

Private Sub cmdForeColor_Click()

    'Set new tip fore color
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.Flags = cdlCCRGBInit
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the form's foreground color to selected color
    textTip.ForeColor = CommonDialog1.Color
    textTitle.ForeColor = CommonDialog1.Color
    TipForeColor = CommonDialog1.Color
    
    ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    BackColorRGB.Text = "&H" & Hex(TipBackColor)
    
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    BackColorRGB.Text = "&H" & Hex(TipBackColor)

End Sub

Private Sub cmdBackColor_Click()

    'Set new tip back color
    'Set Cancel to True
    CommonDialog1.CancelError = True
    
    On Error GoTo ErrHandler
    
    'Set the Flags property
    CommonDialog1.Flags = cdlCCRGBInit
    
    'Display the Color Dialog box
    CommonDialog1.ShowColor
    
    'Set the form's background color to selected color
    textTip.BackColor = CommonDialog1.Color
    textTitle.BackColor = CommonDialog1.Color
  
    'Since 0 is Black (no RGB), and the API thinks 0 is
    'the default color ("off" yeleow),
    'we need to "fudge" Black a bit (yes set bit "1" to "1",)
    'I couldn't resist the pun!
    
        If CommonDialog1.Color = 0 Then
            TipBackColor = &H80000008
        Else
            TipBackColor = CommonDialog1.Color
        End If
    
    ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    BackColorRGB.Text = "&H" & Hex(TipBackColor)
    
    Exit Sub

ErrHandler:
    ' User pressed the Cancel button
    ForeColorRGB.Text = "&H" & Hex(TipForeColor)
    BackColorRGB.Text = "&H" & Hex(TipBackColor)

End Sub

Private Sub cmdCodeGen_Click()

        If textTip.Text = "" Then
            MsgBox "Tool Tip Text ERROR" & vbCrLf & vbCrLf & _
            "Tool Tip Text is Blank" & vbCrLf & _
            "For proper Tool Tip operation, Tip Text is needed", vbCritical, "Tool Tip ERROR"
            Exit Sub
        Else
            Call GenCode
        End If
    
End Sub



Private Sub cmdAPIsubcall_Click()

    Clipboard.Clear
    CallSubText.SelStart = (InStr(1, CallSubText.Text, vbCrLf)) + 1
    CallSubText.SelLength = ((Len(CallSubText.Text) + 1) - CallSubText.SelStart)
    Clipboard.SetText CallSubText.SelText & vbCrLf & vbCrLf
    CallSubText.SelStart = 0
    
End Sub

Private Sub cmdTipDIM_Click()

    Clipboard.Clear
    DeclareText.SelStart = (InStr(1, DeclareText.Text, vbCrLf)) + 1
    DeclareText.SelLength = ((Len(DeclareText.Text) + 1) - DeclareText.SelStart)
    Clipboard.SetText DeclareText.SelText & vbCrLf & vbCrLf
    DeclareText.SelStart = 0
    
End Sub

Private Sub Form_Load()
    
    Dim X            As Integer

    'Make Tool Tip objects
    cmdApplyTip.CreateBalloon cmdApply, "Type in new Tip Text, Title and" + vbCrLf + "choose the other parameters." + vbCrLf + "Use more than one line of text if you want." + vbCrLf + "Click to apply your changes " + vbCrLf + "and test the Tool Tip", "Balloon Tip", tipiconinfo
    
    cmdTestTip.CreateTip cmdTest, "Go Ahead, Change ME!"
    
    textTipTip.CreateBalloon textTip, "Enter Tool Tip Text Here" + vbCrLf + "Double Click to restore default colors", "Test Tip Text", tipiconinfo
    
    textTitleTip.CreateBalloon textTitle, "Enter Tool Tip Title Here" + vbCrLf + "By entering a Title," + vbCrLf + "you enable the Tip Icon selection", "Test Title Text", tipiconinfo
    
    ForeColorRGBtip.CreateBalloon ForeColorRGB, "ForeColor RGB Hex Code"
    
    BackColorRGBtip.CreateBalloon BackColorRGB, "Backcolor RGB Hex Code"
    
    cmdCustomizetip.CreateBalloon cmdCustomize, _
    "Put the Tool Tip name and Parent Control name in the text boxes below." & _
    vbCrLf & _
    vbCrLf & _
    "Then Click the Customize buttom to update the code snippet", _
    "Tool Tip Code Builder", _
    tipiconinfo
    
    picCustomizeTip.CreateBalloon picCustomize, _
    "Put the Tool Tip name and Parent Control name in the text boxes below." & _
    vbCrLf & _
    vbCrLf & _
    "Then Click the Customize buttom to update the code snippet", _
    "Tool Tip Code Builder", _
    tipiconinfo
    
    checkCenterTip.CreateTip checkCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
    checkCenterTip.Centered = True
    
    picCenterTip.CreateTip picCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
    picCenterTip.Centered = True
    
    optionStyleTip(0).CreateTip optionStyle(0), "Tip looks like this."
    
    picStyleTip(0).CreateTip picStyle(0), "Tip looks like this."
    
    optionStyleTip(1).CreateBalloon optionStyle(1), "Tip looks like this."
    
    picStyleTip(1).CreateBalloon picStyle(1), "Tip looks like this."
    
    cmdAPIDeclareTip.CreateBalloon cmdAPIDeclare, _
    "Only use this once per Project." & vbCrLf & _
    "" & vbCrLf & _
    "Click to copy to Clipboard," & vbCrLf & _
    "Then Paste into your application.", _
    "API Declare", 2
    cmdAPIDeclareTip.ForeColor = &HFFFFFF
    cmdAPIDeclareTip.BackColor = &HC08000

    
    cmdAPIsubcallTip.CreateBalloon cmdAPIsubcall, _
    "Only use this once per Project." & vbCrLf & _
    "" & vbCrLf & _
    "Click to copy to Clipboard," & vbCrLf & _
    "Then Paste into your application.", _
    "API Sub Call", 2
    cmdAPIsubcallTip.ForeColor = &HFFFFFF
    cmdAPIsubcallTip.BackColor = &HC08000

    
    cmdTipDIMTip.CreateBalloon cmdTipDIM, _
    "One ""Dim"" per Tool Tip" & vbCrLf & _
    "" & vbCrLf & _
    "Click to copy to Clipboard," & vbCrLf & _
    "Then Paste into your application.", _
    "Create New Tool Tip Object", 1

    
    cmdTipCodeTip.CreateBalloon cmdTipCode, _
    "Code for your Custom Tip" & vbCrLf & _
    "" & vbCrLf & _
    "Click to copy to Clipboard," & vbCrLf & _
    "Then Paste into your application.", _
    "Tool Tip Code", 1

    
    'Code below make with this App's CodeGen feature!
    cmdCodeGenTip.CreateBalloon cmdCodeGen, _
    "Click Here to make the code" & vbCrLf & _
    "for your custom ToolTip." & vbCrLf & _
    "" & vbCrLf & _
    "Cut and Paste into your project!", _
    "Create Code", 1

    cmdCodeGenTip.ForeColor = &HEFEFEF
    cmdCodeGenTip.BackColor = &HC08000

    'Set up what controls are active
    cmdApply.Enabled = True

        For X = 0 To 3
            optionIcon(X).Enabled = False
        Next X
        
    
    'Set start values
    ForeColorRGB.Text = "&H" & Hex(textTip.ForeColor)
    BackColorRGB.Text = "&H" & Hex(textTip.BackColor)
    TipForeColor = 0
    TipBackColor = 0
    TipText = textTip.Text
    cmdCustomize.Enabled = False
    cmdCustomize.BackColor = &H8000000F
    
    'Set cursors
    cmdApply.MouseIcon = LoadResPicture(101, 2)
    cmdTest.MouseIcon = LoadResPicture(103, 2)
    checkCenter.MouseIcon = LoadResPicture(102, 2)
    picCenter.MouseIcon = LoadResPicture(102, 2)
    optionStyle(0).MouseIcon = LoadResPicture(101, 2)
    optionStyle(1).MouseIcon = LoadResPicture(101, 2)
    picStyle(0).MouseIcon = LoadResPicture(101, 2)
    picStyle(1).MouseIcon = LoadResPicture(101, 2)
    cmdCodeGen.MouseIcon = LoadResPicture(101, 2)
    cmdForeColor.MouseIcon = LoadResPicture(101, 2)
    cmdBackColor.MouseIcon = LoadResPicture(101, 2)
    iconWarning.MouseIcon = LoadResPicture(101, 2)
    iconInfo.MouseIcon = LoadResPicture(101, 2)
    iconError.MouseIcon = LoadResPicture(101, 2)
    picCustomize.MouseIcon = LoadResPicture(101, 2)
    cmdCustomize.MouseIcon = LoadResPicture(101, 2)
    
        For X = 0 To 3
            optionIcon(X).MouseIcon = LoadResPicture(101, 2)
        Next X
    
    cmdAPIDeclare.MouseIcon = LoadResPicture(101, 2)
    cmdAPIsubcall.MouseIcon = LoadResPicture(101, 2)
    cmdTipDIM.MouseIcon = LoadResPicture(101, 2)
    cmdTipCode.MouseIcon = LoadResPicture(101, 2)
    StatusBar1.MouseIcon = LoadResPicture(101, 2)
    
End Sub


Private Sub Form_Terminate()

    Unload frmAbout
    'Connect.Hide

    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmAbout
    Connect.Hide

   

End Sub

Private Sub iconError_Click()

        If optionIcon(3).Enabled = True Then optionIcon(3).Value = True

End Sub

Private Sub iconInfo_Click()

        If optionIcon(2).Enabled = True Then optionIcon(2).Value = True

End Sub

Private Sub iconWarning_Click()

        If optionIcon(1).Enabled = True Then optionIcon(1).Value = True

End Sub

Private Sub mnuAbout_Click()

    'Bring up the About info window
    frmAbout.Visible = True
    
End Sub

Private Sub mnuExit_Click()

    'From "Files" menu, "Exit"
    Connect.Hide
    
End Sub

Private Sub optionIcon_Click(index As Integer)
    
    'FInd out what Tool Tip Icon was selected

        Select Case index
            Case 0
                TipIcon = tipNoIcon
            Case 1
                TipIcon = tipIconWarning
            Case 2
                TipIcon = tipiconinfo
            Case 3
                TipIcon = tipIconError
        End Select
    
End Sub

Private Sub optionStyle_Click(index As Integer)

    'Find out what Tool Tip Style was selected

        Select Case index
            Case 0
                TipStyle = styleStandard
                checkCenterTip.CreateTip checkCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                checkCenterTip.Centered = True
                picCenterTip.CreateTip picCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                picCenterTip.Centered = True

            Case 1
                TipStyle = styleBalloon
                checkCenterTip.CreateBalloon checkCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                checkCenterTip.Centered = True
                picCenterTip.CreateBalloon picCenter, "Tip Looks like this." & vbCrLf & "Can be Rectangle" & vbCrLf & "or Balloon style"
                picCenterTip.Centered = True

        End Select
        
End Sub

Private Sub picStyle_Click(index As Integer)

        Select Case index
            Case 0
                optionStyle(0).Value = True
            Case 1
                optionStyle(1).Value = True
        End Select

End Sub

Private Sub picCenter_Click()

        If checkCenter.Value = 0 Then
            checkCenter.Value = 1
        Else
            checkCenter.Value = 0
        End If

End Sub

Private Sub StatusBar1_PanelClick(ByVal Panel As MSComctlLib.Panel)

        Select Case Panel.index
            Case 1
                ShellExecute hWnd, vbNullString, _
                "http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=57698&lngWId=1", _
                vbNullString, vbNullString, vbNormalFocus

            Case 2
                Me.WindowState = vbMinimized
    
            Case 3
                Unload Me
        End Select
    
End Sub

Private Sub textTip_Change()

    TipText = textTip.Text

End Sub

Private Sub textTip_DblClick()

    'Restore text controls colors to default
    textTip.ForeColor = &H80000008
    textTitle.ForeColor = &H80000008
    '"0" = default forecolor in API
    TipForeColor = 0
    textTip.BackColor = &H80000018
    textTitle.BackColor = &H80000018
    '"0" = default backcolor in API
    ForeColorRGB.Text = "&H" & Hex(textTip.ForeColor)
    BackColorRGB.Text = "&H" & Hex(textTip.BackColor)
    TipForeColor = 0
    TipBackColor = 0
    
    cmdApply.SetFocus
    
End Sub


Private Sub textTitle_Change()

    Dim X            As Integer
    
    'See if text box is empty

        If textTitle.Text <> "" Then
            'If not empty enable the Icon option buttons and set variable to the text
            TipTitleText = textTitle.Text

                For X = 0 To 3
                    optionIcon(X).Enabled = True
                Next X

        Else
            'Text is empty, disable Icon option buttons and null out text variable
            TipTitleText = vbNullString

                For X = 0 To 3
                    optionIcon(X).Enabled = False
                Next X

        End If
    
End Sub

Private Sub TipNameText_Change()

        If TipNameText.Text <> "" And ParentNameText.Text <> "" Then
            cmdCustomize.Enabled = True
            cmdCustomize.BackColor = &HC0C0C0
        Else
            cmdCustomize.Enabled = False
            cmdCustomize.BackColor = &H8000000F
        End If
        
End Sub

Private Sub ParentNameText_Change()

        If TipNameText.Text <> "" And ParentNameText.Text <> "" Then
            cmdCustomize.Enabled = True
            cmdCustomize.BackColor = &HC0C0C0
        Else
            cmdCustomize.Enabled = False
            cmdCustomize.BackColor = &H8000000F
        End If

End Sub

Public Sub GenCode()

    Dim TipName              As String
    Dim TipParent            As String
    
    TipNameText.SetFocus
    
    'Write API Delare Textbox
    'Clean out any current text

        With APItext
            .SelStart = 0
            .SelLength = Len(APItext.Text) + 1
            .SelText = ""
            .SelColor = &H8000&
            .SelText = "'Add this to your startup form or module delare section" & vbCrLf
            .SelColor = vbBlue
            .SelStart = Len(APItext.Text) + 1
            .SelText = "Private Declare Sub "
            .SelColor = vbBlack
            .SelStart = Len(APItext.Text) + 1
            .SelText = "InitCommonControls "
            .SelColor = vbBlue
            .SelStart = Len(APItext.Text) + 1
            .SelText = "Lib "
            .SelColor = vbBlack
            .SelStart = Len(APItext.Text) + 1
            .SelText = """comctl32.dll"" ()"
        End With
    
    'Write CallSubText
    'Clean out any current text

        With CallSubText
            .SelStart = 0
            .SelLength = Len(CallSubText.Text) + 1
            .SelText = ""
            .SelColor = &H8000&
            .SelText = "'Add this to your startup Form Load or Sub_Main" & vbCrLf
            .SelColor = vbBlack
            .SelStart = Len(CallSubText.Text) + 1
            .SelText = "InitCommonControls "
        End With
    
    'Get variables
    
        If TipNameText.Text = "" Then
            TipName = "<Your Tip Name>"
        Else
            TipName = TipNameText.Text
        End If
        
        If ParentNameText.Text = "" Then
            TipParent = "<Your Parent Control Name>"
        Else
            TipParent = ParentNameText.Text
        End If
        
    'Write out the Declarations section
    'Clean out any current text

        With DeclareText
            .SelStart = 0
            .SelLength = Len(DeclareText.Text) + 1
            .SelText = ""
            .SelColor = &H8000&
            .SelText = "'Add this to your form delare section" & vbCrLf
            .SelColor = vbBlue
            .SelText = "Dim "
            .SelColor = vbBlack
            .SelText = TipName & vbTab
            .SelColor = vbBlue
            .SelText = "As New "
            .SelColor = vbBlack
            .SelText = "  clsTooltips"
        End With
        
    'Replace vbCrLF code (Chr$(10)+Chr$(13)) with " & vbCrLf & " text
    'for proper string format in code generation
    'TipText = ReplaceText(TipText)
    
        With CodeText
            'Write out the Code section
            .SelStart = 0
            .SelLength = Len(CodeText.Text) + 1
            .SelText = ""
            .SelColor = &H8000&
            .SelText = "'Add this to your form code section" & vbCrLf
            
                If TipStyle = styleBalloon Then
                    .SelText = TipName & ".CreateBalloon " & TipParent & ", _" & vbCrLf & """" & ReplaceText(TipText) & """"
                Else
                    .SelText = TipName & ".CreateTip " & TipParent & ", _" & vbCrLf & """" & ReplaceText(TipText) & """"
                End If
        
                If TipTitleText <> "" Then
                    .SelText = ", _" & vbCrLf & """" & ReplaceTitle(TipTitleText) & """, " & Val(TipIcon)
                End If
                
            .SelText = vbCrLf
                
                If TipCentered = True Then
                    .SelText = TipName & ".Centered = "
                    .SelColor = vbBlue
                    .SelText = "True" & vbCrLf
                    .SelColor = vbBlack
                    
                End If
                
                If TipForeColor <> 0 Then
                    .SelColor = vbBlack
                    .SelText = TipName & ".ForeColor = " & "&H" & Hex(TipForeColor) & "&" & vbCrLf
                End If
                
                If TipBackColor <> 0 Then
                    .SelColor = vbBlack
                    .SelText = TipName & ".BackColor = " & "&H" & Hex(TipBackColor) & "&" & vbCrLf
                End If
                
            .SelText = vbCrLf
            
        End With
                                
    

End Sub

Private Function ReplaceText(rText As String)

    'Temp Text holding string

    Dim tempText             As String

    'Replace Quote marks with double quotes for proper text string format
    tempText = Replace(rText, Chr$(34), Chr$(34) & Chr$(34))
    
    'Replace Tool Tip Text with more verbose string. Add "& vbCrLf &"
    'string in place of Chr$(10)+Chr$(13)
    ReplaceText = Replace(tempText, vbCrLf, """ & vbCrLf &  _" & vbCrLf & """")

End Function

Private Function ReplaceTitle(tText As String)

    'Replace Quote marks with double quotes for proper text string format
    ReplaceTitle = Replace(tText, Chr$(34), Chr$(34) & Chr$(34))

End Function
