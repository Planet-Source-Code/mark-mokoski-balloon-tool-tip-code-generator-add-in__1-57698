VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1455
      Left            =   1080
      MouseIcon       =   "Form1.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   720
      Width           =   2655
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Dim Command1Tip As New clsTooltips


Private Sub Form_Load()
InitCommonControls

Command1Tip.CreateBalloon Command1, _
"Command1 Test Tool Tip." & vbCrLf & _
"Multiline is also supported." & vbCrLf & _
"" & vbCrLf & _
"Thanks for voting for my code.", _
"Command1 Tool Tip", 1
Command1Tip.ForeColor = &H80FFFF
Command1Tip.BackColor = &HC08000



End Sub
