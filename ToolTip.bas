Attribute VB_Name = "modToolTips"
    '**************************************************************
    '
    '   Custom Tool Code Generator Add-In
    '
    '   Mark Mokoski
    '   05-DEC-2004
    '
    '   Module with Sub_Main (App startup)
    '
    '**************************************************************

    Option Explicit

    '************************************************************
    ' Constants
    '************************************************************
    
    ' None
    
    '
    '************************************************************
    ' Types
    '************************************************************

    'None

    '************************************************************
    ' API Functions
    '************************************************************
    'Int Common Controls Lib
    Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

    'Shell out API for HTML files, Mail and Web Browser
    Public Declare Function ShellExecute Lib "shell32" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
    'Sample call:
    'ShellExecute hWnd, vbNullString, "mailto:user@domain.com?body=hello%0a%0world", vbNullString, vbNullString, vbNormalFocus
    'In order to be able to put carriage returns or tabs in your text,
    'replace vbCrLf and vbTab with the following HEX codes:
    '%0a%0d = vbCrLf
    '%09 = vbTab
    'These codes also work when sending URLs to a browser (GET, POST, etc.)
    
Sub Main()

    'Int Common Controls Lib for Balloon Tip use
    InitCommonControls

    'CC_Comment Out (12/5/2004):
    '    ' * Test to see if App is already running
    '    ' * If App is running, terminate copy

    '        If App.PrevInstance Then
    '            MsgBox "IP to Comm Port Control application is already running" & vbCrLf & vbCrLf & _
    '            "Only one instance (copy) of program this can be running" & vbCrLf & _
    '            "for proper operation", vbInformation, "Application ERROR"
    '            End
    '        Else
    '            '  MsgBox "This is the first instance of your application"

    '            'Make main form visible
    '            Load frmToolTips
    '            frmToolTips.Visible = True
            
    '        End If

    'End CC_Comment Out

End Sub
