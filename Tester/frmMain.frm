VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TrayIcon.dll Tester"
   ClientHeight    =   3570
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3570
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About ..."
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   3600
      TabIndex        =   5
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Icon"
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdShow 
      Caption         =   "Show Icon"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   975
   End
   Begin VB.CommandButton cmdChangeTip 
      Caption         =   "Change"
      Height          =   255
      Left            =   3840
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtTip 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label lblInstructions 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Width           =   4455
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      Caption         =   "Up"
      Height          =   195
      Index           =   5
      Left            =   1260
      TabIndex        =   11
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label lblLabel 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "DClick"
      Height          =   195
      Index           =   4
      Left            =   1860
      TabIndex        =   10
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Down"
      Height          =   195
      Index           =   3
      Left            =   660
      TabIndex        =   9
      Top             =   1800
      Width           =   405
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Right:"
      Height          =   195
      Index           =   2
      Left            =   180
      TabIndex        =   8
      Top             =   3120
      Width           =   435
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Middle:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   510
   End
   Begin VB.Label lblLabel 
      AutoSize        =   -1  'True
      Caption         =   "Left: "
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   6
      Top             =   2160
      Width           =   390
   End
   Begin VB.Image imgIconMiddle 
      Height          =   480
      Index           =   2
      Left            =   1860
      Picture         =   "frmMain.frx":0442
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image imgIconMiddle 
      Height          =   480
      Index           =   1
      Left            =   1260
      Picture         =   "frmMain.frx":0884
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image imgIconMiddle 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMain.frx":0CC6
      Top             =   2520
      Width           =   480
   End
   Begin VB.Image imgIconRight 
      Height          =   480
      Index           =   2
      Left            =   1860
      Picture         =   "frmMain.frx":0FD0
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgIconRight 
      Height          =   480
      Index           =   1
      Left            =   1260
      Picture         =   "frmMain.frx":1412
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgIconRight 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMain.frx":171C
      Top             =   3000
      Width           =   480
   End
   Begin VB.Image imgIconLeft 
      Height          =   480
      Index           =   2
      Left            =   1860
      Picture         =   "frmMain.frx":1A26
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image imgIconLeft 
      Height          =   480
      Index           =   1
      Left            =   1260
      Picture         =   "frmMain.frx":1E68
      Top             =   2040
      Width           =   480
   End
   Begin VB.Image imgIconLeft 
      Height          =   480
      Index           =   0
      Left            =   660
      Picture         =   "frmMain.frx":22AA
      Top             =   2040
      Width           =   480
   End
   Begin VB.Label lblTip 
      AutoSize        =   -1  'True
      Caption         =   "Tip: "
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   315
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'NOTE the reference in PROJECT->REFERENCES...

Dim bVisible As Boolean '--> Variable which contains the tray icon status

' The "WithEvents" keyword allows us to handle the DLL's events
Dim WithEvents cTrayIcon As clsTrayIcon
Attribute cTrayIcon.VB_VarHelpID = -1

Private Const LABEL_INSTRUCTIONS As String = "If you want to see the Down event, " & _
                                            "hold the click." & vbCrLf & "If you " & _
                                            "want to see the DoubleClick event, " & _
                                            "double-click the tray icon and hold it."

Private Const ABOUT_BOX As String = "I hope this is useful for you. Remember" & vbCrLf & _
                                    "you may use this code or DLL in anything" & vbCrLf & _
                                    "you want." & vbCrLf & vbCrLf & _
                                    "Don't forget to leave comments, report" & vbCrLf & _
                                    "bugs or say anything you want!"

Private Sub cmdAbout_Click()
    MsgBox ABOUT_BOX, vbInformation + vbOKOnly, "About TrayIcon.dll"
End Sub

Private Sub cmdChangeTip_Click()
    ' Change Tip
    cTrayIcon.ChangeTip txtTip.Text
End Sub

Private Sub cmdExit_Click()
    ' Unload the form
    Unload Me
End Sub

Private Sub cmdRemove_Click()
    ' If it is not visible, we don't remove it again
    If (bVisible = False) Then Exit Sub
    
    ' Remove the tray icon
    cTrayIcon.RemoveTrayIcon
    bVisible = False
End Sub

Private Sub cmdShow_Click()
    ' If it is visible, we don't show it again
    If (bVisible = True) Then Exit Sub
    
    ' Show the tray icon
    cTrayIcon.ShowTrayIcon Me.hWnd, txtTip.Text, Me.Icon
    bVisible = True
End Sub

Private Sub cTrayIcon_LeftButtonDoubleClick()
    cTrayIcon.ChangeIcon imgIconLeft(2).Picture
End Sub

Private Sub cTrayIcon_LeftButtonDown()
    cTrayIcon.ChangeIcon imgIconLeft(0).Picture
End Sub

Private Sub cTrayIcon_LeftButtonUp()
    cTrayIcon.ChangeIcon imgIconLeft(1).Picture
End Sub

Private Sub cTrayIcon_MiddleButtonDoubleClick()
    cTrayIcon.ChangeIcon imgIconMiddle(2).Picture
End Sub

Private Sub cTrayIcon_MiddleButtonDown()
    cTrayIcon.ChangeIcon imgIconMiddle(0).Picture
End Sub

Private Sub cTrayIcon_MiddleButtonUp()
    cTrayIcon.ChangeIcon imgIconMiddle(1).Picture
End Sub

Private Sub cTrayIcon_RightButtonDoubleClick()
    cTrayIcon.ChangeIcon imgIconRight(2).Picture
End Sub

Private Sub cTrayIcon_RightButtonDown()
    cTrayIcon.ChangeIcon imgIconRight(0).Picture
End Sub

Private Sub cTrayIcon_RightButtonUp()
    cTrayIcon.ChangeIcon imgIconRight(1).Picture
End Sub

Private Sub Form_Load()
    ' We create a new instance of the DLL
    Set cTrayIcon = New clsTrayIcon
    
    txtTip.Text = "TrayIcon.dll Tester"
    
    ' Call cmdShow_Click
    cmdShow.Value = True
    
    lblInstructions.Caption = LABEL_INSTRUCTIONS
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Call cmdRemove_Click
    cmdRemove.Value = True
    
    ' Clean up memory
    Set cTrayIcon = Nothing
End Sub
