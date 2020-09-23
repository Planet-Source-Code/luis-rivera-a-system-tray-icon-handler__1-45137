Attribute VB_Name = "modTrayIcon"
'--------------------------------------------------------------------
' Module:       modTrayIcon.bas
' Description:  Handles Tray Icon's messages
' Dependencies: clsTrayIcon.cls
'--------------------------------------------------------------------
' Created:      29th April 2003
'--------------------------------------------------------------------
' Author:       Luis Eduardo Rivera <lerivera@southlink.com.ar>
'--------------------------------------------------------------------
Option Explicit

' This API passes message information to the specified window procedure
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" _
        (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, _
        ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const WM_USER = &H400           'Our base CallBack Message

Public Const WM_LBUTTONDOWN = &H201    'LButton down
Public Const WM_RBUTTONDOWN = &H204    'RButton down
Public Const WM_MBUTTONDOWN = &H207    'MButton down

Public Const WM_LBUTTONUP = &H202      'LButton up
Public Const WM_RBUTTONUP = &H205      'RButton up
Public Const WM_MBUTTONUP = &H208      'MButton up

Public Const WM_LBUTTONDBLCLK = &H203  'LDouble-click
Public Const WM_RBUTTONDBLCLK = &H206  'RDouble-click
Public Const WM_MBUTTONDBLCLK = &H209  'MDouble-click

Public cTray            As clsTrayIcon  ' The clsTrayIcon class

Public lPreviousProcess As Long         ' The previous process handle

Public Function WndProc(ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, _
                        ByVal lParam As Long) As Long
    
    Select Case msg     ' Which message did our window receive?
        Case WM_USER + 1    ' If its our custom message,
            Select Case lParam  ' Was it a mouse clicking event?
                Case WM_LBUTTONDOWN, WM_LBUTTONUP, WM_LBUTTONDBLCLK, WM_RBUTTONDOWN, _
                    WM_RBUTTONUP, WM_RBUTTONDBLCLK, WM_MBUTTONDOWN, WM_MBUTTONUP, _
                    WM_MBUTTONDBLCLK ' Yes, it was
                        cTray.ShowEvent lParam ' Let's raise the event
                Case Else               ' No, it was not
                    WndProc = CallWindowProc(lPreviousProcess, hWnd, msg, wParam, lParam)
                                        ' Let the window handle the message
            End Select
        Case Else           ' It's not our custom message, let the window handle it
            WndProc = CallWindowProc(lPreviousProcess, hWnd, msg, wParam, lParam)
    End Select
End Function

