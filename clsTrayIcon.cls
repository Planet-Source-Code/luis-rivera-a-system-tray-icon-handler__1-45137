VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTrayIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'--------------------------------------------------------------------
' Class:        clsTrayIcon.Cls
' Description:  Shows/Hides TrayIcons, handles their mouse messages
' Dependencies: modTrayIcon.bas
'--------------------------------------------------------------------
' Created:      29th April 2003
'--------------------------------------------------------------------
' Author:       Luis Eduardo Rivera <lerivera@southlink.com.ar>
'--------------------------------------------------------------------
Option Explicit

'This UDT will help us handle the Tray Icon
Private Type NOTIFYICONDATA
    cbSize As Long              ' Its size
    hWnd As Long                ' The handle of the window which will receive Windows' _
                                 messages
    uID As Long                 ' Its Unique Identification
    uFlags As Long              ' Its style
    uCallbackMessage As Long    ' The message Windows will send us to process
    hIcon As Long               ' Its Icon
    szTip As String * 64        ' Its tip. NOTE: MUST BE FINISHED WITH NULL
End Type

' This API lets us receive any window's messages
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" _
        (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

' This API handles the Tray Icon
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" _
        (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Boolean

' dwMessage Constants
Private Const NIM_ADD = &H0             'Flag : "ALL NEW nid"
Private Const NIM_MODIFY = &H1          'Flag : "ONLY MODIFYING nid"
Private Const NIM_DELETE = &H2          'Flag : "DELETE THE CURRENT nid"
Private Const NIF_MESSAGE = &H1         'Flag : "Message in nid is valid"
Private Const NIF_ICON = &H2            'Flag : "Icon in nid is valid"
Private Const NIF_TIP = &H4             'Flag : "Tip in nid is valid"

' Subclassing constant
Private Const GWL_WNDPROC = (-4)

' Our class' events
Event LeftButtonUp()                'Raised when user stops clicking with left button
Event LeftButtonDown()              'Raised when user clicks with left button
Event LeftButtonDoubleClick()       'Raised when user double-clicks with left button

Event RightButtonUp()               'Raised when user stops clicking with right button
Event RightButtonDown()             'Raised when user clicks with right button
Event RightButtonDoubleClick()      'Raised when user double-clicks with right button

Event MiddleButtonUp()              'Raised when user stops clicking with middle button
Event MiddleButtonDown()            'Raised when user clicks with middle button
Event MiddleButtonDoubleClick()     'Raised when user double-clicks with middle button

Private nidTrayIcon         As NOTIFYICONDATA   ' Variable which contains our Tray Icon
Private m_lHWnd             As Long             ' Variable which contains the handle _
                                                 of the window we subclass

'***********************************************
'** Function: ChangeIcon
'***********************************************
'** Description: Changes Tray Icon's Icon
'***********************************************
'** Argument:
'**          . lNewIcon: The new icon
'***********************************************
'** Returns 0 if successful or the error number
'***********************************************
Public Function ChangeIcon(ByVal lNewIcon As Long) As Long
On Error GoTo ErrHandler
    Dim nidNewTray As NOTIFYICONDATA        ' We declare a "dummy" Tray Icon _
                                             just to modify the actual one
    
    With nidNewTray
        .cbSize = Len(nidNewTray)           ' We change its size
        .hWnd = m_lHWnd                     ' We set the window we first chose to _
                                             capture every message
        .uID = 1                            ' The same uID as the original one
        .uFlags = NIF_ICON                  ' This flag makes the Shell_NotifyIcon API _
                                             change the icon
        .hIcon = lNewIcon                   ' The new icon
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidNewTray ' We call the API to do the work
    
    ChangeIcon = 0                          ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ChangeIcon = Err.Number                 ' Oops, something went wrong, _
                                             return the error number
End Function

'***********************************************
'** Function: ChangeTip
'***********************************************
'** Description: Changes Tray Icon's Tip
'***********************************************
'** Argument:
'**          . nNewTip: The new tip
'***********************************************
'** Returns 0 if successful or the error number
'***********************************************
Public Function ChangeTip(ByVal nNewTip As String) As Long
On Error GoTo ErrHandler
    Dim nidNewTray As NOTIFYICONDATA            ' We declare a "dummy" Tray Icon _
                                             just to modify the actual one
    
    With nidNewTray
        .cbSize = Len(nidNewTray)           ' We change its size
        .hWnd = m_lHWnd                     ' We set the window we first chose to _
                                             capture every message
        .uID = 1                            ' The same uID as the original one
        .uFlags = NIF_TIP                   ' This flag makes the Shell_NotifyIcon API _
                                             change the tip
        .szTip = nNewTip & Chr$(0)          ' The new tip
    End With
    
    Shell_NotifyIcon NIM_MODIFY, nidNewTray ' We call the API to do the work
    
    ChangeTip = 0                           ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ChangeTip = Err.Number                  ' Oops, something went wrong, _
                                             return the error number
End Function

'***********************************************
'** Function: RemoveTrayIcon
'***********************************************
'** Description: Removes the Tray Icon
'***********************************************
'** Arguments: (none)
'***********************************************
'** Returns 0 if successful or the error number
'***********************************************
Public Function RemoveTrayIcon() As Long
On Error GoTo ErrHandler
    
    SetWindowLong m_lHWnd, GWL_WNDPROC, lPreviousProcess   ' Stops subclassing the window
    Shell_NotifyIcon NIM_DELETE, nidTrayIcon        ' Removes the Tray Icon
    
    RemoveTrayIcon = 0                              ' Everything went OK, return 0
    Exit Function
ErrHandler:
    RemoveTrayIcon = Err.Number                     ' Oops, something went wrong, _
                                                     return the error number
End Function

'***********************************************
'** Function: ShowTrayIcon
'***********************************************
'** Description: Shows the Tray Icon
'***********************************************
'** Argument:
'**          . lHWnd: The handle of the window
'**                  which creates the tray icon
'**          . sTip: The tray icon's tip
'**          . lIcon: The tray icon's icon
'***********************************************
'** Returns 0 if successful or the error number
'***********************************************
Public Function ShowTrayIcon(ByVal lHWnd As Long, ByVal sTip As String, _
                            ByVal lIcon As Long) As Long
On Error GoTo ErrHandler
    ' Save the window's handle
    m_lHWnd = lHWnd
    
    ' Subclass the window, but save its process just to make Windows _
     handle the messages we don't want to
    lPreviousProcess = SetWindowLong(m_lHWnd, GWL_WNDPROC, AddressOf WndProc)
    
    With nidTrayIcon
        .cbSize = Len(nidTrayIcon)                      ' Its size
        .hIcon = lIcon                                  ' Its icon
        .hWnd = m_lHWnd                                 ' The handle of the window which _
                                                         will capture its messages
        .szTip = sTip & Chr$(0)                         ' The tip
        .uCallbackMessage = WM_USER + 1                 ' The CallBack value we will _
                                                         use to determine the events
        .uFlags = NIF_ICON Or NIF_MESSAGE Or NIF_TIP
        .uID = 1                                        ' Its unique ID
    End With
    
    Shell_NotifyIcon NIM_ADD, nidTrayIcon               ' We call the API to do the work
    
    ShowTrayIcon = 0                                    ' Everything went OK, return 0
    Exit Function
ErrHandler:
    ShowTrayIcon = Err.Number                           'ERROR: we return its number
End Function

'***********************************************
'** Procedure: ShowEvent
'***********************************************
'** Description: Raises the events captured by
'**             our tray icon
'***********************************************
'** Argument:
'**          . lIdEvent: The event's number
'***********************************************
Friend Sub ShowEvent(ByVal lIdEvent As Long)
    Select Case lIdEvent
        Case WM_LBUTTONDOWN: RaiseEvent LeftButtonDown              'Left click down
        Case WM_LBUTTONUP: RaiseEvent LeftButtonUp                  'Left click up
        Case WM_LBUTTONDBLCLK: RaiseEvent LeftButtonDoubleClick     'Left double click
        Case WM_RBUTTONDOWN: RaiseEvent RightButtonDown             'Right click down
        Case WM_RBUTTONUP: RaiseEvent RightButtonUp                 'Right click up
        Case WM_RBUTTONDBLCLK: RaiseEvent RightButtonDoubleClick    'Right double click
        Case WM_MBUTTONDOWN: RaiseEvent MiddleButtonDown            'Middle click down
        Case WM_MBUTTONUP: RaiseEvent MiddleButtonUp                'Middle click up
        Case WM_MBUTTONDBLCLK: RaiseEvent MiddleButtonDoubleClick   'Middle double click
    End Select
End Sub

Private Sub Class_Initialize()
    Set modTrayIcon.cTray = Me  ' Sets the variable of the module _
                                 to let it handle this class' events
End Sub
