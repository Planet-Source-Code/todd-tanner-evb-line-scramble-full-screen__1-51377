Attribute VB_Name = "modFullScreen"
Option Explicit

Const SW_HIDE = 0
Const SW_SHOWNORMAL = 1

Const SHFS_SHOWTASKBAR = &H1
Const SHFS_HIDETASKBAR = &H2
Const SHFS_SHOWSIPBUTTON = &H4
Const SHFS_HIDESIPBUTTON = &H8
Const SHFS_SHOWSTARTICON = &H10
Const SHFS_HIDESTARTICON = &H20
Const HWND_TOPMOST = -1
Const HWND_NOTOPMOST = -2

Const SWP_SHOWWINDOW = &H40
Const SM_CXSCREEN = &H0
Const SM_CYSCREEN = &H1
Const HHTASKBARHEIGHT = 26

Declare Function GetSystemMetrics Lib "Coredll" ( _
    ByVal nIndex As Long) As Long
    
Declare Function SHFullScreen Lib "aygshell" ( _
    ByVal hwndRequester As Long, _
    ByVal dwState As Long) As Boolean

Declare Function MoveWindow Lib "Coredll" ( _
    ByVal hwnd As Long, _
    ByVal x As Long, _
    ByVal y As Long, _
    ByVal nWidth As Long, _
    ByVal nHeight As Long, _
    ByVal bRepaint As Long) As Long

Declare Function SetForegroundWindow Lib "Coredll" ( _
    ByVal hwnd As Long) As Boolean

Declare Function GetLastError Lib "Coredll" () As Long

Declare Function ShowWindow Lib "Coredll" ( _
    ByVal hwnd As Long, _
    ByVal nCmdShow As Long) As Long

Declare Function FindWindow Lib "Coredll" Alias "FindWindowW" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Public Sub FullScreen(ByVal frmHwnd As Long, ByVal makeFull As Boolean)
    Dim lret
    If Not makeFull Then
        ShowSIP frmHwnd, True
        ShowStart frmHwnd, True
        ShowTaskbar frmHwnd, True
        lret = FindWindow("menu_worker", "")
        If lret <> 0 Then 'window found
            ShowWindow lret, SW_SHOWNORMAL
        End If
        lret = SetForegroundWindow(frmHwnd)
        lret = MoveWindow(frmHwnd, 0, HHTASKBARHEIGHT, _
            GetSystemMetrics(SM_CXSCREEN), _
            GetSystemMetrics(SM_CYSCREEN), True)
    Else
        ShowSIP frmHwnd, False
        ShowStart frmHwnd, False
        ShowTaskbar frmHwnd, False
        'show form full screen
        lret = FindWindow("menu_worker", "")
        If lret <> 0 Then 'window found
            ShowWindow lret, SW_HIDE
        End If
        lret = SetForegroundWindow(frmHwnd)
        lret = MoveWindow(frmHwnd, 0, 0, _
            GetSystemMetrics(SM_CXSCREEN), _
            GetSystemMetrics(SM_CYSCREEN) + HHTASKBARHEIGHT, 0)
    End If
End Sub
Public Sub ShowSIP(ByVal frmHwnd As Long, ByVal ShowIt As Boolean)
    Dim lret
    lret = SetForegroundWindow(frmHwnd)
    If Not ShowIt Then
        lret = SHFullScreen(frmHwnd, SHFS_HIDESIPBUTTON)
    Else
        lret = SHFullScreen(frmHwnd, SHFS_SHOWSIPBUTTON)
    End If
End Sub
Public Sub ShowStart(ByVal frmHwnd As Long, ByVal ShowIt As Boolean)
    Dim lret
    lret = SetForegroundWindow(frmHwnd)
    If Not ShowIt Then
        lret = SHFullScreen(frmHwnd, SHFS_HIDESTARTICON)
    Else
        lret = SHFullScreen(frmHwnd, SHFS_SHOWSTARTICON)
    End If
End Sub
Public Sub ShowTaskbar(ByVal frmHwnd As Long, ByVal ShowIt As Boolean)
    Dim lret
    lret = SetForegroundWindow(frmHwnd)
    If Not ShowIt Then
        lret = SHFullScreen(frmHwnd, SHFS_HIDETASKBAR)
    Else
        lret = SHFullScreen(frmHwnd, SHFS_SHOWTASKBAR)
    End If
End Sub
