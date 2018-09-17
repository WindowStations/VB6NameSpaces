Attribute VB_Name = "modMain"
Option Explicit
'Import VB6 class namespaces here in a public module
Public System As New System
Public My As New My
Public Process As New Process
Public SendKeys As New SendKeys
Public Diagnostics As New Diagnostics
Public Directory As New Directory
Public File As New File
Public Application As New Application
Public Environment As New Environment
Public Threading As New Threading
Public IO As New IO
Public Messagebox As New Messagebox
Private Const MONITORINFOF_PRIMARY     As Long = &H1
'Public User defined types can be used by functions inside private classes
'Functions in the scope of a "Friend" (throughout the project only)
Public Type AutomationElement_
    Name As String
    ClassName As String
    ProcessID As Long
End Type
Public Type SerialPort_
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type
Public Type Version
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Public Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type
Private Type MONITORINFO
    cbSize          As Long
    rcMonitor       As RECT
    rcWork          As RECT
    dwFlags         As Long
End Type
Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'Variables
Public screens() As Screen
Public hMon      As Long
Public Timercollection  As New Collection
Public CTimersCol       As New Collection
Private mTimersColCount As Integer
Public Enum MB_RESULT
    IOK = 1
    ICANCEL = 2
    IABORT = 3
    IRETRY = 4
    IIGNORE = 5
    IYES = 6
    INO = 7
    ITRYAGAIN = 10
    ICONTINUE = 11
End Enum
'
'Screen
Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByRef dwData As Long) As Long
    On Error Resume Next
    Dim rects() As RECT
    ReDim Preserve rects(dwData)
    ReDim Preserve screens(dwData)
    rects(dwData) = lprcMonitor
    Dim MI As MONITORINFO
    MI.cbSize = Len(MI)
    If apiGetMonitorInfo(hMonitor, MI) <> 0 Then
        With rects(dwData)
            
            Dim r As New Rectangle
            r.Left = MI.rcMonitor.Left
            r.Top = MI.rcMonitor.Top
            r.Right = MI.rcMonitor.Right
            r.Bottom = MI.rcMonitor.Bottom
            r.Width_ = (MI.rcMonitor.Right - MI.rcMonitor.Left)
            r.Height = (MI.rcMonitor.Bottom - MI.rcMonitor.Top)
            r.Size.Width_ = r.Width_
            r.Size.Height_ = r.Height
            r.Location.x = r.Left
            r.Location.y = r.Top
            Dim rw As New Rectangle
            With rw
                .Left = MI.rcWork.Left
                .Top = MI.rcWork.Top
                .Right = MI.rcWork.Right
                .Bottom = MI.rcWork.Bottom
                .Width_ = (MI.rcWork.Right - MI.rcWork.Left)
                .Height = (MI.rcWork.Bottom - MI.rcWork.Top)
                .Size.Height_ = rw.Width_
                .Size.Width_ = rw.Height
                .Location.x = rw.Left
                .Location.y = rw.Top
            End With
            Dim sc As New Screen
            If hMon = 0 Then sc.Primary = CBool(MI.dwFlags = MONITORINFOF_PRIMARY)
            If hMon <> 0 And hMon = hMonitor Then sc.Primary = True
            Let sc.Handle = hMonitor
            Set sc.Bounds = r
            Set sc.WorkingArea = rw
            Set screens(dwData) = sc
            
'            MsgBox screens(0).Bounds.Width_
        End With
    End If
    dwData = dwData + 1
    MonitorEnumProc = 1
End Function
'
'Timer
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim t As Timer
    Dim c As Timers
    On Error Resume Next
    Set t = Timercollection("id:" & idEvent)
    If t Is Nothing Then
        Call apiKillTimer(0, idEvent)
    Else
        If t.ParentsColKey > 0 Then
            Set c = CTimersCol("key:" & t.ParentsColKey)
            If c Is Nothing Then
                Call apiKillTimer(0, idEvent)
            Else
                c.RaiseTimer_Event t.Index
            End If
        Else
            t.RaiseTimer_Event
        End If
    End If
    Set t = Nothing
End Sub
Public Function RegisterTimerCollection(ByRef c As Timers) As Integer
    Dim key As String
    mTimersColCount = mTimersColCount + 1
    key = "key:" & mTimersColCount
    CTimersCol.Add c, key
    RegisterTimerCollection = mTimersColCount
End Function


















