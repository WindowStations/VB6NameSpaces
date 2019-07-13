VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Screen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False


Option Explicit

'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'  Persistable = 0  'NotPersistable
'  DataBindingBehavior = 0  'vbNone
'  DataSourceBehavior  = 0  'vbNone
'  MTSTransactionMode  = 0  'NotAnMTSObject
'END
'Attribute VB_Name = "Screen"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const CCDEVICENAME                       As Long = 32
Private Const CCFORMNAME                         As Long = 32
Private Const DM_BITSPERPEL                      As Long = &H40000
Private Const DM_PELSWIDTH                       As Long = &H80000
Private Const DM_PELSHEIGHT                      As Long = &H100000
Private Const CDS_UPDATEREGISTRY                 As Long = &H1
Private Const CDS_TEST                           As Long = &H2
Private Const CDS_FULLSCREEN                     As Long = &H4
Private Const CDS_GLOBAL                         As Long = &H8
Private Const CDS_SET_PRIMARY                    As Long = &H10
Private Const CDS_RESET                          As Long = &H40000000
Private Const CDS_SETRECT                        As Long = &H20000000
Private Const CDS_NORESET                        As Long = &H10000000
Private Const DISP_CHANGE_SUCCESSFUL             As Long = 0
Private Const DISP_CHANGE_RESTART                As Long = 1
Private Const DISP_CHANGE_FAILED                 As Long = -1
Private Const DISP_CHANGE_BADMODE                As Long = -2
Private Const DISP_CHANGE_NOTUPDATED             As Long = -3
Private Const DISP_CHANGE_BADFLAGS               As Long = -4
Private Const DISP_CHANGE_BADPARAM               As Long = -5
Private Const MONITORINFOF_PRIMARY               As Long = &H1
Private Const MONITOR_DEFAULTTONEAREST           As Long = &H2
Private Const MONITOR_DEFAULTTONULL              As Long = &H0
Private Const MONITOR_DEFAULTTOPRIMARY           As Long = &H1
Private Const DISPLAY_DEVICE_ACTIVE              As Long = &H1
Private Const DISPLAY_DEVICE_ATTACHED_TO_DESKTOP As Long = &H1
Private Const DISPLAY_DEVICE_MULTI_DRIVER        As Long = &H2
Private Const DISPLAY_DEVICE_PRIMARY_DEVICE      As Long = &H4
Private Const DISPLAY_DEVICE_MIRRORING_DRIVER    As Long = &H8
Private Const DISPLAY_DEVICE_VGA_COMPATIBLE      As Long = &H10
Private Const DISPLAY_DEVICE_MODESPRUNED         As Long = &H8000000
Private Const DISPLAY_DEVICE_REMOTE              As Long = &H8000000
Private Const DISPLAY_DEVICE_DISCONNECT          As Long = &H8000000
Private Const ENUM_CURRENT_SETTINGS              As Long = -1
Private Const ENUM_REGISTRY_SETTINGS             As Long = -2
Private Type typDevMODE
    dmDeviceName As String * 32
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * 32
    dmUnusedPadding As Integer
    dmBitsPerPixel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    dmICMMethod As Long
    dmICMIntent As Long
    dmMediaType As Long
    dmDitherType As Long
    dmReserved1 As Long
    dmReserved2 As Long
    dmPanningWidth As Long
    dmPanningHeight As Long
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
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Declare Function apiChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (ByRef lptypDevMode As typDevMODE, ByVal dwFlags As Long) As Long
Private Declare Function apiChangeDisplaySettingsEx Lib "user32" Alias "ChangeDisplaySettingsExA" (ByRef lpszDeviceName As Any, ByRef lpDevMode As Any, ByVal hwnd As Long, ByVal dwFlags As Long, ByRef lParam As Any) As Long
Private Declare Function apiEnumDisplayDevices Lib "user32" Alias "EnumDisplayDevicesA" (ByVal lpDevice As String, ByVal iDevNum As Long, ByRef lpDisplayDevice As DISPLAY_DEVICE, ByVal dwFlags As Long) As Long
Private Declare Function apiEnumDisplayMonitors Lib "user32" Alias "EnumDisplayMonitors" (ByVal HDC As Long, ByRef lprcClip As Long, ByVal lpfnEnum As Long, ByRef dwData As Long) As Long
Private Declare Function apiEnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, ByRef lptypDevMode As typDevMODE) As Long
'Public Declare Function EnumDisplaySettings Lib "user32.dll" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As String, ByVal iModeNum As Long, lpDevMode As DEVMODE) As Long
Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function apiMonitorFromPoint Lib "user32" Alias "MonitorFromPoint" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
Private lMonitorCount As Long
Public Bounds         As New Rectangle
Public Primary        As Boolean
Public WorkingArea    As New Rectangle
Public Handle         As Long
Friend Function AllScreens() As Screen()
    On Error Resume Next
    Dim i As Long
    lMonitorCount = 0
    ReDim screens(-1)
    Call apiEnumDisplayMonitors(0, ByVal 0, AddressOf MonitorEnumProc, lMonitorCount)
    AllScreens = screens
End Function
Friend Function DeviceName() As String
    Dim d() As DISPLAY_DEVICE
    d = GetDisplayDevices
    Dim i As Long
    For i = 0 To UBound(d)
        If CBool(d(i).StateFlags And DISPLAY_DEVICE_PRIMARY_DEVICE) Then DeviceName = d(i).DeviceName: Exit For
    Next
End Function
Friend Function GetBounds(ByRef p As Point) As Rectangle
    Dim pt As POINTAPI
    pt.x = p.x: pt.y = p.y
    Dim ret As Long
    ret = apiMonitorFromPoint(pt.x, pt.y, MONITOR_DEFAULTTONULL)
    If ret = 0 Then Exit Function
    Dim r   As New Rectangle
    Dim s() As Screen
    hMon = ret 'save variable in public module so that callback can see alternate monitor specified as primary
    s = My.Computer.Screen.AllScreens
    hMon = 0
    Dim i As Long
    For i = 0 To UBound(s)
        If s(i).Primary = True Then 'match monitor here and return working area
            Set r = s(i).Bounds
            Exit For
        End If
    Next
    Set GetBounds = r
End Function
Friend Function GetWorkingArea(ByRef p As Point) As Rectangle
    Dim pt As POINTAPI
    pt.x = p.x: pt.y = p.y
    Dim ret As Long
    ret = apiMonitorFromPoint(pt.x, pt.y, MONITOR_DEFAULTTONULL)
    If ret = 0 Then Exit Function
    Dim r   As New Rectangle
    Dim s() As Screen
    hMon = ret 'save variable in public module so that callback can see alternate monitor specified as primary
    s = My.Computer.Screen.AllScreens
    hMon = 0
    Dim i As Long
    For i = 0 To UBound(s)
        If s(i).Primary = True Then 'match monitor here and return working area
            Set r = s(i).WorkingArea
            Exit For
        End If
    Next
    Set GetWorkingArea = r
End Function
Friend Function PrimaryScreen() As Screen
    Dim ps  As New Screen
    Dim s() As Screen
    s = My.Computer.Screen.AllScreens
    Dim i As Long
    For i = 0 To UBound(s)
        If s(i).Primary = True Then
            Set ps = s(i)
            Exit For
        End If
    Next
    Set PrimaryScreen = ps
End Function
Friend Function GetDisplayDevices() As DISPLAY_DEVICE()
    Dim d() As DISPLAY_DEVICE
    Dim dd  As DISPLAY_DEVICE
    dd.cb = Len(dd)
    Dim i As Long
    Do
        If apiEnumDisplayDevices(vbNullString, i, dd, 0) = 0 Then Exit Do
        If CBool(dd.StateFlags And DISPLAY_DEVICE_ACTIVE) Then
            ReDim Preserve d(i)
            d(i) = dd
        End If
        i = i + 1
    Loop
    GetDisplayDevices = d
End Function
'Todo for net conversion
'Friend Function FromHandle(ByVal hwnd As Long) As Screen
'    Dim s   As New Screen
'    Dim ret As Long
'    ret = apiMonitorFromWindow(hwnd, 0)
'    If ret = 0 Then Exit Function
'    'todo
'    'set
'    FromHandle = s
'End Function
'Friend Function FromPoint(ByVal p As Point) As Screen
'    Dim s  As New Screen
'    Dim pt As POINTAPI
'    pt.x = p.x: pt.y = p.y
'    Dim ret As Long
'    ret = apiMonitorFromPoint(pt.x, pt.y, MONITOR_DEFAULTTONULL)
'    If ret = 0 Then Exit Function
'    'todo
'    'set
'    FromPoint = s
'End Function
'Friend Function FromRectangle(ByVal r As Rectangle) As Screen
'    Dim s  As New Screen
'    Dim rt As RECT
'    rt.Left = r.Left
'    rt.Top = r.Top
'    rt.Right = r.Right
'    rt.Bottom = r.Bottom
'    Dim ret As Long
'    ret = apiMonitorFromRect(rt, 0)
'    If ret = 0 Then Exit Function
'    'todo
'    'set
'    FromRectangle = s
'End Function
'
'Extras todo
'_____________________________________________________
'Private Function GetPrimary(ByVal h As Long) As Boolean
'    Dim MI As MONITORINFO
'    MI.cbSize = Len(MI)
'    If apiGetMonitorInfo(h, MI) <> 0 Then
'        Primary = CBool(MI.dwFlags = MONITORINFOF_PRIMARY)
'    End If
'End Function
'Friend Function SetResolution(ByVal w As Long, ByVal h As Long)
'    Dim typDevM   As typDevMODE
'    Dim lngResult As Long
'    Dim intAns    As Integer
'    lngResult = apiEnumDisplaySettings(0, 0, typDevM)
'    With typDevM
'        .dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
'        .dmPelsWidth = w
'        .dmPelsHeight = h
'    End With
'    lngResult = apiChangeDisplaySettings(typDevM, CDS_TEST)
'    Select Case lngResult
'        Case DISP_CHANGE_RESTART
'            intAns = MsgBox("You must restart your computer to apply these changes." & vbCrLf & vbCrLf & "Do you want to restart now?", vbYesNo + vbSystemModal, "Screen Resolution")
'        Case DISP_CHANGE_SUCCESSFUL
'            Call apiChangeDisplaySettings(typDevM, CDS_UPDATEREGISTRY)
'            MsgBox "Screen resolution changed", vbInformation, "Resolution Changed"
'        Case Else
'            MsgBox "Mode not supported", vbSystemModal, "Error"
'    End Select
'End Function
'Friend Sub SavePosition(ByVal hwnd As Long, Optional ByVal sWinName As String)
'    On Error Resume Next
'    Dim rc      As RECT
'    Dim sRegKey As String
'    sRegKey = IIf((appTitled <> ""), appTitled, "VB65")
'    If Trim(sWinName) <> "" Then sRegKey = sRegKey & "_" & Trim(sWinName)
'    If apiGetWindowRect(hwnd, rc) = 0 Then Exit Sub
'    SaveSetting sRegKey, "WindowPosition", "Left", rc.Left
'    SaveSetting sRegKey, "WindowPosition", "Top", rc.Top
'End Sub
'Friend Sub LoadPosition(ByVal hwnd As Long, Optional ByVal sWinName As String)
'    On Error Resume Next
'    Dim rc          As RECT
'    Dim Left        As Long
'    Dim Top         As Long
'    Dim hMonitor    As Long
'    Dim MI          As MONITORINFO
'    Dim sRegKey As String
'    sRegKey = IIf((appTitled <> ""), appTitled, "VB65")
'    If Trim(sWinName) <> "" Then sRegKey = sRegKey & "_" & Trim(sWinName)
'    If apiGetWindowRect(hwnd, rc) = 0 Then Exit Sub
'    Left = GetSetting(sRegKey, "WindowPosition", "Left", rc.Left)
'    Top = GetSetting(sRegKey, "WindowPosition", "Top", rc.Left)
'    Call apiOffsetRect(rc, Left - rc.Left, Top - rc.Top)
'    hMonitor = apiMonitorFromRect(rc, MONITOR_DEFAULTTONEAREST)
'    MI.cbSize = Len(MI)
'    If apiGetMonitorInfo(hMonitor, MI) = 0 Then Exit Sub
'    If rc.Left < MI.rcWork.Left Then apiOffsetRect rc, MI.rcWork.Left - rc.Left, 0
'    If rc.Right > MI.rcWork.Right Then apiOffsetRect rc, MI.rcWork.Right - rc.Right, 0
'    If rc.Top < MI.rcWork.Top Then apiOffsetRect rc, 0, MI.rcWork.Top - rc.Top
'    If rc.Bottom > MI.rcWork.Bottom Then apiOffsetRect rc, 0, MI.rcWork.Bottom - rc.Bottom
'    Call apiMoveWindow(hwnd, rc.Left, rc.Top, rc.Right - rc.Left, rc.Bottom - rc.Top, 0)
'End Sub
'Friend Function Bounds() As Rectangle
'    Dim r   As New Rectangle
'    Dim s() As Screen
'    s = My.Computer.Screen.AllScreens
'    Dim i As Long
'    For i = 0 To UBound(s)
'        If s(i).Primary = True Then
'            r = s(i).Bounds
'            Exit For
'        End If
'    Next
'    Bounds = r
'End Function
Friend Sub EnumDispSettings()
    On Error Resume Next
    Dim dm     As typDevMODE
    Dim retval As Long
    dm.dmSize = Len(dm)
    retval = apiEnumDisplaySettings(vbNullString, ENUM_CURRENT_SETTINGS, dm)
    Dim txt As String
    ' txt = txt & "Bits Per Pixel: " & dm.dmBitsPerPixel & vbCrLf
    txt = txt & "Width (Pixels): " & dm.dmPelsWidth & vbCrLf
    txt = txt & "Height (Pixels):" & dm.dmPelsHeight & vbCrLf
    txt = txt & "Display Freq.:  " & dm.dmDisplayFrequency & vbCrLf
    MsgBox txt
End Sub

