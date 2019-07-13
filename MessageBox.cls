VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "MessageBox"
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
'Attribute VB_Name = "MessageBox"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const DESKTOP_SECURE            As Long = 131527 'my secure version is:CREATEMENU,CREATEWINDOW,ENUMERATE,READOBJECTS,SWITCHDESKTOP,WRITEOBJECTS,READ_CONTROL
Private Const DESKTOP_SWITCHDESKTOP     As Long = 256 'Required to activate the desktop using the SwitchDesktop function.
Private Const SND_ASYNC                 As Long = 1
Private Const SND_NOSTOP                As Long = 16
Private Const SND_PURGE                 As Long = 64
Private Const SND_FILENAME              As Long = 131072
Private Const SPI_SETDESKWALLPAPER      As Long = 20
Private Const SPIF_UPDATEINIFILE        As Long = 1
Private Const SPIF_SENDWININICHANGE     As Long = 2
Private Const DESKTOP_LOGON             As String = "Winlogon"
Private Const DESKTOP_WINSTATION0       As String = "WinSta0"
Private Const DESKTOP_DEFAULT           As String = "Default"
Private Const MB_APPLMODAL              As Long = &H0
Private Const MB_OK                     As Long = &H0
Private Const MB_OKCANCEL               As Long = &H1
Private Const MB_ABORTRETRYIGNORE       As Long = &H2
Private Const MB_YESNOCANCEL            As Long = &H3
Private Const MB_YESNO                  As Long = &H4
Private Const MB_RETRYCANCEL            As Long = &H5
Private Const MB_CANCELTRYCONTINUE      As Long = &H6
'
Private Const MB_ICONERROR              As Long = &H10
Private Const MB_ICONHAND               As Long = &H10
Private Const MB_ICONSTOP               As Long = &H10
Private Const MB_ICONQUESTION           As Long = &H20
Private Const MB_ICONEXCLAMATION        As Long = &H30
Private Const MB_ICONASTERISK           As Long = &H40
Private Const MB_ICONWARNING            As Long = &H30
Private Const MB_ICONINFORMATION        As Long = &H40
'
Private Const MB_USERICON               As Long = &H80
Private Const MB_HELP                   As Long = &H4000
Private Const MB_MAX_TIMEOUT            As Long = &HFFFFFFFF '49 days
Private Const MB_SERVICE_NOTIFICATION   As Long = &H200000
Private Const MB_YES_NO_SECURE          As Long = MB_YESNO Or MB_ICONQUESTION Or MB_SERVICE_NOTIFICATION
'
Private Const MB_DEFBUTTON1             As Long = &H0
Private Const MB_DEFBUTTON2             As Long = &H100
Private Const MB_DEFBUTTON3             As Long = &H200
Private Const MB_DEFBUTTON4             As Long = &H300
'
Private Const MB_SYSTEMMODAL            As Long = &H1000
Private Const MB_TASKMODAL              As Long = &H2000
Private Const MB_NOFOCUS                As Long = &H8000
Private Const MB_SETFOREGROUND          As Long = &H10000
Private Const MB_DEFAULT_DESKTOP_ONLY   As Long = &H20000
Private Const MB_TOPMOST                As Long = &H40000
Private Const MB_RIGHT                  As Long = &H80000
Private Const MB_RTLREADING             As Long = &H100000
Private Const WIN32_NULL                As Long = 0
Private Const WIN32_FALSE               As Long = 0
Private Const WIN32_TRUE                As Long = Not WIN32_FALSE 'Not for testing.
Private Const WTS_CURRENT_SERVER_HANDLE As Long = WIN32_NULL
Private Const WTS_CURRENT_SESSION       As Long = -1 'For sending to ourself.
Private Const WTS_ANY_SESSION           As Long = -2
'To be able to send a message an another user's session, you need to have the Message permission.
'For sending from a Windows Service, Scheduled Task, etc.:
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function apiWTSGetActiveConsoleSessionId Lib "kernel32" Alias "WTSGetActiveConsoleSessionId" () As Long
Private Declare Function apiWTSSendMessage Lib "Wtsapi32" Alias "WTSSendMessageW" (ByVal hServer As Long, ByVal SessionId As Long, ByVal pTitle As Long, ByVal TitleLength As Long, ByVal pMessage As Long, ByVal MessageLength As Long, ByVal Style As Long, ByVal timeout As Long, ByRef Response As DialogResult, ByVal bWait As Long) As Long
Private Declare Function apiMessageBeep Lib "user32" Alias "MessageBeep" (ByVal wType As Long) As Long
Private Declare Function apiBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function apiMessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal prmlngWindowHandle As Long, ByVal prmstrMessage As String, ByVal prmstrCaption As String, ByVal prmlngType As Long, ByVal prmwLanguage As Integer, ByVal prmdwMiliseconds As Long) As Long
Private Declare Function apiCloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiCreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lDesktop As String, ByVal lDevice As Long, ByVal devmode As Long, ByVal dwFlags As Long, ByVal desiredAccess As Long, ByRef secAttribute As SECURITY_ATTRIBUTES) As Long
Private Declare Function apiGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
Private Declare Function apiGetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThread As Long) As Long
Private Declare Function apiGetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hwnd As Long) As Long
Private Declare Function apiOpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function apiPaintDesktop Lib "user32" Alias "PaintDesktop" (ByVal HDC As Long) As Long
Private Declare Function apiPlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiSetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiSwitchDesktop Lib "user32" Alias "SwitchDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiSystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private oldDskThread As Long
Private oldDskInput  As Long
Private hwnDsk       As Long
Friend Function Authorize(ByVal txt As String, ByVal title As String, Optional ByVal timeout As Long = 20000) As DialogResult   '(ByVal message As String, Optional ByVal title As String, Optional ByVal dskname As String, Optional ByVal exepath As String)
    On Error Resume Next
    DoEvents 'flush all user interfaces to avoid data loss before desktop switch which suspends execution of the default desktop and all UI on it
    Dim dskname As String
    Dim rn      As Long
    Randomize
    rn = Rnd * (2147483647 - 1) + 1
    dskname = CStr(rn) 'Set desk name to random string
    oldDskThread = apiGetThreadDesktop(apiGetCurrentThreadId)
    oldDskInput = apiOpenInputDesktop(0, False, DESKTOP_SWITCHDESKTOP)
    If CreateDesktop(dskname) = 0 Then Exit Function
    Call PlaySnd("") 'play classic UAC sound or xp log off
    SwitchToDeskTop
    KillCTFMON 'kill off extra process started by Windows, otherwise it will continue to run when desk is closed
    Authorize = Show(txt, title, MB_YES_NO_SECURE, timeout, 0)
    ' Authorize = Send(txt, title, 4) 'Alternate message with WTS
    CloseDeskTop 'Close the desktop we created
    apiSetThreadDesktop (oldDskThread) 'Set the thread desktop back
    apiSwitchDesktop (oldDskInput) 'If switched clear old desk
    DoEvents 'flush all user interfaces to avoid data hanging before desktop switch
End Function
Friend Function Show(ByVal Message As String, Optional ByVal Caption As String = "", Optional ByVal Flags As Long = 0, Optional ByVal TimeOutMilliseconds As Long = 20000, Optional ByVal hwnd As Long = 0) As DialogResult
    On Error GoTo poop
    Show = apiMessageBoxTimeOut(hwnd, Message, Caption, Flags, 0, TimeOutMilliseconds)
    Exit Function
poop:
    Show = -1
End Function
Friend Function Send(ByVal txt As String, ByVal title As String, ByVal timeout As Long) As DialogResult
    Dim dres As DialogResult
    On Error GoTo poop
    Dim ret As Long
    timeout = CInt(timeout / 1000)
    ret = apiWTSSendMessage(WTS_CURRENT_SERVER_HANDLE, WTS_CURRENT_SESSION, StrPtr(title), LenB(title) + 2, StrPtr(txt), LenB(txt) + 2, MB_ICONINFORMATION Or MB_OK Or MB_SERVICE_NOTIFICATION Or MB_SETFOREGROUND, timeout, dres, WIN32_TRUE)
    Send = dres
    Exit Function
poop:
    Send = -1
End Function
Friend Sub Beep(ByVal Frequency As Long, ByVal duration As Long)
    On Error Resume Next
    Call apiBeep(Frequency, duration) '100-32767
End Sub
Friend Sub MessageBeep()
    On Error Resume Next
    Call apiMessageBeep(&HFFFFFFFF)
End Sub
Private Function CreateDesktop(ByVal sDesktopName As String) As Long
    On Error Resume Next
    Dim sa As SECURITY_ATTRIBUTES
    hwnDsk = apiCreateDesktop(sDesktopName, ByVal 0, ByVal 0, 0, DESKTOP_SECURE, sa)
    If hwnDsk = 0 Then CreateDesktop = 0: Exit Function
    CreateDesktop = hwnDsk
End Function
Private Function SwitchToDeskTop() As Long
    On Error Resume Next
    Dim st As Long
    Dim sd As Long
    st = apiSetThreadDesktop(hwnDsk)
    sd = apiSwitchDesktop(hwnDsk)
    If sd <> 0 Then SwitchToDeskTop = 1
End Function
Private Sub CloseDeskTop()
    On Error Resume Next
    apiCloseDesktop (hwnDsk)
End Sub
Private Sub KillCTFMON()
    Dim objshell
    Set objshell = CreateObject("Wscript.Shell")
    objshell.Run "taskkill /IM ctfmon.exe", 0, True
End Sub
Private Sub PlaySnd(Optional ByVal uacPath As String)
    On Error Resume Next
    Dim medPath As String
    medPath = GetSystemDirectory
    medPath = Left(medPath, Len(medPath) - 9)
    medPath = medPath & "\media\"
    If uacPath = "on" Then
        If Dir(medPath & "Windows User Account Control.wav") = "" Then
            uacPath = medPath & "Windows XP Logon Sound.wav"
        Else
            uacPath = medPath & "Windows User Account Control.wav"
        End If
    Else
        If Dir(medPath & "Windows User Account Control.wav") = "" Then
            uacPath = medPath & "Windows XP Logoff Sound.wav"
        Else
            uacPath = medPath & "Windows User Account Control.wav"
        End If
    End If
    Call apiPlaySound(vbNullString, 0, SND_FILENAME Or SND_ASYNC)
    Call apiPlaySound(uacPath, 0, SND_FILENAME Or SND_ASYNC)
End Sub
Private Function GetSystemDirectory() As String
    On Error Resume Next
    Dim ret As Long
    GetSystemDirectory = Space(260)
    ret = apiGetSystemDirectory(GetSystemDirectory, 260)
    GetSystemDirectory = Left(GetSystemDirectory, ret)
End Function
Private Function SetDesktopWallpaper(Optional ByVal imgPath As String = "") As Long
    On Error Resume Next
    SetDesktopWallpaper = apiSystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imgPath, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
End Function
Private Function PaintDesktop(ByVal hwnd As Long) As Long
    On Error Resume Next
    PaintDesktop = apiPaintDesktop(apiGetWindowDC(hwnd))
End Function

