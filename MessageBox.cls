VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DESKTOP_SECURE             As Long = 131527 'my secure version is:CREATEMENU,CREATEWINDOW,ENUMERATE,READOBJECTS,SWITCHDESKTOP,WRITEOBJECTS,READ_CONTROL
Private Const DESKTOP_SWITCHDESKTOP      As Long = 256 'Required to activate the desktop using the SwitchDesktop function.
Private Const SND_ASYNC                  As Long = 1
Private Const SND_NOSTOP                 As Long = 16
Private Const SND_PURGE                  As Long = 64
Private Const SND_FILENAME               As Long = 131072
Private Const SPI_SETDESKWALLPAPER       As Long = 20
Private Const SPIF_UPDATEINIFILE         As Long = 1
Private Const SPIF_SENDWININICHANGE      As Long = 2
Private Const DESKTOP_LOGON              As String = "Winlogon"
Private Const DESKTOP_WINSTATION0        As String = "WinSta0"
Private Const DESKTOP_DEFAULT            As String = "Default"
Private Const MB_OK                   As Long = &H0
Private Const MB_OKCANCEL             As Long = &H1
Private Const MB_ABORTRETRYIGNORE     As Long = &H2
Private Const MB_YESNOCANCEL          As Long = &H3
Private Const MB_YESNO                As Long = &H4
Private Const MB_RETRYCANCEL          As Long = &H5
Private Const MB_MAX_TIMEOUT          As Long = &HFFFFFFFF '49 days
Private Const MB_ICONERROR            As Long = &H10 ' MB_ICONSTOP MB_ICONHAND
Private Const MB_ICONQUESTION         As Long = &H20
Private Const MB_ICONWARNING          As Long = &H30 '  MB_ICONEXCLAMATION
Private Const MB_ICONINFORMATION      As Long = &H40 '  MB_ICONASTERISK
Private Const MB_SERVICE_NOTIFICATION As Long = &H200000
Private Const MB_YES_NO_SECURE        As Long = MB_YESNO Or MB_ICONQUESTION Or MB_SERVICE_NOTIFICATION
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function apiMessageBeep Lib "user32" Alias "MessageBeep" (ByVal wType As Long) As Long
Private Declare Function apiBeep Lib "kernel32" Alias "Beep" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Function apiMessageBoxTimeOut Lib "user32" Alias "MessageBoxTimeoutA" (ByVal prmlngWindowHandle As Long, ByVal prmstrMessage As String, ByVal prmstrCaption As String, ByVal prmlngType As Long, ByVal prmwLanguage As Integer, ByVal prmdwMiliseconds As Long) As Long
Private Declare Function apiCloseDesktop Lib "user32" Alias "CloseDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiCreateDesktop Lib "user32" Alias "CreateDesktopA" (ByVal lDesktop As String, ByVal lDevice As Long, ByVal devmode As Long, ByVal dwFlags As Long, ByVal desiredAccess As Long, ByRef secAttribute As SECURITY_ATTRIBUTES) As Long
Private Declare Function apiGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Long) As Long
Private Declare Function apiGetProcessWindowStation Lib "user32" Alias "GetProcessWindowStation" () As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetThreadDesktop Lib "user32" Alias "GetThreadDesktop" (ByVal dwThread As Long) As Long
Private Declare Function apiGetWindowDC Lib "user32" Alias "GetWindowDC" (ByVal hWnd As Long) As Long
Private Declare Function apiOpenInputDesktop Lib "user32" Alias "OpenInputDesktop" (ByVal dwFlags As Long, ByVal fInherit As Boolean, ByVal dwDesiredAccess As Long) As Long
Private Declare Function apiPaintDesktop Lib "user32" Alias "PaintDesktop" (ByVal hDC As Long) As Long
Private Declare Function apiPlaySound Lib "winmm" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiSetThreadDesktop Lib "user32" Alias "SetThreadDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiSwitchDesktop Lib "user32" Alias "SwitchDesktop" (ByVal hDesktop As Long) As Long
Private Declare Function apiSystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lParam As String, ByVal fuWinIni As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private oldDskThread As Long
Private oldDskInput  As Long
Private hwnDsk       As Long

Friend Function Authorize(ByVal message As String, ByVal title As String, Optional ByVal timeout As Long = 20000) As MB_RESULT   '(ByVal message As String, Optional ByVal title As String, Optional ByVal dskname As String, Optional ByVal exepath As String)
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
    Authorize = Show(message, title, MB_YES_NO_SECURE, timeout, 0)
    CloseDeskTop 'Close the desktop we created
    apiSetThreadDesktop (oldDskThread) 'Set the thread desktop back
    apiSwitchDesktop (oldDskInput) 'If switched clear old desk
    DoEvents 'flush all user interfaces to avoid data hanging before desktop switch
End Function
Friend Function Show(ByVal message As String, Optional ByVal Caption As String = "", Optional ByVal Flags As Long = 0, Optional ByVal TimeOutMilliseconds As Long = 20000, Optional ByVal hWnd As Long = 0) As MB_RESULT
    On Error GoTo poop
    Show = apiMessageBoxTimeOut(hWnd, message, Caption, Flags, 0, TimeOutMilliseconds)
    Exit Function
poop:
    Show = -1
End Function
Friend Sub Beep(ByVal frequency As Long, ByVal duration As Long)
    On Error Resume Next
    Call apiBeep(frequency, duration) '100-32767
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
    'Get path to windows media folder, for stock UAC sound
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
    'Clear sound, and then play
    Call apiPlaySound(vbNullString, 0, SND_FILENAME Or SND_ASYNC)
    Call apiPlaySound(uacPath, 0, SND_FILENAME Or SND_ASYNC)
End Sub
Private Function GetSystemDirectory() As String
    On Error Resume Next
    Dim ret As Long
    GetSystemDirectory = Space(260)    'Create a buffer
    ret = apiGetSystemDirectory(GetSystemDirectory, 260)  'Get sysdir
    GetSystemDirectory = Left(GetSystemDirectory, ret) 'Remove chr$(0)'s
End Function
Private Function SetDesktopWallpaper(Optional ByVal imgPath As String = "") As Long
    On Error Resume Next
    SetDesktopWallpaper = apiSystemParametersInfo(SPI_SETDESKWALLPAPER, 0, imgPath, SPIF_SENDWININICHANGE Or SPIF_UPDATEINIFILE)
End Function
Private Function PaintDesktop(ByVal hWnd As Long) As Long
    On Error Resume Next
    PaintDesktop = apiPaintDesktop(apiGetWindowDC(hWnd))
End Function
