VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Process"
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
'Attribute VB_Name = "Process"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const ANYSIZE_ARRAY             As Long = 1
Private Const DESKTOP_WINSTATION0       As String = "WinSta0"
Private Const DESKTOP_LOGON             As String = "Winlogon"
Private Const DESKTOP_DEFAULT           As String = "Default"
Private Const ERROR_SUCCESS             As Long = 0
Private Const ERROR_NOT_ALL_ASSIGNED    As Long = 1300
Private Const MAX_PATH                  As Integer = 260
Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const PROCESS_TERMINATE         As Long = &H1
Private Const PROCESS_QUERY_INFORMATION As Long = &H400
Private Const SE_DEBUG_NAME             As String = "SeDebugPrivilege"
Private Const SE_PRIVILEGE_ENABLED      As Long = &H2
Private Const SEE_MASK_CLASSKEY         As Long = &H3
Private Const SEE_MASK_CLASSNAME        As Long = &H1
Private Const SEE_MASK_CONNECTNETDRV    As Long = &H80
Private Const SEE_MASK_DOENVSUBST       As Long = &H200
Private Const SEE_MASK_FLAG_DDEWAIT     As Long = &H100
Private Const SEE_MASK_FLAG_NO_UI       As Long = &H400
Private Const SEE_MASK_HOTKEY           As Long = &H20
Private Const SEE_MASK_ICON             As Long = &H10
Private Const SEE_MASK_IDLIST           As Long = &H4
Private Const SEE_MASK_INVOKEIDLIST     As Long = &HC
Private Const SEE_MASK_NOCLOSEPROCESS   As Long = &H40
Private Const TH32CS_SNAPPROCESS        As Long = 2
Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const TOKEN_QUERY               As Long = &H8
Private Const WTS_CURRENT_SERVER_HANDLE As Long = 0
Private Const CREATE_DEFAULT_ERROR_MODE As Long = &H4000000
Private Const CREATE_NEW_CONSOLE        As Long = &H10
Private Const CREATE_NEW_PROCESS_GROUP  As Long = &H200
Private Const LOGON_WITH_PROFILE        As Long = 1
Private Const LOGON_NETCREDENTIALS_ONLY As Long = &H2
Private Const LOGON32_LOGON_INTERACTIVE As Long = 2
Private Const LOGON32_PROVIDER_DEFAULT  As Long = 0
Private Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Private Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Byte
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type LUID
    LowPart As Long
    HighPart As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid As LUID
    Attributes As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Private Type WTS_PROCESS_INFO
    SessionId As Long
    ProcessID As Long
    pProcessName As Long
    pUserSid As Long
End Type
Private Type STARTUPINFOW
    cbSize As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Private Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hwnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Sub apiCopyMemoryWTSLONG Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As WTS_PROCESS_INFO, ByRef source As Long, ByVal Length As Long)
Private Declare Sub apiCopyMemoryBYTELONG Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Byte, ByRef source As Long, ByVal Length As Long)
Private Declare Function apiCreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByRef lProcessID As Long) As Long
Private Declare Function apiCreateProcess Lib "kernel32" Alias "CreateProcessW" (ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByRef lpProcessAttributes As Long, ByRef lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByRef lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFOW, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Sub apisExitProcess Lib "kernel32" Alias "ExitProcess" (ByVal uExitCode As Long)
Private Declare Function apiGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
Private Declare Function apiGetCurrentProcessId Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Function apiGetCurrentProcessToken Lib "kernel32" Alias "GetCurrentProcessToken" () As Long
Private Declare Function apiGetExitCodeProcess Lib "kernel32" Alias "GetExitCodeProcess" (ByVal hProcess As Long, ByRef lpExitCode As Long) As Long
Private Declare Function apiGetPriorityClass Lib "kernel32" Alias "GetPriorityClass" (ByVal hProcess As Long) As Long
Private Declare Function apiGetProcessAffinityMask Lib "kernel32" Alias "GetProcessAffinityMask" (ByVal hProcess As Long, ByRef lpProcessAffinityMask As Long, ByRef SystemAffinityMask As Long) As Long
Private Declare Function apiGetProcessHandleCount Lib "kernel32" Alias "GetProcessHandleCount" (ByVal hProcess As Long, ByRef pdwHandleCount As Long) As Long
Private Declare Function apiGetProcessId Lib "kernel32" Alias "GetProcessId" (ByVal hProcess As Long) As Long
Private Declare Function apiGetProcessTimes Lib "kernel32" (ByVal hProcess As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Long
Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function apiProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function apiProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, ByRef uProcess As PROCESSENTRY32) As Long
Private Declare Function apiSetPriorityClass Lib "kernel32" Alias "SetPriorityClass" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Private Declare Function apiSetProcessAffinityMask Lib "kernel32" Alias "SetProcessAffinityMask" (ByVal hProcess As Long, ByVal dwProcessAffinityMask As Long) As Long
Private Declare Function apiTerminateProcess Lib "kernel32" Alias "TerminateProcess" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function apiWaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiAdjustTokenPrivileges Lib "advapi32" Alias "AdjustTokenPrivileges" (ByVal tokenhandle As Long, ByVal DisableAllPrivileges As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef ReturnLength As Long) As Long
Private Declare Function apiLookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, ByRef lpLuid As LUID) As Long
Private Declare Function apiLookupAccountSid Lib "advapi32" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal Name As String, ByRef cbName As Long, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Integer) As Long
Private Declare Function apiOpenProcessToken Lib "advapi32" Alias "OpenProcessToken" (ByVal ProcessHandle As Long, ByVal desiredAccess As Long, ByRef tokenhandle As Long) As Long
Private Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function apiGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function apiWTSEnumerateProcesses Lib "Wtsapi32" Alias "WTSEnumerateProcessesA" (ByVal hServer As Long, ByVal Reserved As Long, ByVal Version As Long, ByRef ppProcessInfo As Long, ByRef pCount As Long) As Long
Private Declare Sub apiWTSFreeMemory Lib "Wtsapi32" Alias "WTSFreeMemory" (ByVal pMemory As Long)
Private Declare Function apiShellExecuteEx Lib "shell32" Alias "ShellExecuteEx" (ByRef SEI As SHELLEXECUTEINFO) As Boolean
Private Declare Function apiGetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Private Declare Function apiCreateProcessWithLogon Lib "advapi32" Alias "CreateProcessWithLogonW" (ByVal lpUserName As Long, ByVal lpDomain As Long, ByVal lpPassword As Long, ByVal dwLogonFlags As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInfo As PROCESS_INFORMATION) As Long
Private Declare Function apiCreateProcessAsUser Lib "advapi32" Alias "CreateProcessAsUserA" (ByVal hToken As Long, ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, ByRef lpStartupInfo As STARTUPINFO, ByRef lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function apiLogonUser Lib "advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, ByRef phToken As Long) As Long
Private Declare Function apiQueryFullProcessImageName Lib "kernel32" Alias "QueryFullProcessImageNameA" (ByVal hProcess As Long, ByVal dwFlags As Long, ByVal lpExeName As String, ByRef lpdwSize As Long) As Long
Private Enum ProcessPriorityClass
    NORMALPRIORITY = &H20
    IDLE = &H40
    HIGH = &H80
    REALTIME = &H100
    ABOVENORMAL = 32768
    BELOWNORMAL = 16384
End Enum
Private mvarProcessName      As String
Private mvarId               As Long
Private mvarHandle           As Long
Private mvarMainWindowTitle  As String
Private mvarMainWindowHandle As Long
Private mvarUserName         As String
Private mvarSessionID        As Long
Private mvarStartInfo        As ProcessStartInfo
Friend Property Set StartInfo(ByVal vData As ProcessStartInfo)
    Set mvarStartInfo = vData
End Property
Friend Property Get StartInfo() As ProcessStartInfo
    Set StartInfo = mvarStartInfo
End Property
Friend Property Let ProcessName(ByVal vData As String)
    mvarProcessName = vData
End Property
Friend Property Get ProcessName() As String
    ProcessName = mvarProcessName
End Property
Friend Property Let Id(ByVal vData As Long)
    mvarId = vData
End Property
Friend Property Get Id() As Long
    Id = mvarId
End Property
Friend Property Get Handle() As Long
    Handle = mvarHandle
End Property
Friend Property Let Handle(ByVal vData As Long)
    mvarHandle = vData
End Property
Friend Property Let MainWindowTitle(ByVal vData As String)
    mvarMainWindowTitle = vData
End Property
Friend Property Get MainWindowTitle() As String
    MainWindowTitle = mvarMainWindowTitle
End Property
Friend Property Let MainWindowHandle(ByVal vData As Long)
    mvarMainWindowHandle = vData
End Property
Friend Property Get MainWindowHandle() As Long
    MainWindowHandle = mvarMainWindowHandle
End Property
Friend Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property
Friend Property Get UserName() As String
    UserName = mvarUserName
End Property
Friend Property Get SessionId() As Long
    SessionId = mvarSessionID
End Property
Friend Property Let SessionId(ByVal vData As Long)
    mvarSessionID = vData
End Property
Friend Function Start(Optional ByVal FileName As String = "", Optional ByVal Arguments As String = "") As Long
    On Error Resume Next
    If FileName <> "" Then
       With StartInfo
         .FileName = FileName
         .Arguments = Arguments
         .WorkingDirectory = vbNullChar
       End With
    End If
    Dim SEI As SHELLEXECUTEINFO
    Dim r   As Boolean
    With SEI
        .cbSize = Len(SEI)
        .fMask = SEE_MASK_NOCLOSEPROCESS ' SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        .hwnd = 0
        .lpVerb = StartInfo.Verb
        .lpFile = StartInfo.FileName
        .lpParameters = Arguments
        .lpDirectory = StartInfo.WorkingDirectory 'vbNullChar
        .nShow = Math.Abs(CLng(Not StartInfo.CreateNoWindow))
        .hInstApp = 0
        .lpIDList = 0
    End With
    r = apiShellExecuteEx(SEI)
    If r = False Then Exit Function
    Start = apiGetProcessId(SEI.hProcess)
End Function
'Private Sub cmdStart_Click()
''    txtFilePath.Text = "C:\Windows\Notepad.exe"
''    txtUser.Text = "Owner"
''    txtPassword.Text = " "
'  Call StartProcessLoggedAs(txtFilePath.Text, txtUser.Text, txtPassword.Text)
'End Sub
Friend Function StartAs(ByVal fPath As String, ByVal User As String, ByVal Password As String) As Long
    On Error Resume Next
    Dim StartInfo          As STARTUPINFO
    Dim ProcessInfo        As PROCESS_INFORMATION
    Dim lpDomain           As String
    Dim lpCommandLine      As String
    Dim lpCurrentDirectory As String
    Dim lpUserName         As String
    Dim lpPassword         As String
    lpUserName = User
    lpDomain = ""
    lpPassword = Password
    lpCommandLine = vbNullString
    lpCurrentDirectory = vbNullString
    StartInfo.cb = LenB(StartInfo)
    StartInfo.dwFlags = 0
    StartAs = apiCreateProcessWithLogon(StrPtr(User), StrPtr(lpDomain), StrPtr(Password), LOGON_WITH_PROFILE, StrPtr(fPath), StrPtr(lpCommandLine), CREATE_DEFAULT_ERROR_MODE Or CREATE_NEW_CONSOLE Or CREATE_NEW_PROCESS_GROUP, ByVal 0, StrPtr(lpCurrentDirectory), StartInfo, ProcessInfo)
    Call apiCloseHandle(ProcessInfo.hThread)
    Call apiCloseHandle(ProcessInfo.hProcess)
End Function
Friend Function RunAsUser(ByVal UserName As String, ByVal Password As String, ByVal DomainName As String, ByVal CommandLine As String, ByVal CurrentDirectory As String) As Long
    Dim ret    As Long
    Dim hToken As Long
    Dim si     As STARTUPINFO
    Dim pi     As PROCESS_INFORMATION
    ret = apiLogonUser(UserName, DomainName, Password, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, hToken)
    If ret = 0 Then
        RunAsUser = Err.LastDllError ' 1314 error code, if the user associatedy does not have "Act as part of the operating system" permission
        MsgBox "LogonUser() failed with error " & Err.LastDllError, vbExclamation
        Exit Function
    End If
    si.cb = Len(si)
    ret = apiCreateProcessAsUser(hToken, 0, CommandLine, 0, 0, False, CREATE_DEFAULT_ERROR_MODE, 0, CurrentDirectory, si, pi)
    If ret = 0 Then
        RunAsUser = Err.LastDllError
        ' 1314 error code, if the user does not have permissions   "Replace a process level token" and "Increase Quotoas"
        MsgBox "CreateProcessAsUser() failed with error " & Err.LastDllError, vbExclamation
        apiCloseHandle hToken
        Exit Function
    End If
    apiCloseHandle hToken
    apiCloseHandle pi.hThread
    apiCloseHandle pi.hProcess
    RunAsUser = 0
End Function
Friend Function WaitForExit(ByVal phandle As Long, ByVal dwMilliseconds As Long) As Long
    On Error Resume Next
    Call apiWaitForSingleObject(phandle, 0)
    WaitForExit = CBool(apiWaitForSingleObject(phandle, dwMilliseconds))
    Dim ExitCode As Long: ExitCode = -1
    Call apiGetExitCodeProcess(phandle, ExitCode)
    WaitForExit = ExitCode
    Call apiCloseHandle(phandle)
End Function
Friend Function WaitForInputIdle(ByVal phandle As Long, ByVal dwMilliseconds As Long) As Boolean
    On Error GoTo poop
    Call apiWaitForInputIdle(phandle, 0) 'avoid lockout bug with infinite time parameter and invalid process handle
    WaitForInputIdle = CBool(apiWaitForInputIdle(phandle, dwMilliseconds))
poop:
End Function
Friend Function GetProcesses() As Process()
    On Error Resume Next
    Dim p()  As Process
    Dim ret  As Long
    Dim cnt  As Long
    Dim buff As Long
    ret = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0, 1, buff, cnt)
    If ret <> 0 Then
        Dim udtProcessInfo As WTS_PROCESS_INFO
        Dim i              As Integer: i = 0
        Dim src            As Long
        src = buff
        Do
            On Error GoTo skip
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal src, LenB(udtProcessInfo))
            Dim pName As String
            If udtProcessInfo.ProcessID = 0 Then pName = "System Idle Process"
            If udtProcessInfo.ProcessID <> 0 Then pName = GetStringFromLP(udtProcessInfo.pProcessName)
            ReDim Preserve p(i)
            Dim proc As New Process
            proc.ProcessName = pName
            proc.Id = udtProcessInfo.ProcessID
            proc.Handle = 0
            proc.MainWindowHandle = 0 '() = mwh
            proc.MainWindowTitle = "" '() = mwt
            proc.UserName = GetUserName(udtProcessInfo.pUserSid)
            proc.SessionId = udtProcessInfo.SessionId
            Set p(i) = proc
            Set proc = Nothing
            i = i + 1
skip:
            src = src + LenB(udtProcessInfo)
            If i = cnt Then Exit Do
        Loop
    End If
    Call apiWTSFreeMemory(buff)
    GetProcesses = p
End Function
Friend Function GetProcessesByName(ByVal Name As String) As Process()
    On Error Resume Next
    Dim p() As Process
    GetProcessesByName = p
    Dim ret  As Long
    Dim cnt  As Long
    Dim buff As Long
    Name = Trim(LCase(Name))
    If Right(Name, 4) = ".exe" Then Name = Left(Name, Len(Name) - 4)
    ret = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0, 1, buff, cnt)
    If ret <> 0 Then
        Dim udtProcessInfo As WTS_PROCESS_INFO
        Dim i              As Integer: i = 0
        Dim n              As Integer: n = 0
        Dim src            As Long
        src = buff
        Do
            On Error GoTo skip
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal src, LenB(udtProcessInfo))
            Dim pName As String
            If udtProcessInfo.ProcessID = 0 Then pName = "System Idle Process"
            If udtProcessInfo.ProcessID <> 0 Then pName = GetStringFromLP(udtProcessInfo.pProcessName)
            If pName = Name Then
                ReDim Preserve p(n)
                p(n).ProcessName = pName
                p(n).Id = udtProcessInfo.ProcessID
                p(n).Handle = 0
                p(n).MainWindowHandle = 0 '() = mwh
                p(n).MainWindowTitle = "" '() = mwt
                p(n).UserName = GetUserName(udtProcessInfo.pUserSid)
                p(n).SessionId = udtProcessInfo.SessionId
                n = n + 1
            End If
            i = i + 1
skip:
            src = src + LenB(udtProcessInfo)
            If i = cnt Then Exit Do
        Loop
    End If
    Call apiWTSFreeMemory(buff)
    GetProcessesByName = p
End Function
Friend Function GetProcessById(ByVal pid As Long) As Process
    On Error Resume Next
    Dim p As Process
    GetProcessById = p
    Dim ret  As Long
    Dim cnt  As Long
    Dim buff As Long
    ret = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0, 1, buff, cnt)
    If ret <> 0 Then
        Dim udtProcessInfo As WTS_PROCESS_INFO
        Dim i              As Integer: i = 0
        Dim src            As Long
        src = buff
        Do
            On Error GoTo skip
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal src, LenB(udtProcessInfo))
            If udtProcessInfo.ProcessID = pid Then
                Dim pName As String
                If udtProcessInfo.ProcessID = 0 Then pName = "System Idle Process"
                If udtProcessInfo.ProcessID <> 0 Then pName = GetStringFromLP(udtProcessInfo.pProcessName)
                p.ProcessName = pName
                p.Id = udtProcessInfo.ProcessID
                p.Handle = 0
                p.MainWindowHandle = 0 '() = mwh
                p.MainWindowTitle = "" '() = mwt
                p.UserName = GetUserName(udtProcessInfo.pUserSid)
                p.SessionId = udtProcessInfo.SessionId
                Exit Do
            End If
            i = i + 1
skip:
            src = src + LenB(udtProcessInfo)
            If i = cnt Then Exit Do
        Loop
    End If
    Call apiWTSFreeMemory(buff)
    GetProcessById = p
End Function
Friend Function GetCurrentProcess() As Process
    On Error Resume Next
    Dim p As Process
    GetCurrentProcess = p
    p = GetProcessById(apiGetCurrentProcessId)
    GetCurrentProcess = p
End Function
Friend Function Kill_ByPID(ByVal pid As Long) As Long
    On Error Resume Next
    Kill_ByPID = -1
    If pid = 0 Then Exit Function
    If AdjustToken = 0 Then Exit Function
    Dim hProcess As Long
    hProcess = apiOpenProcess(PROCESS_TERMINATE, 0, pid)
    If hProcess = 0 Then Exit Function
    Dim ExitCode As Long: ExitCode = -1
    Call apiTerminateProcess(hProcess, ExitCode)
    Call apiCloseHandle(hProcess)
    Kill_ByPID = ExitCode
End Function
Friend Function Kill_ByHWND(ByVal hwnd As Long) As Long
    On Error Resume Next
    Kill_ByHWND = -1
    Dim pid As Long
    pid = GetPIDFromHWND(hwnd)
    If pid = -1 Then Exit Function
    If AdjustToken = 0 Then Exit Function
    Dim hProcess As Long
    hProcess = apiOpenProcess(PROCESS_TERMINATE, 0, pid)
    If hProcess = 0 Then Exit Function
    Dim ExitCode As Long: ExitCode = -1
    Call apiTerminateProcess(hProcess, ExitCode)
    Call apiCloseHandle(hProcess)
    Kill_ByHWND = ExitCode
End Function
Friend Sub Kill_All(ByVal Name As String)
    On Error Resume Next
    Name = Trim(LCase(Name))
    If Right(Name, 4) = ".exe" Then Name = Left(Name, Len(Name) - 4)
    If AdjustToken = 0 Then Exit Sub ' we aquire the right to adjust the token to kill, or abort
    Dim p() As Process
    p = GetProcessesByName(Name)
    Call TerminateProcesses(p, Trim(LCase(Name))) 'kill all matching process by pid
End Sub
Private Function TerminateProcesses(ByRef p() As Process, ByVal pName As String) As Long()
    Dim exits() As Long
    If Trim(pName) = "" Then Exit Function
    pName = Trim(LCase(pName))
    If Right(pName, 4) = ".exe" Then pName = Left(pName, Len(pName) - 4)
    Dim i As Integer: i = 0
    Dim n As Integer: n = 0
    For i = 0 To UBound(p) ' loop through arrays
        On Error GoTo skipFor
        If pName = p(i).ProcessName Then
            Dim hProcess As Long
            hProcess = apiOpenProcess(PROCESS_TERMINATE, 0, p(i).Id) 'Open the process to kill
            If hProcess <> 0 Then
                On Error GoTo skipTerminate
                Dim ExitCode As Long: ExitCode = -1
                Call apiTerminateProcess(hProcess, ExitCode) 'Obtained process handle, kill the process
                ReDim Preserve exits(n)
                exits(n) = ExitCode
                n = n + 1
skipTerminate:
                Call apiCloseHandle(hProcess)
            End If
        End If
skipFor:
    Next i
End Function
Friend Function GetProcessIDs(ByVal pName As String) As Long()
    On Error Resume Next
    Dim pIDs() As Long
    GetProcessIDs = pIDs
    pName = Trim(LCase(pName))
    If Right(pName, 4) = ".exe" Then pName = Left(pName, Len(pName) - 4)
    Dim ret  As Long
    Dim cnt  As Long
    Dim buff As Long
    ret = apiWTSEnumerateProcesses(WTS_CURRENT_SERVER_HANDLE, 0, 1, buff, cnt)
    If ret <> 0 Then
        Dim p              As Long
        Dim udtProcessInfo As WTS_PROCESS_INFO
        Dim tmpname        As String
        p = buff
        Dim i As Long: i = 1
        Dim n As Long: n = 0
        Do
            On Error GoTo skip
            Call apiCopyMemoryWTSLONG(udtProcessInfo, ByVal p, LenB(udtProcessInfo))
            If udtProcessInfo.ProcessID <> 0 Then
                tmpname = GetStringFromLP(udtProcessInfo.pProcessName)
                If LCase(tmpname) = LCase(pName) Then
                    ReDim Preserve pIDs(n)
                    pIDs(n) = udtProcessInfo.ProcessID
                    n = n + 1
                End If
            End If
skip:
            p = p + LenB(udtProcessInfo)
            If i = cnt Then Exit Do
            i = i + 1
        Loop
    End If
    Call apiWTSFreeMemory(buff)
    GetProcessIDs = pIDs
End Function
Friend Function GetPIDFromHWND(ByVal hwnd As Long) As Long
    On Error Resume Next
    GetPIDFromHWND = -1
    Dim pid As Long
    Dim ret As Long
    ret = apiGetWindowThreadProcessId(hwnd, pid)
    If pid <> 0 Then GetPIDFromHWND = pid
End Function
Friend Function SetPriorityClass(ByVal ppc As ProcessPriorityClass, Optional ByVal hProcess As Long) As Long
    On Error Resume Next
    If hProcess = 0 Then hProcess = apiGetCurrentProcess
    SetPriorityClass = apiSetPriorityClass(hProcess, ppc)
End Function
Friend Function GetPriorityClass(Optional ByVal hProcess As Long) As ProcessPriorityClass
    On Error Resume Next
    If hProcess = 0 Then hProcess = apiGetCurrentProcess
    GetPriorityClass = (apiGetPriorityClass(hProcess))
End Function
Friend Function SetProcessAffinityByName(ByVal pName As String, Optional ByVal one As Boolean, Optional ByVal two As Boolean, Optional ByVal three As Boolean, Optional ByVal four As Boolean, Optional ByVal five As Boolean, Optional ByVal six As Boolean, Optional ByVal seven As Boolean, Optional ByVal eight As Boolean) As Long
    On Error Resume Next
    Const AffinityMask As Long = &HF
    Dim BitMasks()     As Long
    Dim mask           As Long
    Dim hProcess       As Long
    Dim pIDs()         As Long
    pName = Trim(LCase(pName))
    If Right(pName, 4) = ".exe" Then pName = Left(pName, Len(pName) - 4)
    pIDs = GetProcessIDs(pName)
    Dim pid As Variant
    For Each pid In pIDs
        On Error GoTo skipFor
        hProcess = apiOpenProcess(PROCESS_ALL_ACCESS, 0, CLng(pid))
        If hProcess <> 0 Then
            On Error GoTo skipSetAffinity
            BitMasks() = GetBitMasks(AffinityMask)     'mask = BitMasks(0) Or BitMasks(2) Or  BitMasks(4) 'CPUs 0, 2, and 4
            If one = True Then mask = mask Or BitMasks(0)
            If two = True Then mask = mask Or BitMasks(1)
            If three = True Then mask = mask Or BitMasks(2)
            If four = True Then mask = mask Or BitMasks(3)
            If five = True Then mask = mask Or BitMasks(4)
            If six = True Then mask = mask Or BitMasks(5)
            If seven = True Then mask = mask Or BitMasks(6)
            If eight = True Then mask = mask Or BitMasks(7)
            Dim ret As Long
            ret = apiSetProcessAffinityMask(hProcess, mask)
            SetProcessAffinityByName = ret
skipSetAffinity:
            Call apiCloseHandle(hProcess)
        End If
skipFor:
    Next
End Function
Friend Function SetProcessAffinityByPID(ByVal pid As Long, Optional ByVal one As Boolean, Optional ByVal two As Boolean, Optional ByVal three As Boolean, Optional ByVal four As Boolean, Optional ByVal five As Boolean, Optional ByVal six As Boolean, Optional ByVal seven As Boolean, Optional ByVal eight As Boolean) As Long
    On Error Resume Next
    Const AffinityMask As Long = &HF
    Dim BitMasks()     As Long
    Dim mask           As Long
    Dim hProcess       As Long
    hProcess = apiOpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    If hProcess = 0 Then Exit Function
    BitMasks() = GetBitMasks(AffinityMask)     'mask = BitMasks(0) Or BitMasks(2) Or  BitMasks(4) 'CPUs 0, 2, and 4
    If one = True Then mask = mask Or BitMasks(0)
    If two = True Then mask = mask Or BitMasks(1)
    If three = True Then mask = mask Or BitMasks(2)
    If four = True Then mask = mask Or BitMasks(3)
    If five = True Then mask = mask Or BitMasks(4)
    If six = True Then mask = mask Or BitMasks(5)
    If seven = True Then mask = mask Or BitMasks(6)
    If eight = True Then mask = mask Or BitMasks(7)
    Dim ret As Long
    ret = apiSetProcessAffinityMask(hProcess, mask)
    SetProcessAffinityByPID = ret
    Call apiCloseHandle(hProcess)
End Function
Friend Function GetProcessAffinityByPID(ByVal pid As Long) As Long
    On Error Resume Next
    Dim hProcess As Long
    hProcess = apiOpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    If hProcess = 0 Then Exit Function
    Dim mask    As Long
    Dim sysmask As Long
    Dim ret     As Long
    ret = apiGetProcessAffinityMask(hProcess, mask, sysmask)
    GetProcessAffinityByPID = mask
    Call apiCloseHandle(hProcess)
End Function
Friend Function GetProcessAffinityByHWND(ByVal hwnd As String) As Long
    On Error Resume Next
    Dim pid As Long
    pid = GetPIDFromHWND(hwnd)
    If pid = 0 Then Exit Function
    Dim hProcess As Long
    hProcess = apiOpenProcess(PROCESS_ALL_ACCESS, 0, pid)
    If hProcess <> 0 Then
        Dim mask    As Long
        Dim sysmask As Long
        Dim ret     As Long
        ret = apiGetProcessAffinityMask(hProcess, mask, sysmask)
        GetProcessAffinityByHWND = mask
        Call apiCloseHandle(hProcess)
    End If
End Function
Friend Function GetProcessImageName(ByVal pid As Long) As String
    On Error Resume Next
    Dim hProcess As Long
    hProcess = apiOpenProcess(PROCESS_QUERY_INFORMATION, 0, pid)
    If hProcess <> 0 Then
        Dim buff As String
        Dim ch   As Long
        buff = String(MAX_PATH, Chr(0))
        ch = MAX_PATH
        Call apiQueryFullProcessImageName(hProcess, 0, buff, ch)
        buff = Left(buff, ch)
        Call apiCloseHandle(hProcess)
    End If
    GetProcessImageName = buff
End Function
Private Function GetUserName(Sid As Long) As String
    On Error Resume Next
    Dim retname   As String
    Dim retdomain As String
    retname = String(255, 0)
    retdomain = String(255, 0)
    Call apiLookupAccountSid(vbNullString, Sid, retname, 255, retdomain, 255, 0)
    GetUserName = Left(retdomain, InStr(retdomain, vbNullChar) - 1) & "\" & Left(retname, InStr(retname, vbNullChar) - 1)
End Function
Private Function GetStringFromLP(ByVal sptr As Long) As String
    On Error Resume Next
    Dim b    As Byte
    Dim s    As String
    Dim buff As String
    Do
        On Error GoTo skip
        Call apiCopyMemoryBYTELONG(b, ByVal sptr, 1) ' Get the byte/character that StrPtr is pointing to.
skip:
        If b = 0 Then Exit Do  ' If you've found a null character, then you're done.
        s = Chr(b)  ' Get the character for the byte's value
        buff = buff & s 'Add it to the string
        sptr = sptr + 1  ' Increment the pointer to next byte/char
    Loop
    GetStringFromLP = buff
    GetStringFromLP = LCase(GetStringFromLP)
    If Right(GetStringFromLP, 4) = ".exe" Then GetStringFromLP = Left(GetStringFromLP, Len(GetStringFromLP) - 4)
End Function
Private Function GetPriorityName(ByVal ppc As ProcessPriorityClass) As String
    Select Case ppc
        Case 32
            GetPriorityName = "Normal" 'basic priority
        Case 64
            GetPriorityName = "Idle" 'only runs when the system is idle
        Case 128
            GetPriorityName = "High" 'time critical
        Case 256
            GetPriorityName = "Realtime" 'highest possible
        Case 16384
            GetPriorityName = "Below normal"
        Case 32768
            GetPriorityName = "Above normal"
    End Select
End Function
Private Function GetBitMasks(ByVal inValue As Long) As Long()
    On Error Resume Next
    Dim RetArr()  As Long
    Dim NumRet    As Long
    Dim LoopBits  As Long
    Dim BitMask   As Long
    Const HighBit As Long = &H80000000
    ReDim RetArr(0 To 31) As Long
    Do
        BitMask = 2 ^ LoopBits
        If (inValue And BitMask) Then
            RetArr(NumRet) = BitMask
            NumRet = NumRet + 1
        End If
        If LoopBits = 30 Then Exit Do
        LoopBits = LoopBits + 1
    Loop
    If (inValue And HighBit) Then
        RetArr(NumRet) = HighBit
        NumRet = NumRet + 1
    End If
    If (NumRet > 0) Then ' Trim unused array items and return array
        If (NumRet < 32) Then ReDim Preserve RetArr(0 To NumRet - 1) As Long
        GetBitMasks = RetArr
    End If
End Function
Private Function AdjustToken() As Long
    On Error Resume Next
    Dim cProcess As Long
    cProcess = apiGetCurrentProcess
    If cProcess = 0 Then AdjustToken = 0: Exit Function  'check to see if we can get a handle to our own process
    Dim tokenhandle       As Long
    Dim tLuid             As LUID
    Dim tokenprivilege    As TOKEN_PRIVILEGES
    Dim newtokenprivilege As TOKEN_PRIVILEGES
    Dim buff              As Long
    On Error GoTo poop
    Call apiLookupPrivilegeValue("", SE_DEBUG_NAME, tLuid)
    tokenprivilege.PrivilegeCount = 1
    tokenprivilege.TheLuid = tLuid
    tokenprivilege.Attributes = SE_PRIVILEGE_ENABLED
    Call apiOpenProcessToken(cProcess, TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, tokenhandle)
    AdjustToken = apiAdjustTokenPrivileges(tokenhandle, False, tokenprivilege, Len(newtokenprivilege), newtokenprivilege, buff) 'If  Err.LastDllError = ERROR_NOT_ALL_ASSIGNED Then
poop:
    Call apiCloseHandle(cProcess)
End Function
'
'
'
'Public Function ProcessKillWMI(ByVal name As String)
'    Dim PROCESS As Object
'    For Each PROCESS In GetObject("winmgmts:").ExecQuery("select name from Win32_Process where name='" & name & ".exe'")
'        PROCESS.Terminate (0)
'    Next
'End Function
'Function ArrayString(ParamArray tokens()) As String()
'    ReDim arr(UBound(tokens)) As String
'    Dim i As Long
'    For i = 0 To UBound(tokens)
'        arr(i) = tokens(i)
'    Next
'    ArrayString = arr
'End Function
'Public Function GetProcessById2(ByVal pid As Long) As PROCESS
'    On Error Resume Next
'    Dim p As PROCESS
'    GetProcessById2 = p
'    Dim hSnap As Long
'    hSnap = apiCreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)  'now take a snapshot of the processes
'    If hSnap = 0 Then Exit Function
'    Dim pEntry As PROCESSENTRY32
'    pEntry.dwSize = Len(pEntry)
'    Dim x As Long
'    x = apiProcessFirst(hSnap, pEntry) 'starting with the first process, ie [system]
'    Dim i As Long
'    i = -1
'    Do
'        On Error GoTo skipdo
'        If x = 0 Then Exit Do
'        i = i + 1
'        If pEntry.th32ProcessID = pid Then
'            p.ProcessName = LCase(Replace(Left(pEntry.szexeFile, InStr(pEntry.szexeFile, Chr(0))), Chr(0), ""))  'magic line to strip outer buffer
'            p.id = pEntry.th32ProcessID
'            p.HANDLE = 0
'            p.MainWindowHandle() = mwh
'            p.MainWindowTitle() = mwt
'            Exit Do
'        End If
'        x = apiProcessNext(hSnap, pEntry)
'skipdo:
'    Loop
'    GetProcessById2 = p
'End Function
'Private Function GetProcesses2() As PROCESS()
'    On Error Resume Next
'    Dim p() As PROCESS
'    GetProcesses2 = p
'    Dim hSnap As Long
'    hSnap = apiCreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
'    If hSnap = 0 Then Exit Function
'    Dim pEntry As PROCESSENTRY32
'    pEntry.dwSize = Len(pEntry)
'    Dim x As Long
'    x = apiProcessFirst(hSnap, pEntry)
'    Dim i As Long
'    i = -1
'    Do
'        On Error GoTo skipdo
'        If x = 0 Then Exit Do
'        i = i + 1
'        ReDim Preserve p(i)
'        p(i).ProcessName = LCase(Replace(Left(pEntry.szexeFile, InStr(pEntry.szexeFile, Chr(0))), Chr(0), ""))  'magic line to strip outer buffer
'        p(i).id = pEntry.th32ProcessID
'        p(i).HANDLE = 0
'        p(i).MainWindowHandle() = mwh
'        p(i).MainWindowTitle() = mwt
'        x = apiProcessNext(hSnap, pEntry)
'skipdo:
'    Loop
'    GetProcesses2 = p
'End Function
'Public Function GetProcessesByName2(ByVal name As String) As PROCESS()
'    On Error Resume Next
'    Dim p() As PROCESS
'    GetProcessesByName2 = p
'    Dim hSnap As Long
'    hSnap = apiCreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
'    If hSnap = 0 Then Exit Function
'    Dim pEntry As PROCESSENTRY32
'    pEntry.dwSize = Len(pEntry)
'    Dim x As Long
'    x = apiProcessFirst(hSnap, pEntry)
'    Dim i As Long
'    i = -1
'    Dim upper As Long
'    Do
'        On Error GoTo skipdo
'        If x = 0 Then Exit Do
'        i = i + 1
'        Dim pname As String
'        pname = LCase(Replace(Left(pEntry.szexeFile, InStr(pEntry.szexeFile, Chr(0))), Chr(0), "")) 'magic line to strip outer buffer
'        If pname = name Then
'            ReDim Preserve p(i)
'            p(i).ProcessName = pname
'            p(i).id = pEntry.th32ProcessID
'            p(i).HANDLE = 0
'             p(i).MainWindowHandle() = mwh
'             p(i).MainWindowTitle() = mwt
'        End If
'        x = apiProcessNext(hSnap, pEntry)
'skipdo:
'    Loop
'    GetProcessesByName2 = p
'End Function
'Public Sub Kill_All2(ByVal name As String)
'    On Error Resume Next
'    If AdjustToken = 0 Then Exit Sub ' we aquire the right to adjust the token to kill, or abort
'    Dim p() As PROCESS
'    p = GetProcessesByName2(Trim(LCase(name)))
'    Call TerminateProcesses(p, Trim(LCase(name)))  'kill all matching process by pid
'End Sub
'
'Public Function Start(ByVal FileName As String, ByVal arguments As String) As Long
''    On Error Resume Next
''    Dim psi   As STARTUPINFOW
''    Dim pInfo As PROCESS_INFORMATION
''    psi.cbSize = Len(psi)
''    psi.lpTitle = StrPtr(DESKTOP_DEFAULT)
''    psi.lpDesktop = StrPtr(DESKTOP_DEFAULT)
''
''    Dim cmd As String
''    cmd = StrPtr(fileName & " " & arguments)
''    Start = apiCreateProcess(StrPtr(fileName), StrPtr(cmd), ByVal 0, ByVal 0, 1, 0, ByVal 0, ByVal 0, psi, pInfo)
''  'apiCloseHandle (pInfo.hprocess)
''  'apiCloseHandle (pInfo.hThread)
'End Function

