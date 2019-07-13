VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ServiceController"
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
'Attribute VB_Name = "ServiceController"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const SERVICE_DISABLED                 As Long = &H4
Private Const SERVICE_DEMAND_START             As Long = &H3
Private Const SERVICE_AUTO_START               As Long = &H2
Private Const SERVICE_SYSTEM_START             As Long = &H1
Private Const SERVICE_BOOT_START               As Long = &H0
Private Const SERVICE_WIN32_OWN_PROCESS        As Long = &H10
Private Const SERVICE_WIN32_SHARE_PROCESS      As Long = &H20
Private Const SERVICE_WIN32                    As Long = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS
Private Const SERVICES_ACTIVE_DATABASE         As String = "ServicesActive"
Private Const SERVICE_CONTROL_STOP             As Long = &H1
Private Const SERVICE_CONTROL_PAUSE            As Long = &H2
Private Const SERVICE_CONTROL_CONTINUE         As Long = &H3
'Private Const SERVICE_STOPPED                  As Long = &H1
'Private Const SERVICE_START_PENDING            As Long = &H2
'Private Const SERVICE_STOP_PENDING             As Long = &H3
'Private Const SERVICE_RUNNING                  As Long = &H4
'Private Const SERVICE_CONTINUE_PENDING         As Long = &H5
'Private Const SERVICE_PAUSE_PENDING            As Long = &H6
'Private Const SERVICE_PAUSED                   As Long = &H7
Private Const STANDARD_RIGHTS_REQUIRED         As Long = &HF0000
Private Const SC_MANAGER_CONNECT               As Long = &H1
Private Const SC_MANAGER_CREATE_SERVICE        As Long = &H2
Private Const SC_MANAGER_ENUMERATE_SERVICE     As Long = &H4
Private Const SC_MANAGER_LOCK                  As Long = &H8
Private Const SC_MANAGER_QUERY_LOCK_STATUS     As Long = &H10
Private Const SC_MANAGER_MODIFY_BOOT_CONFIG    As Long = &H20
Private Const SC_MANAGER_ALL_ACCESS            As Long = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
Private Const SERVICE_QUERY_CONFIG             As Long = &H1
Private Const SERVICE_CHANGE_CONFIG            As Long = &H2
Private Const SERVICE_QUERY_STATUS             As Long = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS     As Long = &H8
Private Const SERVICE_ERROR_NORMAL             As Long = &H1
Private Const SERVICE_START                    As Long = &H10
Private Const SERVICE_STOP                     As Long = &H20
Private Const SERVICE_PAUSE_CONTINUE           As Long = &H40
Private Const SERVICE_INTERROGATE              As Long = &H80
Private Const SERVICE_USER_DEFINED_CONTROL     As Long = &H100
Private Const SERVICE_ALL_ACCESS               As Long = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)
Private Const SERVICE_ACTIVE                   As Long = &H1
Private Const SERVICE_INACTIVE                 As Long = &H2
Private Const GENERIC_READ                     As Long = &H80000000
Private Const ERROR_INSUFFICIENT_BUFFER        As Long = 122
Private Const ERROR_SERVICE_ALREADY_RUNNING    As Long = 1056
Private Const ERROR_SERVICE_MARKED_FOR_DELETE  As Long = 1072
Private Const ERROR_SERVICE_EXISTS             As Long = 1073
Private Const ERROR_SERVICE_DATABASE_LOCKED    As Long = 1055
Private Const ERROR_SERVICE_DEPENDENCY_DELETED As Long = 1075
Private Const ERROR_SERVICE_DEPENDENCY_FAIL    As Long = 1068
Private Const ERROR_SERVICE_DISABLED           As Long = 1058
Private Const ERROR_SERVICE_LOGON_FAILED       As Long = 1069
Private Const ERROR_SERVICE_NO_THREAD          As Long = 1054
Private Const ERROR_SERVICE_REQUEST_TIMEOUT    As Long = 1053
Private Const ERROR_SERVICE_CANNOT_ACCEPT_CTRL As Long = 1061
Private Const ERROR_SERVICE_NOT_ACTIVE         As Long = 1062
Private Const ERROR_MORE_DATA                  As Long = 234
Private Const ERROR_INVALID_HANDLE             As Long = 6
Private Const ERROR_INVALID_NAME               As Long = 123
Private Const ERROR_INVALID_PARAMETER          As Long = 87
Private Const ERROR_INVALID_SERVICE_ACCOUNT    As Long = 1057
Private Const ERROR_INVALID_SERVICE_CONTROL    As Long = 1052
Private Const ERROR_PATH_NOT_FOUND             As Long = 3
Private Const ERROR_ACCESS_DENIED              As Long = 5
Private Const ERROR_CIRCULAR_DEPENDENCY        As Long = 1059
Private Const ERROR_DUP_NAME                   As Long = 52
Private Const ERROR_DEPENDENT_SERVICES_RUNNING As Long = 1051
Private Const ERROR_SHUTDOWN_IN_PROGRESS       As Long = 1115
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type
Private Type ENUM_SERVICE_STATUS
    lpServiceName As Long
    lpDisplayName As Long
    ServiceStatus As SERVICE_STATUS
End Type
Private Declare Function apiCloseServiceHandle Lib "advapi32" Alias "CloseServiceHandle" (ByVal hSCObject As Long) As Long
Private Declare Function apiControlService Lib "advapi32" Alias "ControlService" (ByVal HService As Long, ByVal dwControl As Long, ByRef lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function apiCreateService Lib "advapi32" Alias "CreateServiceA" (ByVal hMan As Long, ByVal lpServiceName As String, ByVal lpDisplayName As String, ByVal dwDesiredAccess As Long, ByVal dwServiceType As Long, ByVal dwStartType As Long, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagId As String, ByVal lpDependencies As String, ByVal lp As String, ByVal lpPassword As String) As Long
Private Declare Function apiDeleteService Lib "advapi32" Alias "DeleteService" (ByVal hServ As Long) As Long
Private Declare Function apiEnumServicesStatus Lib "advapi32" Alias "EnumServicesStatusA" (ByVal hMan As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, ByRef lpServices As Any, ByVal cbBufSize As Long, ByRef pcbBytesNeeded As Long, ByRef lpServicesReturned As Long, ByRef lpResumeHandle As Long) As Long
Private Declare Function apiEnumServicesStatusSERVICE Lib "advapi32" Alias "EnumServicesStatusA" (ByVal hMan As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, ByRef lpServices As ENUM_SERVICE_STATUS, ByVal cbBufSize As Long, ByRef pcbBytesNeeded As Long, ByRef lpServicesReturned As Long, ByRef lpResumeHandle As Long) As Long
Private Declare Function apiGetLastError Lib "kernel32" Alias "GetLastError" () As Long
Private Declare Function apiOpenSCManager Lib "advapi32" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByRef szDest As String, ByRef szcSource As Long) As Long
Private Declare Function apiOpenService Lib "advapi32" Alias "OpenServiceA" (ByVal hMan As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function apiQueryServiceStatus Lib "advapi32" Alias "QueryServiceStatus" (ByVal HService As Long, ByRef lpServiceStatus As SERVICE_STATUS) As Long
Private Declare Function apiStartService Lib "advapi32" Alias "StartServiceA" (ByVal hServ As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Private StrInstallRemoveError As String
Private mvarServiceName        As String
Private mvarDisplayName        As String
Private mvarMachineName        As String
Private mvarServiceHandle        As Long
Private mvarStatus As Long
Public Enum e_ServiceModeStart
    e_ServiceType_Disabled = 4
    e_ServiceType_Manual = 3
    e_ServiceType_Automatic = 2
    e_ServiceType_SystemStart = 1
    e_ServiceType_BootTime = 0
End Enum
Friend Property Let DisplayName(nValue As String)
    mvarDisplayName = nValue
End Property
Friend Property Get DisplayName() As String
    DisplayName = mvarDisplayName
End Property
Friend Property Let ServiceName(nValue As String)
    mvarServiceName = nValue
End Property
Friend Property Get ServiceName() As String
    ServiceName = mvarServiceName
End Property
Friend Property Let MachineName(nValue As String)
    mvarMachineName = nValue
End Property
Friend Property Get MachineName() As String
    MachineName = mvarMachineName
End Property
Friend Property Let ServiceHandle(nValue As Long)
    mvarServiceHandle = nValue
End Property
Friend Property Get ServiceHandle() As Long
    ServiceHandle = mvarServiceHandle
End Property
Friend Property Get InstallRemoveServiceError() As String
    InstallRemoveServiceError = StrInstallRemoveError
End Property
Friend Property Let Status(nValue As Long)
    mvarStatus = nValue
End Property
Friend Property Get Status() As Long
    Status = mvarStatus
End Property
Friend Function GetStatus(Optional ByVal Computer As String = vbNullString) As ServiceControllerStatus
    Dim ss         As SERVICE_STATUS
    Dim hMgr      As Long
    Dim hServ       As Long
    Dim hServStatus As Long
    If Trim(mvarDisplayName) = "" Then Exit Function
    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hMgr <> 0 Then
        hServ = apiOpenService(hMgr, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            hServStatus = apiQueryServiceStatus(hServ, ss)
            If hServStatus <> 0 Then GetStatus = ss.dwCurrentState
            apiCloseServiceHandle hServ
        End If
        apiCloseServiceHandle hMgr
    End If
End Function
Friend Function Start(Optional ByVal Computer As String = vbNullString) As Boolean
    Dim res       As Long
    res = 0
    Dim ss         As SERVICE_STATUS
    Dim hMgr As Long
    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hMgr <> 0 Then
        Dim hServ  As Long
        hServ = apiOpenService(hMgr, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            res = apiStartService(hServ, 0, 0)
            Dim hServStatus As Long
            Dim qss         As SERVICE_STATUS
            hServStatus = apiQueryServiceStatus(hServ, qss)
            apiCloseServiceHandle hServ
            If hServStatus <> 0 Then Status = qss.dwCurrentState
        End If
        apiCloseServiceHandle hMgr
    End If
    Start = CBool(res)
End Function
Friend Function Stop_(Optional ByVal Computer As String) As Boolean
    Dim res       As Long
    res = 0
    Dim ss         As SERVICE_STATUS
    Dim hMgr As Long
    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hMgr <> 0 Then
        Dim hServ  As Long
        hServ = apiOpenService(hMgr, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            res = apiControlService(hServ, SERVICE_CONTROL_STOP, ss)
            Dim hServStatus As Long
            Dim qss         As SERVICE_STATUS
            hServStatus = apiQueryServiceStatus(hServ, qss)
            apiCloseServiceHandle hServ
            If hServStatus <> 0 Then Status = qss.dwCurrentState
        End If
        apiCloseServiceHandle hMgr
    End If
    Stop_ = CBool(res)
End Function
Friend Function Pause(Optional ByVal Computer As String = vbNullString) As Boolean
    Dim res       As Long
    res = 0
    Dim ss         As SERVICE_STATUS
    Dim hMgr As Long
    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hMgr <> 0 Then
        Dim hServ  As Long
        hServ = apiOpenService(hMgr, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            res = apiControlService(hServ, SERVICE_CONTROL_PAUSE, ss)
            Dim hServStatus As Long
            Dim qss         As SERVICE_STATUS
            hServStatus = apiQueryServiceStatus(hServ, qss)
            apiCloseServiceHandle hServ
            If hServStatus <> 0 Then Status = qss.dwCurrentState
        End If
        apiCloseServiceHandle hMgr
    End If
    Pause = CBool(res)
End Function
Friend Function Continue(Optional ByVal Computer As String = vbNullString) As Boolean
    Dim res       As Long
    res = 0
    Dim ss         As SERVICE_STATUS
    Dim hMgr As Long
    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hMgr <> 0 Then
        Dim hServ  As Long
        hServ = apiOpenService(hMgr, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            res = apiControlService(hServ, SERVICE_CONTROL_CONTINUE, ss)
            Dim hServStatus As Long
            Dim qss         As SERVICE_STATUS
            hServStatus = apiQueryServiceStatus(hServ, qss)
            apiCloseServiceHandle hServ
            If hServStatus <> 0 Then Status = qss.dwCurrentState
        End If
        apiCloseServiceHandle hMgr
    End If
    Continue = CBool(res)
End Function
Friend Function GetServices(Optional ByVal Computer As String = vbNullString) As Collection
    Dim Services                   As New Collection
    Dim strServiceName             As String * 250
    Dim lngBytesNeeded             As Long
    Dim lngServicesReturned        As Long
    Dim hNextUnreadEntry           As Long
    Dim lngStructsNeeded           As Long
    Dim lngResult                  As Long
    Dim lpEnumServiceStatus()      As ENUM_SERVICE_STATUS
    Dim i                          As Long
    Dim Hscm                       As Long
    Dim lngServiceStatusInfoBuffer As Long
    Hscm = apiOpenSCManager(Computer, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If Hscm <> 0 Then
        hNextUnreadEntry = 0
        lngResult = apiEnumServicesStatus(Hscm, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, ByVal &H0, &H0, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)
        If Not Err.LastDllError = ERROR_MORE_DATA Then
            MsgBox "LastDLLError = " & CStr(Err.LastDllError)
            Exit Function
        End If
        lngStructsNeeded = lngBytesNeeded / Len(lpEnumServiceStatus(0)) + 1
        ReDim lpEnumServiceStatus(lngStructsNeeded - 1)
        lngServiceStatusInfoBuffer = lngStructsNeeded * Len(lpEnumServiceStatus(0))
        hNextUnreadEntry = 0
        lngResult = apiEnumServicesStatusSERVICE(Hscm, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, lpEnumServiceStatus(0), lngServiceStatusInfoBuffer, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)
        If lngResult = 0 Then
            MsgBox "EnumServicesStatus failed. LastDllError = " & CStr(Err.LastDllError)
            Exit Function
        End If
        For i = 0 To lngServicesReturned - 1
            lngResult = apilstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpServiceName)
            Services.Add StripTerminator(strServiceName)
            lngResult = apilstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpDisplayName)
            Services.Add StripTerminator(strServiceName)
            Select Case lpEnumServiceStatus(i).ServiceStatus.dwCurrentState
                Case ServiceControllerStatus.Stopped
                    Services.Add "STOPPED"
                Case ServiceControllerStatus.StartPending
                    Services.Add "START PENDING"
                Case ServiceControllerStatus.StopPending
                    Services.Add "STOP PENDING"
                Case ServiceControllerStatus.Running
                    Services.Add "RUNNING"
                Case ServiceControllerStatus.ContinuePending
                    Services.Add "CONTINUE PENDING"
                Case ServiceControllerStatus.PausePending
                    Services.Add "PAUSE PENDING"
                Case ServiceControllerStatus.Paused
                    Services.Add "PAUSED"
            End Select
        Next i
        apiCloseServiceHandle (Hscm)
    End If
    Set GetServices = Services
End Function
Friend Function CanPauseAndContinue() As Boolean
 'todo
End Function
Friend Function CanShutdown() As Boolean
 'todo
End Function
Friend Function CanStop() As Boolean
 'todo
End Function
Friend Function DependentServices() As ServiceController()
 'todo
End Function
Friend Function GetDevices() As ServiceController()
 'todo
End Function



'---------------------------------------------
Friend Function ServiceInstalled() As Boolean
    On Error GoTo xc
    Dim colServices As Object
    Dim objService  As Object
    Set colServices = GetObject("winmgmts:").ExecQuery("Select Name from Win32_Service where Name = '" & mvarServiceName & "'")
    For Each objService In colServices
        If Len(objService.Name) Then
            ServiceInstalled = True
        End If
    Next
    Set colServices = Nothing
    Exit Function
xc:
    ServiceInstalled = False
End Function
Friend Function Install(ByVal ServiceFilePath As String, ByVal ServiceDisplayName As String, Optional ByVal Computer As String = vbNullString, Optional ByVal ServiceStartMode As e_ServiceModeStart = e_ServiceType_Automatic) As Boolean
    Dim Hscm     As Long
    Dim hServ As Long
    Dim ret      As Long
    Dim iFlags   As Long
    Select Case ServiceStartMode
        Case 0: iFlags = SERVICE_BOOT_START
        Case 1: iFlags = SERVICE_SYSTEM_START
        Case 2: iFlags = SERVICE_AUTO_START
        Case 3: iFlags = SERVICE_DEMAND_START
        Case 4: iFlags = SERVICE_DISABLED
    End Select
    Install = False
    Hscm = apiOpenSCManager(Computer, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If Hscm <> 0 Then
        hServ = apiCreateService(Hscm, mvarServiceName, ServiceDisplayName, SERVICE_ALL_ACCESS, SERVICE_WIN32_OWN_PROCESS, iFlags, SERVICE_ERROR_NORMAL, ServiceFilePath, vbNullString, vbNullString, vbNullString, vbNullString, vbNullString)
        If hServ <> 0 Then
            Install = True
            StrInstallRemoveError = "Service installed sucessfully"
        Else
            ret = apiGetLastError
            Select Case ret
                Case ERROR_ACCESS_DENIED
                    StrInstallRemoveError = "The handle to the specified service control manager database does not have SC_MANAGER_CREATE_SERVICE access."
                Case ERROR_CIRCULAR_DEPENDENCY
                    StrInstallRemoveError = "A circular service dependency was specified."
                Case ERROR_DUP_NAME
                    StrInstallRemoveError = "The display name already exists in the service control manager database either as a service name or as another display name."
                Case ERROR_INVALID_HANDLE
                    StrInstallRemoveError = "The handle to the specified service control manager database is invalid."
                Case ERROR_INVALID_NAME
                    StrInstallRemoveError = "The specified service name is invalid."
                Case ERROR_INVALID_PARAMETER
                    StrInstallRemoveError = "A parameter that was specified is invalid."
                Case ERROR_INVALID_SERVICE_ACCOUNT
                    StrInstallRemoveError = "The user account name specified in the lpServiceStartName parameter does not exist."
                Case ERROR_SERVICE_EXISTS
                    StrInstallRemoveError = "The specified service already exists in this database."
                Case Else
                    StrInstallRemoveError = "Unknown error"
            End Select
        End If
    End If
    apiCloseServiceHandle Hscm
    apiCloseServiceHandle hServ
End Function
Friend Function Remove(Optional ByVal Computer As String) As Boolean
    Dim Hcsm               As Long
    Dim hServ           As Long
    Dim ret                As Long
    Dim Error              As Long
    Dim StrRemServiceError As String
    ret = 0
    Hcsm = apiOpenSCManager(Computer, vbNullString, SC_MANAGER_ALL_ACCESS)
    If Hcsm <> 0 Then
        hServ = apiOpenService(Hcsm, mvarServiceName, SERVICE_ALL_ACCESS)
        If hServ <> 0 Then
            ret = apiDeleteService(hServ)
        Else
            Error = apiGetLastError
            Select Case Error
                Case ERROR_ACCESS_DENIED
                    StrRemServiceError = "The handle to the specified service control manager database does not have SC_MANAGER_CREATE_SERVICE access."
                Case ERROR_CIRCULAR_DEPENDENCY
                    StrRemServiceError = "A circular service dependency was specified."
                Case ERROR_DUP_NAME
                    StrRemServiceError = "The display name already exists in the service control manager database either as a service name or as another display name."
                Case ERROR_INVALID_HANDLE
                    StrRemServiceError = "The handle to the specified service control manager database is invalid."
                Case ERROR_INVALID_NAME
                    StrRemServiceError = "The specified service name is invalid."
                Case ERROR_INVALID_PARAMETER
                    StrRemServiceError = "A parameter that was specified is invalid."
                Case ERROR_INVALID_SERVICE_ACCOUNT
                    StrRemServiceError = "The user account name specified in the lpServiceStartName parameter does not exist."
                Case ERROR_SERVICE_EXISTS
                    StrRemServiceError = "The specified service already exists in this database."
                Case Else
                    StrRemServiceError = "Unknown error"
            End Select
        End If
    End If
    apiCloseServiceHandle Hcsm
    apiCloseServiceHandle hServ
    Remove = CBool(ret)
    If Remove = True Then StrInstallRemoveError = "Service removed sucessfully"
End Function
Private Function StripTerminator(ByVal sInp As String) As String
    Dim zPos As Integer
    zPos = InStr(1, sInp, Chr(0))
    If zPos > 0 Then
        StripTerminator = Left(sInp, zPos - 1)
    Else
        StripTerminator = sInp
    End If
End Function
'Friend Function ServiceRunning(Optional ByVal Computer As String = vbNullString) As Boolean
'    Dim Status          As SERVICE_STATUS
'    Dim hMgr       As Long
'    Dim hServ        As Long
'    Dim hServRunning As Long
'    If Trim(strServiceName) = "" Then Exit Function
'    hMgr = apiOpenSCManager(Computer, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
'    If hMgr = 0 Then Exit Function
'    hServ = apiOpenService(hMgr, strServiceName, SERVICE_ALL_ACCESS)
'    If hServ <> 0 Then
'        hServRunning = apiQueryServiceStatus(hServ, Status)
'        If hServRunning <> 0 Then
'            Select Case Status.dwCurrentState
'                Case ServiceControllerStatus.Running
'                    ServiceRunning = True
'                Case ServiceControllerStatus.stopped
'                    ServiceRunning = False
'                Case ServiceControllerStatus.StopPending
'                    ServiceRunning = False
'                Case ServiceControllerStatus.Paused
'                    ServiceRunning = False
'                Case ServiceControllerStatus.StartPending
'                    ServiceRunning = False
'                Case ServiceControllerStatus.PausePending
'                    ServiceRunning = False
'                Case Else
'                    ServiceRunning = False
'            End Select
'        End If
'        apiCloseServiceHandle hServ
'    End If
'    apiCloseServiceHandle hMgr
'End Function

