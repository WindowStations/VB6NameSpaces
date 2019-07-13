VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Environment"
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
'Attribute VB_Name = "Environment"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"OSVersion"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const ERR_PATHNOTFOUND        As Long = 76
Private Const MAX_COMPUTERNAME_LENGTH As Long = 31
Private Const UNLEN                   As Long = 256
Private Const UOI_FLAGS               As Long = 1
Private Const WSF_VISIBLE             As Long = 1
Private Const WM_SETTINGCHANGE        As Long = &H1A
Private Const HWND_BROADCAST          As Long = &HFFFF
Private Const SMTO_NORMAL             As Long = &H0
Private Const ERROR_ENVVAR_NOT_FOUND  As Long = 203
Private Const MAX_PATH                As Long = 260
Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type
Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function apiGetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, ByRef lSize As Long) As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiLookupAccountName Lib "advapi32" Alias "LookupAccountNameA" (ByVal lpSystemName As String, ByVal lpAccountName As String, ByRef Sid As Long, ByRef cbSid As Long, ByVal ReferencedDomainName As String, ByRef cbReferencedDomainName As Long, ByRef peUse As Integer) As Long
Private Declare Function apiGetSystemInfo Lib "kernel32" (ByRef lpSystemInfo As SYSTEM_INFO) As Long
Private Declare Function apiGetUserNameEx Lib "secur32" Alias "GetUserNameExA" (ByVal NameFormat As EXTENDED_NAME_FORMAT, ByVal lpNameBuffer As String, ByRef nSize As Long) As Long
Private Declare Function apiGetUserObjectInformation Lib "user32" Alias "GetUserObjectInformationA" (ByVal hObj As Long, ByVal nIndex As Long, ByRef pvInfo As Any, ByVal nLength As Long, ByRef lpnLengthNeeded As Long) As Long
Private Declare Function apiGetCommandLine Lib "kernel32" Alias "GetCommandLineA" () As String
Private Declare Function apiSHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwnd As Long, ByVal csidl As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal pszPath As String) As Long
Private Declare Function apiGetLogicalDrives Lib "kernel32" Alias "GetLogicalDrives" () As Long
Private Declare Function apiGetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiExpandEnvironmentStrings Lib "kernel32" Alias "ExpandEnvironmentStringsA" (ByVal lpSrc As String, ByVal lpDst As String, ByVal nSize As Long) As Long
Private Declare Function apiSetEnvironmentVariable Lib "kernel32" Alias "SetEnvironmentVariableA" (ByVal lpName As String, ByVal lpValue As String) As Long
Private Declare Function apiGetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private mvarExitCode  As Long
Private mvarOSVersion As OSVersion
Private Enum EXTENDED_NAME_FORMAT
    NameUnknown = 0
    NameFullyQualifiedDN = 1
    NameSamCompatible = 2
    NameDisplay = 3
    NameUniqueId = 6
    NameCanonical = 7
    NameUserPrincipal = 8
    NameCanonicalEx = 9
    NameServicePrincipal = 10
End Enum
Public Enum SpecialFolder
    ApplicationData_ = &H1A
    CommonApplicationData_ = &H23
    CommonProgramFiles_ = &H2B
    Cookies_ = &H21
    Desktop_ = &H0
    DesktopDirectory_ = &H10
    Favorites_ = &H6
    History_ = &H22
    InternetCache_ = &H20
    LocalApplicationData_ = &H1C
    MyComputer_ = &H11
    MyMusic_ = &HD
    MyPictures_ = &H27
    Personal_ = &H5
    ProgramFiles_ = &H26
    Programs_ = &H2
    Recent_ = &H8
    SendTo_ = &H9
    StartMenu_ = &HB
    Startup_ = &H7
    SystemDir_ = &H25
    Templates_ = &H15
    Windows_ = &H24
End Enum
Private Sub Class_Initialize()
    Set mvarOSVersion = New OSVersion
End Sub
Private Sub Class_Terminate()
    Set mvarOSVersion = Nothing
End Sub
Friend Property Get CommandLine() As String
    CommandLine = apiGetCommandLine
End Property
Friend Function CurrentDirectory() As String
End Function
Friend Function Exit_(ByVal Code As Long) As String
End Function
Friend Property Get ExitCode() As Long
    ExitCode = mvarExitCode
End Property
Friend Property Let ExitCode(ByVal vData As Long)
    mvarExitCode = vData
End Property
Friend Function GetCommandLineArgs() As String()
    Dim s() As String
    'todo
    GetCommandLineArgs = s
End Function
Friend Function GetFolderPath(ByVal path As SpecialFolder) As String
    Dim buff As String
    buff = String(MAX_PATH, 0)
    Call apiSHGetFolderPath(0, path, 0, 0, buff)
    GetFolderPath = buff
End Function
Friend Function GetLogicalDrives() As String()
    Dim ret    As Long
    Dim i      As Long
    Dim drv    As String
    Dim drvs() As String
    ret = apiGetLogicalDrives
    Do
        If (ret And 2 ^ i) <> 0 Then drv = drv + " " + Chr(65 + i)
        ReDim Preserve drvs(i)
        drvs(i) = drv
        If i = 25 Then Exit Do
        i = i + 1
    Loop
    GetLogicalDrives = drvs
End Function
Friend Function HasShutdownStarted() As Boolean
    HasShutdownStarted = False 'todo
End Function
Friend Function MachineName() As String
    Dim sBuffer As String
    Dim lSize   As Long
    sBuffer = Space(260)
    lSize = Len(sBuffer)
    Call apiGetComputerName(sBuffer, lSize)
    MachineName = Left(sBuffer, lSize)
End Function
Friend Property Get OSVersion() As OSVersion
    Set OSVersion = mvarOSVersion
End Property
Friend Property Get ProcessorCount() As Long
    Dim sinfo As SYSTEM_INFO
    Call apiGetSystemInfo(sinfo)
    ProcessorCount = sinfo.dwNumberOfProcessors
End Property
Friend Property Get SystemDirectory() As String
    Dim ret As String
    ret = String(260, 0)
    Dim lSize As Long
    lSize = apiGetSystemDirectory(ret, Len(ret))
    If lSize > 0 Then SystemDirectory = Left(ret, lSize)
End Property
Friend Function TickCount() As Long
    TickCount = apiGetTickCount
End Function
Friend Property Get UserDomainName() As String
    Const NameSamCompatible As Long = 2
    Dim lSize               As Long
    Dim dName               As String
    Dim Sid                 As String
    lSize = 1024
    dName = String(lSize, 0)
    If apiGetUserNameEx(NameSamCompatible, dName, lSize) <> 0 Then
        Dim i As Long
        i = InStr(dName, "\")
        If i > 0 Then UserDomainName = Left(dName, i - 1): Exit Property
    End If
    lSize = 1024
    Sid = String(lSize, 0)
    If apiLookupAccountName(vbNullString, UserName, StrPtr(Sid), lSize, dName, lSize, 0) = 0 Then Exit Property
    UserDomainName = Left(dName, lSize)
End Property
Friend Function UserInteractive() As Boolean
    'todo
End Function
Friend Function UserName() As String
    Dim buff  As String
    Dim lSize As Long
    buff = Space(260)
    lSize = Len(buff)
    Call apiGetUserName(buff, lSize)
    UserName = Left(buff, lSize)
End Function
Friend Function WorkingSet() As Long
    'todo
End Function
Public Function GetEnvironmentVariable(Name As String) As String
    On Error GoTo hComponentFailure
    Dim lRet As Long
    Dim sRet As String
    lRet = 0
    sRet = ""
    lRet = apiExpandEnvironmentStrings(Name, sRet, lRet)
    If lRet = 0 Then
        On Error GoTo 0
        ' Err.Raise eErrEnvironment_CannotGetEnvironmentVariable, "VB65" & ".CEnvironment", S_ERR_CannotGetEnvironmentVariable
    End If
    sRet = String(lRet - 1, 0)
    lRet = apiExpandEnvironmentStrings(Name, sRet, lRet)
    If lRet = 0 Then
        On Error GoTo 0
        ' Err.Raise eErrEnvironment_CannotGetEnvironmentVariable, "VB65" & ".CEnvironment", S_ERR_CannotGetEnvironmentVariable
    End If
    If Right(sRet, 1) = vbNullChar Then sRet = Left(sRet, Len(sRet) - 1)
    GetEnvironmentVariable = sRet
    Exit Function
hComponentFailure:
    ' Err.Raise eErrEnvironment_ComponentFailure, "VB65" & ".CEnvironment", S_ERR_ComponentFailure
End Function
Public Sub SetEnvironmentVariable(Name As String, Value As String)
    On Error GoTo hComponentFailure
    If apiSetEnvironmentVariable(Name, Value) = 0 Then
        On Error GoTo 0
        'Err.Raise eErrEnvironment_CannotSetEnvironmentVariable, "VB65" & ".CEnvironment", S_ERR_CannotSetEnvironmentVariable
    End If
    Exit Sub
hComponentFailure:
    ' Err.Raise eErrEnvironment_ComponentFailure, "VB65" & ".CEnvironment", S_ERR_ComponentFailure
End Sub

