VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Environment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
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
Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function apiGetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function apiGetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, ByRef lSize As Long) As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'Public Function CreateFileMapping(ByVal hFile As Long, ByVal lpSecurityAttributes As Long, ByVal flProtect As Long, ByVal MaximumSizeHigh As Long, ByVal MaximumSizeLow As Long, Name As String) As Long: End Function
'Public Function ExpandEnvironmentStrings(lpSrc As String, lpDst As String, ByVal nSize As Long) As Long: End Function
'Public Function GetComputerName(Buffer As String, Size As Long) As Long: End Function
'Public Function GetCurrentDirectory(ByVal BufferLength As Long, Buffer As String) As Long: End Function
'Public Function GetEnvironmentStrings() As String: End Function
'Public Function GetEnvironmentVariable(Name As String, Buffer As String, ByVal Size As Long) As Long: End Function
'Public Function GetFileAttributes(FileName As String) As Long: End Function
'Public Function GetFileAttributesEx(FileName As String, ByVal fInfoLevelId As Long, lpFileInformation As CorType.WIN32_FILE_ATTRIBUTE_DATA) As Long: End Function
'Public Function GetFullPathName(FileName As String, ByVal BufferLength As Long, Buffer As String, ByVal lpFilePart As Long) As Long: End Function
'Public Function GetLocaleInfo(ByVal Locale As Long, ByVal LCType As Long, lpLCData As String, ByVal cchData As Long) As Long: End Function
'Public Function GetLongPathName(ShortPath As String, LongPath As String, ByVal LongPathBufferSize As Long) As Long: End Function
'Public Function GetPrivateProfileSection(AppName As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long: End Function
'Public Function GetPrivateProfileString(AppName As String, KeyName As String, Default As String, ReturnedString As String, ByVal Size As Long, FileName As String) As Long: End Function
'Public Function GetSystemDirectory(Buffer As String, ByVal Size As Long) As Long: End Function
'Public Function GetTempFileName(PathName As String, PrefixString As String, ByVal Unique As Long, TempFileName As String) As Long: End Function
'Public Function GetTempPath(ByVal BufferLength As Long, Buffer As String) As Long: End Function
'Public Function MoveFile(ExistingFileName As String, NewFileName As String) As Long: End Function
'Public Function RemoveDirectory(lpPathName As String) As Long: End Function
'Public Function SetCurrentDirectory(PathName As String) As Long: End Function
'Public Function SetEnvironmentVariable(Name As String, Value As String) As Long: End Function
'Public Function SetFileAttributes(FileName As String, ByVal dwFileAttributes As Long) As Long: End Function
'Public Function WritePrivateProfileSection(AppName As String, Value As String, FileName As String) As Long: End Function
'Public Function WritePrivateProfileString(AppName As String, KeyName As String, Value As String, FileName As String) As Long: End Function
Private mvarExitCode  As Long 'local copy
Private mvarOSVersion As OSVersion
Public Enum SpecialFolder
    ApplicationData = &H1A
    CommonApplicationData = &H23
    CommonProgramFiles = &H2B
    Cookies = &H21
    Desktop = &H0
    DesktopDirectory = &H10
    Favorites = &H6
    History = &H22
    InternetCache = &H20
    LocalApplicationData = &H1C
    MyComputer = &H11
    MyMusic = &HD
    MyPictures = &H27
    Personal = &H5
    ProgramFiles = &H26
    Programs = &H2
    Recent = &H8
    SendTo = &H9
    StartMenu = &HB
    Startup = &H7
    SystemDir = &H25
    Templates = &H15
    Windows = &H24
End Enum
Private Sub Class_Initialize()
    Set mvarOSVersion = New OSVersion
End Sub
Private Sub Class_Terminate()
    Set mvarOSVersion = Nothing
End Sub
Public Property Get OSVersion() As OSVersion
    Set OSVersion = mvarOSVersion
End Property
Friend Property Let ExitCode(ByVal vData As Long)
    mvarExitCode = vData
End Property
Friend Property Get ExitCode() As Long
    ExitCode = mvarExitCode
End Property
Friend Function MachineName() As String
    Dim sBuffer As String
    Dim lSize   As Long
    sBuffer = Space(260)
    lSize = Len(sBuffer)
    Call apiGetComputerName(sBuffer, lSize)
    MachineName = Left(sBuffer, lSize)
End Function
Friend Function UserName() As String
    Dim sBuffer As String
    Dim lSize   As Long
    sBuffer = Space(260)
    lSize = Len(sBuffer)
    Call apiGetUserName(sBuffer, lSize)
    UserName = Left(sBuffer, lSize)
End Function
Public Property Get SystemDirectory() As String
    Dim ret As String
    ret = String(260, 0)
    Dim Size As Long
    Size = apiGetSystemDirectory(ret, Len(ret))
    If Size > 0 Then SystemDirectory = Left(ret, Size)
End Property
Friend Function TickCount() As Long
    TickCount = apiGetTickCount
End Function
Friend Function CurrentDirectory() As String
End Function



