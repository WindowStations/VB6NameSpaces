VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Directory"
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
'Attribute VB_Name = "Directory"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const FO_MOVE            As Long = &H1
Private Const FOF_SIMPLEPROGRESS As Long = &H100
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type SHQUERYRBINFO
    cbSize As Long
    lsizelow As Long
    lsizehigh As Long
    lcountlow As Long
    lcounthigh As Long
End Type
Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type
Private Type SHITEMID
    cb As Long
    abID As Byte
End Type
Private Type ITEMIDLIST
    mkid As SHITEMID
End Type
'Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'Private Declare Function SetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function apiCreateDirectory Lib "kernel32" Alias "CreateDirectoryA" (ByVal lpPathName As String, ByRef lpSecurityAttributes As Long) As Long
Private Declare Function apiGetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiGetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetSystemWindowsDirectory Lib "kernel32" Alias "GetSystemWindowsDirectoryA" (ByVal lpBuffer As String, ByVal uSize As Long) As Long
Private Declare Function apiSetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function apiRemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function apiSHCreateDirectoryEx Lib "shell32" Alias "SHCreateDirectoryEx" (ByVal hwnd As Long, ByVal pszPath As String, ByRef psa As SECURITY_ATTRIBUTES) As Long
Private Declare Function apiPathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsDirectoryEmpty Lib "shlwapi" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function apiMakeSureDirectoryPathExists Lib "imagehlp" Alias "MakeSureDirectoryPathExists" (ByVal DirPath As String) As Long
Private Declare Function apiGetUserProfileDirectory Lib "userenv" Alias "GetUserProfileDirectory" (ByVal hToken As Long, ByVal lpProfileDir As String, ByRef lpcchSize As Long) As Long
Private Declare Function apiSHQueryRecycleBin Lib "shell32" Alias "SHQueryRecycleBin" (ByVal sRootPath As String, ByRef lpRBInfo As SHQUERYRBINFO) As Long
Private Declare Function apiSHEmptyRecycleBin Lib "shell32" Alias "SHEmptyRecycleBin" (ByVal hwnd As Long, ByVal sRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function apiSHGetPathFromIDList Lib "shell32" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function apiSHGetSpecialFolderLocation Lib "shell32" Alias "SHGetSpecialFolderLocation" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByRef pidl As ITEMIDLIST) As Long
Private Declare Function apiSHFileOperation Lib "shell32" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Const E_ERR_BASE = 18160 + vbObjectError
Private Const S_ERR_CannotGetPath    As String = "Cannot get path"
Private Const S_ERR_ComponentFailure As String = "CSpecialFolders component failure"
Private Enum ESpecialFolder
    eSpecialFolder_Programs = &H2
    eSpecialFolder_Personal = &H5
    eSpecialFolder_Favorites = &H6
    eSpecialFolder_Startup = &H7
    eSpecialFolder_Recent = &H8
    eSpecialFolder_SendTo = &H9
    eSpecialFolder_StartMenu = &HB
    eSpecialFolder_DesktopDir = &H10
    eSpecialFolder_Templates = &H15
    eSpecialFolder_CommonStartMenu = &H16
    eSpecialFolder_CommonPrograms = &H17
    eSpecialFolder_CommonStartup = &H18
    eSpecialFolder_CommonDesktopDir = &H19
    eSpecialFolder_AppData = &H1A
    eSpecialFolder_InternetCache = &H20
    eSpecialFolder_Cookies = &H21
    eSpecialFolder_History = &H22
End Enum
Private Enum EErrSpecialFolders
    eErrSpecialFolders_CannotGetPath = E_ERR_BASE + 1
    eErrSpecialFolders_ComponentFailure
End Enum
Public Function Create(ByVal path As String) As Long
    On Error Resume Next
    If apiPathIsDirectory(path) <> 0 Then
    End If
    Create = apiCreateDirectory(path, ByVal &H0)
End Function
Public Function Delete(ByVal path As String) As Long
    On Error Resume Next
    If apiPathIsDirectory(path) = 0 Then Return
    Delete = apiRemoveDirectory(path)
End Function
Public Function Move(ByVal pathsource As String, ByVal destSource As String) As Long
    On Error Resume Next
    If apiPathIsDirectory(pathsource) = 0 Then Return
    Dim SHDirOp As SHFILEOPSTRUCT
    With SHDirOp
        .hwnd = 0
        .wFunc = FO_MOVE
        .pFrom = pathsource
        .fFlags = FOF_SIMPLEPROGRESS
    End With
    Move = apiSHFileOperation(SHDirOp) 'Prompts the user to confirm deletion of the directory
End Function
Public Function Exists(ByVal path As String) As Boolean
    On Error Resume Next
    Exists = CBool(apiPathIsDirectory(path))
End Function
'Public Function QueryRecycleBin(ByRef q As SHQUERYRBINFO) As Long
'    On Error Resume Next
'    Dim ret As Long
'    ret = apiSHQueryRecycleBin(vbNullString, q)
'    QueryRecycleBin = ret
'End Function
'Public Function EmptyRecycleBin() As Boolean
'    On Error Resume Next
'    Dim ret As Long
'    ret = apiSHEmptyRecycleBin(0, vbNullString, SHERB_NOPROGRESSUI)  'all recycle bins with 2nd null param
'    EmptyRecycleBin = ret
'End Function
Private Function GetWindowsDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetWindowsDirectory(s, Len(s))
    GetWindowsDirectory = Left(s, ret) ' Left(s, InStr(1, s, vbNullString))
    End Function
Private Function GetCurrentDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetCurrentDirectory(Len(s), s)
    GetCurrentDirectory = Left(s, ret)
End Function
Friend Function GetSystemDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetSystemDirectory(s, Len(s))
    If ret <> 0 Then GetSystemDirectory = Left(s, ret)
End Function
Private Function GetSystemWindowsDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetSystemWindowsDirectory(s, 260)
    If ret <> 0 Then GetSystemWindowsDirectory = Left(s, ret)
End Function
Private Function GetPath(ByVal SpecialFolder As ESpecialFolder) As String
    On Error GoTo hComponentFailure
    Dim IDL   As ITEMIDLIST
    Dim lRet  As Long
    Dim sPath As String
    Dim sTemp As String
    lRet = apiSHGetSpecialFolderLocation(0, SpecialFolder, IDL) ' Get the location into our ITEMIDLIST
    If lRet = 0 Then
        sPath = Space(260) ' Build a buffer string to handle the returned value
        lRet = apiSHGetPathFromIDList(ByVal IDL.mkid.cb, ByVal sPath)
        If lRet <> 0 Then
            sTemp = Left(sPath, apilstrlen(sPath)) ' Make sure we append a '\' to the path
            sTemp = sTemp & IIf(Right(sTemp, 1) = "\", "", "\")
            GetPath = sTemp
        End If
    Else
        On Error GoTo 0
        Err.Raise eErrSpecialFolders_CannotGetPath, "VB65" & ".CSpecialFolders", S_ERR_CannotGetPath
    End If
    Exit Function
hComponentFailure:
    Err.Raise eErrSpecialFolders_ComponentFailure, "VB65" & ".CSpecialFolders", S_ERR_ComponentFailure
End Function

