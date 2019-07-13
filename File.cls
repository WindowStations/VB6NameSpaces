VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "File"
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
'Attribute VB_Name = "File"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const GENERIC_WRITE                         As Long = &H40000000
Private Const OPEN_EXISTING                         As Long = 3
Private Const OPEN_ALWAYS                           As Long = 4
Private Const FILE_SHARE_READ                       As Long = &H1
Private Const FILE_SHARE_WRITE                      As Long = &H2
Private Const OFS_MAXPATHNAME                       As Long = 128
Private Const OF_CREATE                             As Long = &H1000
Private Const OF_READ                               As Long = &H0
Private Const OF_WRITE                              As Long = &H1
Private Const FO_MOVE                               As Long = &H1
Private Const FO_COPY                               As Long = &H2
Private Const FO_DELETE                             As Long = &H3
Private Const FO_RENAME                             As Long = &H4
Private Const FOF_CREATEPROGRESSDLG                 As Long = &H0
Private Const FOF_MULTIDESTFILES                    As Long = &H1
Private Const FOF_CONFIRMMOUSE                      As Long = &H2
Private Const FOF_SILENT                            As Long = &H4
Private Const FOF_RENAMEONCOLLISION                 As Long = &H8
Private Const FOF_NOCONFIRMATION                    As Long = &H10
Private Const FOF_WANTMAPPINGHANDLE                 As Long = &H20
Private Const FOF_ALLOWUNDO                         As Long = &H40
Private Const FOF_FILESONLY                         As Long = &H80
Private Const FOF_SIMPLEPROGRESS                    As Long = &H100
Private Const FOF_NOCONFIRMMKDIR                    As Long = &H200
Private Const MOVEFILE_REPLACE_EXISTING             As Long = &H1
Private Const MOVEFILE_COPY_ALLOWED                 As Long = &H2
Private Const MOVEFILE_DELAY_UNTIL_REBOOT           As Long = &H4
Private Const MOVEFILE_WRITE_THROUGH                As Long = &H8
Private Const MOVEFILE_CREATE_HARDLINK              As Long = 16
Private Const MOVEFILE_FAIL_IF_NOT_TRACKABLE        As Long = 32
Private Const COPY_FILE_FAIL_IF_EXISTS              As Long = &H1
Private Const COPY_FILE_RESTARTABLE                 As Long = &H2
Private Const COPY_FILE_OPEN_SOURCE_FOR_WRITE       As Long = &H4
Private Const COPY_FILE_ALLOW_DECRYPTED_DESTINATION As Long = &H8
Private Const COPY_FILE_COPY_SYMLINK                As Long = &H800
Private Const COPY_FILE_NO_BUFFERING                As Long = &H1000
Private Const PROGRESS_CANCEL                       As Long = 1
Private Const PROGRESS_CONTINUE                     As Long = 0
Private Const PROGRESS_QUIET                        As Long = 3
Private Const PROGRESS_STOP                         As Long = 2
Private Const SHERB_NOCONFIRMATION                  As Long = &H1
Private Const SHERB_NOPROGRESSUI                    As Long = &H2
Private Const SHERB_NOSOUND                         As Long = &H4
Private Const E_ERR_BASE  As Long = 18160 + vbObjectError
Private Const S_ERR_CannotGetPath    As String = "Cannot get path"
Private Const S_ERR_ComponentFailure As String = "CSpecialFolders component failure"
Private Const MAX_LENGTH             As Long = 512
Private Const INVALID_HANDLE_VALUE   As Long = -1
Private Const E_ERR_BASE2 = 17420 + vbObjectError
Private Const DRIVE_CDROM     As Long = 5
Private Const DRIVE_FIXED     As Long = 3
Private Const DRIVE_RAMDISK   As Long = 6
Private Const DRIVE_REMOTE    As Long = 4
Private Const DRIVE_REMOVABLE As Long = 2
Private Const FILE_ATTRIBUTE_ARCHIVE  As Long = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED  As Long = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY  As Long = &H10
Private Const FILE_ATTRIBUTE_HIDDEN  As Long = &H2
Private Const FILE_ATTRIBUTE_NORMAL  As Long = &H80
Private Const FILE_ATTRIBUTE_READONLY  As Long = &H1
Private Const FILE_ATTRIBUTE_SYSTEM  As Long = &H4
Private Const FILE_BEGIN   As Long = 0
Private Const FILE_CURRENT As Long = 1
Private Const FILE_END     As Long = 2
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
Private Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_LENGTH
    cAlternate As String * 14
End Type
Private Type BY_HANDLE_FILE_INFORMATION
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    dwVolumeSerialNumber As Long
    nFileSizeHigh As Long
    nFileSizeLow As Long
    nNumberOfLinks As Long
    nFileIndexHigh As Long
    nFileIndexLow As Long
End Type
Private Declare Function apiCopyFileEx Lib "kernel32" Alias "CopyFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal lpProgressRoutine As Long, ByRef lpData As Any, ByRef pbCancel As Long, ByVal dwCopyFlags As Long) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function apiCreateHardLink Lib "kernel32" Alias "CreateHardLinkA" (ByVal lpFileName As String, ByVal lpExistingFileName As String, ByRef lpSecurityAttributes As Any) As Long
Private Declare Function apiDeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function apiFindClose Lib "kernel32" Alias "FindClose" (ByVal hFindFile As Long) As Long
Private Declare Function apiFindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function apiFindNextFile Lib "kernel32" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function apiGetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function apiGetFileInformationByHandle Lib "kernel32" Alias "GetFileInformationByHandle" (ByVal hFile As Long, lpFileInformation As BY_HANDLE_FILE_INFORMATION) As Long
Private Declare Function apiGetFileSize Lib "kernel32" Alias "GetFileSize" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function apiGetFileTime Lib "kernel32" Alias "GetFileTime" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function apiGetFileTitle Lib "comdlg32" Alias "GetFileTitleA" (ByVal lpszFile As String, ByVal lpszTitle As String, ByVal cbBuf As Integer) As Integer
Private Declare Function apiGetFileType Lib "kernel32" Alias "GetFileType" (ByVal hFile As Long) As Long
Private Declare Function apiGetFileVersionInfo Lib "Version" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwHandle As Long, ByVal dwLen As Long, ByRef lpData As Any) As Long
Private Declare Function apiGetFileVersionInfoSize Lib "Version" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, ByRef lpdwHandle As Long) As Long
Private Declare Function apiGetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function apiGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function apiGetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function apiMoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function apiOpenFile Lib "kernel32" Alias "OpenFile" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Private Declare Function apiPathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private Declare Function apiSetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function apiSetFilePointer Lib "kernel32" Alias "SetFilePointer" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long
Private Declare Function apiSetFileTime Lib "kernel32" Alias "SetFileTime" (ByVal hFile As Long, ByRef lpCreationTime As FILETIME, ByRef lpLastAccessTime As FILETIME, ByRef lpLastWriteTime As FILETIME) As Long
Private Declare Function apiSHFileOperation Lib "shell32" Alias "SHFileOperationA" (ByRef lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Friend Function Exists(ByVal path As String) As Boolean
    On Error Resume Next
    Exists = CBool(apiPathFileExists(path))
End Function
Friend Function Create(ByVal path As String) As Long
    On Error Resume Next
    If apiPathFileExists(path) <> 0 Then Return
    Create = apiCreateFile(path, GENERIC_WRITE, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_ALWAYS, 0, 0) 'Open the file
    If Create = 0 Then Exit Function
    Call apiCloseHandle(Create)
End Function
Friend Function Copy(ByVal pathsource As String, ByVal destSource As String) As Long
    On Error GoTo poop
    Copy = apiCopyFileEx(pathsource, destSource, ByVal 0, ByVal 0, False, COPY_FILE_RESTARTABLE)
    Exit Function
poop:
    Call Copy_SH(pathsource, destSource)
End Function
Friend Function Delete(ByVal path As String) As Long
    On Error GoTo poop
    If apiPathFileExists(path) = 0 Then Return
    Delete = apiDeleteFile(path)
    Exit Function
poop:
    Call Delete_SH(path)
End Function
Friend Function Move(ByVal pathsource As String, ByVal destSource As String) As Long
    On Error GoTo poop
    If apiPathFileExists(pathsource) = 0 Then Return
    Move = apiMoveFileEx(pathsource, destSource, MOVEFILE_COPY_ALLOWED)
    Exit Function
poop:
    Call Move_SH(pathsource, destSource)
End Function
Friend Function Replace(ByVal pathsource As String, ByVal destSource As String, ByVal backSource As String) As Long
    On Error GoTo poop
    If apiPathFileExists(pathsource) = 0 Then Return
    Replace = apiMoveFileEx(pathsource, destSource, MOVEFILE_REPLACE_EXISTING)
    Exit Function
poop:
    Call Rename_SH(pathsource, destSource)
End Function
Friend Function SetSHFILESTRUCT(ByRef SHDirOp As SHFILEOPSTRUCT, ByVal fo As Long, ByVal fof As Long, ByVal pathsource As String, ByVal destSource As String)
    With SHDirOp
        .wFunc = fo
        .fFlags = CInt(fof)
        .pFrom = pathsource
        .pTo = destSource
        .hwnd = 0
    End With
End Function
Friend Function FileOperation(ByRef SHDirOp As SHFILEOPSTRUCT) As Long
    On Error Resume Next
    FileOperation = apiSHFileOperation(SHDirOp)
End Function
Friend Function Copy_SH(ByVal pathsource As String, ByVal pathdestination As String) As Long
    On Error GoTo poop
    Dim SHDirOp As SHFILEOPSTRUCT
    Dim r       As Long
    Call SetSHFILESTRUCT(SHDirOp, FO_COPY, FOF_SILENT, pathsource, pathdestination)
    r = FileOperation(SHDirOp)
    Copy_SH = r
poop:
End Function
Friend Function Delete_SH(ByVal pathsource As String) As Long
    On Error GoTo poop
    Dim SHDirOp As SHFILEOPSTRUCT
    Dim r       As Long
    Call SetSHFILESTRUCT(SHDirOp, FO_DELETE, FOF_SILENT, pathsource, vbNullString)
    r = FileOperation(SHDirOp)
    Delete_SH = r
poop:
End Function
Friend Function Move_SH(ByVal pathsource As String, ByVal pathdestination As String) As Long
    On Error GoTo poop
    Dim SHDirOp As SHFILEOPSTRUCT
    Dim r       As Long
    Call SetSHFILESTRUCT(SHDirOp, FO_MOVE, FOF_SILENT, pathsource, pathdestination)
    r = FileOperation(SHDirOp)
    Move_SH = r
poop:
End Function
Friend Function Rename_SH(ByVal pathsource As String, ByVal pathdestination As String) As Long
    On Error GoTo poop
    Dim SHDirOp As SHFILEOPSTRUCT
    Dim r       As Long
    Call SetSHFILESTRUCT(SHDirOp, FO_RENAME, FOF_SILENT, pathsource, pathdestination)
    r = FileOperation(SHDirOp)
    Rename_SH = r
poop:
End Function

