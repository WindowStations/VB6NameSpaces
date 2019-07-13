VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Path"
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
'Attribute VB_Name = "Path"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiGetFullPathName Lib "kernel32" Alias "GetFullPathNameA" (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, ByVal lpFilePart As String) As Long
Private Declare Function apiGetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function apiGetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiGetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Private Declare Function apiGetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function apiPathAddExtension Lib "shlwapi" Alias "PathAddExtensionA" (ByVal pszPath As String, ByVal pszExt As String) As Long
Private Declare Function apiPathAppend Lib "shlwapi" Alias "PathAppendA" (ByVal pszPath As String, ByVal pMore As String) As Long
Private Declare Function apiPathCombine Lib "shlwapi" Alias "PathCombineA" (ByVal szDest As String, ByVal lpszDir As String, ByVal lpszFile As String) As Long
Private Declare Function apiPathFileExists Lib "shlwapi" Alias "PathFileExistsA" (ByVal pszPath As String) As Long
Private Declare Function apiPathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pszPath As String) As Long
Private Declare Function apiPathFindFileName Lib "shlwapi" Alias "PathFindFileNameA" (ByVal pszPath As String) As Long
Private Declare Function apiPathFindOnPath Lib "shlwapi" Alias "PathFindOnPathA" (ByVal pszPath As String, ByVal ppszOtherDirs As String) As Long
Private Declare Function apiPathGetDriveNumber Lib "shlwapi" Alias "PathGetDriveNumberA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsContentType Lib "shlwapi" Alias "PathIsContentTypeA" (ByVal pszPath As String, ByVal pszContentType As String) As Long
Private Declare Function apiPathIsDirectory Lib "shlwapi" Alias "PathIsDirectoryA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsDirectoryEmpty Lib "shlwapi" Alias "PathIsDirectoryEmptyA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsFileSpec Lib "shlwapi" Alias "PathIsFileSpecA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsLFNFileSpec Lib "shlwapi" Alias "PathIsLFNFileSpecA" (ByVal lpName As String) As Long
Private Declare Function apiPathIsNetworkPath Lib "shlwapi" Alias "PathIsNetworkPathA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsPrefix Lib "shlwapi" Alias "PathIsPrefixA" (ByVal pszPrefix As String, ByVal pszPath As String) As Long
Private Declare Function apiPathIsRelative Lib "shlwapi" Alias "PathIsRelativeA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsRoot Lib "shlwapi" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Private Declare Function apiPathIsSameRoot Lib "shlwapi" Alias "PathIsSameRootA" (ByVal pszPath1 As String, ByVal pszPath2 As String) As Long
Private Declare Function apiPathIsSystemFolder Lib "shlwapi" Alias "PathIsSystemFolderA" (ByVal pszPath As String, ByVal dwAttrb As Long) As Long
Private Declare Function apiPathIsURL Lib "shlwapi" Alias "PathIsURLA" (ByVal pszPath As String) As Long
Private Declare Function apiPathMakePretty Lib "shlwapi" Alias "PathMakePrettyA" (ByVal pszPath As String) As Long
Private Declare Function apiPathMakeSystemFolder Lib "shlwapi" Alias "PathMakeSystemFolderA" (ByVal pszPath As String) As Long
Private Declare Function apiPathRenameExtension Lib "shlwapi" Alias "PathRenameExtensionA" (ByVal pszPath As String, ByVal pszExt As String) As Long
Private Declare Function apiPathRelativePathTo Lib "shlwapi" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Declare Function apiPathRemoveBackslash Lib "shlwapi" Alias "PathRemoveBackslashA" (ByVal pszPath As String) As Long
Private Declare Function apiPathRemoveFileSpec Lib "shlwapi" Alias "PathRemoveFileSpecA" (ByVal pszPath As String) As Long
Private Declare Function apiPathSearchAndQualify Lib "shlwapi" Alias "PathSearchAndQualifyA" (ByVal pszPath As String, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
Private Declare Function apiPathStripToRoot Lib "shlwapi" Alias "PathStripToRootA" (ByRef pszPath As String) As Long
Private Declare Function apiPathStripPath Lib "shlwapi" Alias "PathStripPathA" (ByVal strPath As String) As Long
Private Declare Function apiPathRemoveExtension Lib "shlwapi" Alias "PathRemoveExtensionA" (ByVal pszPath As String) As Long
Friend Function GetTempPath() As String
    Dim path As String
    Dim ret  As Long
    path = Space(260)
    ret = apiGetTempPath(260, path)
    path = Left(path, ret)
    GetTempPath = path
End Function
Friend Function Combine(ByVal ldest As String, ByVal ldir As String, ByVal lfile As String) As String
    Call apiPathCombine(ldest, ldir, lfile)
End Function
Friend Function GetDirectoryName(ByVal path As String) As String

End Function
Friend Function GetExtension(ByVal path As String) As String
   Dim s As String * 260
   s = path
   apiPathFindExtension (path)
   GetExtension = s
End Function
Friend Function GetFileName(ByVal path As String) As String
   Dim s As String * 260
   s = path
   Call apiPathStripPath(s)
   GetFileName = s
End Function
Friend Function GetFileNameWithoutExtension(ByVal path As String) As String
   GetFileNameWithoutExtension = apiPathRemoveExtension(GetFileName(path))
End Function
Friend Function GetFullPath(ByVal path As String) As String
Dim buff As String * 260
Call apiGetFullPathName(path, 260, buff, "")
End Function
Friend Function GetPathRoot(ByVal path As String) As String
    Call apiPathStripToRoot(path)
    GetPathRoot = path
End Function
Friend Function GetRandomFileName() As String
End Function
Friend Function GetTempFileName() As String
    Dim fpth As String
    Dim ret  As Long
    fpth = Space(260)
     ret = apiGetTempFileName(GetTempPath, "api", 0, fpth)
    fpth = Left(fpth, InStr(fpth, vbNullChar) - 1)  ' extract data from the variable
    fpth = Left(fpth, ret)  ' extract data from the variable?
    GetTempFileName = fpth
End Function
Friend Function HasExtension(ByVal path As String) As Boolean
    On Error Resume Next
    Dim ret As Long
    Dim tmp As String * 260
    tmp = path
    '
    '
    ret = apiPathAddExtension(path, "removethisextension")
    If ret = 0 Then
        HasExtension = True
    Else
        HasExtension = Not CBool(apiPathRemoveExtension(path))
    End If
End Function
Friend Function IsPathRooted(ByVal path As String) As Boolean
    IsPathRooted = CBool(apiPathIsRoot(path))
End Function
'Private Sub ShortName()
'    Dim ShortName As String
'    Dim slength As Long
'    ShortName = Space(260)
'    slength = apiGetShortPathName("C:\My Documents\ReadMeFirst.txt", ShortName, 260)
'    ShortName = Left(ShortName, slength)
'    MsgBox "Equivalent: " & ShortName
'End Sub
'
'
'
'
'' Generate a temporary file (path)\api????.TMP, where (path)
'' is Windows's temporary file directory and ???? is a randomly assigned unique value.
'' Then display the name of the created file on the screen.
'Dim temppath As String  ' receives name of temporary file path
'Dim tempfile As String  ' receives name of temporary file
'Dim slength As Long  ' receives length of string returned for the path
'Dim lastfour As Long  ' receives hex value of the randomly assigned ????
'
'' Get Windows's temporary file path
'temppath = Space(255)  ' initialize the buffer to receive the path
'slength = GetTempPath(255, temppath)  ' read the path name
'temppath = Left(temppath, slength)  ' extract data from the variable
'
'' Get a uniquely assigned random file
'tempfile = Space(255)  ' initialize buffer to receive the filename
'lastfour = GetTempFileName(temppath, "api", 0, tempfile)  ' get a unique temporary file name
'' (Note that the file is also created for you in this case.)
'tempfile = Left(tempfile, InStr(tempfile, vbNullChar) - 1)  ' extract data from the variable
'Debug.Print "Temporary filename: "; tempfile


