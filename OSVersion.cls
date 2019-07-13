VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "OSVersion"
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
'Attribute VB_Name = "OSVersion"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiGetVersionEx Lib "kernel32" Alias "GetVersionExA" (ByRef lpVersionInformation As Version) As Long
'Friend Function VersionString() As OSVERSIONINFO
'    Dim osvi As OSVERSIONINFO
'    osvi.dwOSVersionInfoSize = Len(osvi)
'    Dim ret As Long
'    Dim maj As Long
'    Dim plat As Long
'    Dim min As Long
'    Dim ver As String
'    Dim build As Long
'    Dim title As String
'    ret = apiGetVersionEx(osvi)
'    If ret = 0 Then Exit Function
'    maj = osvi.dwMajorVersion
'    plat = osvi.dwPlatformId
'    min = osvi.dwMinorVersion
'    ver = osvi.szCSDVersion
'    build = osvi.dwBuildNumber
'    Dim nt As String: nt = ""
'    If plat = 2 Then nt = "NT"
'    If plat = 1 And min = 0 Then title = "95": GoTo titled
'    If plat = 1 And min = 10 Then title = "98": GoTo titled
'    If plat = 1 And min = 90 Then title = "ME": GoTo titled
'    If maj = 5 And min = 0 Then title = "2000": GoTo titled
'    If maj = 5 And min = 1 Then title = "XP": GoTo titled
'    If min = 6 And min = 0 Then title = "Vista": GoTo titled
'    If maj = 6 And min = 1 Then title = "7": GoTo titled
'titled:
'    Dim v As String
'    v = "Microsoft Windows" & " " & nt & " " & maj & "." & min & "." & build & " " & ver
'    osvi.szCSDVersion = Trim(v)
'    GetOSVersionInfo = osvi
'End Function
Friend Function Version() As Version
    Dim osvi As Version
    osvi.dwOSVersionInfoSize = Len(osvi)
    Dim ret   As Long
    Dim maj   As Long
    Dim plat  As Long
    Dim min   As Long
    Dim ver   As String
    Dim build As Long
    Dim title As String
    ret = apiGetVersionEx(osvi)
    If ret = 0 Then Exit Function
    maj = osvi.dwMajorVersion
    plat = osvi.dwPlatformId
    min = osvi.dwMinorVersion
    ver = osvi.szCSDVersion
    build = osvi.dwBuildNumber
    Dim nt As String: nt = ""
    If plat = 2 Then nt = "NT"
    If plat = 1 And min = 0 Then title = "95": GoTo titled
    If plat = 1 And min = 10 Then title = "98": GoTo titled
    If plat = 1 And min = 90 Then title = "ME": GoTo titled
    If maj = 5 And min = 0 Then title = "2000": GoTo titled
    If maj = 5 And min = 1 Then title = "XP": GoTo titled
    If min = 6 And min = 0 Then title = "Vista": GoTo titled
    If maj = 6 And min = 1 Then title = "7": GoTo titled
titled:
    Dim v As String
    v = "Microsoft Windows" & " " & nt & " " & maj & "." & min & "." & build & " " & ver
    osvi.szCSDVersion = Trim(v)
    Version = osvi
End Function
Friend Function VersionString() As String
    Dim osvi As Version
    osvi.dwOSVersionInfoSize = Len(osvi)
    Dim ret   As Long
    Dim maj   As Long
    Dim plat  As Long
    Dim min   As Long
    Dim ver   As String
    Dim build As Long
    Dim title As String
    ret = apiGetVersionEx(osvi)
    If ret = 0 Then Exit Function
    maj = osvi.dwMajorVersion
    plat = osvi.dwPlatformId
    min = osvi.dwMinorVersion
    ver = osvi.szCSDVersion
    build = osvi.dwBuildNumber
    Dim nt As String: nt = ""
    If plat = 2 Then nt = "NT"
    If plat = 1 And min = 0 Then title = "95": GoTo titled
    If plat = 1 And min = 10 Then title = "98": GoTo titled
    If plat = 1 And min = 90 Then title = "ME": GoTo titled
    If maj = 5 And min = 0 Then title = "2000": GoTo titled
    If maj = 5 And min = 1 Then title = "XP": GoTo titled
    If min = 6 And min = 0 Then title = "Vista": GoTo titled
    If maj = 6 And min = 1 Then title = "7": GoTo titled
titled:
    Dim v As String
    v = "Microsoft Windows" & " " & nt & " " & maj & "." & min & "." & build & " " & ver
    VersionString = Trim(v)
End Function

