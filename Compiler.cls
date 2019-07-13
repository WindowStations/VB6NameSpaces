VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Compiler"
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
'Attribute VB_Name = "Compiler"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiWinExec Lib "kernel32" Alias "WinExec" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function apiGetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function apiGetCurrentDirectory Lib "kernel32" Alias "GetCurrentDirectoryA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiSetCurrentDirectory Lib "kernel32" Alias "SetCurrentDirectoryA" (ByVal lpPathName As String) As Long
Private mvarPath As String
Private mvarArguments As String
Private mvarVB6EXE As String
Private mvarOutputDirectory As String
Private mvarCurDir As String
Private Sub Class_Initialize()
   Dim wd As String
    wd = GetWindowsDirectory
    VB6EXE = wd & ":\Program Files (x86)\Microsoft Visual Studio\VB98\VB6.exe"
    Dim pth As String
    pth = Left(path, 0)
    mvarCurDir = GetCurrentDirectory
    Call apiSetCurrentDirectory(Left(path, InStrRev(path, "\"))) '("C:\Program Files (x86)\Microsoft Visual Studio\VB98\me")
End Sub
Private Sub Class_Terminate()
    Call apiSetCurrentDirectory(mvarCurDir)
End Sub
Friend Property Let path(ByVal vData As String)
    mvarPath = vData
End Property
Friend Property Get path() As String
    path = mvarPath
End Property
Friend Property Let Arguments(ByVal vData As String)
    mvarArguments = vData
End Property
Friend Property Get Arguments() As String
    Arguments = mvarArguments
End Property
Friend Property Let VB6EXE(ByVal vData As String)
    mvarVB6EXE = vData
End Property
Friend Property Get VB6EXE() As String
    VB6EXE = mvarVB6EXE
End Property
Friend Property Let OutputDirectory(ByVal vData As String)
    mvarOutputDirectory = vData
End Property
Friend Property Get OutputDirectory() As String
    OutputDirectory = mvarOutputDirectory
End Property
Friend Sub Compile()
   On Error Resume Next
   
   Call apiWinExec(Chr(34) & VB6EXE & Chr(34) & " " & Trim(Arguments) & " " & Chr(34) & path & Chr(34) & " /outdir " & Chr(34) & Trim(OutputDirectory) & Chr(34), 1)
End Sub
Friend Sub Run()
  ' Call apiWinExec(pthFile & " /runexit " & Chr(34) & pthProject & Chr(34), 1)
   
End Sub
Private Function GetWindowsDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetWindowsDirectory(s, Len(s))
    GetWindowsDirectory = Left(s, InStr(1, s, vbNullString)) 'Left(s, InStr(1, s, vbNullString))
    End Function
Private Function GetCurrentDirectory() As String
    On Error Resume Next
    Dim s As String
    s = String(260, 0)
    Dim ret As Long
    ret = apiGetCurrentDirectory(Len(s), s)
    GetCurrentDirectory = Left(s, ret)
End Function

