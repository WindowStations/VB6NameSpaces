VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Thread"
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
'Attribute VB_Name = "Thread"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const CALLBACK_THREAD               As Long = &H20000
Private Const SYNCHRONIZE                   As Long = &H100000
Private Const WAIT_FAILED                   As Long = &HFFFFFFFF
Private Const WAIT_OBJECT_0                 As Long = 0
Private Const WAIT_TIMEOUT                  As Long = &H102
Private Const FORMAT_MESSAGE_FROM_SYSTEM    As Long = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function apiSleepEx Lib "kernel32" Alias "SleepEx" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiCreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function apiReleaseMutex Lib "kernel32" Alias "ReleaseMutex" (ByVal hMutex As Long) As Long
Private Declare Function apiCreateThread Lib "kernel32" Alias "CreateThread" (ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByRef lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadID As Long) As Long
Private Declare Function apiCreateRemoteThread Lib "kernel32" Alias "CreateRemoteThread" (ByVal hProcess As Long, ByRef lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByRef lpStartAddress As Long, ByRef lpParameter As Any, ByVal dwCreationFlags As Long, ByRef lpThreadID As Long) As Long
Private Declare Function apiResumeThread Lib "kernel32" Alias "ResumeThread" (ByVal hThread As Long) As Long
Private Declare Function apiSuspendThread Lib "kernel32" Alias "SuspendThread" (ByVal hThread As Long) As Long
Private Declare Function apiTerminateThread Lib "kernel32" Alias "TerminateThread" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Private Declare Function apiOpenThread Lib "kernel32" Alias "OpenThread " (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwThreadId As Long) As Long
'Private Declare Function ConvertThreadToFiber Lib "kernel32" (ByRef lpParameter As Any) As Long
'Private Declare Sub apiCoCreateFreeThreadedMarshaler Lib "ole32" (ByVal punkOuter As Long, ByVal ppunkMarshal As Long)
'Private Declare Sub apiCoMarshalInterThreadInterfaceInStream Lib "ole32" (ByVal riid As Long, ByVal pUnk As Long, ByRef ppStm As Long)
'Private Declare Function apiEnumThreadWindows Lib "user32" (ByVal dwThreadId As Long, ByVal lpfn As Long, ByVal lParam As Long) As Long
'Private Declare Function apiDisableThreadLibraryCalls Lib "kernel32" (ByVal hLibModule As Long) As Long
'Private Declare Sub apiExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function apiGetCurrentThread Lib "kernel32" Alias "GetCurrentThread" () As Long
Private Declare Function apiGetCurrentProcess Lib "kernel32" Alias "GetCurrentProcess" () As Long
'Private Declare Function apiGetCurrentThreadId Lib "kernel32" () As Long
'Private Declare Function apiGetExitCodeThread Lib "kernel32" (ByVal hThread As Long, ByRef lpExitCode As Long) As Long
'Private Declare Function apiPostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'Private Declare Function apiSetThreadLocale Lib "kernel32" (ByVal Locale As Long) As Long
'Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Private Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, ByRef lpContext As CONTEXT) As Long
'Private Declare Function GetThreadTimes Lib "kernel32" (ByVal hThread As Long, ByRef lpCreationTime As FILETIME, ByRef lpExitTime As FILETIME, ByRef lpKernelTime As FILETIME, ByRef lpUserTime As FILETIME) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private hMutex As Long
Friend Function CreateMutex()
    hMutex = apiCreateMutex(0, True, "TestMutex")
    Dim i As Long
    Do
        i = i + 1
        apiSleep (1)
        DoEvents
        If i = 5000 Then Exit Do
    Loop
    apiReleaseMutex (hMutex)
    MsgBox "yup"
End Function
Friend Function ReleaseMutex()
End Function
Friend Function Sleep(Optional ByVal dwMilliseconds As Long = 0) As Boolean
    On Error Resume Next
    Sleep = Not CBool(apiSleep(dwMilliseconds))
End Function
Friend Function SleepEx(Optional ByVal dwMilliseconds As Long = 0, Optional ByVal bAlertable As Long) As Long
    SleepEx = apiSleepEx(dwMilliseconds, bAlertable)
End Function
Friend Sub CreateThread()
    Dim sa As SECURITY_ATTRIBUTES
    hThread = apiCreateRemoteThread(apiGetCurrentProcess, sa, ByVal 0, AddressOf AsyncThread, ByVal 0, ByVal 0, hThreadID)
    apiCloseHandle hThread
End Sub
Friend Sub TerminateThread()
    If hThread <> 0 Then apiTerminateThread hThread, 0
End Sub

