VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Thread"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function apiSleepEx Lib "kernel32" Alias "SleepEx" (ByVal dwMilliseconds As Long, ByVal bAlertable As Long) As Long
Private Declare Function apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long) As Long
Friend Function Sleep(Optional ByVal dwMilliseconds As Long = 0) As Boolean
    On Error Resume Next
    Sleep = Not CBool(apiSleep(dwMilliseconds))
End Function
Friend Function SleepEx(Optional ByVal dwMilliseconds As Long = 0, Optional ByVal bAlertable As Long) As Long
    SleepEx = apiSleepEx(dwMilliseconds, bAlertable)
End Function

