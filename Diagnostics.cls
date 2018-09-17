VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagnostics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarProcess As Process
Private mvarStopWatch As StopWatch
Private Sub Class_Initialize()
    Set mvarProcess = New Process
    Set mvarStopWatch = New StopWatch
End Sub
Private Sub Class_Terminate()
    Set mvarProcess = Nothing
    Set mvarStopWatch = Nothing
End Sub
Public Property Get Process() As Process
    Set Process = mvarProcess
End Property
Public Property Get StopWatch() As StopWatch
    Set StopWatch = mvarStopWatch
End Property

