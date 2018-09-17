VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StopWatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function apiQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef x As Currency) As Boolean
Private Declare Function apiQueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef x As Currency) As Boolean
Private Tic()           As Currency
Private Toc()           As Currency
Private Freq            As Currency
Private overhead        As Currency
Private bTimerRunning() As Boolean
Private StopWatchCount  As Long
Friend Function StartNew() As Long
    StopWatchCount = StopWatchCount + 1
    ReDim Preserve Tic(StopWatchCount)
    ReDim Preserve Toc(StopWatchCount)
    ReDim Preserve bTimerRunning(StopWatchCount)
    StartNew = StopWatchCount
End Function
Friend Function Start(stopwatchhandle As Long) As Long
    If ((stopwatchhandle > 0) And (stopwatchhandle <= StopWatchCount)) Then
        Start = 1
        bTimerRunning(stopwatchhandle) = True
        apiQueryPerformanceFrequency Freq
        apiQueryPerformanceCounter Tic(stopwatchhandle)
        apiQueryPerformanceCounter Toc(stopwatchhandle)
        overhead = Toc(stopwatchhandle) - Tic(stopwatchhandle)
        apiQueryPerformanceCounter Tic(stopwatchhandle)
    Else
        Start = 0
    End If
End Function
Friend Function ElapsedTicks(stopwatchhandle As Long) As Double
    If ((stopwatchhandle > 0) And (stopwatchhandle <= StopWatchCount)) Then
        If bTimerRunning(stopwatchhandle) = True Then
            apiQueryPerformanceCounter Toc(stopwatchhandle)
            ElapsedTicks = (Toc(stopwatchhandle) - Tic(stopwatchhandle) - overhead) / Freq
        Else
            ElapsedTicks = 0
        End If
    Else
        ElapsedTicks = -1
    End If
End Function
Friend Function GetStopWatchStatus(stopwatchhandle) As Long
    If ((stopwatchhandle > 0) And (stopwatchhandle <= StopWatchCount)) Then
        If bTimerRunning(stopwatchhandle) = True Then
            GetStopWatchStatus = 1
        Else
            GetStopWatchStatus = 0
        End If
    Else
        GetStopWatchStatus = -1
    End If
End Function
Friend Function Reset()
    'stop timer reset to zero
    'todo
End Function
Friend Function ElapsedMilliseconds(stopwatchhandle As Long) As Long
    ElapsedMilliseconds = Math.Round(1000 * ElapsedTicks(stopwatchhandle))
End Function
'Friend Function GetStopWatcheNumber() As Long
'    GetStopWatcheNumber = StopWatchCount 'If for some reason you want to know how many stopwatches you have initialized...
'End Function

