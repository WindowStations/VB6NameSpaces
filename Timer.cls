VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Timer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const DEBUGMODE As Boolean = False
Private Declare Function apiSetTimer Lib "user32" Alias "SetTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private timerId         As Long
Private isActive        As Boolean
Public Interval         As Long
Public Tag              As Variant
Public Event Timer()
Private mIndex As Long
Private mParentsColKey
Private Sub Class_Initialize()
    Interval = 2000
End Sub
Private Sub Class_Terminate()
    TimerDestroy
End Sub
Friend Property Let Index(lIndex As Long)
    mIndex = lIndex
End Property
Friend Property Get Index() As Long
    Index = mIndex
End Property
Friend Property Let ParentsColKey(key As Long)
    mParentsColKey = key
End Property
Friend Property Get ParentsColKey() As Long
    ParentsColKey = mParentsColKey
End Property
Property Get Enabled() As Boolean
    Enabled = isActive
End Property
Property Let Enabled(activateIt As Boolean)
    If activateIt And isActive Then
        Reset
    ElseIf activateIt And Not isActive Then
        Enable
    ElseIf isActive And Not activateIt Then
        Disable
    ElseIf Not isActive And Not activateIt Then
        'dumb ass
    End If
End Property
Sub EnableTimer(mInterval As Long)
    If Interval = mInterval And isActive Then TimerDestroy
    Interval = mInterval
    Enable
End Sub
Sub Enable()
    If isActive Then TimerDestroy
    TimerCreate
End Sub
Sub Disable()
    TimerDestroy
End Sub
Sub Reset()
    TimerDestroy
    TimerCreate
End Sub
Public Sub RaiseTimer_Event()
    RaiseEvent Timer
End Sub
Private Function TimerCreate() As Boolean
    If Interval <= 0 Then Exit Function
    timerId = apiSetTimer(0, 0, Interval, AddressOf TimerProc)
    If timerId Then
        TimerCreate = True
        Timercollection.Add Me, "id:" & timerId
        isActive = True
    Else
        TimerCreate = False
        timerId = 0
        isActive = False
    End If
End Function
Private Function TimerDestroy() As Long
    Dim i As Integer
    Dim f As Boolean
    If TimerExists() = True Then
        f = apiKillTimer(0, timerId)
        Timercollection.Remove "id:" & timerId
        TimerDestroy = True
        isActive = False
    End If
End Function
Private Function TimerExists() As Boolean
    On Error Resume Next
    Dim c As Timer
    Set c = Timercollection("id:" & timerId)
    If Err.Number = 0 Then
        TimerExists = True
    Else
    End If
    Set c = Nothing
End Function
