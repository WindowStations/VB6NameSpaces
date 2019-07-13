VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Timers"
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
'Attribute VB_Name = "Timers"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Timer"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Public Event TimerEvent(ByVal Index As Integer)
Private mCount    As Integer
Private mCol      As New Collection
Private mClassKey As Integer
Private mvarTimer As Timer
Private Sub Class_Initialize()
    mClassKey = RegisterTimerCollection(Me)
    Set mvarTimer = New Timer
End Sub
Private Sub Class_Terminate()
    Set mvarTimer = Nothing
    Dim t As Timer
    For Each t In mCol
        Set t = Nothing
    Next
    Set mCol = Nothing
    CTimersCol.Remove "key:" & mClassKey
End Sub
Friend Property Get Timer() As Timer
    Set Timer = mvarTimer
End Property
Friend Property Set Timer(vData As Timer)
    Set mvarTimer = vData
End Property
Property Get Count() As Integer
    Count = mCount
End Property
Function Add() As Integer
    Dim t As New Timer
    mCount = mCount + 1
    t.Index = mCount
    t.ParentsColKey = mClassKey
    mCol.Add t
    Add = mCount
End Function
Public Function BulkLoad(num As Integer)
    Dim i As Long
    If num > 0 And num < 200 Then
        For i = 0 To num
            Add
        Next
    Else
        Err.Raise vbObjectError, "CTimers.BulkLoad", "No more than 200 times allowed per class"
    End If
End Function
Public Function Item(ByVal Index As Integer) As Timer
    If Index > Count Then Err.Raise vbObjectError, "CTimers.Item", "Index does not Exist"
    Set Item = mCol(Index)
End Function
Friend Function RaiseTimer_Event(ByVal Index As Integer)
    RaiseEvent TimerEvent(Index)
End Function

