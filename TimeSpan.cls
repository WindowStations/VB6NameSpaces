VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "TimeSpan"
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
'Attribute VB_Name = "TimeSpan"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarDays              As Long
Private mvarHours             As Long
Private mvarMinutes           As Long
Private mvarSeconds           As Long
Private mvarMilliseconds      As Long
Private mvarTicks             As Long
Private mvarTotalDays         As Long
Private mvarTotalHours        As Long
Private mvarTotalMilliseconds As Long
Private mvarTotalMinutes      As Long
Private mvarTotalSeconds      As Long
Friend Property Get Days() As Long
    Days = mvarDays
End Property
Friend Property Get Hours() As Long
    Hours = mvarHours
End Property
Friend Property Get Minutes() As Long
    Minutes = mvarMinutes
End Property
Friend Property Get Seconds() As Long
    Seconds = mvarSeconds
End Property
Friend Property Get Milliseconds() As Long
    Milliseconds = mvarMilliseconds
End Property
Friend Property Get Ticks() As Long
    Ticks = mvarTicks
End Property
Friend Property Get TotalDays() As Long
    TotalDays = mvarTotalDays
End Property
Friend Property Get TotalHours() As Long
    TotalHours = mvarTotalHours
End Property
Friend Property Get TotalMilliseconds() As Long
    TotalMilliseconds = mvarTotalMilliseconds
End Property
Friend Property Get TotalMinutes() As Long
    TotalMinutes = mvarTotalMinutes
End Property
Friend Property Get TotalSeconds() As Long
    TotalSeconds = mvarTotalSeconds
End Property
Friend Function TicksPerDay() As Long
    TicksPerDay = 864000000000#
End Function
Friend Function TicksPerHour() As Long
    TicksPerHour = 36000000000#
End Function
Friend Function TicksPerMillisecond() As Long
    TicksPerMillisecond = 10000
End Function
Friend Function TicksPerMinute() As Long
    TicksPerMinute = 600000000
End Function
Friend Function Add(ByVal ts As TimeSpan) As TimeSpan
    mvarTotalMilliseconds = mvarTotalMilliseconds + ts.TotalMilliseconds
End Function
Friend Function Compare(ByVal ts1 As TimeSpan, ByVal ts2 As TimeSpan) As Long
End Function
Friend Function Subtract(ByVal ts As TimeSpan) As TimeSpan
End Function
Friend Function duration() As TimeSpan
End Function
Friend Function FromDays(ByVal d As Double) As TimeSpan
End Function
Friend Function FromHours(ByVal d As Double) As TimeSpan
End Function
Friend Function FromMilliseconds(ByVal d As Double) As TimeSpan
End Function
Friend Function FromMinutes(ByVal d As Double) As TimeSpan
End Function
Friend Function FromSeconds(ByVal d As Double) As TimeSpan
End Function
Friend Function FromTicks(ByVal d As Long) As TimeSpan
End Function

