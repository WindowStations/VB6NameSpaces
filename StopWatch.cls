VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "StopWatch"
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
'Attribute VB_Name = "StopWatch"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiQueryPerformanceCounter Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef x As Currency) As Long
Private Declare Function apiQueryPerformanceFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef x As Currency) As Long
Private overhead             As Currency
Private mvarHandle           As Long
Private mvarFrequency        As Currency
Private mvarIsHighResolution As Boolean
Friend Property Get IsHighResolution() As Currency
    IsHighResolution = mvarIsHighResolution
End Property
Friend Property Get Frequency() As Currency
    Frequency = mvarFrequency
End Property
Friend Property Get Handle() As Long
    Handle = mvarHandle
End Property
Friend Property Let Handle(ByVal vData As Long)
    mvarHandle = vData
End Property
Friend Function StartNew() As StopWatch
    Dim sw As New StopWatch
    swCount = swCount + 1
    ReDim Preserve swTic(swCount)
    ReDim Preserve swToc(swCount)
    ReDim Preserve swRunning(swCount)
    sw.Handle = swCount
    Set StartNew = sw
End Function
Friend Sub Start()
    If HandleCreated(mvarHandle) = False Then Exit Sub
    If swRunning(mvarHandle) = True Then Exit Sub
    swRunning(mvarHandle) = True
    mvarIsHighResolution = (apiQueryPerformanceFrequency(mvarFrequency) <> 0)
    'apiQueryPerformanceFrequency mvarFrequency
    apiQueryPerformanceCounter swTic(mvarHandle)
    apiQueryPerformanceCounter swToc(mvarHandle)
    overhead = swToc(mvarHandle) - swTic(mvarHandle)
    apiQueryPerformanceCounter swTic(mvarHandle)
End Sub
Friend Function ElapsedTicks() As Double
    ElapsedTicks = -1
    If HandleCreated(mvarHandle) = False Then Exit Function
    If swRunning(mvarHandle) = False Then Exit Function
    apiQueryPerformanceCounter swToc(mvarHandle)
    ElapsedTicks = (swToc(mvarHandle) - swTic(mvarHandle) - overhead) / mvarFrequency
End Function
Friend Function ElapsedMilliseconds() As Long
    ElapsedMilliseconds = Math.Round(1000 * ElapsedTicks())
End Function
Friend Function Elapsed() As TimeSpan
    Dim milli As Long
    milli = Math.Round(1000 * ElapsedTicks())
    Dim ts As New TimeSpan
    '    With ts
    '      'todo set TimeSpan class properties here when class is complete
    '    End With
    Elapsed = ts
End Function
Friend Function IsRunning() As Boolean
    IsRunning = False
    If HandleCreated(mvarHandle) = False Then Exit Function
    If swRunning(mvarHandle) = True Then IsRunning = True
End Function
Friend Function GetTimestamp() As Long
    GetTimestamp = -1
    If HandleCreated(mvarHandle) = False Then Exit Function
    GetTimestamp = 0
    apiQueryPerformanceCounter swToc(mvarHandle)
    GetTimestamp = (swToc(mvarHandle) - swTic(mvarHandle) - overhead) / mvarFrequency
End Function
Private Function HandleCreated(ByVal hwnd As Long) As Boolean
    HandleCreated = CBool((hwnd > 0) And (hwnd <= swCount))
End Function
Friend Function Reset()
    swRunning(mvarHandle) = False
    swTic(mvarHandle) = 0
    swToc(mvarHandle) = 0
    overhead = 0
    apiQueryPerformanceFrequency mvarFrequency
    apiQueryPerformanceCounter swTic(mvarHandle)
    apiQueryPerformanceCounter swToc(mvarHandle)
    overhead = swToc(mvarHandle) - swTic(mvarHandle)
    apiQueryPerformanceCounter swTic(mvarHandle)
End Function
Friend Function Stop_()
    swRunning(mvarHandle) = False
    swTic(mvarHandle) = 0
    swToc(mvarHandle) = 0
    overhead = 0
End Function

