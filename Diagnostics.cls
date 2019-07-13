VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Diagnostics"
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
'Attribute VB_Name = "Diagnostics"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"ProcessWindowStyle"
'Attribute VB_Ext_KEY = "Member1" ,"ProcessStartInfo"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member2" ,"StopWatch"
'Attribute VB_Ext_KEY = "Member3" ,"Process"
'Option Explicit
Private mvarProcess          As Process
Private mvarStopWatch        As StopWatch
Private mvarProcessStartInfo As ProcessStartInfo
Private mvarProcessWindowStyle As ProcessWindowStyle
Private Sub Class_Initialize()
    Set mvarProcess = New Process
    Set mvarStopWatch = New StopWatch
    Set mvarProcessStartInfo = New ProcessStartInfo
End Sub
Private Sub Class_Terminate()
    Set mvarProcess = Nothing
    Set mvarStopWatch = Nothing
    Set mvarProcessStartInfo = Nothing
End Sub
Friend Property Get Process() As Process
    Set Process = mvarProcess
End Property
Friend Property Get StopWatch() As StopWatch
    Set StopWatch = mvarStopWatch
End Property
Friend Property Get ProcessStartInfo() As ProcessStartInfo
    Set ProcessStartInfo = mvarProcessStartInfo
End Property
Friend Property Get ProcessWindowStyle() As ProcessWindowStyle
    Set ProcessWindowStyle = mvarProcessWindowStyle
End Property

