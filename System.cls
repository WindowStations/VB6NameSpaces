VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "System"
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
'Attribute VB_Name = "System"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Diagnostics"
'Attribute VB_Ext_KEY = "Member1" ,"Drawing"
'Attribute VB_Ext_KEY = "Member2" ,"Environment"
'Attribute VB_Ext_KEY = "Member3" ,"IO"
'Attribute VB_Ext_KEY = "Member4" ,"Net"
'Attribute VB_Ext_KEY = "Member5" ,"ServiceProcess"
'Attribute VB_Ext_KEY = "Member6" ,"StringBuilder"
'Attribute VB_Ext_KEY = "Member7" ,"Threading"
'Attribute VB_Ext_KEY = "Member8" ,"Timers"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member9" ,"TimeSpan"
'Attribute VB_Ext_KEY = "Member10" ,"Windows"
'Option Explicit
Private mvarDiagnostics    As Diagnostics
Private mvarDrawing        As Drawing
Private mvarEnvironment    As Environment
Private mvarIO             As IO
Private mvarThreading      As Threading
Private mvarTimers         As Timers
Private mvarWindows        As Windows
Private mvarNet            As Net
Private mvarServiceProcess As ServiceProcess

Private Sub Class_Initialize()
    Set mvarDiagnostics = New Diagnostics
    Set mvarDrawing = New Drawing
    Set mvarEnvironment = New Environment
    Set mvarIO = New IO
    Set mvarThreading = New Threading
    Set mvarTimers = New Timers
    Set mvarWindows = New Windows
    Set mvarNet = New Net
    Set mvarServiceProcess = New ServiceProcess
End Sub
Private Sub Class_Terminate()
    Set mvarDiagnostics = Nothing
    Set mvarDrawing = Nothing
    Set mvarEnvironment = Nothing
    Set mvarIO = Nothing
    Set mvarThreading = Nothing
    Set mvarTimers = Nothing
    Set mvarWindows = Nothing
    Set mvarNet = Nothing
    Set mvarServiceProcess = Nothing
End Sub
Friend Property Get Diagnostics() As Diagnostics
    Set Diagnostics = mvarDiagnostics
End Property
Friend Property Get Drawing() As Drawing
    Set Drawing = mvarDrawing
End Property
Friend Property Get Environment() As Environment
    Set Environment = mvarEnvironment
End Property
Friend Property Get IO() As IO
    Set IO = mvarIO
End Property
Friend Property Get Threading() As Threading
    Set Threading = mvarThreading
End Property
Friend Property Get Timers() As Timers
    Set Timers = mvarTimers
End Property
Friend Property Get Windows() As Windows
    Set Windows = mvarWindows
End Property
Friend Property Get Net() As Net
    Set Net = mvarNet
End Property
Friend Property Get ServiceProcess() As ServiceProcess
    Set ServiceProcess = mvarServiceProcess
End Property


