VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "System"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarDiagnostics As Diagnostics
Private mvarDrawing As Drawing
Private mvarEnvironment As Environment
Private mvarIO As IO
Private mvarStringBuilder As StringBuilder
Private mvarThreading As Threading
Private mvarTimers As Timers
Private mvarTimeSpan As TimeSpan
Private mvarWindows As Windows
Private Sub Class_Initialize()
    Set mvarDiagnostics = New Diagnostics
    Set mvarDrawing = New Drawing
    Set mvarEnvironment = New Environment
    Set mvarIO = New IO
    Set mvarStringBuilder = New StringBuilder
    Set mvarThreading = New Threading
    Set mvarTimers = New Timers
    Set mvarTimeSpan = New TimeSpan
    Set mvarWindows = New Windows
End Sub
Private Sub Class_Terminate()
    Set mvarDiagnostics = Nothing
    Set mvarDrawing = Nothing
    Set mvarEnvironment = Nothing
    Set mvarIO = Nothing
    Set mvarStringBuilder = Nothing
    Set mvarThreading = Nothing
    Set mvarTimers = Nothing
    Set mvarTimeSpan = Nothing
    Set mvarWindows = Nothing
End Sub
Public Property Get Diagnostics() As Diagnostics
    Set Diagnostics = mvarDiagnostics
End Property
Public Property Get Drawing() As Drawing
    Set Drawing = mvarDrawing
End Property
Public Property Get Environment() As Environment
    Set Environment = mvarEnvironment
End Property
Public Property Get IO() As IO
    Set IO = mvarIO
End Property
Public Property Get StringBuilder() As StringBuilder
    Set StringBuilder = mvarStringBuilder
End Property
Public Property Get Threading() As Threading
    Set Threading = mvarThreading
End Property
Public Property Get Timers() As Timers
    Set Timers = mvarTimers
End Property
Public Property Get TimeSpan() As TimeSpan
    Set TimeSpan = mvarTimeSpan
End Property
Public Property Get Windows() As Windows
    Set Windows = mvarWindows
End Property
