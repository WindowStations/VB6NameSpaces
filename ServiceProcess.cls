VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ServiceProcess"
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
'Attribute VB_Name = "ServiceProcess"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"ServiceController"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarServiceController As ServiceController
Public Enum ServiceControllerStatus
  Stopped = 1
  StartPending = 2
  StopPending = 3
  Running = 4
  ContinuePending = 5
  PausePending = 6
  Paused = 7
End Enum
Private Sub Class_Initialize()
    Set mvarServiceController = New ServiceController
End Sub
Private Sub Class_Terminate()
    Set mvarServiceController = Nothing
End Sub
Friend Property Get ServiceController() As ServiceController
    Set ServiceController = mvarServiceController
End Property


