VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Win32"
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
'Attribute VB_Name = "Win32"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"SystemEvents"
'Attribute VB_Ext_KEY = "Member1" ,"RegistryKey"
'Attribute VB_Ext_KEY = "Member2" ,"Registry"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarRegistry     As Registry
Private mvarRegistryKey  As RegistryKey
Private mvarSystemEvents As SystemEvents
Private Sub Class_Initialize()
    Set mvarRegistry = New Registry
    Set mvarRegistryKey = New RegistryKey
    Set mvarSystemEvents = New SystemEvents
End Sub
Private Sub Class_Terminate()
    Set mvarRegistry = Nothing
    Set mvarRegistryKey = Nothing
    Set mvarSystemEvents = Nothing
End Sub
Friend Property Get Registry() As Registry
    Set Registry = mvarRegistry
End Property
Friend Property Get RegistryKey() As RegistryKey
    Set RegistryKey = mvarRegistryKey
End Property
Friend Property Get SystemEvents() As SystemEvents
    Set SystemEvents = mvarSystemEvents
End Property

