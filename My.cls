VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "My"
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
'Attribute VB_Name = "My"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = True
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Application"
'Attribute VB_Ext_KEY = "Member1" ,"Computer"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member2" ,"User"
'Option Explicit
Private mvarComputer    As Computer
Private mvarUser        As User
Private mvarApplication As Application
Private Sub Class_Initialize()
    Set mvarComputer = New Computer
    Set mvarUser = New User
    Set mvarApplication = New Application
End Sub
Private Sub Class_Terminate()
    Set mvarComputer = Nothing
    Set mvarUser = Nothing
    Set mvarApplication = Nothing
End Sub
Friend Property Get Computer() As Computer
    Set Computer = mvarComputer
End Property
Friend Property Get User() As User
    Set User = mvarUser
End Property
Friend Property Get Application() As Application
    Set Application = mvarApplication
End Property

