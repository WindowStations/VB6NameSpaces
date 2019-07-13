VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Windows"
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
'Attribute VB_Name = "Windows"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Management"
'Attribute VB_Ext_KEY = "Member1" ,"Forms"
'Attribute VB_Ext_KEY = "Member2" ,"Automation"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarAutomation As Automation
Private mvarForms      As Forms
Private mvarManagement As Management
Private Sub Class_Initialize()
    Set mvarAutomation = New Automation
    Set mvarForms = New Forms
    Set mvarManagement = New Management
End Sub
Private Sub Class_Terminate()
    Set mvarAutomation = Nothing
    Set mvarForms = Nothing
    Set mvarManagement = Nothing
End Sub
Friend Property Get Automation() As Automation
    Set Automation = mvarAutomation
End Property
Friend Property Get Forms() As Forms
    Set Forms = mvarForms
End Property
Friend Property Get Management() As Management
    Set Management = mvarManagement
End Property

