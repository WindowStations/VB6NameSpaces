VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Windows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarAutomation As Automation
Private mvarForms As Forms
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
Public Property Get Automation() As Automation
    Set Automation = mvarAutomation
End Property
Public Property Get Forms() As Forms
    Set Forms = mvarForms
End Property
Public Property Get Management() As Management
    Set Management = mvarManagement
End Property

