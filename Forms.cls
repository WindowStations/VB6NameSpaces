VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Forms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarApplication As Application
Private mvarMessagebox As Messagebox
Private mvarSendKeys As SendKeys
Private Sub Class_Initialize()
    Set mvarApplication = New Application
    Set mvarMessagebox = New Messagebox
    Set mvarSendKeys = New SendKeys
End Sub
Private Sub Class_Terminate()
    Set mvarApplication = Nothing
    Set mvarMessagebox = Nothing
    Set mvarSendKeys = Nothing
End Sub
Public Property Get Application() As Application
    Set Application = mvarApplication
End Property
Public Property Get Messagebox() As Messagebox
    Set Messagebox = mvarMessagebox
End Property
Public Property Get SendKeys() As SendKeys
    Set SendKeys = mvarSendKeys
End Property
