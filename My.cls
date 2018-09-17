VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "My"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarComputer As Computer
Private mvarUser As User
Private Sub Class_Initialize()
    Set mvarComputer = New Computer
    Set mvarUser = New User
End Sub
Private Sub Class_Terminate()
    Set mvarComputer = Nothing
    Set mvarUser = Nothing
End Sub
Public Property Get Computer() As Computer
    Set Computer = mvarComputer
End Property
Public Property Get User() As User
    Set User = mvarUser
End Property

