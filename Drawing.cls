VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Drawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mvarPoint As Point
Private mvarRectangle As Rectangle
Private mvarSize As Size

Private Sub Class_Initialize()
    Set mvarPoint = New Point
    Set mvarRectangle = New Rectangle
    Set mvarSize = New Size
End Sub
Private Sub Class_Terminate()
    Set mvarPoint = Nothing
    Set mvarRectangle = Nothing
    Set mvarSize = Nothing
End Sub
Public Property Get Point() As Point
    Set Point = mvarPoint
End Property
Public Property Get Rectangle() As Rectangle
    Set Rectangle = mvarRectangle
End Property
Public Property Get Size() As Size
    Set Size = mvarSize
End Property
