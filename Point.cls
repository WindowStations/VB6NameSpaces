VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarX     As Long
Private mvarY     As Long
Public Property Let y(ByVal vData As Long)
    mvarY = vData
End Property
Public Property Get y() As Long
    y = mvarY
End Property
Public Property Let x(ByVal vData As Long)
    mvarX = vData
End Property
Public Property Get x() As Long
    x = mvarX
End Property
Friend Function Add(ByVal pt As Point, ByVal sz As Size) As Point
    'todo
End Function
Friend Function Subtract(ByVal pt As Point, ByVal sz As Size) As Point
    'todo
End Function

