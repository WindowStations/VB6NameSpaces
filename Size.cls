VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Size"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarWidth_  As Long
Private mvarHeight_ As Long
Public Property Let Height_(ByVal vData As Long)
    mvarHeight_ = vData
End Property
Public Property Get Height_() As Long
    Height_ = mvarHeight_
End Property
Public Property Let Width_(ByVal vData As Long)
    mvarWidth_ = vData
End Property
Public Property Get Width_() As Long
    Width_ = mvarWidth_
End Property
Friend Function Add(ByVal sz1 As Size, ByVal sz2 As Size) As Size
    'todo
End Function
Friend Function Subtract(ByVal sz1 As Size, ByVal sz2 As Size) As Size
    'todo
End Function
