VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Size"
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
'Attribute VB_Name = "Size"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarWidth_  As Long
Private mvarHeight_ As Long
Friend Property Let Height_(ByVal vData As Long)
    mvarHeight_ = vData
End Property
Friend Property Get Height_() As Long
    Height_ = mvarHeight_
End Property
Friend Property Let Width_(ByVal vData As Long)
    mvarWidth_ = vData
End Property
Friend Property Get Width_() As Long
    Width_ = mvarWidth_
End Property
Friend Function Add(ByVal sz1 As Size, ByVal sz2 As Size) As Size
    'todo
End Function
Friend Function Subtract(ByVal sz1 As Size, ByVal sz2 As Size) As Size
    'todo
End Function

