VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Point"
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
'Attribute VB_Name = "Point"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarX As Long
Private mvarY As Long
Friend Property Let y(ByVal vData As Long)
mvarY = vData
End Property
Friend Property Get y() As Long
y = mvarY
End Property
Friend Property Let x(ByVal vData As Long)
mvarX = vData
End Property
Friend Property Get x() As Long
x = mvarX
End Property
Friend Function Add(ByVal pt As Point, ByVal sz As Size) As Point
   'todo
End Function
Friend Function Subtract(ByVal pt As Point, ByVal sz As Size) As Point
   'todo
End Function

Private Sub Class_Initialize()
   mvarX = x
   mvarY = y
End Sub

