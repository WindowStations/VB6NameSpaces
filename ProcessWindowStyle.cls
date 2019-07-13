VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ProcessWindowStyle"
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
'Attribute VB_Name = "ProcessWindowStyle"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Friend Property Get Normal() As Long
     Normal = 0
End Property
Friend Property Get Hidden() As Long
     Hidden = 1
End Property
Friend Property Get Minimized() As Long
     Minimized = 2
End Property
Friend Property Get Maximized() As Long
     Maximized = 3
End Property
'  Friend Function Normal() As Long
'    Normal = 0
'  End Function
'   Friend Function Hidden() As Long
'    Hidden = 1
'  End Function
'   Friend Function Minimized() As Long
'    Minimized = 2
'  End Function
'Friend Function Maximized() As Long
'    Maximized = 3
'End Function

