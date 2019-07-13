VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Computer"
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
'Attribute VB_Name = "Computer"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Screen"
'Attribute VB_Ext_KEY = "Member1" ,"Network"
'Attribute VB_Ext_KEY = "Member2" ,"Mouse"
'Attribute VB_Ext_KEY = "Member3" ,"Keyboard"
'Attribute VB_Ext_KEY = "Member4" ,"Clock"
'Attribute VB_Ext_KEY = "Member5" ,"Clipboard"
'Attribute VB_Ext_KEY = "Member6" ,"Audio"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarAudio     As Audio
Private mvarClipboard As Clipboard
Private mvarClock     As Clock
Private mvarKeyboard  As Keyboard
Private mvarMouse     As Mouse
Private mvarNetwork   As Network
Private mvarScreen    As Screen
Private Sub Class_Initialize()
    Set mvarAudio = New Audio
    Set mvarClipboard = New Clipboard
    Set mvarClock = New Clock
    Set mvarKeyboard = New Keyboard
    Set mvarMouse = New Mouse
    Set mvarNetwork = New Network
    Set mvarScreen = New Screen
End Sub
Private Sub Class_Terminate()
    Set mvarAudio = Nothing
    Set mvarClipboard = Nothing
    Set mvarClock = Nothing
    Set mvarKeyboard = Nothing
    Set mvarMouse = Nothing
    Set mvarNetwork = Nothing
    Set mvarScreen = Nothing
End Sub
Friend Property Get Audio() As Audio
    Set Audio = mvarAudio
End Property
Friend Property Get Clipboard() As Clipboard
    Set Clipboard = mvarClipboard
End Property
Friend Property Get Clock() As Clock
    Set Clock = mvarClock
End Property
Friend Property Get Keyboard() As Keyboard
    Set Keyboard = mvarKeyboard
End Property
Friend Property Get Mouse() As Mouse
    Set Mouse = mvarMouse
End Property
Friend Property Get Network() As Network
    Set Network = mvarNetwork
End Property
Friend Property Get Screen() As Screen
    Set Screen = mvarScreen
End Property

