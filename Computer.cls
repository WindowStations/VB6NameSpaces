VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Computer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarAudio As Audio
Private mvarClipboard As Clipboard
Private mvarClock As Clock
Private mvarKeyboard As Keyboard
Private mvarMouse As Mouse
Private mvarNetwork As Network
Private mvarScreen As Screen
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
Public Property Get Audio() As Audio
    Set Audio = mvarAudio
End Property
Public Property Get Clipboard() As Clipboard
    Set Clipboard = mvarClipboard
End Property
Public Property Get Clock() As Clock
    Set Clock = mvarClock
End Property
Public Property Get Keyboard() As Keyboard
    Set Keyboard = mvarKeyboard
End Property
Public Property Get Mouse() As Mouse
    Set Mouse = mvarMouse
End Property
Public Property Get Network() As Network
    Set Network = mvarNetwork
End Property
Public Property Get Screen() As Screen
    Set Screen = mvarScreen
End Property

