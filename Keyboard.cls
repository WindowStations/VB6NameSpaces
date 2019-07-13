VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Keyboard"
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
'Attribute VB_Name = "Keyboard"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"SendKeys"
'Option Explicit
Private Const VK_CAPITAL           As Long = &H14
Private Const VK_NUMLOCK           As Long = &H90
Private Const VK_SCROLL            As Long = &H91
Private Const SPI_GETKEYBOARDSPEED As Long = 10
Private Const SPI_GETKEYBOARDDELAY As Long = 22
Private Type kBytes
    k(0 To 255) As Byte
End Type
Private Declare Function apiGetKeyState Lib "user32" Alias "GetKeyState" (ByVal nVirtKey As Long) As Long
Private Declare Function apiGetKeyboardState Lib "user32" Alias "GetKeyboardState" (ByRef kbArray As kBytes) As Long
Private Declare Function apiSetKeyboardState Lib "user32" Alias "SetKeyboardState" (ByRef kbArray As kBytes) As Long
Private Declare Function apiGetCaretBlinkTime Lib "user32" Alias "GetCaretBlinkTime" () As Long
Private Declare Function apiGetKeyboardType Lib "user32" Alias "GetKeyboardType" (ByVal nTypeFlag As Long) As Long
Private Declare Function apiSystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private kArr         As kBytes
Private mvarCapsLock As Boolean
Private mvarSendKeys       As Sendkeys
Private mvarAltKeyDown As Boolean
Private mvarCtrlKeyDown As Boolean
Private mvarShiftKeyDown As Boolean
Public Enum kState
    kDown = 1
    kToggled = 2
    kUp = 4
End Enum
Private Sub Class_Initialize()
    Set mvarSendKeys = New Sendkeys
End Sub
Private Sub Class_Terminate()
    Set mvarSendKeys = Nothing
End Sub
Friend Property Get Sendkeys() As Sendkeys
    Set Sendkeys = mvarSendKeys
End Property
Friend Property Get AtlKeyDown() As Boolean
  AtlKeyDown = (KeyState(Keys.Menu_) = kState.kDown)
End Property
Friend Property Get CtrlKeyDown() As Boolean
  CtrlKeyDown = (KeyState(Keys.ControlKey_) = kState.kDown)
End Property
Friend Property Get ShiftKeyDown() As Boolean
  ShiftKeyDown = (KeyState(Keys.ShiftKey_) = kState.kDown)
End Property
Friend Property Let CapsLock(ByVal b As Boolean)
    apiGetKeyboardState kArr
    kArr.k(VK_CAPITAL) = CInt(b)
    apiSetKeyboardState kArr
End Property
Friend Property Get CapsLock() As Boolean
    CapsLock = IIf(apiGetKeyState(VK_CAPITAL) = 1, True, False)
End Property
Friend Property Let NumLock(ByVal b As Boolean)
    apiGetKeyboardState kArr
    kArr.k(VK_NUMLOCK) = CInt(b)
    apiSetKeyboardState kArr
End Property
Friend Property Get NumLock() As Boolean
    NumLock = IIf(apiGetKeyState(VK_NUMLOCK) = 1, True, False)
End Property
Friend Property Let ScrollLock(ByVal b As Boolean)
    apiGetKeyboardState kArr
    kArr.k(VK_SCROLL) = CInt(b)
    apiSetKeyboardState kArr
End Property
Friend Property Get ScrollLock() As Boolean
    ScrollLock = IIf(apiGetKeyState(VK_SCROLL) = 1, True, False)
End Property
Friend Property Get CaretBlinkTime() As Long
    On Error GoTo skip
    Dim cbt As Long
    cbt = apiGetCaretBlinkTime
    If cbt = 0 Then GoTo skip
    CaretBlinkTime = cbt
    Exit Property
skip:
End Property
Friend Property Get FunctionKeys() As Long
    On Error GoTo skip
    Dim kt As Long
    kt = apiGetKeyboardType(2)
    If kt <> 0 Then
        Select Case kt
            Case 1
                FunctionKeys = 10
            Case 2
                FunctionKeys = 12
            Case 3
                FunctionKeys = 10
            Case 4
                FunctionKeys = 12
            Case 5
                FunctionKeys = 10
            Case 6
                FunctionKeys = 24
            Case Else
                FunctionKeys = kt
        End Select
    End If
    Exit Property
skip:
End Property
Friend Property Get KeyState(ByVal Key As Keys) As kState
    On Error GoTo skip
    Dim ks As Integer
    ks = apiGetKeyState(Key)
    If ks And &H8000 Then
        KeyState = kState.kDown
    ElseIf ks And &H1 Then
        KeyState = kState.kToggled
    Else
        KeyState = kState.kUp
    End If
    Exit Property
skip:
End Property
Friend Property Get RepeatDelay() As Long
    On Error GoTo skip
    Dim spi       As Long
    Dim d As Long
    spi = apiSystemParametersInfo(SPI_GETKEYBOARDDELAY, 0, d, 0)
    If spi = 0 Then GoTo skip
    RepeatDelay = (d + 1) * 250
    Exit Property
skip:
End Property
Friend Property Get RepeatSpeed() As Long
    On Error GoTo skip
    Dim spi  As Long
    Dim rs   As Long
    spi = apiSystemParametersInfo(SPI_GETKEYBOARDSPEED, 0, rs, 0)
    RepeatSpeed = rs 'RepeatSpeed = 30 - (rs * 0.5)'per sec
    Exit Property
skip:
End Property
Friend Property Get TypeInfo() As String
    On Error GoTo skip
    Dim kt As Long
    kt = apiGetKeyboardType(0)
    If kt <> 0 Then
        Select Case kt
            Case 1
                TypeInfo = "PC or compatible 83-key keyboard"
            Case 2
                TypeInfo = "Olivetti 102-key keyboard"
            Case 3
                TypeInfo = "AT or compatible 84-key keyboard"
            Case 4
                TypeInfo = "Enhanced(IBM) 101-102-key keyboard"
            Case 5
                TypeInfo = "Nokia 1050 keyboard"
            Case 6
                TypeInfo = "Nokia 9140 keyboard"
            Case 7
                TypeInfo = "Japanese keyboard"
            Case Else
                TypeInfo = "Unknown keyboard"
        End Select
    End If
    Exit Property
skip:
End Property

