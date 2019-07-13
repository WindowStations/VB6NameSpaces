VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Mouse"
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
'Attribute VB_Name = "Mouse"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const SM_CMOUSEBUTTONS     As Long = 43
Private Const SM_MOUSEWHEELPRESENT As Long = 75
Private Const SM_SWAPBUTTON        As Long = 23
Private Declare Function apiGetCapture Lib "user32" Alias "GetCapture" () As Long
Private Declare Function apiGetDoubleClickTime Lib "user32" Alias "GetDoubleClickTime" () As Long
Private Declare Function apiReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Long
Private Declare Function apiSetCapture Lib "user32" Alias "SetCapture" (ByVal hwnd As Long) As Long
Private Declare Function apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare Function apiSwapMouseButton Lib "user32" Alias "SwapMouseButton" (ByVal bSwap As Long)
Private Declare Function apiShowCursor Lib "user32" Alias "ShowCursor" (ByVal bShow As Long) As Long
Private Declare Function apiSetDoubleClickTime Lib "user32" Alias "SetDoubleClickTime" (ByVal wCount As Long) As Long

Private Sub SetDblClickTime(ByVal dwMilliseconds As Long)
    apiSetDoubleClickTime dwMilliseconds
End Sub
Friend Sub HideCursor()
    Dim ret As Long
    ret = apiShowCursor(False)
End Sub
Friend Sub ShowCursor()
    Dim ret As Long
    ret = apiShowCursor(True)
End Sub
Friend Property Get MouseButtonSwapped() As Boolean
    MouseButtonSwapped = CBool(apiGetSystemMetrics(SM_SWAPBUTTON))
End Property
Friend Property Let MouseButtonSwapped(ByVal b As Boolean)
    On Error GoTo er
    apiSwapMouseButton b
er:
End Property
Friend Property Get MouseExists() As Boolean
    MouseExists = CBool(apiGetSystemMetrics(SM_CMOUSEBUTTONS))
End Property
Friend Property Get MouseWheel() As Boolean
    MouseWheel = CBool(apiGetSystemMetrics(SM_MOUSEWHEELPRESENT))
End Property
Friend Property Get MouseButtonCount() As Long
    MouseButtonCount = apiGetSystemMetrics(SM_CMOUSEBUTTONS)
End Property
Friend Property Get DoubleClickTime() As Long
    DoubleClickTime = apiGetDoubleClickTime
End Property

