VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Keyboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Declare Function apiGetCapture Lib "user32" Alias "GetCapture" () As Long
Private Declare Function apiGetDoubleClickTime Lib "user32" Alias "GetDoubleClickTime" () As Long
Private Declare Function apiReleaseCapture Lib "user32" Alias "ReleaseCapture" () As Long
Private Declare Function apiSetCapture Lib "user32" Alias "SetCapture" (ByVal hwnd As Long) As Long
Private Declare Function apiSetDoubleClickTime Lib "user32" Alias "SetDoubleClickTime" (ByVal wCount As Long) As Long
Private Declare Function apiSwapMouseButton Lib "user32" Alias "SwapMouseButton" (ByVal bSwap As Long) As Long

