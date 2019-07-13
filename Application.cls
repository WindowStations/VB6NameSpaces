VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Application"
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
'Attribute VB_Name = "Application"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const PM_REMOVE   As Long = &H1
Private Const QS_ALLQUEUE As Long = 511
Private Type POINTAPI
   x As Long
   y As Long
End Type
Private Type MSG
   hwnd As Long
   Message As Long
   wParam As Long
   lParam As Long
   time As Long
   pt As POINTAPI
End Type
Private Declare Function apiPeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function apiTranslateMessage Lib "user32" Alias "TranslateMessage" (ByRef lpMsg As MSG) As Long
Private Declare Function apiDispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long
Private Declare Function apiGetInputState Lib "user32" Alias "GetInputState" () As Long
Private Declare Function apiGetQueueStatus Lib "user32" Alias "GetQueueStatus" (ByVal fuFlags As Long) As Long
Private mvarCompanyName    As String
Private mvarExecutablePath As String
Private mvarProductName    As String
Private mvarProductVersion As String
Friend Property Get ProductVersion() As String
ProductVersion = mvarProductVersion 'todo with api
End Property
Friend Property Get ProductName() As String
ProductName = mvarProductName 'todo with api
End Property
Friend Property Get ExecutablePath() As String
ExecutablePath = mvarExecutablePath 'todo with api
End Property
Friend Property Get CompanyName() As String
CompanyName = mvarCompanyName 'todo with api
End Property
Friend Function Exit_()
   End 'todo with api
End Function
Friend Function StartupPath()
   StartupPath = CurDir 'todo with API
   
End Function
Friend Function DoEventsQueued(Optional ByVal checkQueue As Boolean = True, Optional ByVal returnQueued As Boolean = False) As Long
   DoEventsQueued = -1
   If checkQueue = True And IsQueued = False Then Exit Function
   Dim d As Long
   d = DoEvents_
   If returnQueued = False Then DoEventsQueued = d
   If returnQueued = True Then DoEventsQueued = apiGetQueueStatus(QS_ALLQUEUE)
End Function 'If specified this function returns true if there are still message in the queue, otherwise false will return the number of messages dispatched to applications
Friend Function GetQueueStatus(ByVal QS_ As Long) As Long
   On Error Resume Next
   GetQueueStatus = apiGetQueueStatus(QS_ALLQUEUE)
End Function
Friend Function GetInputState() As Long
   On Error Resume Next
   GetInputState = apiGetInputState
End Function
Friend Function IsQueued() As Boolean
   If apiGetInputState <> 0 Then IsQueued = True: Exit Function
   If apiGetQueueStatus(QS_ALLQUEUE) <> 0 Then IsQueued = True: Exit Function
   IsQueued = False
End Function
'The alternative function for DoEvents:
Friend Function DoEvents_() As Long
   Dim i    As Long
   Dim cMsg As MSG
   Do
      On Error GoTo abort
      If apiPeekMessage(cMsg, 0, 0, 0, PM_REMOVE) = 0 Then Exit Do
      Call apiTranslateMessage(cMsg)
      Call apiDispatchMessage(cMsg)
      i = i + 1
      If i = 400 Then Exit Do 'Let's not hang for to long, just in case, we exit at upper limit
   Loop
abort:
   DoEvents_ = i
End Function 'returns the number of dispatched messages before finding an empty queue

