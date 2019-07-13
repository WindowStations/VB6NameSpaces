VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "User"
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
'Attribute VB_Name = "User"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiGetUserName Lib "advapi32" Alias "GetUserNameA" (ByVal lpBuffer As String, ByRef nSize As Long) As Long
Private Declare Function apiIsUserAnAdmin Lib "shell32" Alias "IsUserAnAdmin" () As Long
Private mvarCurrentPrincipal As Object
Friend Function Name() As String
    On Error Resume Next
    Dim buff  As String
    Dim lSize As Long
    buff = Space(260)
    lSize = Len(buff)
    Call apiGetUserName(buff, lSize)
    Name = Left(buff, lSize)
End Function
Friend Function IsInRole(ByVal role As String) As Boolean
    IsInRole = CBool(apiIsUserAnAdmin)
End Function
Friend Function IsAuthenticated() As Boolean
    'todo
End Function
Friend Sub InitializeWithWindowsUser()
    'todo
End Sub
Friend Property Get CurrentPrincipal() As Object
    CurrentPrincipal = mvarCurrentPrincipal 'todo
End Property
Friend Property Let CurrentPrincipal(ByVal vData As Object)
    mvarCurrentPrincipal = vData 'todo
End Property

