VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ProcessStartInfo"
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
'Attribute VB_Name = "ProcessStartInfo"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarCreateNoWindow As Boolean
Private mvarLoadUserProfile As Boolean
Private mvarUseShellExecute As Boolean
Private mvarArguments As String
Private mvarDomain As String
Private mvarFileName As String
Private mvarUserName As String
Private mvarVerb As String
Private mvarWorkingDirectory As String
Private mvarPassword As String
Private mvarWindowStyle As Long
Friend Property Get CreateNoWindow() As Boolean
    CreateNoWindow = mvarCreateNoWindow
End Property
Friend Property Let CreateNoWindow(ByVal vData As Boolean)
    mvarCreateNoWindow = vData
End Property
Friend Property Get LoadUserProfile() As Boolean
    LoadUserProfile = mvarLoadUserProfile
End Property
Friend Property Let LoadUserProfile(ByVal vData As Boolean)
    mvarLoadUserProfile = vData
End Property
Friend Property Get UseShellExecute() As Boolean
    UseShellExecute = mvarUseShellExecute
End Property
Friend Property Let UseShellExecute(ByVal vData As Boolean)
    mvarUseShellExecute = vData
End Property
Friend Property Get Password() As String
    Password = mvarPassword
End Property
Friend Property Let Password(ByVal vData As String)
    mvarPassword = vData
End Property
Friend Property Get WorkingDirectory() As String
    WorkingDirectory = mvarWorkingDirectory
End Property
Friend Property Let WorkingDirectory(ByVal vData As String)
    mvarWorkingDirectory = vData
End Property
Friend Property Get Verb() As String
    Verb = mvarVerb
End Property
Friend Property Let Verb(ByVal vData As String)
    mvarVerb = vData
End Property
Friend Property Get UserName() As String
    UserName = mvarUserName
End Property
Friend Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property
Friend Property Get FileName() As String
    FileName = mvarFileName
End Property
Friend Property Let FileName(ByVal vData As String)
    mvarFileName = vData
End Property
Friend Property Get Domain() As String
    Domain = mvarDomain
End Property
Friend Property Let Domain(ByVal vData As String)
    mvarDomain = vData
End Property
Friend Property Get Arguments() As String
    Arguments = mvarArguments
End Property
Friend Property Let Arguments(ByVal vData As String)
    mvarArguments = vData
End Property
Friend Property Get WindowStyle() As Long
    WindowStyle = mvarWindowStyle
End Property
Friend Property Let WindowStyle(ByVal vData As Long)
     mvarWindowStyle = vData
End Property

