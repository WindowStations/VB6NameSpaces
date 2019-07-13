VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "IO"
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
'Attribute VB_Name = "IO"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"StreamWriter"
'Attribute VB_Ext_KEY = "Member1" ,"StreamReader"
'Attribute VB_Ext_KEY = "Member2" ,"Ports"
'Attribute VB_Ext_KEY = "Member3" ,"path"
'Attribute VB_Ext_KEY = "Member4" ,"File"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member5" ,"DriveInfo"
'Attribute VB_Ext_KEY = "Member6" ,"Directory"
'Option Explicit
Private mvarDirectory As Directory
Private mvarDriveInfo As DriveInfo
Private mvarFile      As File
Private mvarPath      As path
Private mvarPorts     As Ports
Private mvarStreamReader As StreamReader
Private mvarStreamWriter As StreamWriter
Private Sub Class_Initialize()
    Set mvarDirectory = New Directory
    Set mvarDriveInfo = New DriveInfo
    Set mvarFile = New File
    Set mvarPath = New path
    Set mvarPorts = New Ports
    Set mvarStreamReader = New StreamReader
    Set mvarStreamWriter = New StreamWriter
End Sub
Private Sub Class_Terminate()
    Set mvarDirectory = Nothing
    Set mvarDriveInfo = Nothing
    Set mvarFile = Nothing
    Set mvarPath = Nothing
    Set mvarPorts = Nothing
    Set mvarStreamReader = Nothing
    Set mvarStreamWriter = Nothing
End Sub
Friend Property Get Directory() As Directory
    Set Directory = mvarDirectory
End Property
Friend Property Get DriveInfo() As DriveInfo
    Set DriveInfo = mvarDriveInfo
End Property
Friend Property Get File() As File
    Set File = mvarFile
End Property
Friend Property Get path() As path
    Set path = mvarPath
End Property
Friend Property Get Ports() As Ports
    Set Ports = mvarPorts
End Property
Friend Property Get StreamReader() As StreamReader
    Set StreamReader = mvarStreamReader
End Property
Friend Property Get StreamWriter() As StreamWriter
    Set StreamWriter = mvarStreamWriter
End Property

