VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mvarDirectory As Directory
Private mvarDriveInfo As DriveInfo
Private mvarFile As File
Private mvarPath As Path
Private mvarPorts As Ports
Private Sub Class_Initialize()
    Set mvarDirectory = New Directory
    Set mvarDriveInfo = New DriveInfo
    Set mvarFile = New File
    Set mvarPath = New Path
    Set mvarPorts = New Ports
End Sub
Private Sub Class_Terminate()
    Set mvarDirectory = Nothing
    Set mvarDriveInfo = Nothing
    Set mvarFile = Nothing
    Set mvarPath = Nothing
    Set mvarPorts = Nothing
End Sub
Public Property Get Directory() As Directory
    Set Directory = mvarDirectory
End Property
Public Property Get DriveInfo() As DriveInfo
    Set DriveInfo = mvarDriveInfo
End Property
Public Property Get File() As File
    Set File = mvarFile
End Property
Public Property Get Path() As Path
    Set Path = mvarPath
End Property
Public Property Get Ports() As Ports
    Set Ports = mvarPorts
End Property
