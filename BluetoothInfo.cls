VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "BluetoothInfo"
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
'Attribute VB_Name = "BluetoothInfo"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mAddress       As String
Private mClassofDevice As Long
Private mConnected     As Boolean
Private mRemembered    As Boolean
Private mAuthenticated As Boolean
Private mLastSeen      As Variant
Private mLastUsed      As Variant
Private mName          As String
Friend Property Get Address() As String
    Address = mAddress
End Property
Friend Property Get ClassofDevice() As Long
    ClassofDevice = mClassofDevice
End Property
Friend Property Get Connected() As Boolean
    Connected = mConnected
End Property
Friend Property Get Remembered() As Boolean
    Remembered = mRemembered
End Property
Friend Property Get Authenticated() As Boolean
    Authenticated = mAuthenticated
End Property
Friend Property Get LastSeen() As Variant
    LastSeen = mLastSeen
End Property
Friend Property Get LastUsed() As Variant
    LastUsed = mLastUsed
End Property
Friend Property Get Name() As String
    Name = mName
End Property
Friend Sub Assign(ByVal Address As String, ByVal ulClassofDevice As Long, ByVal fConnected As Long, ByVal fRemembered As Long, ByVal fAuthenticated As Long, ByVal LastSeen As Variant, ByVal LastUsed As Variant, ByVal Name As String)
    mAddress = Address
    mClassofDevice = ulClassofDevice
    mConnected = CBool(fConnected)
    mRemembered = CBool(fRemembered)
    mAuthenticated = CBool(fAuthenticated)
    mLastSeen = LastSeen
    mLastUsed = LastUsed
    mName = Name
End Sub

