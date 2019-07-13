VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "NetworkAdapters"
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
'Attribute VB_Name = "NetworkAdapters"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mCol      As Collection
Dim objWMIService As Object
Private Sub Class_Initialize()
    Set mCol = New Collection
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    GetNetWorkAdapter
End Sub
Private Sub Class_Terminate()
    Set mCol = Nothing
End Sub
Friend Property Get Item(ByVal vntIndexKey As Variant) As NetworkAdapter
    Set Item = mCol(vntIndexKey)
End Property
Friend Property Get Count() As Long
    Count = mCol.Count
End Property
Private Sub GetNetWorkAdapter()
    On Local Error Resume Next
    Dim nicSet As Object
    Dim nic    As Object
    Dim dtb    As String
    Dim d      As String
    Dim t      As String
    Dim bias   As Long
    Dim d1     As Long
    Set nicSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_NetworkAdapterConfiguration")
    DoEvents
    For Each nic In nicSet
        On Error GoTo skip
        Add nic, Val(nic.Index) + 1
skip:
    Next
End Sub
Private Function Add(ByVal NicObject As Object, Optional Index As String) As NetworkAdapter
    Dim objNewMember As Object
    Set objNewMember = New NetworkAdapter
    Set objNewMember.NicObject = NicObject
    If Index > 0 Then
        mCol.Add objNewMember, Index
    Else
        mCol.Add objNewMember, Index
    End If
    Set Add = objNewMember
    Set objNewMember = Nothing
End Function
Private Sub RenewAll()
    Dim objNetworkSettings As Object
    Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
    objNetworkSettings.RenewDHCPLeaseAll
End Sub
Private Sub ReleaseAll()
    Dim objNetworkSettings As Object
    Set objNetworkSettings = objWMIService.Get("Win32_NetworkAdapterConfiguration")
    objNetworkSettings.ReleaseDHCPLeaseAll
End Sub

