VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "NetworkAdapter"
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
'Attribute VB_Name = "NetworkAdapter"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mName                        As String
Private mDHCPEnabled                 As String
Private mDHCPLeaseExpires            As Variant
Private mDHCPLeaseObtained           As Variant
Private mDHCPServer                  As Variant
Private mDNSDomain                   As Variant
Private mDNSEnabledForWINSResolution As Boolean
Private mIPEnabled                   As Boolean
Private mIndex                       As Long
Private mIPFilterSecurityEnabled     As Boolean
Private mIPPortSecurityEnabled       As Boolean
Private mIPUseZeroBroadcast          As Boolean
Private mIPv4Address                 As String
Private mMacAddress                  As String
Private mWINSEnableLMHostsLookup     As Boolean
Private mWINSHostLookupFile          As String
Private mWINSPrimaryServer           As String
Private mWINSSecondaryServer         As String
Private mDefaultIPGateway            As Variant
Private mSubnetMask                  As String
Private mIPv6Address                 As String
Private mNic                         As Object
Private mNic2                        As Object
Private mStatus                      As String
Private mManufacturer                As String
Private mMaxSpeed                    As Variant
Friend Property Set NicObject(x As Object)
    Set mNic = x
End Property
Friend Property Set NicObject2(x As Object)
    Set mNic2 = x
End Property
Friend Property Get NicObject() As Object
    Set NicObject = mNic
End Property
Friend Property Get NicObject2() As Object
    Set NicObject2 = mNic2
End Property
Friend Property Get Name() As String
    On Error Resume Next
    Name = mNic.Description
End Property
Friend Property Get ServiceName() As String
    On Error Resume Next
    ServiceName = mNic.ServiceName
End Property
Friend Property Get DHCPEnabled() As Boolean
    On Error Resume Next
    DHCPEnabled = mNic.DHCPEnabled
End Property
Friend Property Get DHCPLeaseExpires() As Variant
    On Error Resume Next
    Dim dtb  As String
    Dim bias As Long
    Dim d    As String
    Dim t    As String
    dtb = mNic.DHCPLeaseExpires
    bias = SplitDateTimeBias(dtb, d, t)
    DHCPLeaseExpires = Format(d, "short date")
End Property
Friend Property Get DHCPLeaseObtained() As Variant
    On Error Resume Next
    Dim dtb  As String
    Dim bias As Long
    Dim d    As String
    Dim t    As String
    dtb = mNic.DHCPLeaseObtained
    bias = SplitDateTimeBias(dtb, d, t)
    DHCPLeaseObtained = Format(d, "short date")
End Property
Friend Property Get DHCPServer() As Variant
    On Error Resume Next
    DHCPServer = mNic.DHCPServer
End Property
Property Get DefaultIPGateway() As Variant
    On Error Resume Next
    DefaultIPGateway = mNic.DefaultIPGateway
End Property
Friend Property Get DNSDomain() As Variant
    On Error Resume Next
    DNSDomain = mNic.DNSDomain
End Property
Friend Property Get DNSEnabledForWINSResolution() As Boolean
    On Error Resume Next
    DNSEnabledForWINSResolution = mNic.DNSEnabledForWINSResolution
End Property
Friend Property Get IPEnabled() As Boolean
    On Error Resume Next
    IPEnabled = mNic.IPEnabled
End Property
Friend Property Get Index() As Long
    On Error Resume Next
    Index = mNic.Index
End Property
Friend Property Get IPFilterSecurityEnabled() As Boolean
    On Error Resume Next
    IPFilterSecurityEnabled = mNic.IPFilterSecurityEnabled
End Property
Friend Property Get IPPortSecurityEnabled() As Boolean
    On Error Resume Next
    IPPortSecurityEnabled = mNic.IPPortSecurityEnabled
End Property
Friend Property Get IPUseZeroBroadcast() As Boolean
    On Error Resume Next
    IPUseZeroBroadcast = mNic.IPUseZeroBroadcast
End Property
Friend Property Get IPv4Address() As String
    On Error Resume Next
    IPv4Address = mNic.IPAddress(0)
End Property
Friend Property Get IPv6Address() As String
    On Error Resume Next
    IPv6Address = mNic.IPv6Address(1)
End Property
Friend Property Get MacAddress() As String
    On Error Resume Next
    MacAddress = mNic.MacAddress
End Property
Friend Property Get WINSEnableLMHostsLookup() As Boolean
    On Error Resume Next
    WINSEnableLMHostsLookup = mNic.WINSEnableLMHostsLookup
End Property
Friend Property Get WINSHostLookupFile() As String
    On Error Resume Next
    WINSHostLookupFile = mNic.WINSHostLookupFile
End Property
Property Get WINSPrimaryServer() As String
    On Error Resume Next
    WINSPrimaryServer = mNic.WINSPrimaryServer
End Property
Friend Property Get WINSSecondaryServer() As String
    On Error Resume Next
    WINSSecondaryServer = mNic.WINSSecondaryServer
End Property
Friend Property Get SubnetMask() As String
    On Error Resume Next
    SubnetMask = mNic.SubnetMask
End Property
Private Function SplitDateTimeBias(ByVal leasedate As String, ByRef leasedatepart As String, ByRef leasetimepart As String) As Long
    Dim pos  As Long
    Dim bias As Long
    pos = InStr(leasedate, ".")
    If pos > 0 Then
        bias = StripTimeZoneBias(leasedate)
        leasedatepart = Left(leasedate, 8)
        leasetimepart = Mid(leasedate, 9, pos - Len(leasedatepart) - 1)
        leasedatepart = InsertInString(leasedatepart, "-", 5, "")
        leasedatepart = InsertInString(leasedatepart, "-", 8, "")
        leasetimepart = InsertInString(leasetimepart, ":", 3, "")
        leasetimepart = InsertInString(leasetimepart, ":", 6, "")
        SplitDateTimeBias = bias
    Else
    End If
End Function
Private Function InsertInString(ByVal sOriginal As String, ByRef sReplace As String, ByRef nField As Long, ByRef sDelimeter As String) As String
    Dim nCount As Long
    Dim nStart As Long
    Dim nLast  As Long
    Do 'While InStr(nStart + 1, sOriginal, sDelimeter) > 0
        If InStr(nStart + 1, sOriginal, sDelimeter) <= 0 Then Exit Do
        nStart = InStr(nStart + 1, sOriginal, sDelimeter)
        nCount = nCount + 1
        If nCount >= nField Then Exit Do
        nLast = nStart
    Loop
    Select Case nCount
        Case 1
            InsertInString = sReplace & Mid(sOriginal, nStart)
        Case Is >= nField
            InsertInString = Mid(sOriginal, 1, nLast) & sReplace & Mid(sOriginal, nStart)
        Case Else
            InsertInString = sOriginal & String((nField - 1) - nCount, sDelimeter) & sReplace
    End Select
End Function
Private Function StripTimeZoneBias(leasedate As String) As Long
    Dim pos As Long
    Dim tmp As String
    pos = InStr(leasedate, "-")
    If pos = 0 Then
        pos = InStr(leasedate, "+")
        If pos = 0 Then StripTimeZoneBias = 0
    Else
        tmp = Mid(leasedate, pos, Len(leasedate))
        leasedate = Mid(leasedate, 1, pos - 1)
        StripTimeZoneBias = CLng(tmp)
    End If
End Function
Friend Sub ReleaseDHCPLease()
    mNic.ReleaseDHCPLease
End Sub
Friend Sub RenewDHCPLease()
    mNic.RenewDHCPLease
End Sub

