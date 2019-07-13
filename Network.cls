VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Network"
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
'Attribute VB_Name = "Network"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"FTP"
'Attribute VB_Ext_KEY = "Member1" ,"HTTP"
'Attribute VB_Ext_KEY = "Member2" ,"NetworkAdapters"
'Option Explicit
Private Const MAX_HOSTNAME_LEN               As Long = 132
Private Const MAX_DOMAIN_NAME_LEN            As Long = 132
Private Const MAX_SCOPE_ID_LEN               As Long = 260
Private Const MAX_ADAPTER_NAME_LENGTH        As Long = 260
Private Const MAX_ADAPTER_ADDRESS_LENGTH     As Long = 8
Private Const MAX_ADAPTER_DESCRIPTION_LENGTH As Long = 132
Private Const ERROR_BUFFER_OVERFLOW          As Long = 111
Private Const MIB_IF_TYPE_ETHERNET           As Long = 6
Private Const MIB_IF_TYPE_TOKENRING          As Long = 9
Private Const MIB_IF_TYPE_FDDI               As Long = 15
Private Const MIB_IF_TYPE_PPP                As Long = 23
Private Const MIB_IF_TYPE_LOOPBACK           As Long = 24
Private Const MIB_IF_TYPE_SLIP               As Long = 28
Private Const IP_SUCCESS                     As Long = 0
Private Const SOCKET_ERROR                   As Long = -1
Private Const MAX_WSADescription             As Long = 256
Private Const MAX_WSASYSStatus               As Long = 128
Private Const MIN_SOCKETS_REQD               As Long = 1
Private Const WS_VERSION_REQD                As Long = &H101
Private Const WS_VERSION_MAJOR               As Long = WS_VERSION_REQD \ &H100 And &HFF
Private Const WS_VERSION_MINOR               As Long = WS_VERSION_REQD And &HFF
Private Const WSADescription_Len             As Long = 256
Private Const WSASYS_Status_Len              As Long = 128
Private Const AF_INET                        As Long = 2
Private Const INFINITE                       As Long = -1
Private Const SYNCHRONIZE                    As Long = &H100000
Private Type IP_ADDR_STRING
    lNext As Long
    IPAddress As String * 16
    IpMask As String * 16
    lContext As Long
End Type
Private Type FIXED_INFO
    HostName As String * MAX_HOSTNAME_LEN
    DomainName As String * MAX_DOMAIN_NAME_LEN
    CurrentDnsServer As Long
    DnsServerList As IP_ADDR_STRING
    NodeType As Long
    ScopeId As String * MAX_SCOPE_ID_LEN
    EnableRouting As Long
    EnableProxy As Long
    EnableDns As Long
End Type
Private Type IP_ADAPTER_INFO
    lNext As Long
    ComboIndex As Long
    AdapterName As String * MAX_ADAPTER_NAME_LENGTH
    Description As String * MAX_ADAPTER_DESCRIPTION_LENGTH
    AddressLength As Long
    Address(MAX_ADAPTER_ADDRESS_LENGTH - 1) As Byte
    Index As Long
    lType As Long
    DHCPEnabled As Long
    CurrentIpAddress As Long
    IpAddressList As IP_ADDR_STRING
    GatewayList As IP_ADDR_STRING
    DHCPServer As IP_ADDR_STRING
    HaveWins As Byte
    PrimaryWinsServer As IP_ADDR_STRING
    SecondaryWinsServer As IP_ADDR_STRING
    LeaseObtained As Long
    LeaseExpires As Long
End Type
Private Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type
Private Type WSADATA
    wVersion As Integer
    wHighVersion As Integer
    szDescription(0 To MAX_WSADescription) As Byte
    szSystemStatus(0 To MAX_WSASYSStatus) As Byte
    wMaxSockets As Long
    wMaxUDPDG As Long
    dwVendorInfo As Long
End Type
Private Declare Sub apiCopyMemoryHOSTLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As HOSTENT, ByRef xSource As Long, ByVal nBytes As Long)
Private Declare Sub apiCopyMemoryLongLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nBytes As Long)
Private Declare Sub apiCopyMemoryStringLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef xDest As Long, ByRef xSource As Long, ByVal nBytes As Long)
Private Declare Function apiStrLen Lib "kernel32" Alias "lstrlenA" (ByRef lpString As Any) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hHandle As Long) As Long
Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function apiWaitForSingleObject Lib "kernel32" Alias "WaitForSingleObject" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiGetHostByName Lib "wsock32" Alias "gethostbyname" (ByVal HostName As String) As Long
Private Declare Function apiWSAStartup Lib "wsock32" Alias "WSAStartup" (ByVal wVersionRequired As Long, ByRef lpWSADATA As WSADATA) As Long
Private Declare Function apiWSACleanup Lib "wsock32" Alias "WSACleanup" () As Long
Private Declare Function apiInetAddr Lib "wsock32" Alias "inet_addr" (ByVal s As String) As Long
Private Declare Function apiinet_addr Lib "wsock32" Alias "inet_addr" (ByVal cP As String) As Long
Private Declare Function apiGetHostByAddr Lib "wsock32" Alias "gethostbyaddr" (ByRef haddr As Long, ByVal hnlen As Long, ByVal addrtype As Long) As Long
Private Declare Function apiGetNetworkParams Lib "iphlpapi" Alias "GetNetworkParams" (ByRef FixedInfo As Any, ByRef pOutBufLen As Long) As Long
Private Declare Function apiGetAdaptersInfo Lib "iphlpapi" Alias "GetAdaptersInfo" (ByRef IpAdapterInfo As Any, ByRef pOutBufLen As Long) As Long
Private Declare Function apiDoFileDownload Lib "shdocvw" Alias "DoFileDownload" (ByVal lpszFile As String) As Long
Private Declare Function apiInternetGetConnectedStateEx Lib "wininet" Alias "InternetGetConnectedStateEx" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Function apiGetRTTAndHopCount Lib "iphlpapi" Alias "GetRTTAndHopCount" (ByVal lDestIPAddr As Long, ByRef lHopCount As Long, ByVal lMaxHops As Long, ByRef lRTT As Long) As Long
Private ccon      As Boolean
Private sConnType As String * 255
Private an        As String
Private nt        As String
Private ire       As Boolean
Private wp        As Boolean
Private ed        As Boolean
Private Sub Class_Initialize()
    CheckInternetConnection
    InitializeSocket
    GetNetworkInfo
End Sub
Private Sub Class_Terminate()
    CloseSocket
End Sub
Friend Property Get NetworkAdapters() As NetworkAdapters
    Dim Netp As New NetworkAdapters
    Set NetworkAdapters = Netp
End Property
Friend Property Get ProxyEnable() As Boolean
    ProxyEnable = wp
End Property
Friend Property Get DnsEnable() As Boolean
    DnsEnable = ed
End Property
Friend Property Get NodeType() As String
    NodeType = nt
End Property
Friend Property Get HTTP() As HTTP
    Dim tmp As New HTTP
    Set HTTP = tmp
End Property
Friend Property Get FTP() As FTP
    Dim tmp As New FTP
    Set FTP = tmp
End Property
Friend Property Get IsWebConnected() As Boolean
    IsWebConnected = ccon
End Property
Friend Property Get ConnectionType() As String
    ConnectionType = sConnType
End Property
Friend Function Ping(sIPadr As String) As Boolean
    Dim lIPadr     As Long
    Dim lHopsCount As Long
    Dim lRTT       As Long
    Dim lMaxHops   As Long
    Dim lResult    As Long
    Const success = 1
    lMaxHops = 20
    lIPadr = apiinet_addr(sIPadr)
    Ping = (apiGetRTTAndHopCount(lIPadr, lHopsCount, lMaxHops, lRTT) = success)
End Function
Friend Sub DownloadFile(URL As String)
    Dim d As Long
    d = apiDoFileDownload(URL)
End Sub
Private Sub CheckInternetConnection()
    Dim ret As Long
    ret = apiInternetGetConnectedStateEx(ret, sConnType, 254, 0)
    ccon = CBool(ret)
End Sub
Friend Sub Refresh()
    CheckInternetConnection
    GetNetworkInfo
End Sub
Private Function InitializeSocket() As Boolean
    Dim WSAD As WSADATA 'attempt to initialize the socket
    InitializeSocket = apiWSAStartup(WS_VERSION_REQD, WSAD) = IP_SUCCESS
End Function
Private Sub CloseSocket()
    apiWSACleanup 'If apiWSACleanup() <> 0 Then MsgBox "Error calling apiWSACleanup.", vbCritical, ""
End Sub
Friend Function GetIPFromHostName(ByVal sHostName As String) As String
    Dim nBytes       As Long    'converts a host name to an IP address.
    Dim ptrHosent    As Long
    Dim hstHost      As HOSTENT
    Dim ptrName      As Long
    Dim ptrAddress   As Long
    Dim ptrIPAddress As Long
    Dim sAddress     As String 'declare this as Dim sAddress(1) As String if you want 2 ip addresses returned
    If InitializeSocket() = True Then
        ptrHosent = apiGetHostByName(sHostName & vbNullChar)
        If ptrHosent <> 0 Then
            apiCopyMemoryHOSTLong hstHost, ByVal ptrHosent, LenB(hstHost)
            apiCopyMemoryLongLong ptrIPAddress, ByVal hstHost.hAddrList, 4
            sAddress = Space(4)
            'if you want multiple domains returned,            'fill all items in sAddress array with 4 spaces
            apiCopyMemoryStringLong ByVal sAddress, ByVal ptrIPAddress, hstHost.hLength
            'change this to
            'CopyMemory ByVal sAddress(0), ByVal ptrIPAddress, hstHost.hLength if you want an array of ip addresses returned
            '(some domains have more than one ip address associated with it)  get the IP address
            GetIPFromHostName = IPToText(sAddress)
            'if you are using multiple addresses, you need IPToText(sAddress(0)) & "," & IPToText(sAddress(1)) etc
        End If
    Else
        MsgBox "Failed to open Socket.", vbExclamation, ""
    End If
End Function
Private Function IPToText(ByVal IPAddress As String) As String
    IPToText = CStr(Asc(IPAddress)) & "." & CStr(Asc(Mid(IPAddress, 2, 1))) & "." & CStr(Asc(Mid(IPAddress, 3, 1))) & "." & CStr(Asc(Mid(IPAddress, 4, 1)))
End Function
Friend Function GetHostNameFromIP(ByVal sIPAddress As String) As String
    Dim ptrHosent As Long
    Dim hAddress  As Long
    Dim sHost     As String
    Dim nBytes    As Long 'try to open the socket
    If InitializeSocket() = True Then 'convert string address to long datatype
        hAddress = apiInetAddr(sIPAddress)
        If hAddress <> SOCKET_ERROR Then 'check if an error ocucred
            'obtain a pointer to the HOSTENT structure that contains the name and address corresponding to the given network address.
            ptrHosent = apiGetHostByAddr(hAddress, 4, AF_INET)
            If ptrHosent <> 0 Then  'convert address and get resolved hostname
                apiCopyMemoryLongLong ptrHosent, ByVal ptrHosent, 4
                nBytes = apiStrLen(ByVal ptrHosent)
                If nBytes > 0 Then 'fill the IP address buffer
                    sHost = Space(nBytes)
                    apiCopyMemoryStringLong ByVal sHost, ByVal ptrHosent, nBytes
                    GetHostNameFromIP = sHost
                End If
            Else
                MsgBox "Call to Get Host by address failed.", vbExclamation, ""
            End If
            CloseSocket
        Else
            MsgBox "Invalid IP address", vbExclamation, ""
        End If
    Else
        MsgBox "Failed to open Socket", vbExclamation, ""
    End If
End Function
Private Sub GetNetworkInfo()
    Dim Error               As Long
    Dim FixedInfoSize       As Long
    Dim AdapterInfoSize     As Long
    Dim i                   As Integer
    Dim PhysicalAddress     As String
    Dim NewTime             As Date
    Dim AdapterInfo         As IP_ADAPTER_INFO
    Dim AddrStr             As IP_ADDR_STRING
    Dim FixedInfo           As FIXED_INFO
    Dim buffer              As IP_ADDR_STRING
    Dim pAddrStr            As Long
    Dim pAdapt              As Long
    Dim Buffer2             As IP_ADAPTER_INFO
    Dim FixedInfoBuffer()   As Byte
    Dim AdapterInfoBuffer() As Byte
    FixedInfoSize = 0 ' Get the main IP configuration information for this machine  using a FIXED_INFO structure.
    Error = apiGetNetworkParams(ByVal 0, FixedInfoSize)
    If Error <> 0 Then
        If Error <> ERROR_BUFFER_OVERFLOW Then
            MsgBox "GetNetworkParams sizing failed with error " & Error, vbExclamation, ""
            Exit Sub
        End If
    End If
    ReDim FixedInfoBuffer(FixedInfoSize - 1)
    Error = apiGetNetworkParams(FixedInfoBuffer(0), FixedInfoSize)
    Select Case FixedInfo.NodeType
        Case 1
            nt = "Broadcast"
        Case 2
            nt = "Peer to peer"
        Case 4
            nt = "Mixed"
        Case 8
            nt = "Hybrid"
        Case Else
            nt = "Unknown node type"
    End Select
    If FixedInfo.EnableRouting Then
        ire = True
    Else
        ire = False
    End If
    If FixedInfo.EnableProxy Then
        wp = True
    Else
        wp = False
    End If
    If FixedInfo.EnableDns Then
        ed = True
    Else
        ed = False
    End If
End Sub ' Enumerate all of the adapter specific information using the IP_ADAPTER_INFO structure. Note: IP_ADAPTER_INFO contains a linked list of adapter entries.

