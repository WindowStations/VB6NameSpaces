VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Bluetooth"
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
'Attribute VB_Name = "Bluetooth"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const BLUETOOTH_MAX_NAME_SIZE = 248
Private Const ERROR_NO_MORE_ITEMS     As Long = 259
Private Const ERROR_REVISION_MISMATCH As Long = 1306
Private Const WIN32_NULL              As Long = 0
Private Const WIN32_FALSE             As Long = 0
Private Const WIN32_TRUE              As Long = Not WIN32_FALSE 'Not to be used for testing for True.
Private Const WIN32_EPOCH             As Date = #1/1/1601#
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
Private Type BLUETOOTH_ADDRESS
    rgBytes(0 To 5) As Byte
End Type
Private Type BLUETOOTH_DEVICE_INFO
    dwSize As Long
    UndocumentedPadding As Long
    Address As BLUETOOTH_ADDRESS 'VB6 will add 2 bytes of padding after this.
    ulClassofDevice As Long      'Bluetooth "Class of Device."
    fConnected As Long
    fRemembered As Long
    fAuthenticated As Long       'Device authenticated/paired/bonded.
    stLastSeen As SYSTEMTIME     'Last time the device was seen.
    stLastUsed As SYSTEMTIME     'Last time the device was used for other than RNR, inquiry, or SDP.
    szName(1 To (BLUETOOTH_MAX_NAME_SIZE + 1) * 2 - 4) As Byte
End Type
Private Type BLUETOOTH_DEVICE_SEARCH_PARAMS
    dwSize As Long               'Size, in bytes, of the structure.
    fReturnAuthenticated As Long 'Search should return authenticated Bluetooth devices.
    fReturnRemembered As Long    'Search should return remembered Bluetooth devices.
    fReturnUnknown As Long       'Search should return unknown Bluetooth devices.
    fReturnConnected As Long     'Search should return connected Bluetooth devices.
    fIssueInquiry As Long        'New inquiry should be issued.
    cTimeoutMultiplier As Byte   'Time out for the inquiry in increments of 1.28 seconds.
    'The maximum value is 48. When greater than 48 is used,the calling function immediately fails and returns E_INVALIDARG.
    hRadio As Long               'NULL to perform the inquiry on all local Bluetooth radios.
End Type
Private Type TIME_ZONE_INFORMATION
    bias As Long
    StandardName(0 To 31) As Integer
    StandardDate(0 To 7) As Integer
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate(0 To 7) As Integer
    DaylightBias As Long
End Type
Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Function apiGetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" (ByRef TimeZoneInformation As TIME_ZONE_INFORMATION) As TIME_ZONE_IDS
Private Declare Function apiBluetoothFindFirstDevice Lib "bthprops.cpl" Alias "BluetoothFindFirstDevice" (ByRef btsp As BLUETOOTH_DEVICE_SEARCH_PARAMS, ByRef btdi As BLUETOOTH_DEVICE_INFO) As Long
Private Declare Function apiBluetoothFindNextDevice Lib "bthprops.cpl" Alias "BluetoothFindNextDevice" (ByVal hFind As Long, ByRef btdi As BLUETOOTH_DEVICE_INFO) As Long
Private Declare Function apiBluetoothFindDeviceClose Lib "bthprops.cpl" Alias "BluetoothFindDeviceClose" (ByVal hFind As Long) As Long
Private Declare Function apiSystemTimeToVariantTime Lib "oleaut32" Alias "SystemTimeToVariantTime" (ByRef SYSTEMTIME As SYSTEMTIME, ByRef vtime As Date) As Long
Private Items              As New Collection
Private CurrentLocalBias   As Single
Private mTimeoutMultiplier As Byte
Private Enum TIME_ZONE_IDS
    TIME_ZONE_ID_UNKNOWN = 0
    TIME_ZONE_ID_STANDARD = 1
    TIME_ZONE_ID_DAYLIGHT = 2
End Enum
Private Sub Class_Initialize()
    TimeoutMultiplier = 5
End Sub
Friend Property Get Count() As Long
    Count = Items.Count
End Property
Friend Property Get Item(ByVal Index As Long) As BluetoothInfo
    Set Item = Items.Item(Index)
End Property
Friend Property Get TimeoutMultiplier() As Byte
    TimeoutMultiplier = mTimeoutMultiplier
End Property
Friend Property Let TimeoutMultiplier(ByVal RHS As Byte)
    If RHS > 48 Then Err.Raise 5, TypeName(Me)
    mTimeoutMultiplier = RHS
End Property
'Public Function NewEnum() As IUnknown
'    Set NewEnum = Items.[_NewEnum]
'End Function
Friend Sub Search()
    Dim tzi          As TIME_ZONE_INFORMATION
    Dim btsp         As BLUETOOTH_DEVICE_SEARCH_PARAMS
    Dim btdi         As BLUETOOTH_DEVICE_INFO
    Dim hFind        As Long
    Dim LastDllError As Long
    If apiGetTimeZoneInformation(tzi) < TIME_ZONE_ID_DAYLIGHT Then
        CurrentLocalBias = CSng(tzi.bias + tzi.StandardBias) / 60#
    Else
        CurrentLocalBias = CSng(tzi.bias + tzi.DaylightBias) / 60#
    End If
    With btsp
        .dwSize = LenB(btsp)
        .fReturnAuthenticated = WIN32_TRUE
        .fReturnUnknown = WIN32_TRUE
        .fReturnConnected = WIN32_TRUE
        .fIssueInquiry = WIN32_TRUE
        .cTimeoutMultiplier = TimeoutMultiplier
        '.hRadio = WIN32_NULL 'Defaults to this when allocated or zeroed.
    End With
    btdi.dwSize = LenB(btdi)
    Set Items = New Collection
    hFind = apiBluetoothFindFirstDevice(btsp, btdi)
    If hFind = WIN32_NULL Then
        LastDllError = Err.LastDllError
        If LastDllError <> ERROR_NO_MORE_ITEMS Then Err.Raise &H80047100, TypeName(Me), "apiBluetoothFindFirstDevice error " & CStr(LastDllError)
    Else
        AddItem btdi
        Do
            If apiBluetoothFindNextDevice(hFind, btdi) = WIN32_NULL Then
                LastDllError = Err.LastDllError
                If LastDllError <> ERROR_NO_MORE_ITEMS Then Err.Raise &H80047104, TypeName(Me), "BluetoothFindNextDevice error " & CStr(LastDllError)
                Exit Do
            End If
            AddItem btdi
        Loop
    End If
    Call apiBluetoothFindDeviceClose(hFind)
End Sub
Private Sub AddItem(ByRef btdi As BLUETOOTH_DEVICE_INFO)
    Dim i             As Long
    Dim Address       As String
    Dim dtLastSeen    As Date
    Dim dtLastUsed    As Date
    Dim LastSeen      As Variant
    Dim LastUsed      As Variant
    Dim Length        As Long
    Dim BluetoothInfo As BluetoothInfo
    With btdi
        Address = "00:00:00:00:00:00"
        For i = 0 To 5
            Mid(Address, i * 3 + 1, 2) = Right("0" & Hex(.Address.rgBytes(i)), 2)
        Next
        If apiSystemTimeToVariantTime(.stLastSeen, dtLastSeen) = 0 Then
            LastSeen = Null
        ElseIf dtLastSeen = WIN32_EPOCH Then
            LastSeen = Null
        Else
            LastSeen = DateAdd("h", -CurrentLocalBias, dtLastSeen)
        End If
        If apiSystemTimeToVariantTime(.stLastUsed, dtLastUsed) = 0 Then
            LastUsed = Null
        ElseIf dtLastUsed = WIN32_EPOCH Then
            LastUsed = Null
        Else
            LastUsed = dtLastUsed
        End If
        Length = apilstrlen(VarPtr(.szName(1)))
        Set BluetoothInfo = New BluetoothInfo
        BluetoothInfo.Assign Address, .ulClassofDevice, .fConnected, .fRemembered, .fAuthenticated, LastSeen, LastUsed, Left(.szName, Length)
        Items.Add BluetoothInfo
    End With
End Sub

