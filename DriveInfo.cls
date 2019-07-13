VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "DriveInfo"
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
'Attribute VB_Name = "DriveInfo"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const FS_CASE_IS_PRESERVED         As Long = &H2
Private Const FS_CASE_SENSITIVE            As Long = &H1
Private Const FS_UNICODE_STORED_ON_DISK    As Long = &H4
Private Const FS_PERSISTENT_ACLS           As Long = &H8
Private Const FS_FILE_COMPRESSION          As Long = &H10
Private Const FS_VOLUME_IS_COMPRESSED      As Long = &H8000
Private Const FILE_NAMED_STREAMS           As Long = &H40000
Private Const FILE_SUPPORTS_ENCRYPTION     As Long = &H20000
Private Const FILE_SUPPORTS_OBJECT_IDS     As Long = &H10000
Private Const FILE_SUPPORTS_REPARSE_POINTS As Long = &H80
Private Const FILE_SUPPORTS_SPARSE_FILES   As Long = &H40
Private Const FILE_VOLUME_QUOTAS           As Long = &H20
Private Const OPEN_EXISTING                As Long = 3
Private Const FILE_SHARE_READ              As Long = &H1
Private Const FILE_SHARE_WRITE             As Long = &H2
Private Const GENERIC_READ                 As Long = &H80000000
Private Const IOCTL_STORAGE_QUERY_PROPERTY As Long = &H2D1400
Private Const DRIVE_REMOVABLE              As Long = 2
Private Const DRIVE_FIXED                  As Long = 3
Private Const DRIVE_REMOTE                 As Long = 4
Private Const DRIVE_CDROM                  As Long = 5
Private Const DRIVE_RAMDISK                As Long = 6
Private Const RESOURCETYPE_ANY             As Long = &H0
Private Const RESOURCE_CONNECTED           As Long = &H1
Private Type STORAGE_PROPERTY_QUERY
    PropertyId                              As STORAGE_PROPERTY_ID
    QueryType                               As STORAGE_QUERY_TYPE
    AdditionalParameters                    As Byte
End Type
Private Type DEVICE_INFORMATION
    Valid                                   As Boolean
    BusType                                 As STORAGE_BUS_TYPE
    Removable                               As Boolean
    VendorID                                As String
    ProductID                               As String
    ProductRevision                         As String
End Type
Private Type STORAGE_DEVICE_DESCRIPTOR
    Version                                 As Long
    Size                                    As Long
    DeviceType                              As Byte
    DeviceTypeModifier                      As Byte
    RemovableMedia                          As Byte
    CommandQueueing                         As Byte
    VendorIdOffset                          As Long
    ProductIdOffset                         As Long
    ProductRevisionOffset                   As Long
    SerialNumberOffset                      As Long
    BusType                                 As Integer
    RawPropertiesLength                     As Long
    RawDeviceProperties                     As Byte
End Type
Public Enum STORAGE_BUS_TYPE
    BusTypeUnknown = 0
    BusTypeScsi
    BusTypeAtapi
    BusTypeAta
    BusType1394
    BusTypeSsa
    BusTypeFibre
    BusTypeUsb
    BusTypeRAID
    BusTypeMaxReserved = &H7F
End Enum
Private Enum STORAGE_PROPERTY_ID
    StorageDeviceProperty = 0
    StorageAdapterProperty
    StorageDeviceIdProperty
End Enum
Private Enum STORAGE_QUERY_TYPE
    PropertyStandardQuery = 0
    PropertyExistsQuery
    PropertyMaskQuery
    PropertyQueryMaxDefined
End Enum
Private Enum DRIVE_STATUS
    DRIVE_DOESNT_EXIST = 1
    DRIVE_NOT_READY = 2
    DRIVE_READY = 3
End Enum
Private Type NETRESOURCE
    dwScope As Long
    dwType As Long
    dwDisplayType As Long
    dwUsage As Long
    lpLocalName As Long
    lpRemoteName As Long
    lpComment As Long
    lpProvider As Long
End Type
Private Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Declare Sub apiCpyMem Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDst As Any, ByRef pSrc As Any, ByVal cb As Long)
Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function apiDeviceIoControl Lib "kernel32" (ByVal hDevice As Long, ByVal dwIoControlCode As Long, ByRef lpInBuffer As Any, ByVal nInBufferSize As Long, ByRef lpOutBuffer As Any, ByVal nOutBufferSize As Long, ByRef lpBytesReturned As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function apiCloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function apiGetLogicalDrives Lib "kernel32" Alias "GetLogicalDrives" () As Long
Private Declare Function apiGetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiGetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, ByRef lpSectorsPerCluster As Long, ByRef lpBytesPerSector As Long, ByRef lpNumberOfFreeClusters As Long, ByRef lpTotalNumberOfClusters As Long) As Long
Private Declare Function apiGetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToCaller As ULARGE_INTEGER, ByRef lpTotalNumberOfBytes As ULARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Private Declare Function apiGetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function apiSetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function apiWNetOpenEnum Lib "mpr" Alias "WNetOpenEnumA" (ByVal dwScope As Long, ByVal dwType As Long, ByVal dwUsage As Long, ByRef lpNetResource As Any, ByRef lphEnum As Long) As Long
Private Declare Function apiWNetEnumResource Lib "mpr" Alias "WNetEnumResourceA" (ByVal hEnum As Long, ByRef lpcCount As Long, ByRef lpBuffer As Any, ByRef lpBufferSize As Long) As Long
Private Declare Function apiWNetCloseEnum Lib "mpr" (ByVal hEnum As Long) As Long
Public IsReady            As Boolean
Public dType              As Variant
Public AvailableFreeSpace As Long
Public Name               As String
Public DriveFormat        As String
Public RootDirectory      As String
Public TotalFreeSpace     As Long
Public TotalSize          As Long
Public VolumeLabel        As String
Dim Root                  As String
Dim volume_name           As String
Dim serial_number         As Long
Dim max_component_length  As Long
Dim file_system_flags     As Long
Dim file_system_name      As String
Dim drive_type            As String
Dim pos                   As Integer
Dim Capacity              As String
Dim dused                 As String
Dim dfree                 As String
Dim pcapacity             As String
Dim pdused                As String
Dim pdfree                As String
Dim h_name                As String
Dim bus_type              As String
Dim p_ver                 As String
Dim rem_                  As String
Dim cDrive                As String
Dim cReady                As Boolean
Dim DevInfo               As DEVICE_INFORMATION
Private Enum DriveType
    CDRom
    Unknown
End Enum
'Private mvarIsReady As Boolean
'Friend Property Let IsReady(ByVal vData As String)
'    mvarIsReady = vData
'End Property
'Friend Property Get IsReady() As String
'    IsReady = mvarIsReady
'End Property
'driveinfo class
'Dim di As IO.DriveInfo
'di.GetDrives() ' driveinfo()
'di.DriveType= IO.DriveType.CDRom
'di.AvailableFreeSpace =long
'di.DriveFormat fat32 ntfs
'di.IsReady 'boolean
'di.Name= string
'di.RootDirectory
'di.TotalFreeSpace  = long
'di.TotalSize ='long
'di.VolumeLabel = string
Friend Function GetDrives() As DriveInfo()
    'getall all drives and info
End Function
Friend Sub GetLogicalDriveStrings()
    Dim strSave As String
    strSave = String(260, Chr(0))
    Call apiGetLogicalDriveStrings(260, strSave)
    Dim keer As Long
    For keer = 1 To 100
        If Left(strSave, InStr(1, strSave, Chr(0))) = Chr(0) Then Exit For
        MsgBox Left(strSave, InStr(1, strSave, Chr(0)) - 1)
        strSave = Right(strSave, Len(strSave) - InStr(1, strSave, Chr(0)))
    Next
End Sub
Friend Function GetLogicalDrives() As String()
    Dim ret    As Long
    Dim i      As Long
    Dim drv    As String
    Dim drvs() As String
    ret = apiGetLogicalDrives
    Do
        If (ret And 2 ^ i) <> 0 Then drv = drv + " " + Chr(65 + i)
        ReDim Preserve drvs(i)
        drvs(i) = drv
        If i = 25 Then Exit Do
        i = i + 1
    Loop
    GetLogicalDrives = drvs
End Function
Friend Function SetVolumeLabel(ByVal drivePath As String, ByVal newDriveName As String)
    Dim ret As Long
    ret = apiSetVolumeLabel(drivePath, newDriveName) '"C:\", "new name here"
End Function
Friend Function GetVolumeInformation(ByVal drivePath As String)
    Dim volname    As String
    Dim sn         As Long
    Dim snstr      As String
    Dim maxcomplen As Long
    Dim sysflags   As Long
    Dim sysname    As String
    Dim ret        As Long
    volname = Space(260)
    sysname = Space(260)
    ret = apiGetVolumeInformation("C:\", volname, Len(volname), sn, maxcomplen, sysflags, sysname, Len(sysname))
    volname = Left(volname, InStr(volname, vbNullChar) - 1)
    sysname = Left(sysname, InStr(sysname, vbNullChar) - 1)
    snstr = Trim(Hex(sn))
    snstr = String(8 - Len(snstr), "0") & snstr
    snstr = Left(snstr, 4) & "-" & Right(snstr, 4)
    Dim txt As String
    txt = txt & "Volume Name: " & volname & vbCrLf
    txt = txt & "Serial Number: " & snstr & vbCrLf
    txt = txt & "File System: " & sysname & vbCrLf
    MsgBox txt
End Function
'----------------------------
Private Function GetShareName(DriveLetter As String) As String
    Dim hEnum         As Long
    Dim NetInfo(1023) As NETRESOURCE
    Dim entries       As Long
    Dim nStatus       As Long
    Dim LocalName     As String
    Dim UNCName       As String
    Dim i             As Long
    Dim r             As Long
    On Error GoTo xc
    DriveLetter = Left(DriveLetter, 2)
    nStatus = apiWNetOpenEnum(RESOURCE_CONNECTED, RESOURCETYPE_ANY, 0, ByVal 0, hEnum)
    GetShareName = ""
    If ((nStatus = 0) And (hEnum <> 0)) Then
        entries = 1024
        nStatus = apiWNetEnumResource(hEnum, entries, NetInfo(0), CLng(Len(NetInfo(0))) * 1024)
        If nStatus = 0 Then
            For i = 0 To entries - 1
                LocalName = ""
                If NetInfo(i).lpLocalName <> 0 Then
                    LocalName = Space(apilstrlen(NetInfo(i).lpLocalName) + 1)
                    r = apilstrcpy(LocalName, NetInfo(i).lpLocalName)
                End If
                If Len(LocalName) <> 0 Then LocalName = Left(LocalName, (Len(LocalName) - 1))
                If UCase(LocalName) = UCase(DriveLetter) Then
                    UNCName = ""
                    If NetInfo(i).lpRemoteName <> 0 Then
                        UNCName = Space(apilstrlen(NetInfo(i).lpRemoteName) + 1)
                        r = apilstrcpy(UNCName, NetInfo(i).lpRemoteName)
                    End If
                    If Len(UNCName) <> 0 Then UNCName = Left(UNCName, (Len(UNCName) - 1))
                    GetShareName = UNCName
                    Exit For
                End If
            Next i
        End If
    End If
    nStatus = apiWNetCloseEnum(hEnum)
xc:
End Function
Private Function GetDevInfo(ByVal strDrive As String) As DEVICE_INFORMATION
    Dim hDrive         As Long
    Dim udtQuery       As STORAGE_PROPERTY_QUERY
    Dim dwOutBytes     As Long
    Dim lngResult      As Long
    Dim btBuffer(9999) As Byte
    Dim udtOut         As STORAGE_DEVICE_DESCRIPTOR
    hDrive = apiCreateFile("\\.\" & Left(strDrive, 1) & ":", 0, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal 0, OPEN_EXISTING, 0, 0)
    If hDrive = -1 Then Exit Function
    With udtQuery
        .PropertyId = StorageDeviceProperty
        .QueryType = PropertyStandardQuery
    End With
    lngResult = apiDeviceIoControl(hDrive, IOCTL_STORAGE_QUERY_PROPERTY, udtQuery, LenB(udtQuery), btBuffer(0), UBound(btBuffer) + 1, dwOutBytes, ByVal 0)
    If lngResult Then
        apiCpyMem udtOut, btBuffer(0), Len(udtOut)
        With GetDevInfo
            .Valid = True
            .BusType = udtOut.BusType
            .Removable = CBool(udtOut.RemovableMedia)
            If udtOut.ProductIdOffset > 0 Then _
               .ProductID = StringCopy(VarPtr(btBuffer(udtOut.ProductIdOffset)))
            If udtOut.ProductRevisionOffset > 0 Then _
               .ProductRevision = StringCopy(VarPtr(btBuffer(udtOut.ProductRevisionOffset)))
            If udtOut.VendorIdOffset > 0 Then
                .VendorID = StringCopy(VarPtr(btBuffer(udtOut.VendorIdOffset)))
            End If
        End With
    Else
        GetDevInfo.Valid = False
    End If
    apiCloseHandle hDrive
End Function
Private Function StringCopy(ByVal pBuffer As Long) As String
    Dim tmp As String
    tmp = Space(apilstrlen(ByVal pBuffer))
    apilstrcpy ByVal tmp, ByVal pBuffer
    StringCopy = Trim(tmp)
End Function
Private Sub GetDriveInformation(GetDriveLetter As String)
    On Error GoTo xc
    Dim lDriveType         As Long, lFreeSpace As Long
    Dim lSectorsPerCluster As Long, lBytesPerSector As Long
    Dim lFreeClusters      As Long, lTotalClusters As Long
    If Len(GetDriveLetter) <> 3 Then
        GetDriveLetter = Left(GetDriveLetter, 1) & ":\"
    End If
    cDrive = GetDriveLetter
    Root = GetDriveLetter
    volume_name = Space(1024)
    file_system_name = Space(1024)
    Call apiGetVolumeInformation(Root, volume_name, Len(volume_name), serial_number, max_component_length, file_system_flags, file_system_name, Len(file_system_name))
    pos = InStr(volume_name, Chr(0))
    volume_name = Left(volume_name, pos - 1)
    volume_name = Trim(volume_name)
    pos = InStr(file_system_name, Chr(0))
    file_system_name = Left(file_system_name, pos - 1)
    file_system_name = Trim(file_system_name)
    Select Case apiGetDriveType(GetDriveLetter)
        Case DRIVE_REMOVABLE
            drive_type = "Removable Drive"
        Case DRIVE_FIXED
            drive_type = "Fixed Drive"
        Case DRIVE_REMOTE
            drive_type = "Remote Drive"
        Case DRIVE_CDROM
            drive_type = "Optical Drive"
        Case DRIVE_RAMDISK
            drive_type = "RAM Disk"
    End Select
    lDriveType = apiGetDriveType(GetDriveLetter)
    If lDriveType = 1 Then
        cReady = False
    Else
        lFreeSpace = apiGetDiskFreeSpace(GetDriveLetter, lSectorsPerCluster, lBytesPerSector, lFreeClusters, lTotalClusters)
        cReady = IIf(Err.LastDllError > 0, False, True)
    End If
    If GetDevInfo(GetDriveLetter).Valid Then
        Select Case GetDevInfo(GetDriveLetter).BusType
            Case BusTypeUsb:        bus_type = "USB"
            Case BusType1394:       bus_type = "1394"
            Case BusTypeAta:        bus_type = "ATA"
            Case BusTypeAtapi:      bus_type = "ATAPI"
            Case BusTypeFibre:      bus_type = "Fibre"
            Case BusTypeRAID:       bus_type = "RAID"
            Case BusTypeScsi:       bus_type = "SCSI"
            Case BusTypeSsa:        bus_type = "SSA"
            Case BusTypeUnknown:    bus_type = "Unknown"
        End Select
        h_name = Trim(GetDevInfo(GetDriveLetter).ProductID)
        p_ver = GetDevInfo(GetDriveLetter).ProductRevision
        rem_ = GetDevInfo(GetDriveLetter).Removable
    End If
    GetDiskSpace GetDriveLetter
xc:
End Sub
'di.GetDrives() ' driveinfo()
'di.DriveType= IO.DriveType.CDRom
'di.AvailableFreeSpace =long
'di.DriveFormat fat32 ntfs
'di.IsReady 'boolean
'di.Name= string
'di.RootDirectory
'di.TotalFreeSpace  = long
'di.TotalSize ='long
'di.VolumeLabel = string
Property Get DriveLetter() As String
    DriveLetter = Root
End Property
Property Get ShareName() As String
    ShareName = GetShareName(Root)
End Property
Property Get DriveName() As String
    DriveName = volume_name
End Property
Property Get SerialNumber() As String
    SerialNumber = serial_number
End Property
Property Get FileSystem() As String
    FileSystem = file_system_name
End Property
Property Get DriveType() As String
    DriveType = drive_type
End Property
Property Get BusType() As String
    BusType = bus_type
End Property
Property Get HardwareName() As String
    HardwareName = h_name
End Property
Property Get ProductVersion() As String
    ProductVersion = p_ver
End Property
Property Get IsRemovable() As Boolean
    IsRemovable = rem_
End Property
Property Get DiskCapacity() As String
    DiskCapacity = Capacity
End Property
Property Get FreeDiskSpace() As String
    FreeDiskSpace = dfree
End Property
Property Get UsedDiskSpace() As String
    UsedDiskSpace = dused
End Property
Property Get FreeDiskSpacePercent() As String
    FreeDiskSpacePercent = pdfree
End Property
Property Get UsedDiskSpacePercent() As String
    UsedDiskSpacePercent = pdused
End Property
Private Sub GetDiskSpace(drive As String)
    Dim bytes_avail As ULARGE_INTEGER
    Dim bytes_total As ULARGE_INTEGER
    Dim bytes_free  As ULARGE_INTEGER
    Dim dbl_total   As Double
    Dim dbl_free    As Double
    apiGetDiskFreeSpaceEx drive, bytes_avail, bytes_total, bytes_free
    dbl_total = LargeIntegerToDouble(bytes_total.LowPart, bytes_total.HighPart)
    dbl_free = LargeIntegerToDouble(bytes_free.LowPart, bytes_free.HighPart)
    Capacity = FormatBytes(dbl_total)
    dfree = FormatBytes(dbl_free)
    dused = FormatBytes(dbl_total - dbl_free)
    pdfree = Format(dbl_free / dbl_total, "percent")
    pdused = Format((dbl_total - dbl_free) / dbl_total, "percent")
End Sub
Private Function FormatBytes(ByVal num_bytes As Double) As String
    Const ONE_KB As Double = 1024
    Const ONE_MB As Double = ONE_KB * 1024
    Const ONE_GB As Double = ONE_MB * 1024
    Const ONE_TB As Double = ONE_GB * 1024
    Const ONE_PB As Double = ONE_TB * 1024
    Const ONE_EB As Double = ONE_PB * 1024
    Const ONE_ZB As Double = ONE_EB * 1024
    Const ONE_YB As Double = ONE_ZB * 1024
    Dim Value    As Double
    Dim txt      As String
    If num_bytes <= 999 Then
        FormatBytes = Format(num_bytes, "0") & " bytes"
    ElseIf num_bytes <= ONE_KB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_KB) & " KB"
    ElseIf num_bytes <= ONE_MB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_MB) & " MB"
    ElseIf num_bytes <= ONE_GB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_GB) & " GB"
    ElseIf num_bytes <= ONE_TB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_TB) & " TB"
    ElseIf num_bytes <= ONE_PB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_PB) & " PB"
    ElseIf num_bytes <= ONE_EB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_EB) & " EB"
    ElseIf num_bytes <= ONE_ZB * 999 Then
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_ZB) & " ZB"
    Else
        FormatBytes = ThreeNonZeroDigits(num_bytes / ONE_YB) & " YB"
    End If
End Function
Private Function ThreeNonZeroDigits(ByVal Value As Double) As String
    If Value >= 100 Then
        ThreeNonZeroDigits = Format(CInt(Value))
    ElseIf Value >= 10 Then
        ThreeNonZeroDigits = Format(Value, "0.0")
    Else
        ThreeNonZeroDigits = Format(Value, "0.00")
    End If
End Function
Private Function LargeIntegerToDouble(low_part As Long, high_part As Long) As Double
    Dim result As Double
    result = high_part
    If high_part < 0 Then result = result + 2 ^ 32
    result = result * 2 ^ 32
    result = result + low_part
    If low_part < 0 Then result = result + 2 ^ 32
    LargeIntegerToDouble = result
End Function
Sub Refresh()
    GetDriveInformation cDrive
End Sub
Property Get Ready() As Boolean
    Ready = cReady
End Property


