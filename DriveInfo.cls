VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DriveInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const FS_CASE_IS_PRESERVED = &H2
Private Const FS_CASE_SENSITIVE = &H1
Private Const FS_UNICODE_STORED_ON_DISK = &H4
Private Const FS_PERSISTENT_ACLS = &H8
Private Const FS_FILE_COMPRESSION = &H10
Private Const FS_VOLUME_IS_COMPRESSED = &H8000
Private Const FILE_NAMED_STREAMS = &H40000
Private Const FILE_SUPPORTS_ENCRYPTION = &H20000
Private Const FILE_SUPPORTS_OBJECT_IDS = &H10000
Private Const FILE_SUPPORTS_REPARSE_POINTS = &H80
Private Const FILE_SUPPORTS_SPARSE_FILES = &H40
Private Const FILE_VOLUME_QUOTAS = &H20
Private Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type
Private Declare Function apiGetLogicalDrives Lib "kernel32" Alias "GetLogicalDrives" () As Long
Private Declare Function apiGetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function apiGetDiskFreeSpace Lib "kernel32.dll" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, ByRef lpSectorsPerCluster As Long, ByRef lpBytesPerSector As Long, ByRef lpNumberOfFreeClusters As Long, ByRef lpTotalNumberOfClusters As Long) As Long
Private Declare Function apiGetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpDirectoryName As String, ByRef lpFreeBytesAvailableToCaller As ULARGE_INTEGER, ByRef lpTotalNumberOfBytes As ULARGE_INTEGER, ByRef lpTotalNumberOfFreeBytes As ULARGE_INTEGER) As Long
Private Declare Function apiGetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Private Declare Function apiSetVolumeLabel Lib "kernel32" Alias "SetVolumeLabelA" (ByVal lpRootPathName As String, ByVal lpVolumeName As String) As Long
Private Declare Function apiGetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

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

Friend Function GetLogicalDrives()
    Dim ret As Long
    Dim i As Long
    Dim drvs As String
    ret = apiGetLogicalDrives
    Do
       If (ret And 2 ^ i) <> 0 Then drvs = drvs + " " + Chr(65 + i)
       If i = 25 Then Exit Do
       i = i + 1
    Loop
    MsgBox drvs
End Function
Friend Function SetVolumeLabel(ByVal drivePath As String, ByVal newDriveName As String)
   Dim ret As Long
  ret = apiSetVolumeLabel(drivePath, newDriveName) '"C:\", "new name here"
End Function
Friend Function GetVolumeInformation(ByVal drivePath As String)
    Dim volname    As String   ' receives volume name of C:
    Dim sn         As Long          ' receives serial number of C:
    Dim snstr      As String     ' display form of serial number
    Dim maxcomplen As Long  ' receives maximum component length
    Dim sysflags   As Long    ' receives file system flags
    Dim sysname    As String   ' receives the file system name
    Dim ret        As Long      ' return value
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

