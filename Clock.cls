VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
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
Private Type TIME_ZONE_INFORMATION
    Bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Declare Function apiCompareFileTime Lib "kernel32" Alias "CompareFileTime" (ByRef lpFileTime1 As FILETIME, ByRef lpFileTime2 As FILETIME) As Long
Private Declare Function apiGetLocalTime Lib "kernel32" Alias "GetLocalTime" (ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function apiGetSystemTime Lib "kernel32" Alias "GetSystemTime" (ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function apiSetSystemTime Lib "kernel32" Alias "SetSystemTime" (ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function apiSetSystemTimeAdjustment Lib "kernel32" Alias "SetSystemTimeAdjustment" (ByVal dwTimeAdjustment As Long, ByVal bTimeAdjustmentDisabled As Boolean) As Long
Private Declare Function apiFileTimeToLocalFileTime Lib "kernel32" Alias "FileTimeToLocalFileTime" (ByRef lpFileTime As FILETIME, ByRef lpLocalFileTime As FILETIME) As Long
Private Declare Function apiFileTimeToSystemTime Lib "kernel32" Alias "FileTimeToSystemTime" (ByRef lpFileTime As FILETIME, ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function apiGetSystemTimeAsFileTime Lib "kernel32" Alias "GetSystemTimeAsFileTime" (ByRef lpSystemTimeAsFileTime As FILETIME) As Long
Private Declare Function apiGetTimeZoneInformation Lib "kernel32" Alias "GetTimeZoneInformation" (ByRef lpTimeZoneInformation As TIME_ZONE_INFORMATION) As Long
Private Declare Function apiLocalFileTimeToFileTime Lib "kernel32" Alias "LocalFileTimeToFileTime" (ByRef lpLocalFileTime As FILETIME, ByRef lpFileTime As FILETIME) As Long
Private Declare Function apiSystemTimeToFileTime Lib "kernel32" Alias "SystemTimeToFileTime" (ByRef lpSystemTime As SYSTEMTIME, ByRef lpFileTime As FILETIME) As Long
'Private Declare Function apitimeGetSystemTime Lib "winmm" Alias "timeGetSystemTime" (ByRef lpTime As MMTIME, ByVal uSize As Long) As Long
'Private Declare Function apitimeGetTime Lib "winmm" Alias "timeGetTime" () As Long
Friend Function localtime() As SYSTEMTIME
    Dim lpSystemTime As SYSTEMTIME
    Call apiGetLocalTime(lpSystemTime)
    localtime = lpSystemTime
End Function
Friend Function gmtTime() As SYSTEMTIME
    Dim lpSystemTime As SYSTEMTIME
    Call apiGetSystemTime(lpSystemTime)
    gmtTime = lpSystemTime
End Function
Friend Function TickCount() As Long
    TickCount = apiGetTickCount
End Function
