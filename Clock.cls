VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Clock"
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
'Attribute VB_Name = "Clock"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
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
    bias As Long
    StandardName(0 To 31) As Integer
    StandardDate As SYSTEMTIME
    StandardBias As Long
    DaylightName(0 To 31) As Integer
    DaylightDate As SYSTEMTIME
    DaylightBias As Long
End Type
Private Type MMTIME
    wType As Long
    u As Long
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
Private Declare Function apiGetDateFormat Lib "kernel32" Alias "GetDateFormatA" (ByVal Locale As Long, ByVal dwFlags As Long, ByRef lpDate As SYSTEMTIME, ByVal lpFormat As String, ByVal lpDateStr As String, ByVal cchDate As Long) As Long
Private Declare Function apiEnumDateFormats Lib "kernel32" Alias "EnumDateFormatsA" (ByVal lpDateFmtEnumProc As Long, ByVal Locale As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiDosDateTimeToFileTime Lib "kernel32" Alias "DosDateTimeToFileTime" (ByVal wFatDate As Long, ByVal wFatTime As Long, ByRef lpFileTime As FILETIME) As Long
Private Declare Function apiFileTimeToDosDateTime Lib "kernel32" Alias "FileTimeToDosDateTime" (ByRef lpFileTime As FILETIME, ByVal lpFatDate As Long, ByVal lpFatTime As Long) As Long
Private Declare Function apitimeGetSystemTime Lib "winmm" Alias "timeGetSystemTime" (ByRef lpTime As MMTIME, ByVal uSize As Long) As Long
Private Declare Function apitimeGetTime Lib "winmm" Alias "timeGetTime" () As Long
Private Declare Function apitimeBeginPeriod Lib "winmm" Alias "timeBeginPeriod" (ByVal uPeriod As Long) As Long
Private Declare Function apitimeEndPeriod Lib "winmm" Alias "timeEndPeriod" (ByVal uPeriod As Long) As Long
Private Declare Function apiVarMonthName Lib "oleaut32" Alias "VarMonthName" (ByVal lMonth As Long, ByVal fAddrev As Long, ByVal dwFlags As Long, ByRef pbstrOut As String) As Long
Private Declare Function apiVarWeekdayName Lib "oleaut32" Alias "VarWeekdayName" (ByVal lWeekDay As Long, ByVal fAddrev As Long, ByVal lFirstDay As Long, ByVal dwFlags As Long, ByRef pbstrOut As String) As Long
Private Declare Sub apiVarBoolFromDate Lib "oleaut32" Alias "VarBoolFromDate" (ByVal dateIn As Double, ByRef pboolOut As Integer)
Private Declare Sub apiVarBstrFromDate Lib "oleaut32" Alias "VarBstrFromDate" (ByVal dateIn As Double, ByVal lcid As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Private Declare Sub apiVarDateFromBool Lib "oleaut32" Alias "VarDateFromBool" (ByVal boolIn As Integer, ByRef pdateOut As Double)
Private Declare Sub apiVarDateFromDec Lib "oleaut32" Alias "VarDateFromDec" (ByRef pdecIn As Single, ByRef pdateOut As Double)
Private Declare Sub apiVarDateFromStr Lib "oleaut32" Alias "VarDateFromStr" (ByRef strIn As Byte, ByVal lcid As Long, ByVal dwFlags As Long, ByRef pdateOut As Double)
Private Declare Sub apiVarDecFromDate Lib "oleaut32" Alias "VarDecFromDate" (ByVal dateIn As Double, ByRef pdecOut As Variant)
Private Declare Sub apiVarFormatDateTime Lib "oleaut32" Alias "VarFormatDateTime" (ByRef pvarIn As Variant, ByValamedFormat As Long, ByVal dwFlags As Long, ByVal pbstrOut As Long)
Private Declare Function apiVariantTimeToDosDateTime Lib "oleaut32" Alias "VariantTimeToDosDateTime" (ByVal vtime As Double, ByRef pwDosDate As Integer, ByRef pwDosTime As Integer) As Long
Private Declare Function apiDosDateTimeToVariantTime Lib "oleaut32" Alias "DosDateTimeToVariantTime" (ByVal wDosDate As Integer, ByVal wDosTime As Integer, ByRef pvtime As Double) As Long
Private Declare Function apiVariantTimeToSystemTime Lib "oleaut32" Alias "VariantTimeToSystemTime" (ByRef vtime As Double, ByRef lpSystemTime As SYSTEMTIME) As Long
Private Declare Function apiCoDosDateTimeToFileTime Lib "ole32" Alias "CoDosDateTimeToFileTime" (ByVal nDosDate As Integer, ByVal nDosTime As Integer, ByRef lpFileTime As FILETIME) As Long
Private Declare Function apiCoFileTimeToDosDateTime Lib "ole32" Alias "CoFileTimeToDosDateTime" (ByRef lpFileTime As FILETIME, ByRef lpDosDate As Integer, ByRef lpDosTime As Integer) As Long
'
'Public Function FindEOCM(FindDate As Date) As Long
'    FindEOCM = DateSerial(Year(FindDate), Month(FindDate) + 1, 0)
'End Function
'
'Public Function FindEOPM(FindDate As Date) As Long
'    FindEOPM = DateSerial(Year(FindDate), Month(FindDate), 0)
'End Function
'
'Public Function WeekNumber(InDate As Date) As Integer
'  Dim DayNo     As Integer
'  Dim StartDays As Integer
'  Dim StopDays  As Integer
'  Dim StartDay  As Integer
'  Dim StopDay   As Integer
'  Dim VNumber   As Integer
'  Dim ThurFlag  As Boolean
'    DayNo = Days(InDate)
'    StartDay = Weekday(DateSerial(Year(InDate), 1, 1)) - 1
'    StopDay = Weekday(DateSerial(Year(InDate), 12, 31)) - 1
'    StartDays = 7 - (StartDay - 1)  ' Number of days belonging to first calendar week
'    StopDays = 7 - (StopDay - 1) ' Number of days belonging to last calendar week
'    If StartDay = 4 Or StopDay = 4 Then ThurFlag = True Else ThurFlag = False ' Test to see if the year will have 53 weeks or not
'    VNumber = (DayNo - StartDays - 4) / 7
'    If StartDays >= 4 Then ' If first week has 4 or more days, it will be calendar week 1    ' If first week has less than 4 days, it will belong to last year's    ' last calendar week
'        WeekNumber = Fix(VNumber) + 2
'    Else
'        WeekNumber = Fix(VNumber) + 1
'    End If    ' Handle years whose last days will belong to coming year's first    ' calendar week
'    If WeekNumber > 52 And ThurFlag = False Then WeekNumber = 1
'
'    If WeekNumber = 0 Then ' Handle years whose first days will belong to the last year's    ' last calendar week
'        WeekNumber = WeekNumber(DateSerial(Year(InDate) - 1, 12, 31))
'    End If
'End Function
'Public Function WeekNum2Date(iYear As Integer, iWeekNum As Integer, eWeekDay As DayConstants) As Date
'    Dim iStartDay As Integer
'    iStartDay = eWeekDay - 1
'    If iStartDay = 0 Then iStartDay = 7
'    WeekNum2Date = DateSerial(iYear, 1, 7 * (iWeekNum - 1) + iStartDay)
'End Function
'
'Public Sub GetMonthNames(ByVal bAbbreviate As Boolean, MonthNames() As String)
'    Dim lLoop       As Long
'    Dim fAbbrevFlag As Long
'    Dim sResult     As String
'    ReDim MonthNames(1 To 12)
'    fAbbrevFlag = Abs(bAbbreviate)
'    For lLoop = 1 To 12
'        sResult = Space(32)
'        apiVarMonthName lLoop, fAbbrevFlag, 0, sResult
'        MonthNames(lLoop) = StrConv(sResult, vbFromUnicode)
'    Next
'End Sub
'Public Sub GetWeekDayNames(ByVal bAbbreviate As Boolean, WeekDayNames() As String, Optional fFirstDay As DayConstants = ocalMonday)
'    Dim lLoop       As Long
'    Dim fAbbrevFlag As Long
'    Dim lFirstDay   As Long
'    Dim sResult     As String
'    ReDim WeekDayNames(1 To 7)
'    fAbbrevFlag = Abs(bAbbreviate)
'    lFirstDay = CLng(fFirstDay)
'    For lLoop = 1 To 7
'        sResult = Space(32)
'        apiVarWeekdayName lLoop, fAbbrevFlag, lFirstDay, 0, sResult
'        WeekDayNames(lLoop) = StrConv(sResult, vbFromUnicode)
'    Next
'End Sub
Private Function Days(DayNo As Date) As Integer
    Days = DayNo - DateSerial(Year(DayNo), 1, 0)
End Function
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

