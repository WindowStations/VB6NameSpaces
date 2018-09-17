VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SerialPort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type API_PORT_INFO_2
    pPortName As Long
    pMonitorName As Long
    pDescription As Long
    fPortType As Long
    Reserved As Long
End Type
Private Declare Function apiDeletePort Lib "winspool.drv" Alias "DeletePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Private Declare Function apiConfigurePort Lib "winspool.drv" Alias "ConfigurePortA" (ByVal pName As String, ByVal hwnd As Long, ByVal pPortName As String) As Long
Private Declare Function apiEnumPorts Lib "winspool.drv" Alias "EnumPortsA" (ByVal pName As String, ByVal Level As Long, ByVal lpbPorts As Long, ByVal cbBuf As Long, ByRef pcbNeeded As Long, ByRef pcReturned As Long) As Long
Private Declare Function apilstrlenW Lib "kernel32" Alias "lstrlenW" (ByVal lpString As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pTo As Any, ByRef uFrom As Long, ByVal lSize As Long)
Private Declare Function apiHeapAlloc Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function apiGetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
Private Declare Function apiHeapFree Lib "kernel32" Alias "HeapFree" (ByVal hHeap As Long, ByVal dwFlags As Long, ByRef lpMem As Any) As Long
' specify a server name (example //WIN2KWKSTN)
Friend Function GetPortNames(Optional ByVal ServerName As String = "") As String()
    On Error Resume Next
    Dim p()      As String
    Dim needed   As Long
    Dim returned As Long
    Dim buff     As Long
    Dim ret      As Long
    ret = apiEnumPorts(ServerName, 2, buff, 0, needed, returned)
    buff = apiHeapAlloc(apiGetProcessHeap(), 0, needed)
    ret = apiEnumPorts(ServerName, 2, buff, needed, needed, returned)
    If ret <> 0 Then
        Dim PortsStruct(0 To 100) As API_PORT_INFO_2
        Call apiCopyMemory(PortsStruct(0), ByVal buff, needed)
        Dim i As Integer
        For i = 0 To returned - 1
            On Error GoTo skip
            ReDim Preserve p(i)
            p(i) = LPSTRtoSTRING(PortsStruct(i).pPortName)
skip:
        Next
    End If
    If buff <> 0 Then
        Call apiHeapFree(apiGetProcessHeap(), 0, buff)
    End If
    GetPortNames = p
End Function
Friend Function GetSerialPorts(Optional ByVal ServerName As String = "") As SerialPort_()
    On Error Resume Next
    Dim port()   As SerialPort_
    Dim needed   As Long
    Dim returned As Long
    Dim buff     As Long
    Dim ret      As Long
    ret = apiEnumPorts(ServerName, 2, buff, 0, needed, returned)
    buff = apiHeapAlloc(apiGetProcessHeap(), 0, needed)
    ret = apiEnumPorts(ServerName, 2, buff, needed, needed, returned)
    If ret <> 0 Then
        Dim PortsStruct(0 To 100) As API_PORT_INFO_2
        Call apiCopyMemory(PortsStruct(0), ByVal buff, needed)
        Dim i As Integer
        For i = 0 To returned - 1
            ReDim Preserve port(i)
            On Error GoTo skip
            port(i).pDescription = LPSTRtoSTRING(PortsStruct(i).pDescription)
            port(i).pPortName = LPSTRtoSTRING(PortsStruct(i).pPortName)
            port(i).pMonitorName = LPSTRtoSTRING(PortsStruct(i).pMonitorName)
            port(i).fPortType = PortsStruct(i).fPortType
skip:
        Next
    End If
    If buff <> 0 Then Call apiHeapFree(apiGetProcessHeap(), 0, buff)
    GetSerialPorts = port
End Function
Private Function LPSTRtoSTRING(ByVal lngPointer As Long) As String
    Dim lngLength As Long
    lngLength = apilstrlenW(lngPointer) * 2
    LPSTRtoSTRING = String(lngLength, 0)
    Call apiCopyMemory(ByVal StrPtr(LPSTRtoSTRING), ByVal lngPointer, lngLength)
    LPSTRtoSTRING = CutString(StrConv(LPSTRtoSTRING, vbUnicode))
End Function
Private Function CutString(strName As String) As String
    Dim x As Integer
    x = InStr(strName, vbNullChar)
    If x > 0 Then CutString = Left(strName, x - 1) Else CutString = strName
End Function

