VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "SerialPort"
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
'Attribute VB_Name = "SerialPort"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
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
Private Declare Sub apiCopyMemoryPORT Lib "kernel32" Alias "RtlMoveMemory" (ByRef pTo As API_PORT_INFO_2, ByRef uFrom As Long, ByVal lSize As Long)
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pTo As Long, ByRef uFrom As Long, ByVal lSize As Long)
Private Declare Function apiHeapAlloc Lib "kernel32" Alias "HeapAlloc" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function apiGetProcessHeap Lib "kernel32" Alias "GetProcessHeap" () As Long
Private Declare Function apiHeapFree Lib "kernel32" Alias "HeapFree" (ByVal hHeap As Long, ByVal dwFlags As Long, ByRef lpMem As Long) As Long
'--------------------------------
Private Const LINE_BREAK                 As Long = 1
Private Const LINE_DTR                   As Long = 2
Private Const LINE_RTS                   As Long = 3
Private Const LINE_CTS                   As Long = &H10
Private Const LINE_DSR                   As Long = &H20
Private Const LINE_RING                  As Long = &H40
Private Const LINE_RLSD                  As Long = &H80
Private Const LINE_CD                    As Long = &H80
Private Const ERROR_IO_INCOMPLETE        As Long = 996
Private Const ERROR_IO_PENDING           As Long = 997
Private Const GENERIC_READ               As Long = &H80000000
Private Const GENERIC_WRITE              As Long = &H40000000
Private Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Private Const FILE_FLAG_OVERLAPPED       As Long = &H40000000
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const OPEN_EXISTING              As Long = 3
Private Const MS_CTS_ON                  As Long = &H10
Private Const MS_DSR_ON                  As Long = &H20
Private Const MS_RING_ON                 As Long = &H40
Private Const MS_RLSD_ON                 As Long = &H80
Private Const PURGE_RXABORT              As Long = &H2
Private Const PURGE_RXCLEAR              As Long = &H8
Private Const PURGE_TXABORT              As Long = &H1
Private Const PURGE_TXCLEAR              As Long = &H4
Private Const CLRBREAK                   As Long = 9
Private Const CLRDTR                     As Long = 6
Private Const CLRRTS                     As Long = 4
Private Const SETBREAK                   As Long = 8
Private Const SETDTR                     As Long = 5
Private Const SETRTS                     As Long = 3
Private Const MAX_PORTS                  As Long = 4
Private Type DCB
    DCBlength As Long
    BaudRate As Long
    fBitFields As Long
    wReserved As Integer
    XonLim As Integer
    XoffLim As Integer
    ByteSize As Byte
    Parity As Byte
    StopBits As Byte
    XonChar As Byte
    XoffChar As Byte
    ErrorChar As Byte
    EofChar As Byte
    EvtChar As Byte
    wReserved1 As Integer
End Type
Private Type COMMTIMEOUTS
    ReadIntervalTimeout As Long
    ReadTotalTimeoutMultiplier As Long
    ReadTotalTimeoutConstant As Long
    WriteTotalTimeoutMultiplier As Long
    WriteTotalTimeoutConstant As Long
End Type
Private Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type
Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Type COMSTAT
    fBitFields As Long
    cbInQue As Long
    cbOutQue As Long
End Type
Private Type COMM_ERROR
    lngErrorCode As Long
    strFunction As String
    strErrorMessage As String
End Type
Private Type COMM_PORT
    lngHandle As Long
    blnPortOpen As Boolean
    udtDCB As DCB
End Type
Private Declare Function apiBuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, ByRef lpDCB As DCB) As Long
Private Declare Function apiClearCommError Lib "kernel32" Alias "ClearCommError" (ByVal hFile As Long, ByRef lpErrors As Long, ByRef lpStat As COMSTAT) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal hObject As Long) As Long
Private Declare Function apiCreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function apiEscapeCommFunction Lib "kernel32" Alias "EscapeCommFunction" (ByVal nCid As Long, ByVal nFunc As Long) As Long
Private Declare Function apiFormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Long, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByRef Arguments As Long) As Long
Private Declare Function apiGetCommModemStatus Lib "kernel32" Alias "GetCommModemStatus" (ByVal hFile As Long, ByRef lpModemStat As Long) As Long
Private Declare Function apiGetCommState Lib "kernel32" Alias "GetCommState" (ByVal nCid As Long, ByRef lpDCB As DCB) As Long
Private Declare Function apiGetLastError Lib "kernel32" Alias "GetLastError" () As Long
Private Declare Function apiGetOverlappedResult Lib "kernel32" Alias "GetOverlappedResult" (ByVal hFile As Long, ByRef lpOverlapped As OVERLAPPED, ByRef lpNumberOfBytesTransferred As Long, ByVal bWait As Long) As Long
Private Declare Function apiPurgeComm Lib "kernel32" Alias "PurgeComm" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiReadFile Lib "kernel32" Alias "ReadFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private Declare Function apiSetCommState Lib "kernel32" Alias "SetCommState" (ByVal hCommDev As Long, ByRef lpDCB As DCB) As Long
Private Declare Function apiSetCommTimeouts Lib "kernel32" Alias "SetCommTimeouts" (ByVal hFile As Long, ByRef lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function apiSetupComm Lib "kernel32" Alias "SetupComm" (ByVal hFile As Long, ByVal dwInQueue As Long, ByVal dwOutQueue As Long) As Long
Private Declare Function apiWriteFile Lib "kernel32" Alias "WriteFile" (ByVal hFile As Long, ByVal lpBuffer As String, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Private udtCommOverlap           As OVERLAPPED
Private udtCommError             As COMM_ERROR
Private udtPorts(1 To MAX_PORTS) As COMM_PORT
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
        Call apiCopyMemoryPORT(PortsStruct(0), ByVal buff, needed)
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
    Dim Port()   As SerialPort_
    Dim needed   As Long
    Dim returned As Long
    Dim buff     As Long
    Dim ret      As Long
    ret = apiEnumPorts(ServerName, 2, buff, 0, needed, returned)
    buff = apiHeapAlloc(apiGetProcessHeap(), 0, needed)
    ret = apiEnumPorts(ServerName, 2, buff, needed, needed, returned)
    If ret <> 0 Then
        Dim PortsStruct(0 To 100) As API_PORT_INFO_2
        Call apiCopyMemoryPORT(PortsStruct(0), ByVal buff, needed)
        Dim i As Integer
        For i = 0 To returned - 1
            ReDim Preserve Port(i)
            On Error GoTo skip
            Port(i).pDescription = LPSTRtoSTRING(PortsStruct(i).pDescription)
            Port(i).pPortName = LPSTRtoSTRING(PortsStruct(i).pPortName)
            Port(i).pMonitorName = LPSTRtoSTRING(PortsStruct(i).pMonitorName)
            Port(i).fPortType = PortsStruct(i).fPortType
skip:
        Next
    End If
    If buff <> 0 Then Call apiHeapFree(apiGetProcessHeap(), 0, buff)
    GetSerialPorts = Port
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
'
'
'
'
'----------------------------------------------------
Private Function GetSystemMessage(ByVal lngErrorCode As Long) As String
    Dim intPos     As Integer
    Dim strMessage As String, strMsgBuff As String * 256
    Call apiFormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0, lngErrorCode, 0, strMsgBuff, 255, 0)
    intPos = InStr(1, strMsgBuff, vbNullChar)
    If intPos > 0 Then
        strMessage = Trim(Left(strMsgBuff, intPos - 1))
    Else
        strMessage = Trim(strMsgBuff)
    End If
    GetSystemMessage = strMessage
End Function
' CommOpen - Opens/Initializes serial port.
'   intPortID   - Port ID used when port was opened.
'   strPort     - COM port name. (COM1, COM2, COM3, COM4)
'   strSettings - Communication settings.
'                 Example: "baud=9600 parity=N data=8 stop=1"
Private Function CommOpen(intPortID As Integer, strPort As String, strSettings As String) As Long
    Dim lngStatus       As Long
    Dim udtCommTimeOuts As COMMTIMEOUTS
    On Error GoTo Routine_Error
    If udtPorts(intPortID).blnPortOpen = True Then
        lngStatus = -1
        With udtCommError
            .lngErrorCode = lngStatus
            .strFunction = "CommOpen"
            .strErrorMessage = "Port in use."
        End With
        GoTo Routine_Exit
    End If
    Dim sa As SECURITY_ATTRIBUTES
    udtPorts(intPortID).lngHandle = apiCreateFile(strPort, GENERIC_READ Or GENERIC_WRITE, 0, sa, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If udtPorts(intPortID).lngHandle = -1 Then
        lngStatus = SetCommError("CommOpen (CreateFile)")
        GoTo Routine_Exit
    End If
    udtPorts(intPortID).blnPortOpen = True
    lngStatus = apiSetupComm(udtPorts(intPortID).lngHandle, 1024, 1024)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetupComm)")
        GoTo Routine_Exit
    End If
    lngStatus = apiPurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (PurgeComm)")
        GoTo Routine_Exit
    End If
    With udtCommTimeOuts
        .ReadIntervalTimeout = -1
        .ReadTotalTimeoutMultiplier = 0
        .ReadTotalTimeoutConstant = 1000
        .WriteTotalTimeoutMultiplier = 0
        .WriteTotalTimeoutMultiplier = 1000
    End With
    lngStatus = apiSetCommTimeouts(udtPorts(intPortID).lngHandle, udtCommTimeOuts)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommTimeouts)")
        GoTo Routine_Exit
    End If
    lngStatus = apiGetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (GetCommState)")
        GoTo Routine_Exit
    End If
    lngStatus = apiBuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (BuildCommDCB)")
        GoTo Routine_Exit
    End If
    lngStatus = apiSetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommOpen (SetCommState)")
        GoTo Routine_Exit
    End If
    lngStatus = 0
Routine_Exit:
    CommOpen = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommOpen"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
Private Function SetCommError(strFunction As String) As Long
    With udtCommError
        .lngErrorCode = Err.LastDllError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        SetCommError = .lngErrorCode
    End With
End Function
Private Function SetCommErrorEx(strFunction As String, lngHnd As Long) As Long
    Dim lngErrorFlags As Long
    Dim udtCommStat   As COMSTAT
    With udtCommError
        .lngErrorCode = apiGetLastError
        .strFunction = strFunction
        .strErrorMessage = GetSystemMessage(.lngErrorCode)
        Call apiClearCommError(lngHnd, lngErrorFlags, udtCommStat)
        .strErrorMessage = .strErrorMessage & "  COMM Error Flags = " & Hex(lngErrorFlags)
        SetCommErrorEx = .lngErrorCode
    End With
End Function
' CommSet - Modifies the serial port settings.' Parameters:  intPortID   - Port ID used when port was opened.   strSettings - Communication settings.
'Example: "baud=9600 parity=N data=8 stop=1"
Private Function CommSet(intPortID As Integer, strSettings As String) As Long
    Dim lngStatus As Long
    On Error GoTo Routine_Error
    lngStatus = apiGetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (GetCommState)")
        GoTo Routine_Exit
    End If
    lngStatus = apiBuildCommDCB(strSettings, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (BuildCommDCB)")
        GoTo Routine_Exit
    End If
    lngStatus = apiSetCommState(udtPorts(intPortID).lngHandle, udtPorts(intPortID).udtDCB)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSet (SetCommState)")
        GoTo Routine_Exit
    End If
    lngStatus = 0
Routine_Exit:
    CommSet = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSet"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommClose - Close the serial port. intPortID   - Port ID used when port was opened.
Private Function CommClose(intPortID As Integer) As Long
    Dim lngStatus As Long
    On Error GoTo Routine_Error
    If udtPorts(intPortID).blnPortOpen Then
        lngStatus = apiCloseHandle(udtPorts(intPortID).lngHandle)
        If lngStatus = 0 Then
            lngStatus = SetCommError("CommClose (CloseHandle)")
            GoTo Routine_Exit
        End If
        udtPorts(intPortID).blnPortOpen = False
    End If
    lngStatus = 0
Routine_Exit:
    CommClose = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommClose"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommFlush - Flush the send and receive serial port buffers. intPortID   - Port ID used when port was opened.
Private Function CommFlush(intPortID As Integer) As Long
    Dim lngStatus As Long
    On Error GoTo Routine_Error
    lngStatus = apiPurgeComm(udtPorts(intPortID).lngHandle, PURGE_TXABORT Or PURGE_RXABORT Or PURGE_TXCLEAR Or PURGE_RXCLEAR)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommFlush (PurgeComm)")
        GoTo Routine_Exit
    End If
    lngStatus = 0
Routine_Exit:
    CommFlush = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommFlush"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommRead - Read serial port input buffer.
'   intPortID   - Port ID used when port was opened.
'   strData     - Data buffer.
'   lngSize     - Maximum number of bytes to be read.
Private Function CommRead(intPortID As Integer, strData As String, lngSize As Long) As Long
    Dim lngStatus     As Long
    Dim lngRdSize     As Long, lngBytesRead As Long
    Dim lngRdStatus   As Long, strRdBuffer As String * 1024
    Dim lngErrorFlags As Long, udtCommStat As COMSTAT
    On Error GoTo Routine_Error
    strData = ""
    lngBytesRead = 0
    DoEvents
    lngStatus = apiClearCommError(udtPorts(intPortID).lngHandle, lngErrorFlags, udtCommStat)
    If lngStatus = 0 Then
        lngBytesRead = -1
        lngStatus = SetCommError("CommRead (ClearCommError)")
        GoTo Routine_Exit
    End If
    If udtCommStat.cbInQue > 0 Then
        If udtCommStat.cbInQue > lngSize Then
            lngRdSize = udtCommStat.cbInQue
        Else
            lngRdSize = lngSize
        End If
    Else
        lngRdSize = 0
    End If
    If lngRdSize Then
        lngRdStatus = apiReadFile(udtPorts(intPortID).lngHandle, strRdBuffer, lngRdSize, lngBytesRead, udtCommOverlap)
        If lngRdStatus = 0 Then
            lngStatus = apiGetLastError
            If lngStatus = ERROR_IO_PENDING Then
                ' Wait for read to complete. This function will timeout according to the  COMMTIMEOUTS.ReadTotalTimeoutConstant variable.
                ' Every time it times out, check for port errors. Loop until operation is complete.
                While apiGetOverlappedResult(udtPorts(intPortID).lngHandle, udtCommOverlap, lngBytesRead, True) = 0
                    lngStatus = apiGetLastError
                    If lngStatus <> ERROR_IO_INCOMPLETE Then
                        lngBytesRead = -1
                        lngStatus = SetCommErrorEx("CommRead (GetOverlappedResult)", udtPorts(intPortID).lngHandle)
                        GoTo Routine_Exit
                    End If
                Wend
            Else
                lngBytesRead = -1
                lngStatus = SetCommErrorEx("CommRead (ReadFile)", udtPorts(intPortID).lngHandle)
                GoTo Routine_Exit
            End If
        End If
        strData = Left(strRdBuffer, lngBytesRead)
    End If
Routine_Exit:
    CommRead = lngBytesRead
    Exit Function
Routine_Error:
    lngBytesRead = -1
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommRead"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommWrite - Output data to the serial port.
'   intPortID   - Port ID used when port was opened.
'   strData     - Data to be transmitted.
Private Function CommWrite(intPortID As Integer, strData As String) As Long
    Dim i         As Integer
    Dim lngStatus As Long, lngSize As Long
    Dim lngWrSize As Long, lngWrStatus As Long
    On Error GoTo Routine_Error
    lngSize = Len(strData)
    lngWrStatus = apiWriteFile(udtPorts(intPortID).lngHandle, strData, lngSize, lngWrSize, udtCommOverlap)
    ' Note that normally the following code will not execute because the driver caches write operations. Small I/O requests (up to several thousand bytes)
    ' will normally be accepted immediately and WriteFile will return true even though an overlapped operation was specified.
    DoEvents
    If lngWrStatus = 0 Then
        lngStatus = apiGetLastError
        If lngStatus = 0 Then
            GoTo Routine_Exit
        ElseIf lngStatus = ERROR_IO_PENDING Then
            ' We should wait for the completion of the write operation so we know  if it worked or not.
            ' This is only one way to do this. It might be beneficial to place the writing operation in a separate thread so that blocking on completion
            ' will not negatively affect the responsiveness of the UI. If the write takes long enough to complete, this function will timeout
            ' according to the CommTimeOuts.WriteTotalTimeoutConstant variable. At that time we can check for errors and then wait some more. Loop until operation is complete.
            While apiGetOverlappedResult(udtPorts(intPortID).lngHandle, udtCommOverlap, lngWrSize, True) = 0
                lngStatus = apiGetLastError
                If lngStatus <> ERROR_IO_INCOMPLETE Then
                    lngStatus = SetCommErrorEx("CommWrite (GetOverlappedResult)", udtPorts(intPortID).lngHandle)
                    GoTo Routine_Exit
                End If
            Wend
        Else
            lngWrSize = -1
            lngStatus = SetCommErrorEx("CommWrite (WriteFile)", udtPorts(intPortID).lngHandle)
            GoTo Routine_Exit
        End If
    End If
    For i = 1 To 10
        DoEvents
    Next
Routine_Exit:
    CommWrite = lngWrSize
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommWrite"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommGetLine - Get the state of selected serial port control lines.
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. CTS, DSR, RING, RLSD (CD)
'   blnState    - Returns state of line (Cleared or Set).
Private Function CommGetLine(intPortID As Integer, intLine As Integer, blnState As Boolean) As Long
    Dim lngStatus    As Long
    Dim lngComStatus As Long, lngModemStatus As Long
    On Error GoTo Routine_Error
    lngStatus = apiGetCommModemStatus(udtPorts(intPortID).lngHandle, lngModemStatus)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommReadCD (GetCommModemStatus)")
        GoTo Routine_Exit
    End If
    If (lngModemStatus And intLine) Then
        blnState = True
    Else
        blnState = False
    End If
    lngStatus = 0
Routine_Exit:
    CommGetLine = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommReadCD"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
' CommSetLine - Set the state of selected serial port control lines.
'   intPortID   - Port ID used when port was opened.
'   intLine     - Serial port line. BREAK, DTR, RTS
'                 Note: BREAK actually sets or clears a "break" condition on
'                 the transmit data line.
'   blnState    - Sets the state of line (Cleared or Set).
Private Function CommSetLine(intPortID As Integer, intLine As Integer, blnState As Boolean) As Long
    Dim lngStatus   As Long
    Dim lngNewState As Long
    On Error GoTo Routine_Error
    If intLine = LINE_BREAK Then
        If blnState Then
            lngNewState = SETBREAK
        Else
            lngNewState = CLRBREAK
        End If
    ElseIf intLine = LINE_DTR Then
        If blnState Then
            lngNewState = SETDTR
        Else
            lngNewState = CLRDTR
        End If
    ElseIf intLine = LINE_RTS Then
        If blnState Then
            lngNewState = SETRTS
        Else
            lngNewState = CLRRTS
        End If
    End If
    lngStatus = apiEscapeCommFunction(udtPorts(intPortID).lngHandle, lngNewState)
    If lngStatus = 0 Then
        lngStatus = SetCommError("CommSetLine (EscapeCommFunction)")
        GoTo Routine_Exit
    End If
    lngStatus = 0
Routine_Exit:
    CommSetLine = lngStatus
    Exit Function
Routine_Error:
    lngStatus = Err.Number
    With udtCommError
        .lngErrorCode = lngStatus
        .strFunction = "CommSetLine"
        .strErrorMessage = Err.Description
    End With
    Resume Routine_Exit
End Function
Private Function CommGetError(strMessage As String) As Long
    With udtCommError
        CommGetError = .lngErrorCode
        strMessage = "Error (" & CStr(.lngErrorCode) & "): " & .strFunction & " - " & .strErrorMessage
    End With
End Function

