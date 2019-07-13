Attribute VB_Name = "modMain"
Option Explicit
Public Const appTitled As String = "Visual Basic 6.5"
'Attribute VB_Name = "modMain"
'Option Explicit 'Import VB6 class namespaces in a public module
'-Framework base classes
Public System                          As New System
Public My                              As New My
Public Microsoft                       As New Microsoft
'-Subset of framework base classes
'My
Public Application                     As New Application
'System
Public Environment                     As New Environment
Public Threading                       As New Threading
Public Diagnostics                     As New Diagnostics
Public IO                              As New IO
Public Drawing                         As New Drawing
Public Net                             As New Net
Public Automation                      As New Automation
'-Framework namespaces (classes) imported directly for ease of access, similar to a default vb.net form
'System.Diagnostics
Public Process                         As New Process
'System.Windows.Forms
Public Sendkeys                        As New Sendkeys
Public MessageBox                      As New MessageBox
Public OpenFileDialog                  As New OpenFileDialog
Public SaveFileDialog                  As New SaveFileDialog
Public ColorDialog                     As New ColorDialog
Public FontDialog                      As New FontDialog
Public PrintDialog                     As New PrintDialog
Public StringEx                        As New StringEx
Public ProcessWindowStyle              As New ProcessWindowStyle
'Public User defined types can be used for these private class modules
'They are dimensioned in a module.  They return multi-variable
' values from Functions in the scope of a Friend (throughout the project only)
Private Const WHDR_DONE                As Long = &H1
Private Const WHDR_PREPARED            As Long = &H2
Private Const CALLBACK_WINDOW          As Long = &H10000
Private Const WAVE_MAPPED              As Long = &H4
Private Const WAVE_MAPPER              As Long = -1
Private Const MMSYSERR_BASE            As Long = 0
Private Const MMSYSERR_ALLOCATED       As Long = (MMSYSERR_BASE + 4)
Private Const MMSYSERR_BADDB           As Long = (MMSYSERR_BASE + 14)
Private Const MMSYSERR_BADDEVICEID     As Long = (MMSYSERR_BASE + 2)
Private Const MMSYSERR_BADERRNUM       As Long = (MMSYSERR_BASE + 9)
Private Const MMSYSERR_DELETEERROR     As Long = (MMSYSERR_BASE + 18)
Private Const MMSYSERR_ERROR           As Long = (MMSYSERR_BASE + 1)
Private Const MMSYSERR_HANDLEBUSY      As Long = (MMSYSERR_BASE + 12)
Private Const MMSYSERR_INVALFLAG       As Long = (MMSYSERR_BASE + 10)
Private Const MMSYSERR_INVALHANDLE     As Long = (MMSYSERR_BASE + 5)
Private Const MMSYSERR_INVALIDALIAS    As Long = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_INVALPARAM      As Long = (MMSYSERR_BASE + 11)
Private Const MMSYSERR_KEYNOTFOUND     As Long = (MMSYSERR_BASE + 15)
Private Const MMSYSERR_LASTERROR       As Long = (MMSYSERR_BASE + 13)
Private Const MMSYSERR_MOREDATA        As Long = (MMSYSERR_BASE + 21)
Private Const MMSYSERR_NODRIVER        As Long = (MMSYSERR_BASE + 6)
Private Const MMSYSERR_NODRIVERCB      As Long = (MMSYSERR_BASE + 20)
Private Const MMSYSERR_NOERROR         As Long = 0
Private Const MMSYSERR_NOMEM           As Long = (MMSYSERR_BASE + 7)
Private Const MMSYSERR_NOTENABLED      As Long = (MMSYSERR_BASE + 3)
Private Const MMSYSERR_NOTSUPPORTED    As Long = (MMSYSERR_BASE + 8)
Private Const MMSYSERR_READERROR       As Long = (MMSYSERR_BASE + 16)
Private Const MMSYSERR_VALNOTFOUND     As Long = (MMSYSERR_BASE + 19)
Private Const MMSYSERR_WRITEERROR      As Long = (MMSYSERR_BASE + 17)
Private Const MMIO_ALLOCBUF            As Long = &H10000
Private Const MMIO_COMPAT              As Long = &H0
Private Const MMIO_CREATE              As Long = &H1000
Private Const MMIO_CREATELIST          As Long = &H40
Private Const MMIO_CREATERIFF          As Long = &H20
Private Const MMIO_DEFAULTBUFFER       As Long = 8192
Private Const MMIO_DELETE              As Long = &H200
Private Const MMIO_DENYNONE            As Long = &H40
Private Const MMIO_DENYREAD            As Long = &H30
Private Const MMIO_DENYWRITE           As Long = &H20
Private Const MMIO_DIRTY               As Long = &H10000000
Private Const MMIO_EMPTYBUF            As Long = &H10
Private Const MMIO_EXCLUSIVE           As Long = &H10
Private Const MMIO_EXIST               As Long = &H4000
Private Const MMIO_FHOPEN              As Long = &H10
Private Const MMIO_FINDCHUNK           As Long = &H10
Private Const MMIO_FINDLIST            As Long = &H40
Private Const MMIO_FINDPROC            As Long = &H40000
Private Const MMIO_FINDRIFF            As Long = &H20
Private Const MMIO_GETTEMP             As Long = &H20000
Private Const MMIO_GLOBALPROC          As Long = &H10000000
Private Const MMIO_INSTALLPROC         As Long = &H10000
Private Const MMIO_OPEN_VALID          As Long = &H3FFFF
Private Const MMIO_PARSE               As Long = &H100
Private Const MMIO_PUBLICPROC          As Long = &H10000000
Private Const MMIO_READ                As Long = &H0
Private Const MMIO_READWRITE           As Long = &H2
Private Const MMIO_REMOVEPROC          As Long = &H20000
Private Const MMIO_RWMODE              As Long = &H3
Private Const MMIO_SHAREMODE           As Long = &H70
Private Const MMIO_TOUPPER             As Long = &H10
Private Const MMIO_UNICODEPROC         As Long = &H1000000
Private Const MMIO_VALIDPROC           As Long = &H11070000
Private Const MMIO_WRITE               As Long = &H1
Private Const MMIOERR_BASE             As Long = 256
Private Const MMIOERR_ACCESSDENIED     As Long = (MMIOERR_BASE + 12)
Private Const MMIOERR_CANNOTCLOSE      As Long = (MMIOERR_BASE + 4)
Private Const MMIOERR_CANNOTEXPAND     As Long = (MMIOERR_BASE + 8)
Private Const MMIOERR_CANNOTOPEN       As Long = (MMIOERR_BASE + 3)
Private Const MMIOERR_CANNOTREAD       As Long = (MMIOERR_BASE + 5)
Private Const MMIOERR_CANNOTSEEK       As Long = (MMIOERR_BASE + 7)
Private Const MMIOERR_CANNOTWRITE      As Long = (MMIOERR_BASE + 6)
Private Const MMIOERR_CHUNKNOTFOUND    As Long = (MMIOERR_BASE + 9)
Private Const MMIOERR_FILENOTFOUND     As Long = (MMIOERR_BASE + 1)
Private Const MMIOERR_INVALIDFILE      As Long = (MMIOERR_BASE + 16)
Private Const MMIOERR_NETWORKERROR     As Long = (MMIOERR_BASE + 14)
Private Const MMIOERR_OUTOFMEMORY      As Long = (MMIOERR_BASE + 2)
Private Const MMIOERR_PATHNOTFOUND     As Long = (MMIOERR_BASE + 11)
Private Const MMIOERR_SHARINGVIOLATION As Long = (MMIOERR_BASE + 13)
Private Const MMIOERR_TOOMANYOPENFILES As Long = (MMIOERR_BASE + 15)
Private Const MMIOERR_UNBUFFERED       As Long = (MMIOERR_BASE + 10)
Private Const MMIOM_CLOSE              As Long = 4
Private Const MMIOM_OPEN               As Long = 3
Private Const MMIOM_READ               As Long = MMIO_READ
Private Const MMIOM_RENAME             As Long = 6
Private Const MMIOM_SEEK               As Long = 2
Private Const MMIOM_USER               As Long = &H8000
Private Const MMIOM_WRITE              As Long = MMIO_WRITE
Private Const MMIOM_WRITEFLUSH         As Long = 5
Private Const SEEK_SET                 As Long = 0
Private Const MM_WOM_CLOSE             As Long = &H3BC
Private Const MM_WOM_DONE              As Long = &H3BD
Private Const MM_WOM_OPEN              As Long = &H3BB
Private Const WM_DESTROY               As Long = &H2
Private Const WM_CLOSE                 As Long = &H10
Private Const SS_SIMPLE                As Long = &HB
Private Const WS_POPUP                 As Long = &H80000000
Private Const GWL_WNDPROC              As Long = -4
Private Const MAX_BUFFER_COUNT         As Long = 32
Public Const ALL_SOUND_BUFFERS         As Long = -1
'Screen Monitor
Private Const CCDEVICENAME             As Long = 32
Private Const CCFORMNAME               As Long = 32
Private Const DM_BITSPERPEL            As Long = &H40000
Private Const DM_PELSWIDTH             As Long = &H80000
Private Const DM_PELSHEIGHT            As Long = &H100000
Private Const CDS_UPDATEREGISTRY       As Long = &H1
Private Const CDS_TEST                 As Long = &H4
Private Const DISP_CHANGE_SUCCESSFUL   As Long = 0
Private Const DISP_CHANGE_RESTART      As Long = 1
Private Const MONITORINFOF_PRIMARY     As Long = &H1
Private Const MONITOR_DEFAULTTONEAREST As Long = &H2
Private Const MONITOR_DEFAULTTONULL    As Long = &H0
Private Const MONITOR_DEFAULTTOPRIMARY As Long = &H1
'Strings
Private Const CSTR_LESS_THAN           As Long = 1
Private Const CSTR_EQUAL               As Long = 2
Private Const CSTR_GREATER_THAN        As Long = 3
Private Const LOCALE_SYSTEM_DEFAULT    As Long = &H400
Private Const LOCALE_USER_DEFAULT      As Long = &H800
Private Const NORM_IGNORECASE          As Long = &H1
Private Const NORM_IGNOREKANATYPE      As Long = &H10000
Private Const NORM_IGNORENONSPACE      As Long = &H2
Private Const NORM_IGNORESYMBOLS       As Long = &H4
Private Const NORM_IGNOREWIDTH         As Long = &H20000
Private Const SORT_STRINGSORT          As Long = &H1000
'
Public Type AutomationElement_
    Name As String
    ClassName As String
    ProcessID As Long
End Type
'
Public Type AdapterInfo
    Name As String
    AdapterIndex As Long
    Type As Long
    Address As String
    IP As String
    Description As String
    GatewayIP As String
End Type
'
Public Type Version
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
'
Public Type SerialPort_
    pPortName As String
    pMonitorName As String
    pDescription As String
    fPortType As Long
    Reserved As Long
End Type
'
Public Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type
'
'Audio
Private Type WAVEHDR
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type
Private Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type
Private Type MMIOINFO
    dwFlags As Long
    fccIOProc As Long
    pIOProc As Long
    wErrorRet As Long
    htask As Long
    cchBuffer As Long
    pchBuffer As String
    pchNext As String
    pchEndRead As String
    pchEndWrite As String
    lBufOffset As Long
    lDiskOffset As Long
    adwInfo(4) As Long
    dwReserved1 As Long
    dwReserved2 As Long
    hmmio As Long
End Type
Private Type WAVEFORMATEX
    wFormatTag As Integer
    nChannels As Integer
    nSamplesPerSec As Long
    nAvgBytesPerSec As Long
    nBlockAlign As Integer
    wBitsPerSample As Integer
    cbSize As Integer
End Type
Private Type SoundBufferInfo
    hWaveOut As Long
    hdr As WAVEHDR
    buf() As Byte
    Status As SoundBufferStatus
    Flags As SoundBufferFlags
End Type
'
'keyboard
Private Type KBDLLHOOKSTRUCT
    vkCode As Long
    scanCode As Long
    Flags As Long
    time As Long
    dwExtraInfo As Long
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type
Public Type WINNAME
    lpText As String
    lpClass As String
End Type
Public Type WINFOCUS
    Foreground As Long
    Focus As Long
End Type
'
'Screen
Public Type DISPLAY_DEVICE
    cb As Long
    DeviceName As String * 32
    DeviceString As String * 128
    StateFlags As Long
    DeviceID As String * 128
    DeviceKey As String * 128
End Type
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type
Private Type MONITORINFO
    cbSize          As Long
    rcMonitor       As RECT
    rcWork          As RECT
    dwFlags         As Long
End Type
Public Type TSAFEARRAYBOUND
    lElements As Long
    lLowest As Long
End Type
Public Type TSAFEARRAY
    iDims As Integer
    iFeatures As Integer
    lElementSize As Long
    lLocks As Long
    lData As Long
    lPointer As Long
    lVarType As Long
    lSorted As Long
    uBounds() As TSAFEARRAYBOUND
End Type
Public Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As TSAFEARRAYBOUND
End Type
'
'API
Private Declare Sub apiZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Function apiwaveOutClose Lib "winmm" Alias "waveOutClose" (ByVal hWaveOut As Long) As Long
Private Declare Function apiwaveOutOpen Lib "winmm" Alias "waveOutOpen" (ByRef lphWaveOut As Long, ByVal uDeviceID As Long, ByRef lpFormat As WAVEFORMATEX, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiwaveOutPrepareHeader Lib "winmm" Alias "waveOutPrepareHeader" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function apiwaveOutUnprepareHeader Lib "winmm" Alias "waveOutUnprepareHeader" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function apiwaveOutWrite Lib "winmm" Alias "waveOutWrite" (ByVal hWaveOut As Long, ByRef lpWaveOutHdr As WAVEHDR, ByVal uSize As Long) As Long
Private Declare Function apiwaveOutPause Lib "winmm" Alias "waveOutPause" (ByVal hWaveOut As Long) As Long
Private Declare Function apiwaveOutReset Lib "winmm" Alias "waveOutReset" (ByVal hWaveOut As Long) As Long
Private Declare Function apiwaveOutRestart Lib "winmm" Alias "waveOutRestart" (ByVal hWaveOut As Long) As Long
Private Declare Function apimmioAdvance Lib "winmm" Alias "mmioAdvance" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function apimmioAscend Lib "winmm" Alias "mmioAscend" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function apimmioClose Lib "winmm" Alias "mmioClose" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function apimmioCreateChunk Lib "winmm" Alias "mmioCreateChunk" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByVal uFlags As Long) As Long
Private Declare Function apimmioDescend Lib "winmm" Alias "mmioDescend" (ByVal hmmio As Long, ByRef lpck As MMCKINFO, ByRef lpckParent As Any, ByVal uFlags As Long) As Long
Private Declare Function apimmioFlush Lib "winmm" Alias "mmioFlush" (ByVal hmmio As Long, ByVal uFlags As Long) As Long
Private Declare Function apimmioGetInfo Lib "winmm" Alias "mmioGetInfo" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function apimmioInstallIOProc Lib "winmm" Alias "mmioInstallIOProcA" (ByVal fccIOProc As Long, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function apimmioInstallIOProcA Lib "winmm" Alias "mmioInstallIOProcA" (ByVal fccIOProc As String, ByVal pIOProc As Long, ByVal dwFlags As Long) As Long
Private Declare Function apimmioOpen Lib "winmm" Alias "mmioOpenA" (ByVal szFileName As String, ByRef lpmmioinfo As Any, ByVal dwOpenFlags As Long) As Long
Private Declare Function apimmioRead Lib "winmm" Alias "mmioRead" (ByVal hmmio As Long, ByRef pch As Any, ByVal cch As Long) As Long
Private Declare Function apimmioRename Lib "winmm" Alias "mmioRenameA" (ByVal szFileName As String, ByVal SzNewFileName As String, ByRef lpmmioinfo As MMIOINFO, ByVal dwRenameFlags As Long) As Long
Private Declare Function apimmioSeek Lib "winmm" Alias "mmioSeek" (ByVal hmmio As Long, ByVal lOffset As Long, ByVal iOrigin As Long) As Long
Private Declare Function apimmioSendMessage Lib "winmm" Alias "mmioSendMessage" (ByVal hmmio As Long, ByVal uMsg As Long, ByVal lParam1 As Long, ByVal lParam2 As Long) As Long
Private Declare Function apimmioSetBuffer Lib "winmm" Alias "mmioSetBuffer" (ByVal hmmio As Long, ByVal pchBuffer As String, ByVal cchBuffer As Long, ByVal uFlags As Long) As Long
Private Declare Function apimmioSetInfo Lib "winmm" Alias "mmioSetInfo" (ByVal hmmio As Long, ByRef lpmmioinfo As MMIOINFO, ByVal uFlags As Long) As Long
Private Declare Function apimmioStringToFOURCC Lib "winmm" Alias "mmioStringToFOURCCA" (ByVal sz As String, ByVal uFlags As Long) As Long
Private Declare Function apimmioWrite Lib "winmm" Alias "mmioWrite" (ByVal hmmio As Long, ByVal pch As String, ByVal cch As Long) As Long
Private Declare Function apimmsystemGetVersion Lib "winmm" Alias "mmsystemGetVersion" () As Long
Private Declare Function apiCallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiCreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function apiDestroyWindow Lib "user32" Alias "DestroyWindow" (ByVal hwnd As Long) As Long
Private Declare Function apiSetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'
'Keyboard
Private Declare Function apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As KBDLLHOOKSTRUCT, ByVal pSource As Long, ByVal cb As Long) As Long
Private Declare Function apiUnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
Private Declare Function apiCallNextKeyHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'
'Screen
Private Declare Function apiGetMonitorInfo Lib "user32" Alias "GetMonitorInfoA" (ByVal hMonitor As Long, ByRef lpmi As MONITORINFO) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiUnionRect Lib "user32" Alias "UnionRect" (ByRef lprcDst As RECT, ByRef lprcSrc1 As RECT, ByRef lprcSrc2 As RECT) As Long
Private Declare Function apiMonitorFromPoint Lib "user32" Alias "MonitorFromPoint" (ByVal x As Long, ByVal y As Long, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromRect Lib "user32" Alias "MonitorFromRect" (ByRef lprc As RECT, ByVal dwFlags As Long) As Long
Private Declare Function apiMonitorFromWindow Lib "user32" Alias "MonitorFromWindow" (ByVal hwnd As Long, ByVal dwFlags As Long) As Long
'
'Timer
Private Declare Function apiKillTimer Lib "user32" Alias "KillTimer" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
'
'string
Private Declare Function apiCharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As String
Private Declare Function apiCharUpper Lib "user32" Alias "CharUpperA" (ByVal lpsz As String) As String
Private Declare Function apiCompareString Lib "kernel32" Alias "CompareStringA" (ByVal Locale As Long, ByVal dwCmpFlags As Long, ByVal lpString1 As String, ByVal cchCount1 As Long, ByVal lpString2 As String, ByVal cchCount2 As Long) As Long
Private Declare Function apiGetThreadLocale Lib "kernel32" Alias "GetThreadLocale" () As Long
Private Declare Function apilstrcmp Lib "kernel32" Alias "lstrcmpA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function apilstrcmpi Lib "kernel32" Alias "lstrcmpiA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function apilstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As Any, ByVal lpString2 As Any) As Long
Private Declare Function apilstrcpyn Lib "kernel32" Alias "lstrcpynA" (ByVal lpString1 As Any, ByVal lpString2 As Any, ByVal iMaxLength As Long) As Long
Private Declare Function apilstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Any) As Long
'
'Variables
Public screens()                       As Screen
Public hMon                            As Long
'
Public hKey                            As Long
Public fWnd                            As Long
Public kSent                           As Long
'
Private Buffers(1 To MAX_BUFFER_COUNT) As SoundBufferInfo
Private hCallbackWnd                   As Long
Private pfnOldWindowProc               As Long
'Public Notifier As SoundManagerNotifier
'
Public Timercollection                 As New Collection
Public CTimersCol                      As New Collection
Private mTimersColCount                As Integer
'
'
'
Public swCount                         As Long 'stopwatch
Public swTic()                         As Currency
Public swToc()                         As Currency
Public swRunning()                     As Boolean
'
'
'
'
Public hThread                         As Long
Public hThreadID                       As Long
'
'
'
'
Public Enum SoundBufferFlags
    BufferFlagNone = 0
    BufferFlagAutoPlay = 1
    BufferFlagFreeWhenDone = 2
    BufferFlagNoNotify = 4
    BufferFlagInstant = BufferFlagAutoPlay Or BufferFlagFreeWhenDone Or BufferFlagNoNotify ' This one's just for convenience
End Enum
Public Enum SoundBufferStatus
    BufferError = -1
    BufferEmpty = 0
    BufferLoaded = 1
    BufferPlaying = 2
End Enum
Public Enum DialogResult
    IOK = 1
    ICANCEL = 2
    IABORT = 3
    IRETRY = 4
    IIGNORE = 5
    IYES = 6
    INO = 7
    ITRYAGAIN = 10
    ICONTINUE = 11
    IDTIMEOUT = 32000
    IDASYNC = 32001
End Enum
'
'
'
'
'
'Audio
Public Sub DestroySoundManager()
    FreeSound ALL_SOUND_BUFFERS  ' Do not forget to call this when you're done.
    If hCallbackWnd = 0 Then Exit Sub
    Call apiSetWindowLong(hCallbackWnd, GWL_WNDPROC, pfnOldWindowProc)
    Call apiDestroyWindow(hCallbackWnd)
End Sub
Private Function FindIndexFromHandle(ByVal hWaveOut As Long) As Long
    Dim BufferIndex As Long ' This should be optimized into a fast lookup routine, but for the small amount of buffers here it doesn't matter. (returns 0 if not found)
    For BufferIndex = 1 To MAX_BUFFER_COUNT
        If Buffers(BufferIndex).hWaveOut = hWaveOut Then
            FindIndexFromHandle = BufferIndex
            Exit Function
        End If
    Next
End Function
Public Function FreeBuffer() As Long
    Dim Index As Long ' Find the first free buffer (returns 0 if none found)
    For Index = 1 To MAX_BUFFER_COUNT
        If Buffers(Index).Status = BufferEmpty Then
            FreeBuffer = Index
            Exit Function
        End If
    Next
End Function
Public Function SoundStatus(ByVal BufferIndex As Long) As SoundBufferStatus
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then
        SoundStatus = BufferError
        Exit Function
    End If
    SoundStatus = Buffers(BufferIndex).Status
End Function
Public Function LoadSoundFile(ByVal BufferIndex As Long, ByVal FileName As String, Optional Flags As SoundBufferFlags = BufferFlagNone) As Boolean
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then Exit Function
    FreeSound BufferIndex  ' Free any sound currently in the buffer
    Dim InputHandle As Long
    Dim DataChunk   As MMCKINFO
    Dim ParentChunk As MMCKINFO
    Dim InputChunk  As MMCKINFO
    Dim EmptyChunk  As MMCKINFO
    Dim MinSize     As Long
    Dim WaveFCC     As Long
    Dim RiffFCC     As Long
    Dim WaveFormat  As WAVEFORMATEX
    MinSize = Len(WaveFormat) - 2
    WaveFCC = apimmioStringToFOURCC("WAVE", 0)
    RiffFCC = apimmioStringToFOURCC("RIFF", 0)
    InputHandle = apimmioOpen(FileName, ByVal 0, MMIO_ALLOCBUF Or MMIO_READ)
    If InputHandle = 0 Then
        MsgBox "Cannot open file"
        InputHandle = 0
        Exit Function
    End If
    If apimmioDescend(InputHandle, ParentChunk, ByVal 0, 0) <> 0 Then
        MsgBox "Cannot descend"
        GoTo CLEARUP_AND_EXIT
    End If
    If ParentChunk.ckid <> RiffFCC Or ParentChunk.fccType <> WaveFCC Then
        MsgBox "Incorrect format"
        GoTo CLEARUP_AND_EXIT
    End If
    InputChunk.ckid = apimmioStringToFOURCC("fmt", 0)
    If apimmioDescend(InputHandle, InputChunk, ParentChunk, MMIO_FINDCHUNK) <> 0 Then
        MsgBox "Could not find fmt chunk"
        GoTo CLEARUP_AND_EXIT
    End If
    If InputChunk.ckSize < MinSize Then
        MsgBox "Not enough data, only " & InputChunk.ckSize & ", wanted " & MinSize
        GoTo CLEARUP_AND_EXIT
    End If
    If apimmioRead(InputHandle, WaveFormat, LenB(WaveFormat)) < MinSize Then
        MsgBox "Not enough data read"
        GoTo CLEARUP_AND_EXIT
    End If
    If apimmioSeek(InputHandle, ParentChunk.dwDataOffset + 4, SEEK_SET) = -1 Then
        MsgBox "Could not seek"
        GoTo CLEARUP_AND_EXIT
    End If
    DataChunk = EmptyChunk
    DataChunk.ckid = apimmioStringToFOURCC("data", 0)
    If apimmioDescend(InputHandle, DataChunk, ParentChunk, MMIO_FINDCHUNK) <> 0 Then
        MsgBox "Could not descend into data"
        GoTo CLEARUP_AND_EXIT
    End If
    If hCallbackWnd = 0 Then ' Make sure we have a callback window
        If CreateCallbackWindow = False Then GoTo CLEARUP_AND_EXIT
    End If
    With Buffers(BufferIndex)
        ReDim .buf(0 To DataChunk.ckSize - 1)
        If apimmioRead(InputHandle, .buf(0), DataChunk.ckSize) <> DataChunk.ckSize Then
            MsgBox "Could not read full buffer"
            GoTo CLEARUP_AND_EXIT
        End If
        Call apiwaveOutOpen(.hWaveOut, WAVE_MAPPER, WaveFormat, hCallbackWnd, 0, CALLBACK_WINDOW)
        .hdr.lpData = VarPtr(.buf(0)) ' Prep the header
        .hdr.dwBufferLength = UBound(.buf) - LBound(.buf) + 1
        Call apiwaveOutPrepareHeader(.hWaveOut, .hdr, LenB(.hdr))
        .Status = BufferLoaded
        .Flags = Flags
        LoadSoundFile = True ' Send notification if needed
        ' If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundLoaded(BufferIndex)
        If Flags And BufferFlagAutoPlay Then ' Check for automatic playback
            PlaySound BufferIndex
        End If
    End With
CLEARUP_AND_EXIT:
    If InputHandle <> 0 Then
        Call apimmioClose(InputHandle, 0)
        InputHandle = 0
    End If
End Function
Public Sub FreeSound(ByVal BufferIndex As Long)
    If BufferIndex = ALL_SOUND_BUFFERS Then ' Handle the "all buffers" flag
        For BufferIndex = 1 To MAX_BUFFER_COUNT
            If Buffers(BufferIndex).Status <> BufferEmpty Then FreeSound BufferIndex
        Next
        Exit Sub
    End If
    If Buffers(BufferIndex).Status = BufferEmpty Then Exit Sub
    StopSound BufferIndex ' If the sound is currently playing then we need to stop it first
    With Buffers(BufferIndex)
        Call apiwaveOutUnprepareHeader(.hWaveOut, .hdr, LenB(.hdr))  ' Unprepare the header
        Call apiwaveOutClose(.hWaveOut) ' Close the handle
        .hWaveOut = 0
        Erase .buf ' Erase the buffer
        Call apiZeroMemory(.hdr, LenB(.hdr))
        .Status = BufferEmpty  ' Set the status to empty
        ' Debug.Print "Sound " & BufferIndex & " Freed"
        ' If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundUnloaded(BufferIndex)
    End With
End Sub
Public Sub StopSound(ByVal BufferIndex As Long)
    If BufferIndex = ALL_SOUND_BUFFERS Then ' Handle the "all buffers" flag
        For BufferIndex = 1 To MAX_BUFFER_COUNT
            StopSound BufferIndex
        Next
        Exit Sub
    End If
    With Buffers(BufferIndex)
        ' Debug.Print .status
        If .Status = BufferPlaying Then apiwaveOutReset (.hWaveOut)
    End With
End Sub
Public Function PlaySound(ByVal BufferIndex As Long) As Boolean
    If BufferIndex < 1 Or BufferIndex > MAX_BUFFER_COUNT Then Exit Function ' Check we've got a valid index
    StopSound BufferIndex
    With Buffers(BufferIndex)
        If .Status <> BufferLoaded Then Exit Function ' The sound must be loaded and not currently playing to be played
        If .hWaveOut = 0 Then Exit Function ' Ensure we have a valid handle loaded
        Call apiwaveOutWrite(.hWaveOut, .hdr, LenB(.hdr)) ' Write the data to the sound device
        .Status = BufferPlaying ' Update status
        ' If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundPlayStart(BufferIndex) ' Notify if required
    End With
    PlaySound = True ' All done!
End Function
Private Function CreateCallbackWindow() As Boolean
    If hCallbackWnd <> 0 Then Exit Function
    hCallbackWnd = apiCreateWindowEx(0, "STATIC", "Soundmanager Window", WS_POPUP Or SS_SIMPLE, 0, 0, 100, 20, 0, 0, 0, ByVal 0)
    If hCallbackWnd = 0 Then Exit Function
    pfnOldWindowProc = apiSetWindowLong(hCallbackWnd, GWL_WNDPROC, AddressOf CallbackWindowProc)
    CreateCallbackWindow = True
End Function
Private Function CallbackWindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Dim BufferIndex As Long
    Select Case uMsg
            '    Case MM_WOM_OPEN
            '    Case MM_WOM_CLOSE
        Case MM_WOM_DONE
            BufferIndex = FindIndexFromHandle(wParam)
            If BufferIndex <> 0 Then
                With Buffers(BufferIndex)
                    ' If Not (Notifier Is Nothing) And Not (CBool(.flags And BufferFlagNoNotify)) Then Call Notifier.SoundPlayEnd(BufferIndex)
                    .Status = BufferLoaded
                    If .Flags And BufferFlagFreeWhenDone Then FreeSound BufferIndex
                End With
            End If
    End Select
    CallbackWindowProc = apiCallWindowProc(pfnOldWindowProc, hwnd, uMsg, wParam, lParam)
End Function
'
'
'
'
'
'
'
'
'Input
Public Function Callback(ByVal Code As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    On Error Resume Next
    Static hStruct As KBDLLHOOKSTRUCT
    Call apiCopyMemory(hStruct, lParam, Len(hStruct))
    '    If fWnd <> 0 Then
    '        If apiGetForegroundWindow <> fWnd Then 'If foreground window is not where it's supposed to be
    '            Sleep (1) ''''''''''''''''''''''''Sleep for one millisecond and flush messages
    '            If apiSetForegroundWindow(fWnd) = 0 Then 'If foreground window cannot be set
    '                If hStruct.dwExtraInfo = -11 Then kSent = kSent - 1 'Uncount this key, since it's blocked, and has been sent from the Send function
    '                Callback = HC_GETNEXT ''''''''If foreground cannot be set then block key no matter if it's a user or internally sent from here
    '                Exit Function
    '            End If
    '        End If
    '    End If
    '    If hStruct.dwExtraInfo <> -11 Then '''''''If key press is not simulated from this module with -11 attached as extrainfomessage, then block this user key
    '        Callback = HC_GETNEXT ''''''''''''''''If key action and stroke blocked then get next key in the hook chain
    '        Exit Function
    '    End If
    Callback = apiCallNextKeyHookEx(hKey, Code, wParam, lParam) 'Call next key hook if no action
End Function
'
'
'
'
'
'Screen
Public Function MonitorEnumProc(ByVal hMonitor As Long, ByVal hdcMonitor As Long, ByRef lprcMonitor As RECT, ByRef dwData As Long) As Long
    On Error Resume Next
    Dim rects() As RECT
    ReDim Preserve rects(dwData)
    ReDim Preserve screens(dwData)
    rects(dwData) = lprcMonitor
    Dim MI As MONITORINFO
    MI.cbSize = Len(MI)
    If apiGetMonitorInfo(hMonitor, MI) <> 0 Then
        With rects(dwData)
            Dim r As New Rectangle
            r.Left = MI.rcMonitor.Left
            r.Top = MI.rcMonitor.Top
            r.Right = MI.rcMonitor.Right
            r.Bottom = MI.rcMonitor.Bottom
            r.Width_ = (MI.rcMonitor.Right - MI.rcMonitor.Left)
            r.Height = (MI.rcMonitor.Bottom - MI.rcMonitor.Top)
            r.Size.Width_ = r.Width_
            r.Size.Height_ = r.Height
            r.Location.x = r.Left
            r.Location.y = r.Top
            Dim rw As New Rectangle
            With rw
                .Left = MI.rcWork.Left
                .Top = MI.rcWork.Top
                .Right = MI.rcWork.Right
                .Bottom = MI.rcWork.Bottom
                .Width_ = (MI.rcWork.Right - MI.rcWork.Left)
                .Height = (MI.rcWork.Bottom - MI.rcWork.Top)
                .Size.Height_ = rw.Height
                .Size.Width_ = rw.Width_
                .Location.x = rw.Left
                .Location.y = rw.Top
            End With
            Dim sc As New Screen
            If hMon = 0 Then sc.Primary = CBool(MI.dwFlags = MONITORINFOF_PRIMARY)
            If hMon <> 0 And hMon = hMonitor Then sc.Primary = True
            Let sc.Handle = hMonitor
            Set sc.Bounds = r
            Set sc.WorkingArea = rw
            Set screens(dwData) = sc
            '            MsgBox screens(0).Bounds.Width_
        End With
    End If
    dwData = dwData + 1
    MonitorEnumProc = 1
End Function
'
'
'
'
'
'
'Strings
Private Sub DoLower()
    Dim s As String
    s = apiCharLower("Hello")
    MsgBox s
End Sub
Private Sub DoUpper()
    Dim s As String
    s = apiCharUpper("Hello")
    MsgBox s
End Sub
Public Function StringLength(ByVal txt As String) As Long
    Dim ret As Long
    ret = apilstrlen(txt)
End Function
Public Function StringCopyLength(ByVal source As String, ByRef target As String, ByVal Length As Long)
    Dim retval As Long
    target = Space(Length)  ' make room in target to receive the copied string
    retval = apilstrcpyn(target, source, Length) 'copies one or more characters from one string into another string, followed by a terminating null character. Either string, instead of being a "real" string, can also be merely a pointer to a string instead. The target string must already have enough space to receive the source string's contents along with the terminating null
    target = Left(target, Len(target) - 1)  ' remove the terminating null character
End Function
Public Function StringCopy(ByVal source As String, ByRef target As String)
    Dim retval As Long
    target = Space(Len(source))
    retval = apilstrcpy(target, source)
End Function
'Dim words(1 To 9) As String
'words(1) = "can't"
'words(2) = "cant"
'words(3) = "cannot"
'words(4) = "pants"
'words(5) = "co-op"
'words(6) = "coop"
'words(7) = "Denver"
'words(8) = "denver"
'words(9) = "denveR"
' Use a case-senitive comparison method to alphabetically sort
' nine words.  The sorting method simply compares each possible pair
' of words; if a pair is out of alphabetical order, they are switched.
Public Sub StringSort(ByRef s() As String)
    Dim tempstr As String
    Dim oc      As Integer
    Dim ic      As Integer
    Dim compval As Long
    ' Sort the strings, swapping any pairs which are out of order.
    For oc = 1 To 8  ' first string of the pair
        For ic = oc + 1 To 9  ' second string of the pair
            compval = apilstrcmp(s(oc), s(ic)) 'index
            ' If words(oc) is greater, swap them.
            If compval > 0 Then
                tempstr = s(oc)
                s(oc) = s(ic)
                s(ic) = tempstr
            End If
        Next ic
    Next oc
    Dim txt As String
    For oc = 1 To 9
        txt = txt & s(oc) & vbCrLf
    Next oc
    MsgBox txt
End Sub
Public Sub StringSorti()
    ' Use a case-insenitive comparison method to alphabetically sort
    ' nine words.  The sorting method simply compares each possible pair
    ' of words; if a pair is out of alphabetical order, they are switched.
    ' (Note how this sort will seemingly arrange the "Denver" trio in
    ' a random order, depending on how the search loops play out --
    ' the three strings are equal in the eyes of the function and therefore
    ' not sorted relative to each other.)
    Dim words(1 To 9) As String  ' the words to sort
    Dim tempstr       As String  ' buffer used to swap strings
    Dim oc            As Integer, ic As Integer  ' counter variables
    Dim compval       As Long  ' result of comparison
    ' Load the nine strings into the array.
    words(1) = "can't"
    words(2) = "cant"
    words(3) = "cannot"
    words(4) = "pants"
    words(5) = "co-op"
    words(6) = "coop"
    words(7) = "Denver"
    words(8) = "denver"
    words(9) = "denveR"
    ' Sort the strings, swapping any pairs which are out of order.
    For oc = 1 To 8  ' first string of the pair
        For ic = oc + 1 To 9  ' second string of the pair
            compval = apilstrcmpi(words(oc), words(ic)) 'If the function returns a negative value, the first string is less than the second string (i.e., the first string comes before the second string in alphabetical order). If the function returns 0, the two strings are equal. If the function returns a positive value, the first string is greater than the second string. In any case, the actual return value is the difference of the first unequal characters encountered.
            If compval > 0 Then
                tempstr = words(oc)
                words(oc) = words(ic)
                words(ic) = tempstr
            End If
        Next ic
    Next oc
    Dim txt As String
    For oc = 1 To 9
        txt = txt & words(oc) & vbCrLf
    Next oc
    MsgBox txt
End Sub
Public Sub StringSortCompareString()
    Dim words(1 To 9) As String
    Dim buff          As String
    Dim oc            As Integer, ic As Integer
    Dim v             As Long
    Dim thlocale      As Long
    thlocale = apiGetThreadLocale
    words(1) = "can't"
    words(2) = "cant"
    words(3) = "cannot"
    words(4) = "pants"
    words(5) = "co-op"
    words(6) = "coop"
    words(7) = "Denver"
    words(8) = "denver"
    words(9) = "denveR"
    ' Sort the strings, swapping any pairs which are out of order.
    For oc = 1 To 8  ' first string of the pair
        For ic = oc + 1 To 9  ' second string of the pair
            v = apiCompareString(thlocale, SORT_STRINGSORT, words(oc), Len(words(oc)), words(ic), Len(words(ic)))
            If v = CSTR_GREATER_THAN Then ' If words(oc) is greater, swap them.
                buff = words(oc)
                words(oc) = words(ic)
                words(ic) = buff
            End If
        Next ic
    Next oc
    Dim txt As String
    For oc = 1 To 9
        txt = txt & words(oc) & vbCrLf
    Next oc
    MsgBox txt
End Sub
'
'
'
'
'
'
'
'Timer
Public Sub TimerProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal idEvent As Long, ByVal dwTime As Long)
    Dim t As Timer
    Dim C As Timers
    On Error Resume Next
    Set t = Timercollection("id:" & idEvent)
    If t Is Nothing Then
        Call apiKillTimer(0, idEvent)
    Else
        If t.ParentsColKey > 0 Then
            Set C = CTimersCol("key:" & t.ParentsColKey)
            If C Is Nothing Then
                Call apiKillTimer(0, idEvent)
            Else
                C.RaiseTimer_Event t.Index
            End If
        Else
            t.RaiseTimer_Event
        End If
    End If
    Set t = Nothing
End Sub
Public Function RegisterTimerCollection(ByRef C As Timers) As Integer
    Dim Key As String
    mTimersColCount = mTimersColCount + 1
    Key = "key:" & mTimersColCount
    CTimersCol.Add C, Key
    RegisterTimerCollection = mTimersColCount
End Function
'
'
'
Public Sub AsyncThread()
    'Let this thread sleep for 10 seconds
    Threading.Thread.Sleep 10000
    hThread = 0
    hThreadID = 0
    MsgBox "thread complete"
End Sub

