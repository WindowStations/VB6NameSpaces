VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Sendkeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum Keys
    VK_None = 0
    VK_LButton = 1
    VK_RButton = 2
    VK_Cancel = 3
    VK_MButton = 4
    VK_XButton1 = 5
    VK_XButton2 = 6
    VK_LButton_XButton2 = 7
    VK_Back = 8
    VK_Tab = 9
    VK_LineFeed = 10
    VK_LButton_LineFeed = 11
    VK_Clear = 12
    VK_Return = 13
    VK_RButton_Clear = 14
    VK_RButton_Return = 15
    VK_ShiftKey = 16
    VK_ControlKey = 17
    VK_Menu = 18
    VK_Pause = 19
    VK_Capital = 20
    VK_KanaMode = 21
    VK_RButton_Capital = 22
    VK_JunjaMode = 23
    VK_FinalMode = 24
    VK_HanjaMode = 25
    VK_RButton_FinalMode = 26
    VK_Escape = 27
    VK_IMEConvert = 28
    VK_IMENonconvert = 29
    VK_IMEAceept = 30
    VK_IMEModeChange = 31
    VK_Space = 32
    VK_PageUp = 33
    VK_Next = 34
    VK_End = 35
    VK_Home = 36
    VK_Left = 37
    VK_Up = 38
    VK_Right = 39
    VK_Down = 40
    VK_Select = 41
    VK_Print = 42
    VK_Execute = 43
    VK_PrintScreen = 44
    VK_Insert = 45
    VK_Delete = 46
    VK_Help = 47
    VK_D0 = 48
    VK_D1 = 49
    VK_D2 = 50
    VK_D3 = 51
    VK_D4 = 52
    VK_D5 = 53
    VK_D6 = 54
    VK_D7 = 55
    VK_D8 = 56
    VK_D9 = 57
    VK_RButton_D8 = 58
    VK_RButton_D9 = 59
    VK_MButton_D8 = 60
    VK_MButton_D9 = 61
    VK_XButton2_D8 = 62
    VK_XButton2_D9 = 63
    VK_64 = 64
    VK_A = 65
    VK_B = 66
    VK_C = 67
    VK_D = 68
    VK_E = 69
    VK_F = 70
    VK_G = 71
    VK_H = 72
    VK_I = 73
    VK_J = 74
    VK_K = 75
    VK_L = 76
    VK_M = 77
    VK_N = 78
    VK_O = 79
    VK_P = 80
    VK_Q = 81
    VK_R = 82
    VK_S = 83
    VK_T = 84
    VK_U = 85
    VK_V = 86
    VK_W = 87
    VK_X = 88
    VK_Y = 89
    VK_Z = 90
    VK_LWin = 91
    VK_RWin = 92
    VK_Apps = 93
    VK_RButton_RWin = 94
    VK_Sleep = 95
    VK_NumPad0 = 96
    VK_NumPad1 = 97
    VK_NumPad2 = 98
    VK_NumPad3 = 99
    VK_NumPad4 = 100
    VK_NumPad5 = 101
    VK_NumPad6 = 102
    VK_NumPad7 = 103
    VK_NumPad8 = 104
    VK_NumPad9 = 105
    VK_Multiply = 106
    VK_Add = 107
    VK_Separator = 108
    VK_Subtract = 109
    VK_Decimal = 110
    VK_Divide = 111
    VK_F1 = 112
    VK_F2 = 113
    VK_F3 = 114
    VK_F4 = 115
    VK_F5 = 116
    VK_F6 = 117
    VK_F7 = 118
    VK_F8 = 119
    VK_F9 = 120
    VK_F10 = 121
    VK_F11 = 122
    VK_F12 = 123
    VK_F13 = 124
    VK_F14 = 125
    VK_F15 = 126
    VK_F16 = 127
    VK_F17 = 128
    VK_F18 = 129
    VK_F19 = 130
    VK_F20 = 131
    VK_F21 = 132
    VK_F22 = 133
    VK_F23 = 134
    VK_F24 = 135
    VK_Back_F17 = 136
    VK_Back_F18 = 137
    VK_Back_F19 = 138
    VK_Back_F20 = 139
    VK_Back_F21 = 140
    VK_Back_F22 = 141
    VK_Back_F23 = 142
    VK_Back_F24 = 143
    VK_NumLock = 144
    VK_Scroll = 145
    VK_RButton_NumLock = 146
    VK_RButton_Scroll = 147
    VK_MButton_NumLock = 148
    VK_MButton_Scroll = 149
    VK_XButton2_NumLock = 150
    VK_XButton2_Scroll = 151
    VK_Back_NumLock = 152
    VK_Back_Scroll = 153
    VK_LineFeed_NumLock = 154
    VK_LineFeed_Scroll = 155
    VK_Clear_NumLock = 156
    VK_Clear_Scroll = 157
    VK_RButton_Clear_NumLock = 158
    VK_RButton_Clear_Scroll = 159
    VK_LShiftKey = 160
    VK_RShiftKey = 161
    VK_LControlKey = 162
    VK_RControlKey = 163
    VK_LMenu = 164
    VK_RMenu = 165
    VK_BrowserBack = 166
    VK_BrowserForward = 167
    VK_BrowserRefresh = 168
    VK_BrowserStop = 169
    VK_BrowserSearch = 170
    VK_BrowserFavorites = 171
    VK_BrowserHome = 172
    VK_VolumeMute = 173
    VK_VolumeDown = 174
    VK_VolumeUp = 175
    VK_MediaNextTrack = 176
    VK_MediaPreviousTrack = 177
    VK_MediaStop = 178
    VK_MediaPlayPause = 179
    VK_LaunchMail = 180
    VK_SelectMedia = 181
    VK_LaunchApplication1 = 182
    VK_LaunchApplication2 = 183
    VK_Back_MediaNextTrack = 184
    VK_Back_MediaPreviousTrack = 185
    VK_Oem1 = 186
    VK_Oemplus = 187
    VK_Oemcomma = 188
    VK_OemMinus = 189
    VK_OemPeriod = 190
    VK_OemQuestion = 191
    VK_Oemtilde = 192
    VK_LButton_Oemtilde = 193
    VK_RButton_Oemtilde = 194
    VK_Cancel_Oemtilde = 195
    VK_MButton_Oemtilde = 196
    VK_XButton1_Oemtilde = 197
    VK_XButton2_Oemtilde = 198
    VK_LButton_XButton2_Oemtilde = 199
    VK_Back_Oemtilde = 200
    VK_Tab_Oemtilde = 201
    VK_LineFeed_Oemtilde = 202
    VK_LButton_LineFeed_Oemtilde = 203
    VK_Clear_Oemtilde = 204
    VK_Return_Oemtilde = 205
    VK_RButton_Clear_Oemtilde = 206
    VK_RButton_Return_Oemtilde = 207
    VK_ShiftKey_Oemtilde = 208
    VK_ControlKey_Oemtilde = 209
    VK_Menu_Oemtilde = 210
    VK_Pause_Oemtilde = 211
    VK_Capital_Oemtilde = 212
    VK_KanaMode_Oemtilde = 213
    VK_RButton_Capital_Oemtilde = 214
    VK_JunjaMode_Oemtilde = 215
    VK_FinalMode_Oemtilde = 216
    VK_HanjaMode_Oemtilde = 217
    VK_RButton_FinalMode_Oemtilde = 218
    VK_OemOpenBrackets = 219
    VK_Oem5 = 220
    VK_Oem6 = 221
    VK_Oem7 = 222
    VK_Oem8 = 223
    VK_Space_Oemtilde = 224
    VK_PageUp_Oemtilde = 225
    VK_OemBackslash = 226
    VK_LButton_OemBackslash = 227
    VK_Home_Oemtilde = 228
    VK_ProcessKey = 229
    VK_MButton_OemBackslash = 230
    VK_Packet = 231
    VK_Down_Oemtilde = 232
    VK_Select_Oemtilde = 233
    VK_Back_OemBackslash = 234
    VK_Tab_OemBackslash = 235
    VK_PrintScreen_Oemtilde = 236
    VK_Back_ProcessKey = 237
    VK_Clear_OemBackslash = 238
    VK_Back_Packet = 239
    VK_D0_Oemtilde = 240
    VK_D1_Oemtilde = 241
    VK_ShiftKey_OemBackslash = 242
    VK_ControlKey_OemBackslash = 243
    VK_D4_Oemtilde = 244
    VK_ShiftKey_ProcessKey = 245
    VK_Attn = 246
    VK_Crsel = 247
    VK_Exsel = 248
    VK_EraseEof = 249
    VK_Play = 250
    VK_Zoom = 251
    VK_NoName = 252
    VK_Pa1 = 253
    VK_OemClear = 254
    VK_LButton_OemClear = 255
End Enum
Const ASFW_ANY        As Long = -1
Const GW_HWNDNEXT     As Long = 2
Const GA_ROOT         As Long = 2
Const HC_GETNEXT      As Long = 1
Const KEYEVENTF_KEYUP As Long = 2
Const LSFW_LOCK       As Long = 1
Const LSFW_UNLOCK     As Long = 2
Const MAX_PATH = 260
Const NEGATIVE As Long = -1
Const PROCESS_QUERY_INFORMATION = 1024
Const PROCESS_VM_READ = 16
Const SW_HIDE                 As Long = 0
Const SW_SHOWNORMAL           As Long = 1
Const SW_SHOWMINIMIZED        As Long = 2
Const WH_KEYBOARD_LL          As Long = 13
Const WM_SETTEXT              As Long = 12
Const WM_KEYDOWN              As Long = 256
Const WM_KEYUP                As Long = 257
Const HWND_DESKTOP            As Long = 0
Const HWND_NOTOPMOST          As Long = -2
Const HWND_TOP                As Long = 0
Const HWND_TOPMOST            As Long = -1
Const GW_HWNDFIRST            As Long = 0
Const MOUSEEVENTF_MOVE        As Long = 1
Const MOUSEEVENTF_LEFTDOWN    As Long = 2
Const MOUSEEVENTF_LEFTUP      As Long = 4
Const MOUSEEVENTF_RIGHTDOWN   As Long = 8
Const MOUSEEVENTF_RIGHTUP     As Long = 16
Const MOUSEEVENTF_MIDDLEDOWN  As Long = 32
Const MOUSEEVENTF_MIDDLEUP    As Long = 64
Const MOUSEEVENTF_XDOWN       As Long = 128
Const MOUSEEVENTF_XUP         As Long = 256
Const MOUSEEVENTF_WHEEL       As Long = 2048
Const MOUSEEVENTF_VIRTUALDESK As Long = 16384
Const MOUSEEVENTF_ABSOLUTE    As Long = 32768
Const QS_ALLQUEUE             As Long = 511
Const SM_CXSCREEN             As Long = 0
Const SM_CYSCREEN             As Long = 1
Const SM_FULLSCREEN           As Long = 65535
Const SWP_NOSIZE              As Long = 1
Const SWP_NOMOVE              As Long = 2
Const SWP_NOACTIVATE          As Long = 16
Const SWP_SHOWWINDOW          As Long = 64
Const WM_COMMAND              As Long = 273 '
Const WM_LBUTTONDBLCLK        As Long = 515
Const WM_LBUTTONDOWN          As Long = 513
Const WM_LBUTTONUP            As Long = 514
Const WM_MBUTTONDBLCLK        As Long = 521
Const WM_MBUTTONDOWN          As Long = 519
Const WM_MBUTTONUP            As Long = 520
Const WM_RBUTTONDBLCLK        As Long = 518
Const WM_RBUTTONDOWN          As Long = 516
Const WM_RBUTTONUP            As Long = 517
Private Const PM_REMOVE   As Long = &H1
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type
Private Type RECT
    rLeft As Long
    rTop As Long
    rRight As Long
    rBottom As Long
End Type
Private Type GUITHREADINFO
    cbSize As Long
    Flags As Long
    hWndActive As Long
    hWndFocus As Long
    hWndCapture As Long
    hWndMenuOwner As Long
    hWndMoveSize As Long
    hWndCaret As Long
    rcCaret As RECT
End Type
Private Type ITEMINFO
    Width As Long
    Height As Long
    Right As Long
    Left As Long
    Top As Long
    Bottom As Long
    Center As POINTAPI
End Type
Private Type MENUINFO
    hwnd As Long
    hMenu As Long
    hSubMenu As Long
End Type
Private Type WINSTATE
    IsIconic As Boolean
    IsHidden As Boolean
    IsDisabled As Boolean
    IsChildHidden As Boolean
    IsChildDisabled As Boolean
End Type
Private Declare Function apiAllowSetForegroundWindow Lib "user32" Alias "AllowSetForegroundWindow" (ByVal dwProcessId As Long) As Boolean
Private Declare Function apiAttachThreadInput Lib "user32" Alias "AttachThreadInput" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Private Declare Function apiCallNextKeyHookEx Lib "user32" Alias "CallNextHookEx" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function apiCharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
Private Declare Function apiChildWindowFromPointEx Lib "user32" Alias "ChildWindowFromPointEx" (ByVal hWndParent As Long, ByVal ptx As Long, ByVal pty As Long, ByVal uFlags As Long) As Long
Private Declare Function apiCloseHandle Lib "kernel32" Alias "CloseHandle" (ByVal Handle As Long) As Long
Private Declare Function apiEnableWindow Lib "user32" Alias "EnableWindow" (ByVal hwnd As Long, ByVal fEnable As Boolean) As Boolean
Private Declare Function apiEnumProcessModules Lib "PSAPI" Alias "EnumProcessModules" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function apiEnumProcesses Lib "PSAPI" Alias "EnumProcesses" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function apiFindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function apiFindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function apiGetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Private Declare Function apiGetAncestor Lib "user32" Alias "GetAncestor" (ByVal hwnd As Long, ByVal gaFlags As Long) As Long
Private Declare Function apiGetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiGetCurrentThreadId Lib "kernel32" Alias "GetCurrentThreadId" () As Long
Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetForegroundWindow Lib "user32" Alias "GetForegroundWindow" () As Long
Private Declare Function apiGetGUIThreadInfo Lib "user32" Alias "GetGUIThreadInfo" (ByVal dwThreadId As Long, ByRef lpGUIThreadInfo As GUITHREADINFO) As Long
Private Declare Function apiGetInputState Lib "user32" Alias "GetInputState" () As Long
Private Declare Function apiGetKeyState Lib "user32" Alias "GetKeyState" (ByVal vKey As Long) As Long
Private Declare Function apiGetModuleFileNameExA Lib "PSAPI" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function apiGetParent Lib "user32" Alias "GetParent" (ByVal hwnd As Long) As Long
Private Declare Function apiGetQueueStatus Lib "user32" Alias "GetQueueStatus" (ByVal fuFlags As Long) As Long
Private Declare Function apiGetTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Private Declare Function apiGetTopWindow Lib "user32" Alias "GetTopWindow" (ByVal hwnd As Long) As Long
Private Declare Function apiGetWindow Lib "user32" Alias "GetWindow" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function apiGetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function apiGetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function apiGetWindowThreadProcessId Lib "user32" Alias "GetWindowThreadProcessId" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function apiIsIconic Lib "user32" Alias "IsIconic" (ByVal hwnd As Long) As Boolean
Private Declare Function apiIsWindow Lib "user32" Alias "IsWindow" (ByVal hwnd As Long) As Boolean
Private Declare Function apiIsWindowEnabled Lib "user32" Alias "IsWindowEnabled" (ByVal hwnd As Long) As Boolean
Private Declare Function apiIsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal hwnd As Long) As Long
Private Declare Function apikeybd_event Lib "user32" Alias "keybd_event" (ByVal vKey As Long, ByVal bScan As Long, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Boolean
Private Declare Function apiLockSetForegroundWindow Lib "user32" Alias "LockSetForegroundWindow" (ByVal uLockCode As Long) As Boolean
Private Declare Function apiLockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hWndLock As Long) As Long
Private Declare Function apiOpenProcess Lib "kernel32" Alias "OpenProcess" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function apiPostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Boolean
Private Declare Function apiSendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Boolean
Private Declare Function apiSetFocus Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private Declare Function apiSetForegroundWindow Lib "user32" Alias "SetForegroundWindow" (ByVal hwnd As Long) As Long
Private Declare Function apiSetWindowsKeyHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Boolean
Private Declare Function apiSwitchToThread Lib "kernel32" Alias "SwitchToThread" () As Long
Private Declare Function apiUnhookWindowsHookEx Lib "user32" Alias "UnhookWindowsHookEx" (ByVal hHook As Long) As Long
Private Declare Function apiVkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar2 As Long) As Long
Private Declare Function apiWaitForInputIdle Lib "user32" Alias "WaitForInputIdle" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function apiWindowFromPoint Lib "user32" Alias "WindowFromPoint" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function apiGetMenu Lib "user32" Alias "GetMenu" (ByVal hwnd As Long) As Long
Private Declare Function apiGetMenuItemCount Lib "user32" Alias "GetMenuItemCount" (ByVal hMenu As Long) As Long
Private Declare Function apiGetMenuItemID Lib "user32" Alias "GetMenuItemID" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function apiGetMenuItemRect Lib "user32" Alias "GetMenuItemRect" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal uItem As Long, ByRef lprcItem As RECT) As Long
Private Declare Function apiGetMenuString Lib "user32" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function apiGetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
Private Declare Function apiGetSubMenu Lib "user32" Alias "GetSubMenu" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function apiGetSystemMetrics Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, ByRef lpRect As RECT) As Boolean
Private Declare Function apiIsMenu Lib "user32" Alias "IsMenu" (ByVal hMenu As Long) As Boolean
Private Declare Function apimouse_event Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long) As Boolean
Private Declare Function apiMoveWindow Lib "user32" Alias "MoveWindow" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Boolean) As Boolean
Private Declare Function apiSetCursorPos Lib "user32" Alias "SetCursorPos" (ByVal x As Long, ByVal y As Long) As Boolean
Private Declare Function apiSetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function apiPeekMessage Lib "user32" Alias "PeekMessageA" (ByRef lpMsg As MSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function apiTranslateMessage Lib "user32" Alias "TranslateMessage" (ByRef lpMsg As MSG) As Long
Private Declare Function apiDispatchMessage Lib "user32" Alias "DispatchMessageA" (ByRef lpMsg As MSG) As Long

Friend Function Flush(Optional ByVal checkQueue As Boolean = True, Optional ByVal returnQueued As Boolean = False) As Long
    Flush = -1
    If checkQueue = True And IsQueued = False Then Exit Function
    Dim d As Long
    d = DoEvents_
    If returnQueued = False Then Flush = d
    If returnQueued = True Then Flush = apiGetQueueStatus(QS_ALLQUEUE)
End Function 'If specified this function returns true if there are still message in the queue, otherwise false will return the number of messages dispatched to applications
Friend Function GetQueueStatus(ByVal QS_ As Long) As Long
    On Error Resume Next
    GetQueueStatus = apiGetQueueStatus(QS_ALLQUEUE)
End Function
Friend Function GetInputState() As Long
    On Error Resume Next
    GetInputState = apiGetInputState
End Function
Friend Function IsQueued() As Boolean
    If apiGetInputState <> 0 Then IsQueued = True: Exit Function
    If apiGetQueueStatus(QS_ALLQUEUE) <> 0 Then IsQueued = True: Exit Function
    IsQueued = False
End Function
'The alternative function for DoEvents:
Friend Function DoEvents_() As Long
    Dim i    As Long
    Dim cMsg As MSG
    Do
        On Error GoTo abort
        If apiPeekMessage(cMsg, 0, 0, 0, PM_REMOVE) = 0 Then Exit Do
        Call apiTranslateMessage(cMsg)
        Call apiDispatchMessage(cMsg)
        i = i + 1
        If i = 400 Then Exit Do 'Let's not hang for to long, just in case, we exit at upper limit
    Loop
abort:
    DoEvents_ = i
End Function 'returns the number of dispatched messages before finding an empty queue
