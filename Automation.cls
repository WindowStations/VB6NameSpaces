VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Automation"
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
'Attribute VB_Name = "Automation"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'ThisWorkbook.VBProject.References.AddFromFile "C:\Windows\SysWOW64\UIAutomationCore.dll"
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
Const WM_COMMAND              As Long = 273
Const WM_LBUTTONDBLCLK        As Long = 515
Const WM_LBUTTONDOWN          As Long = 513
Const WM_LBUTTONUP            As Long = 514
Const WM_MBUTTONDBLCLK        As Long = 521
Const WM_MBUTTONDOWN          As Long = 519
Const WM_MBUTTONUP            As Long = 520
Const WM_RBUTTONDBLCLK        As Long = 518
Const WM_RBUTTONDOWN          As Long = 516
Const WM_RBUTTONUP            As Long = 517
Private Type EVENTCLICK
    mUp As Boolean
    mDown As Boolean
    mButtons As Long
    x As Long
    y As Long
    WFOCUS As WINFOCUS
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
Private Type POINTAPI
    x As Long
    y As Long
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
Private Type tagPOINT
    x As Long
    y As Long
End Type
Private Declare Function apiElementFromPoint Lib "UIAutomationCore" Alias "ElementFromPoint" (ByRef tp As tagPOINT) As IUIAutomationElement
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
Private Declare Function apiShowWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
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
Private Declare Sub apiSleep Lib "kernel32" Alias "Sleep" (ByVal dwMilliseconds As Long)
Private Enum Buttons
    LeftClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
    LeftDoubleClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
    MiddleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
    MiddleDoubleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP + MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
    Move = MOUSEEVENTF_MOVE
    MoveAbsolute = MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE
    RightClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
    RightDoubleClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
    VirtualDesk = MOUSEEVENTF_VIRTUALDESK
    Wheel = MOUSEEVENTF_WHEEL
    xClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
    xDoubleClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP + MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
End Enum
' Gets the handle of the current foreground window and the handle of the child window with keyboard focus.
' If keyFocus is False then the handles retrieved are the windows under the cursor with mouse focus.
' Returns the WINFOCUS structure containing the current focus handles. Depends on the keyFocus parameter.
' "showNames"(Optional) Shows the title and class name of the windows,
' to be used with the GetWinHandles function.
' "keyFocus"(Optional) Returns the handle of the main window, and child window with keyboard focus.
' If false then the function returns the handles of the windows currently under the cursor.</param>
Friend Function GetWinFocus(Optional ByVal showNames As Boolean = False, Optional ByVal keyFocus As Boolean = True) As WINFOCUS
    On Error Resume Next
    If keyFocus = True Then ''''''''''''''''''''''Keyboard focus
        Dim g As GUITHREADINFO '''''''''''''''''''Dimension a thread input structure
        g.cbSize = Len(g) ''''''''''''''''''''''''Initialize structure
        Call apiGetGUIThreadInfo(apiGetWindowThreadProcessId(0, 0), g) 'Retrieve information about the active window
        GetWinFocus.Foreground = g.hWndActive ''''Set handle of active foreground
        GetWinFocus.Focus = g.hWndFocus ''''''''''Set handle of focus window
    Else '''''''''''''''''''''''''''''''''''''''''Mouse focus instead
        Dim p As POINTAPI  '''''''''''''''''''''''Dimension a point for the mouse position
        Call apiGetCursorPos(p)  '''''''''''''''''Get current cursor position
        GetWinFocus.Focus = apiWindowFromPoint(p.x, p.y) 'Get handle of window under cursor
        GetWinFocus.Foreground = apiGetAncestor(GetWinFocus.Focus, GA_ROOT) 'Try to get it's ancestor
        If GetWinFocus.Foreground = 0 Then GetWinFocus.Foreground = GetWinFocus.Focus 'If no ancestor then set main focus to child focus
    End If
    If showNames = True Then Call GetWinAncestory(GetWinFocus.Focus, True) 'Set ancestory using focus window
End Function
' Sends text directly to a window.  Use GetWinHandles or GetWinFocus to get a structure.
' Returns true if successful.
' sText The text to send.
' wFocus The structure of the window.  Use GetWinHandles or GetWinFocus to get a structure.
Friend Function Text(ByVal sText As String, ByRef WFOCUS As WINFOCUS) As Boolean
    On Error Resume Next
    Text = apiSendMessage(WFOCUS.Focus, WM_SETTEXT, 0, sText) 'Return the full result of SendMessage
End Function
' Gets the handle of the main window and the handle of the child window that is to recieve
' keyboard or mouse focus, by using the specified title, class name, or index.
' Returns the WINFOCUS structure containing the specified focus handles.
' "wName1"The main window title or class name.
' "wIndex1"(Optional) The main window index.
' "wName2"(Optional) The first child window.
' "wIndex2"(Optional) The first child index.
' "wName3"(Optional) The second child window.
' "wIndex3"(Optional) The second child index.
' "wName4"(Optional) The third child window.
' "wIndex4"(Optional) The third child index.
' "wName5"(Optional) The fourth child window.
' "wIndex5"(Optional) The fourth child index.
' "wName6"(Optional) The fifth child window.
' "wIndex6"(Optional) The fifth child index.
' "wName7"(Optional) The sixth child window.
' "wIndex7"(Optional) The sixth child index.
' "wName8"(Optional) The seventh child window.
' "wIndex8"(Optional) The seventh child index.
' "wName9"(Optional) The eighth child window.
' "wIndex9"(Optional) The eighth child index.
' "wName10"(Optional) The ninth child window.
' "wIndex10"(Optional) The ninth child index.
Friend Function GetWinHandles(ByVal wName1 As String, Optional ByVal wIndex1 As Long = 1, Optional ByVal wName2 As String = " ", Optional ByVal wIndex2 As Long = 1, Optional ByVal wName3 As String = " ", Optional ByVal wIndex3 As Long = 1, Optional ByVal wName4 As String = " ", Optional ByVal wIndex4 As Long = 1, Optional ByVal wName5 As String = " ", Optional ByVal wIndex5 As Long = 1, Optional ByVal wName6 As String = " ", Optional ByVal wIndex6 As Long = 1, Optional ByVal wName7 As String = " ", Optional ByVal wIndex7 As Long = 1, Optional ByVal wName8 As String = " ", Optional ByVal wIndex8 As Long = 1, Optional ByVal wName9 As String = " ", Optional ByVal wIndex9 As Long = 1, Optional ByVal wName10 As String = " ", Optional ByVal wIndex10 As Long = 1) As WINFOCUS
    On Error Resume Next
    Dim hwnd     As Long
    Dim cwnd     As Long
    Dim i        As Long
    Dim wn       As String
    Dim wf       As WINFOCUS
    Dim wNames() As String '''''''''''''''''''''''Dimension an array string the size of max index
    If wName1 <> " " Then i = 1: ReDim wNames(1) 'Set i to index of parameters
    If wName2 <> " " Then i = 3: ReDim wNames(3)
    If wName3 <> " " Then i = 5: ReDim wNames(5)
    If wName4 <> " " Then i = 7: ReDim wNames(7)
    If wName5 <> " " Then i = 9: ReDim wNames(9)
    If wName6 <> " " Then i = 11: ReDim wNames(11)
    If wName7 <> " " Then i = 13: ReDim wNames(13)
    If wName8 <> " " Then i = 15: ReDim wNames(15)
    If wName9 <> " " Then i = 17: ReDim wNames(17)
    If wName10 <> " " Then i = 19: ReDim wNames(19)
    If wName1 <> " " Then wNames(0) = wName1: wNames(1) = CStr(wIndex1) 'Set array elements
    If wName2 <> " " Then wNames(2) = wName2: wNames(3) = CStr(wIndex2)
    If wName3 <> " " Then wNames(4) = wName3: wNames(5) = CStr(wIndex3)
    If wName4 <> " " Then wNames(6) = wName4: wNames(7) = CStr(wIndex4)
    If wName5 <> " " Then wNames(8) = wName5: wNames(9) = CStr(wIndex5)
    If wName6 <> " " Then wNames(10) = wName6: wNames(11) = CStr(wIndex6)
    If wName7 <> " " Then wNames(12) = wName7: wNames(13) = CStr(wIndex7)
    If wName8 <> " " Then wNames(14) = wName8: wNames(15) = CStr(wIndex8)
    If wName9 <> " " Then wNames(16) = wName9: wNames(17) = CStr(wIndex9)
    If wName10 <> " " Then wNames(18) = wName10: wNames(19) = CStr(wIndex10)
    hwnd = apiFindWindow(vbNullString, wName1) '''Look for handle from title
    If hwnd = 0 Then hwnd = apiFindWindow(wName1, vbNullString) 'If not found by title, then look for handle from class name
    If hwnd <> 0 And CInt(wNames(1)) > 1 Then  '''If searching for window by index, ie more than 1
        Dim nxtwnd As Long
        nxtwnd = hwnd ''''''''''''''''''''''''''''Initialize handle containing the next window in the top level z-order
        Dim Index As Long
        Index = 1 ''''''''''''''''''''''''''''''''Initialize index to 1, since handle was found.
        Dim n As WINNAME  ''''''''''''''''''''''''Create structure for window names
        Do '''''''''''''''''''''''''''''''''''''''Do loop to look for handles matching paramaters
            nxtwnd = apiGetWindow(hwnd, GW_HWNDNEXT) 'Set next window
            If nxtwnd = 0 Then Exit Do '''''''''''Eject if there are no more top level windows
            If wNames(0) <> "" Then ''''''''''''''If title or class name was given
                n = GetWinName(nxtwnd, True, True) 'Get title and classname of the window
                If n.lpText = wNames(0) Or n.lpClass = wNames(0) Then 'If title or class name matches first wNames parameter
                    Index = Index + 1 ''''''''''''Increment Index for the matching window
                    hwnd = nxtwnd ''''''''''''''''Set handle to the handle of the matching window
                    If Index >= CInt(wNames(1)) Then Exit Do 'If the specified index has been reached then exit with last matching handle
                End If
            Else '''''''''''''''''''''''''''''''''Then "" indicates the search is by index only
                Index = Index + 1 ''''''''''''''''Increment Index for the matching window
                hwnd = nxtwnd ''''''''''''''''''''Set handle to the handle of the matching window
                If Index >= CInt(wNames(1)) Then Exit Do 'If index specified has been reached then exit with last matching handle
            End If
        Loop
    End If
    If hwnd = 0 Then '''''''''''''''''''''''''''''If not found by title or class name, then look for special commands
        Dim pid As Long ''''''''''''''''''''''''''Dimension process identification
        Dim p   As POINTAPI  '''''''''''''''''''''''Dimension a point structure
        wn = wName1 ''''''''''''''''''''''''''''''Set string to name for lower case conversion
        Call apiCharLower(wn)  '''''''''''''''''''Does not need to be case sensitive
        If wn = "{focus}" Then '''''''''''''''''''If focus window is specified
            hwnd = GetWinFocus.Focus '''''''''''''Get the handle of the focus window
        ElseIf wn = "{foreground}" Then ''''''''''If foreground window is specified
            hwnd = apiGetForegroundWindow ''''''''Get the handle of the foreground window
        ElseIf wn = "{active}" Then ''''''''''''''If active window is specified
            hwnd = apiGetActiveWindow ''''''''''''Get the handle of the active window
        ElseIf wn = "{desktop}" Then '''''''''''''If desktop window is specified
            hwnd = apiGetDesktopWindow() '''''''''Get the handle of the desktop window
        ElseIf wn = "{top}" Then '''''''''''''''''If top window is specified
            hwnd = apiGetTopWindow(hwnd) '''''''''Get the handle of the top window
        ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
            hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
        ElseIf wn = "{parent}" Then ''''''''''''''If parent window is specified
            hwnd = apiGetParent(wIndex1) '''''''''Get the handle of the parent window
        ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
            hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
        ElseIf wn = "{frompoint}" Then '''''''''''If window under the cursor is specified
            If wIndex1 <> 1 And wIndex2 <> 1 Then 'If point specified
                p.x = wIndex1 ''''''''''''''''''''If x specified
                p.y = wIndex2 ''''''''''''''''''''If y specified
            Else
                Call apiGetCursorPos(p)  ''''''''''Get cursor position as POINTAPI
            End If
            hwnd = apiWindowFromPoint(p.x, p.y) ''Get the handle of the window under the cursor
        ElseIf wn = "{childfrompoint}" Then ''''''If child window under the cursor is specified
            If wIndex2 <> 1 And wIndex3 <> 1 Then 'If point specified
                p.x = wIndex2 ''''''''''''''''''''If x specified
                p.y = wIndex3 ''''''''''''''''''''If y specified
            Else
                Call apiGetCursorPos(p)  '''''''''Get cursor position as POINTAPI
            End If
            hwnd = apiChildWindowFromPointEx(wIndex1, p.x, p.y, wIndex2) 'Get the handle of the child window under the cursor
        ElseIf InStr(1, wName1, ".exe") <> 0 Then 'If process name specified
            pid = ProcessNameToId(wName1) ''''''''Get process Id and wait for input idle
            If pid <> 0 Then hwnd = ProcessIdToHandle(pid) 'If process id found then set handle
        End If
    End If
    If hwnd = 0 Then '''''''''''''''''''''''''''''If no handle found
        pid = Shell(wName1, wIndex1) '''''''''''''Shell(start) the process
        If pid <> 0 Then '''''''''''''''''''''''''If process id found
            hwnd = ProcessIdToHandle(pid) ''''''''Get handle from process identification
            Call ProcessNameToId("", pid) ''''''''Wait for process id to be idle by specifying it in the second parameter
        End If
    End If
    If hwnd = 0 Then hwnd = -1 '''''''''''''''''''If main window not found, then set failure return
    wf.Foreground = hwnd '''''''''''''''''''''''''Set structure return
    wf.Focus = hwnd ''''''''''''''''''''''''''''''Set focus to main handle for now
    If hwnd = -1 Then GetWinHandles = wf: Exit Function ''''''''''''''''''If no handle found then return now
    If i = 1 Then ''''''''''''''''''''''''''''''''If only the main window was specified
        Dim prvswnd As Long
        prvswnd = apiGetForegroundWindow '''''''''Remember the current foreground window
        ForceForeground (hwnd) '''''''''''''''''''Force foreground onto specified main window
        apiSleep (0) ''''''''''''''''''''''''''''''''Sleep a moment
        wf = GetWinFocus() '''''''''''''''''''''''Set structure
        ForceForeground (prvswnd) ''''''''''''''''Force foreground back to where it was
    ElseIf i > 1 Then ''''''''''''''''''''''''''''If there are child windows specified
        cwnd = GetChildWindow(hwnd, wNames(2), CInt(wNames(3))) 'Set first child handle
        wf.Focus = cwnd ''''''''''''''''''''''''''Set focus handle to first child
        If (i - 1) > 1 Then ''''''''''''''''''''''If more than one child specified
            Dim q As Long
            For q = 4 To i - 1 Step 2  '''''''''''Step through array looking for grandchildren.
                cwnd = GetChildWindow(cwnd, wNames(q), CInt(wNames(q + 1))) 'Set new child window
                wf.Focus = cwnd ''''''''''''''''''Set focus handle to youngest grandchild so far
            Next '''''''''''''''''''''''''''''''''Next in array
        End If
    End If
    GetWinHandles = wf ''''''''''''''''''''''''''''''''''''Return WINFOCUS structure
End Function
' Waits for a window to exist, and waits for it to be ready for keyboard and mouse input.
' Returns true if a window exists. Returns false if it does not, or the specified time ends.
' wTitle The title or class name of the window.  Specify an application title or class name.
' dwMilliseconds The number of milliseconds to wait.  Specify an integer.
Friend Function WaitForWindow(ByVal wTitle As String, Optional ByVal dwMilliseconds As Long = 5000) As Boolean
    On Error Resume Next
    Dim tick As Long
    Dim hwnd As Long
    tick = apiGetTickCount '''''''''''''''''''''''Get the number of millisecons since last log on
    hwnd = 0 '''''''''''''''''''''''''''''''''''''Initialize handle
    Do '''''''''''''''''''''''''''''''''''''''''''Loop this thread until it's time
        hwnd = apiFindWindow(vbNullString, wTitle) 'Find window by title
        If hwnd = 0 Then hwnd = apiFindWindow(wTitle, vbNullString) 'If not found by title then find by class name
        If hwnd <> 0 Then WaitForWindow = WaitForWindowIdle(hwnd): Exit Do  'If found, then wait for input idle and return
        If apiGetTickCount >= (tick + dwMilliseconds) Then WaitForWindow = False: Exit Do  'If time is up then return that value
        apiSleep (1) '''''''''''''''''''''''''''''Sleep thread for one millisecond
        DoEvents ''''''''''''''''''''''''''''''''''''Process messages in the queue
    Loop
End Function
Private Function AttachInput(ByVal hwnd As Long, Optional ByVal bAttach As Boolean = True) As Boolean
    On Error Resume Next
    AttachInput = CBool(apiAttachThreadInput(apiGetWindowThreadProcessId(hwnd, 0), apiGetCurrentThreadId, CInt(bAttach)))
End Function
Private Function ForceForeground(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    apiLockSetForegroundWindow (LSFW_UNLOCK) '''''Unlock setforegroundwindow calls
    apiAllowSetForegroundWindow (ASFW_ANY) '''''''Allow setforeground window calls
    '  Call KeyEvent(VK_Menu, False, True) ''''''''''Lift menu key if pressed, and it also allows the foreground window to be set
    ForceForeground = CBool(apiSetForegroundWindow(hwnd)) 'Set foreground window
    apiLockSetForegroundWindow (LSFW_LOCK) '''''''Lock other apps from using setforegroundwindow
End Function
Friend Function GetChildWindow(ByVal hwnd As Long, Optional ByVal wName As String = "", Optional ByVal wIndex As Long = 1) As Long
    On Error Resume Next
    Dim cwnd       As Long
    Dim TextCount  As Long
    Dim ClassCount As Long
    Dim NullCount  As Long
    Dim w          As WINNAME
    Do '''''''''''''''''''''''''''''''''''''''''''Loop through sibling windows
        cwnd = apiFindWindowEx(hwnd, cwnd, vbNullString, vbNullString) 'Set child handle
        If cwnd = 0 Then GetChildWindow = cwnd: Exit Do  'If no more sibling children then return
        w = GetWinName(cwnd, True, True) '''''''''Get the text, and class from that window
        If w.lpText = wName Then TextCount = TextCount + 1 'If text matches the specified wName then increment the text count
        If TextCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If text count is equal to the specified wIndex then return
        If w.lpClass = wName Then ClassCount = ClassCount + 1 'If class name matches the specified wName then increment the class count
        If ClassCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If class count is equal to the specified wIndex then return
        If w.lpText = "" And w.lpClass = "" And wName = "" Then NullCount = NullCount + 1 'If text and class name not specified, then increment by index only
        If NullCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If null count is equal to wIndex then return
    Loop
End Function
Friend Function GetChildWindow2(ByVal hwnd As Long, Optional ByVal wName As String = "", Optional ByVal wIndex As Long = 1) As Long
    On Error Resume Next
    Dim cwnd       As Long
    Dim ClassCount As Long
    Dim NullCount  As Long
    Dim w          As WINNAME
    Do '''''''''''''''''''''''''''''''''''''''''''Loop through sibling windows
        cwnd = apiFindWindowEx(hwnd, cwnd, vbNullString, vbNullString) 'Set child handle
        If cwnd = 0 Then GetChildWindow2 = cwnd: Exit Do  'If no more sibling children then return
        w = GetWinName(cwnd, False, True) '''''''''Get the text, and class from that window
        If InStr(0, w.lpClass, wName) <> 0 Then
            GetChildWindow2 = cwnd
            Exit Do
        End If
    Loop
End Function
Private Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String
    On Error Resume Next
    Dim lngCounter As Long
    strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn) ' Calculate the offset for the item required based on the number of columns the list strList' has i.e. 'lngNumColumns' and from which row the element is to be  selected i.e. 'lngRow'.
    For lngCounter = 0 To lngColumn - 1 ''''''''''Search for the 'lngColumn' item from the list 'strList'.
        strList = Mid(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))  ' Remove each item from the list.
        If Len(strList) = 0 Then GetElement = "": Exit Function 'If list becomes empty before 'lngColumn' is found then just return an empty string.
    Next lngCounter
    GetElement = Left(strList, InStr(strList, strDelimiter) - 1) 'Return the sought list element.
End Function
Private Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer
    On Error Resume Next
    Dim intElementCount As Integer
    If Len(strList) = 0 Then GetNumElements = 0: Exit Function 'If no elements in the list 'strList' then just return 0.
    strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
    While InStr(strList, strDelimiter) > 0 '''''''Count the number of elements in 'strlist'
        intElementCount = intElementCount + 1
        strList = Mid(strList, InStr(strList, strDelimiter) + 1, Len(strList))
    Wend
    GetNumElements = intElementCount '''''''''''''Return the number of elements in 'strList'.
End Function
Private Function GetWinAncestory(ByVal cwnd As Long, ByVal gName As Boolean) As Long()
    On Error Resume Next
    Dim i       As Long
    Dim z       As Long
    Dim cWnd2   As Long
    Dim hwnds() As Long
    Dim w       As WINNAME
    cWnd2 = cwnd '''''''''''''''''''''''''''''''''Remember handle for seconds loop
    w = GetWinName(cwnd, True, True) '''''''''''''Get text and class name of the window
    If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Focus window:  Title | Class name") 'Display to developer if specified
    Do '''''''''''''''''''''''''''''''''''''''''''Loop while counting ancestors
        cwnd = apiGetParent(cwnd) ''''''''''''''''Get new parent handle if any found
        If cwnd = 0 Then Exit Do '''''''''''''''''If there are no more parents then abort without setting anymore to the array
        i = i + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
    Loop
    ReDim hwnds(i) '''''''''''''''''''''''''''''''Re-Dimension array to number of ancestors including the handle specified
    Do '''''''''''''''''''''''''''''''''''''''''''Do look for parents of the window specified.
        hwnds(z) = cWnd2 '''''''''''''''''''''''''Set the indexed array element to the specified handle or it's ancestors
        cWnd2 = apiGetParent(cWnd2) ''''''''''''''Get new parent handle if any found
        If cWnd2 = 0 Then Exit Do ''''''''''''''''If there are no more parents then abort
        z = z + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
        w = GetWinName(cWnd2, True, True) ''''''''Get text and class name of the window
        If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Parent window:  Title | Class name")  'Display to developer if specified
    Loop
    GetWinAncestory = hwnds ''''''''''''''''''''''Return array of integer handles, or handle.
End Function
Friend Function GetWinName(ByVal hwnd As Long, Optional ByVal gText As Boolean = True, Optional ByVal gClass As Boolean = True) As WINNAME
    On Error Resume Next
    Dim tLength As Long
    Dim rValue  As Long
    GetWinName.lpText = "" '''''''''''''''''''''''Initialize string for text name
    GetWinName.lpClass = "" ''''''''''''''''''''''Initialize string for class name
    If gText = True Then '''''''''''''''''''''''''If text is to be retrieved
        tLength = apiGetWindowTextLength(hwnd) + 4 'Get length
        GetWinName.lpText = Strings.Space(tLength) 'Pad with buffer
        rValue = apiGetWindowText(hwnd, GetWinName.lpText, tLength) 'Get text
        GetWinName.lpText = Left(GetWinName.lpText, rValue) 'Strip buffer
    End If
    If gClass = True Then ''''''''''''''''''''''''If class name is to be retrieved
        GetWinName.lpClass = Strings.Space(260) ''Pad with buffer
        rValue = apiGetClassName(hwnd, GetWinName.lpClass, 260) 'Get classname
        GetWinName.lpClass = Left(GetWinName.lpClass, rValue) 'Strip buffer
    End If
End Function
Private Function HandleToProcessId(ByVal hwnd As Long) As Long
    On Error Resume Next
    Dim wnd As Long
    Dim pid As Long
    If apiGetParent(hwnd) <> 0 Then hwnd = apiGetAncestor(hwnd, GA_ROOT)
    wnd = apiGetTopWindow(0) '''''''''''''''''''''Get the top window in the z-order
    Do
        If wnd = 0 Then Exit Do
        If wnd = hwnd Then
            Call apiGetWindowThreadProcessId(wnd, pid)  'Get the window's process id
            HandleToProcessId = pid: Exit Do '''''Set process id and return
        End If
        wnd = apiGetWindow(wnd, GW_HWNDNEXT) '''''Retrieve the next window
    Loop
End Function
Private Function ProcessIdToHandle(ByVal pid As Long) As Long
    On Error Resume Next
    Dim hwnd      As Long
    Dim ProcessID As Long
    hwnd = apiGetTopWindow(0) ''''''''''''''''''''Get the top window in the z-order
    Do
        If apiGetParent(hwnd) = 0 Then '''''''''''If window has no parents
            Call apiGetWindowThreadProcessId(hwnd, ProcessID)  'Get the window's process id
            If ProcessID = pid Then ProcessIdToHandle = hwnd: Exit Do 'If pid matches then return
        End If
        hwnd = apiGetWindow(hwnd, GW_HWNDNEXT) '''Retrieve the next window
        If hwnd = 0 Then Exit Do '''''''''''''''''Exit if there are no more windows
    Loop
End Function
Private Function ProcessNameToId(ByVal ProcessName As String, Optional ByVal pid As Long = 0) As Long
    On Error Resume Next
    Dim pLength            As Long
    Dim cbSizeReturned     As Long
    Dim numElements        As Long
    Dim pIDs()             As Long
    Dim cbSize             As Long
    Dim cbSize2            As Long
    Dim ret                As Long
    Dim pSize              As Long
    Dim hProcess           As Long
    Dim pLoop              As Long
    Dim pName              As String
    Dim mName              As String
    Dim prName             As String
    Dim pModules(1 To 200) As Long
    ProcessName = UCase(Trim(ProcessName))
    pLength = Len(ProcessName)
    cbSize = 8
    cbSizeReturned = 96
    Do While cbSize <= cbSizeReturned
        cbSize = cbSize * 2 ''''''''''''''''''Increment Size
        ReDim pIDs(cbSize / 4) As Long '''''''Allocate Memory for Array
        ret = apiEnumProcesses(pIDs(1), cbSize, cbSizeReturned) 'Get Process ID's
    Loop
    numElements = cbSizeReturned / 4  ''''''''Count number of processes returned
    For pLoop = 1 To numElements '''''''''''''Loop thru each process
        hProcess = apiOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDs(pLoop))  'Get a handle to the Process and Open
        If hProcess <> 0 Then
            ret = apiEnumProcessModules(hProcess, pModules(1), 200, cbSize2) 'Get an array of the module handles for the specified process
            If ret <> 0 Then '''''''''''''''''If the Module Array is retrieved, Get the ModuleFileName
                mName = Space(MAX_PATH) ''''''Buffer with spaces first to allocate memory for byte array
                pSize = 500  '''''''''''''''''Must be set prior to calling API
                ret = apiGetModuleFileNameExA(hProcess, pModules(1), mName, pSize) 'Get Process Name
                pName = Left(mName, ret) '''''Remove trailing spaces
                pName = UCase(Trim(pName)) '''Check for Matching Upper case result
                prName = GetElement(Trim(Replace(pName, Chr(0), "")), "\", 0, 0, GetNumElements(Trim(Replace(pName, Chr(0), "")), "\") - 1)
                If pid = 0 Then ''''''''''''''If getting process id
                    If prName = ProcessName Then 'If process name matches specification
                        ProcessNameToId = pIDs(pLoop) 'Set pid return
                        Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
                    End If
                Else '''''''''''''''''''''''''If getting process handle from id
                    If pid = pIDs(pLoop) Then 'If process id matches specification
                        ProcessNameToId = hProcess 'Set process handle return
                        Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
                    End If
                End If
            End If
        End If
        ret = apiCloseHandle(hProcess) '''''''Close the handle to this process
    Next
End Function
Private Function WaitForWindowIdle(ByVal hwnd As Long) As Boolean
    On Error Resume Next
    Dim pid As Long
    pid = HandleToProcessId(hwnd) ''''''''''''''''Get process id from handle
    If pid <> 0 Then WaitForWindowIdle = CBool(ProcessNameToId("", pid)) 'If process id found then set return value of specified pId, which is a process handle
End Function
' This function sends window messages directly to a window.
' If the asMessage parameter is false then, this function simulates mouse events directly to a window.
' Returns false if the specified window/s cannot be found.
' If the asMessage parameter is false then, this function returns false if a handle cannot be established or
' the top window could not be set, or a window rectangle could not be found, or the cursor could not be set,
' or the mouse event could not be simulated.
' "mButtons" The button to click.  Use the Buttons enumeration for this parameter.
'  "wFocus" The structure of the window to focus on.
' Use GetWinHandles or GetWinFocus to get a structure.
'  "mDown" (Optional) Press mouse button down only.
'  "mUp" (Optional) Lift mouse button up only.
'  "asMessage" (Optional) Send as window message, or as a mouse event.
' It's recommended that you use True in this parameter.
' If false then the x and y parameters may apply.
'  "x" (Optional) The x coordinate only applies if asMessage is false, and mButtons is
' Buttons.Move or Buttons.MoveAbsolute.
'  "y" (Optional) The y coordinate only applies if asMessage is false, and mButtons is
' Buttons.Move or Buttons.MoveAbsolute.
Friend Function Click(ByVal mButtons As Long, ByRef WFOCUS As WINFOCUS, Optional ByVal mDown As Boolean = True, Optional ByVal mUp As Boolean = True, Optional ByVal asMessage As Boolean = True, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1) As Boolean
    On Error Resume Next
    Dim hwnd        As Long
    Dim cwnd        As Long
    Dim zOrder      As Long
    Dim zOrderChild As Long
    Dim p           As POINTAPI
    Dim ws          As WINSTATE
    Click = True
    If WFOCUS.Foreground = -1 Then Click = False: Exit Function
    If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False)
    If asMessage = True Then  ''''''''''''''''''''If sending as message
        Dim WM_DOWN As Long
        Dim WM_UP   As Long
        If mButtons = Buttons.Move Then ''''''''''If moving from current position
            Dim cP As POINTAPI
            Call apiGetCursorPos(cP)  '''''''''''''Get cursor position
            Call apiSetCursorPos(cP.x + x, cP.y + y) 'Add to coordinates, and set cursor there
            Click = True: Exit Function ''''''''''Return success
        ElseIf mButtons = Buttons.MoveAbsolute Then
            Call apiSetCursorPos(x, y) '''''''''''Set to absolute coordinates
            Click = False: Exit Function '''''''''Return success
        ElseIf mButtons = Buttons.LeftClick Or mButtons = Buttons.LeftDoubleClick Then 'If button specified is left
            WM_DOWN = WM_LBUTTONDOWN: WM_UP = WM_LBUTTONUP 'Then set left messages
        ElseIf mButtons = Buttons.RightClick Or mButtons = Buttons.RightDoubleClick Then 'If button specified is right
            WM_DOWN = WM_RBUTTONDOWN: WM_UP = WM_RBUTTONUP 'Then set right messages
        ElseIf mButtons = Buttons.MiddleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If button specified is middle
            WM_DOWN = WM_MBUTTONDOWN: WM_UP = WM_MBUTTONUP ' Then set middle messages
        End If
        hwnd = WFOCUS.Foreground '''''''''''''''''Set main handle to something smaller
        cwnd = WFOCUS.Focus ''''''''''''''''''''''Set child handle(if any) to something smaller
        If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL) 'if minimized then show normal
        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True) 'If disabled then enable
        If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True) 'If child disabled then enable
        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL) 'If hidden then show
        If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL) 'If child hidden then show
        zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
        zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
        If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
        If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
        If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'If button down is specified, then press down
        If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'If button up is specified, then lift up
        If mButtons = Buttons.LeftDoubleClick Or mButtons = Buttons.RightDoubleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If it's a double click
            If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'Down again
            If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'Up again
        End If
        Call MouseAbort(ws, False, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
    Else '''''''''''''''''''''''''''''''''''''''''Else it's an event to send
        Dim repeat  As Long
        Dim rCursor As Boolean
        Dim pts     As POINTAPI
        If mButtons = Buttons.Move Then  '''''''''Move cursor to click point
            Call MouseEvent(mButtons, x, y) ''''''Move cursor to click point
            Exit Function '''''''''''''''''''''''''''''''''Exit thread
        ElseIf mButtons = Buttons.MoveAbsolute Then
            pts = ToScreen(x, y)   ''''''''''''''Convert to screen coordinates
            If pts.x <> 0 Then x = pts.x  ''''''''If x point found then set
            If pts.y <> 0 Then y = pts.y  ''''''''If y point found then set
            Call MouseEvent(mButtons, x, y) ''''''Move cursor to click point
            Exit Function ''''''''''''''''''''''''Exit thread
        End If
        If WFOCUS.Foreground = -1 Then Exit Function  'Exit if return from GetWinHandles is negative
        If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False) 'Get current focus
        hwnd = WFOCUS.Foreground  ''''''''''''''''Set main handle to something smaller
        cwnd = WFOCUS.Focus  '''''''''''''''''''''Set child handle(if any) to something smaller
        If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL) 'if minimized then show normal
        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True) 'If disabled then enable
        If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True) 'If child disabled then enable
        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL) 'If hidden then show
        If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL) 'If child hidden then show
        zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
        zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
        If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
        If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
        If mButtons = 12 Or mButtons = 48 Or mButtons = 768 Then repeat = 1    'If double click
        If mDown = True And mUp = False Then '''''If mouse down
            If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
                mButtons = MOUSEEVENTF_LEFTDOWN 'Set as left down
            ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
                mButtons = MOUSEEVENTF_RIGHTDOWN 'Set as right down
            ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
                mButtons = MOUSEEVENTF_MIDDLEDOWN 'Set as middle down
            End If
        ElseIf mDown = False And mUp = True Then 'If mouse up
            If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
                mButtons = MOUSEEVENTF_LEFTUP '''Set as left up
            ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
                mButtons = MOUSEEVENTF_RIGHTUP ''Set as right up
            ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
                mButtons = MOUSEEVENTF_MIDDLEUP 'Set as middle up
            End If
        End If
        If mButtons <> Buttons.Wheel And mButtons <> Buttons.VirtualDesk Then   'If it's a click
            Dim r As RECT
            If apiGetWindowRect(cwnd, r) = False Then 'If no rectangle is found
                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'If rectangle not found, then exit failure
                Exit Function ''''''''''''''''''''Exit
            End If
            pts.x = CInt(r.rLeft + ((r.rRight - r.rLeft) / 2)) 'Set to the center of the horizon
            pts.y = CInt(r.rTop + ((r.rBottom - r.rTop) / 2)) 'Set to the center of the vertical
            pts = ToScreen(pts.x, pts.y) '''''''''Convert to screen coordinates
            rCursor = True '''''''''''''''''''''''Cursor position changed, remember to return it later
            x = 0 '''''''''''''''''''''''''''''''null for click
            y = 0 '''''''''''''''''''''''''''''''null for click
            Call apiGetCursorPos(p)  '''''''''''''Get the current cursor position, to be returned later
            If MouseEvent(Buttons.MoveAbsolute, pts.x, pts.y) = False Then 'Move cursor to click point
                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if fails to move
                Exit Function ''''''''''''''''''''Exit thread
            End If
        End If
        Dim i As Long
        For i = 1 To repeat + 1  '''''''''''''''''Loop the number of repeats
            If MouseEvent(mButtons, x, y) = False Then  'Do mouse event repeated number of times
                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if failure
                Exit Function ''''''''''''''''''''Exit thread
            End If
        Next
        Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
    End If
End Function
' Clicks on the specified menu item, by sending it a command message.
' If asMessage is false then this function clicks on the specified menu item,
' by simulating an entire chain of events.
' Returns false if window or menu cannot be found.
' If asMessage is false then this function returns false if a window or menu or rectangle cannot be found,
' or an event fails to be simulated.
'  "asMessage"Send as window message, or as a mouse event.
' It's recommended that you use True in this parameter.
'  "wName"Title or class name of the main window.
'  "wIndex"Index of the main window.
'  "mnuName1"Main menu title.
'  "mnuIndex1"Main menu index.
'  "mnuName2"(Optional) Sub menu title.
'  "mnuIndex2"(Optional) Sub menu index.
'  "mnuName3"(Optional) Sub menu title.
'  "mnuIndex3"(Optional) Sub menu index.
'  "mnuName4"(Optional) Sub menu title.
'  "mnuIndex4"(Optional) Sub menu index.
'  "mnuName5"(Optional) Sub menu title.
'  "mnuIndex5"(Optional) Sub menu index.
'  "mnuName6"(Optional) Sub menu title.
'  "mnuIndex6"(Optional) Sub menu index.
'  "mnuName7"(Optional) Sub menu title.
'  "mnuIndex7"(Optional) Sub menu index.
'  "mnuName8"(Optional) Sub menu title.
'  "mnuIndex8"(Optional) Sub menu index.
'  "mnuName9"(Optional) Sub menu title.
'  "mnuIndex9"(Optional) Sub menu index.
'  "mnuName10"Sub menu title.
'  "mnuIndex10"(Optional) Sub menu index.</param>
Friend Function ClickMenu(ByVal asMessage As Boolean, ByVal wName As String, ByVal wIndex As Long, ByVal mnuName1 As String, Optional ByVal mnuIndex1 As Long = 1, Optional ByVal mnuName2 As String = " ", Optional ByVal mnuIndex2 As Long = 1, Optional ByVal mnuName3 As String = " ", Optional ByVal mnuIndex3 As Long = 1, Optional ByVal mnuName4 As String = " ", Optional ByVal mnuIndex4 As Long = 1, Optional ByVal mnuName5 As String = " ", Optional ByVal mnuIndex5 As Long = 1, Optional ByVal mnuName6 As String = " ", Optional ByVal mnuIndex6 As Long = 1, Optional ByVal mnuName7 As String = " ", Optional ByVal mnuIndex7 As Long = 1, Optional ByVal mnuName8 As String = " ", Optional ByVal mnuIndex8 As Long = 1, Optional ByVal mnuName9 As String = " ", Optional ByVal mnuIndex9 As Long = 1, Optional ByVal mnuName10 As String = " ", Optional ByVal mnuIndex10 As Long = 1) As Boolean
    On Error Resume Next
    Dim n          As Long
    Dim mnuNames() As String
    Dim i          As Long
    Dim nPos       As Long
    Dim ws         As WINSTATE
    If asMessage = True Then
        Dim hwnd     As Long
        Dim hMenu    As Long
        Dim hSubMenu As Long
        Dim hId      As Long
        hwnd = GetWinHandles(wName, wIndex).Foreground 'Set handle to specification
        If hwnd = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL) 'If minimized then show it
        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True) 'If disabled then enabled it
        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL) 'If Hidden then show it
        hMenu = apiGetMenu(hwnd) '''''''''''''''''Set handle of main menu
        If hMenu = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
        hSubMenu = hMenu '''''''''''''''''''''''''Initialize sub menu to main menu
        If mnuName1 <> " " Then n = 1: ReDim mnuNames(1) 'Re-dimension mnuNames array
        If mnuName2 <> " " Then n = 3: ReDim mnuNames(3)
        If mnuName3 <> " " Then n = 5: ReDim mnuNames(5)
        If mnuName4 <> " " Then n = 7: ReDim mnuNames(7)
        If mnuName5 <> " " Then n = 9: ReDim mnuNames(9)
        If mnuName6 <> " " Then n = 11: ReDim mnuNames(11)
        If mnuName7 <> " " Then n = 13: ReDim mnuNames(13)
        If mnuName8 <> " " Then n = 15: ReDim mnuNames(15)
        If mnuName9 <> " " Then n = 17: ReDim mnuNames(17)
        If mnuName10 <> " " Then n = 19: ReDim mnuNames(19)
        If n > 0 Then mnuNames(0) = mnuName1: mnuNames(1) = CStr(mnuIndex1) 'Set elements of array to specification
        If n > 2 Then mnuNames(2) = mnuName2: mnuNames(3) = CStr(mnuIndex2)
        If n > 4 Then mnuNames(4) = mnuName3: mnuNames(5) = CStr(mnuIndex3)
        If n > 6 Then mnuNames(6) = mnuName4: mnuNames(7) = CStr(mnuIndex4)
        If n > 8 Then mnuNames(8) = mnuName5: mnuNames(9) = CStr(mnuIndex5)
        If n > 10 Then mnuNames(10) = mnuName6: mnuNames(11) = CStr(mnuIndex6)
        If n > 12 Then mnuNames(12) = mnuName7: mnuNames(13) = CStr(mnuIndex7)
        If n > 14 Then mnuNames(14) = mnuName8: mnuNames(15) = CStr(mnuIndex8)
        If n > 16 Then mnuNames(16) = mnuName9: mnuNames(17) = CStr(mnuIndex9)
        If n > 18 Then mnuNames(18) = mnuName10: mnuNames(19) = CStr(mnuIndex10)
        For i = 0 To (n - 1) Step 2  '''''''''''''Loop through menu tree
            If mnuNames(i) = "" Then '''''''''''''If menu name not specified then
                nPos = CInt(mnuNames(i + 1)) - 1 'Set position as an index
            Else '''''''''''''''''''''''''''''''''Then name was set
                nPos = FindMenuItemPos(hSubMenu, mnuNames(i), CInt(mnuNames(i + 1))) 'Set position to name and index of specified item.
            End If
            If nPos <> -1 Then
                If apiGetSubMenu(hSubMenu, nPos) <> 0 Then hSubMenu = apiGetSubMenu(hSubMenu, nPos) 'If sub menu exits, and is specified then get the handle
            End If
        Next
        If nPos <> -1 Then '''''''''''''''''''''''If final item has a valid position
            hId = apiGetMenuItemID(hSubMenu, nPos) 'Get menu id
            If hId <> -1 Then ''''''''''''''''''''If item has no sub menus
                ClickMenu = Not CBool(apiSendMessage(hwnd, WM_COMMAND, hId, vbNullString)) 'Send command message
            Else
                ClickMenu = False ''''''''''''''''Return failure
            End If
        End If
        If ws.IsIconic = True Then Call apiShowWindow(hwnd, SW_SHOWMINIMIZED)  'If window was minimized then re-minimize it
        If ws.IsDisabled = True Then Call apiEnableWindow(hwnd, False)  'If window was disabled then re-disable it
        If ws.IsHidden = True Then Call apiShowWindow(hwnd, SW_HIDE) 'If window was hidden then re-hide it
    Else
        Dim rOffset     As Long
        Dim tOffset     As Long
        Dim LeftMost    As Long
        Dim TopMost     As Long
        Dim ArrayLength As Long
        Dim m           As MENUINFO
        Dim MI          As ITEMINFO
        Dim p           As POINTAPI
        Dim poi         As POINTAPI
        Dim r           As RECT
        ClickMenu = True '''''''''''''''''''''''''Signal true indicating that the thread was started
        If mnuName1 <> " " Then n = 3: ReDim mnuNames(3)  'Re-dimension mnuNames array
        If mnuName2 <> " " Then n = 5: ReDim mnuNames(5)
        If mnuName3 <> " " Then n = 7: ReDim mnuNames(7)
        If mnuName4 <> " " Then n = 9: ReDim mnuNames(9)
        If mnuName5 <> " " Then n = 11: ReDim mnuNames(11)
        If mnuName6 <> " " Then n = 13: ReDim mnuNames(13)
        If mnuName7 <> " " Then n = 15: ReDim mnuNames(15)
        If mnuName8 <> " " Then n = 17: ReDim mnuNames(17)
        If mnuName9 <> " " Then n = 19: ReDim mnuNames(19)
        If mnuName10 <> " " Then n = 21: ReDim mnuNames(21)
        If n > 0 Then mnuNames(0) = wName: mnuNames(1) = wIndex  'Set elements of array to specification
        If n > 2 Then mnuNames(2) = mnuName1: mnuNames(3) = CStr(mnuIndex1)
        If n > 4 Then mnuNames(4) = mnuName2: mnuNames(5) = CStr(mnuIndex2)
        If n > 6 Then mnuNames(6) = mnuName3: mnuNames(7) = CStr(mnuIndex3)
        If n > 8 Then mnuNames(8) = mnuName4: mnuNames(9) = CStr(mnuIndex4)
        If n > 10 Then mnuNames(10) = mnuName5: mnuNames(11) = CStr(mnuIndex5)
        If n > 12 Then mnuNames(12) = mnuName6: mnuNames(13) = CStr(mnuIndex6)
        If n > 14 Then mnuNames(14) = mnuName7: mnuNames(15) = CStr(mnuIndex7)
        If n > 16 Then mnuNames(16) = mnuName8: mnuNames(17) = CStr(mnuIndex8)
        If n > 18 Then mnuNames(18) = mnuName9: mnuNames(19) = CStr(mnuIndex9)
        If n > 20 Then mnuNames(20) = mnuName10: mnuNames(21) = CStr(mnuIndex10)
        ArrayLength = n ''''''''''''''''''''''''''Set length of array
        m.hwnd = GetWinHandles(wName).Foreground 'Get the handle of the specified window
        If m.hwnd = 0 Then ClickMenu = False: Exit Function 'If window not found, then exit failure
        If apiIsIconic(m.hwnd) = True Then ws.IsIconic = apiShowWindow(m.hwnd, SW_SHOWNORMAL) 'If minimized then show it
        If apiIsWindowEnabled(m.hwnd) = False Then ws.IsDisabled = apiEnableWindow(m.hwnd, True) 'If disabled then enable it
        If apiIsWindowVisible(m.hwnd) = False Then ws.IsHidden = Not apiShowWindow(m.hwnd, SW_SHOWNORMAL) 'If hidden then show it
        Call apiGetWindowRect(m.hwnd, r) '''''''''Set confirmation of window rectangle
        If r.rBottom <> 0 Then
            Call apiMoveWindow(m.hwnd, 0, 0, r.rRight - r.rLeft, r.rBottom - r.rTop, True) 'If rectangle found then move window with coordinates
        End If
        Call apiGetCursorPos(poi)  '''''''''''''''Get the current position of the user's cursor, so that it can be returned
        m.hMenu = apiGetMenu(m.hwnd) '''''''''''''Set handle of the main menu
        If m.hMenu = 0 Then ClickMenu = False: Exit Function '''''''''''''If no handle found then exit sub with failure
        m.hSubMenu = apiGetSubMenu(m.hMenu, 0) '''Set handle of the first sub menu if any
        MI = MenuItemDim(m.hwnd, m.hMenu, 0) '''''Get the dimensions of the menu item
        If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon negative results
        LeftMost = MI.Left '''''''''''''''''''''''Initialize the left most coordinate
        If mnuNames(2) = "" Then '''''''''''''''''If no name specified
            nPos = CInt(mnuNames(3)) - 1 '''''''''Set position by index only
        Else '''''''''''''''''''''''''''''''''''''Otherwise set position by name and index
            nPos = FindMenuItemPos(m.hMenu, mnuNames(2), CInt(mnuNames(3))) 'Find position of the menu item
        End If
        If nPos = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
        MI = MenuItemDim(m.hwnd, m.hMenu, nPos) ''Get item  dimensions
        If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
        p = ToScreen(MI.Center.x, MI.Center.y) '''Convert point to screen coordinates
        If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then ClickMenu = False: Exit Function 'Move mouse, and exit if failure
        rOffset = MI.Left - LeftMost '''''''''''''Initialize offset from the left
        TopMost = MI.Bottom ''''''''''''''''''''''Initialize offset from the top
        If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click mouse, and exit upon failure
        m.hSubMenu = apiGetSubMenu(m.hMenu, nPos) 'Set handle of submenu
        If m.hSubMenu <> 0 Then ''''''''''''''''''If handle found
            If mnuNames(4) = "" Then '''''''''''''''''If no name specified
                nPos = CInt(mnuNames(5)) - 1 '''''''''Set by index only
            Else '''''''''''''''''''''''''''''''''''''Name and index specified
                nPos = FindMenuItemPos(m.hSubMenu, mnuNames(4), CInt(mnuNames(5))) 'Find by name and index
            End If
            If nPos = -1 Then ClickMenu = False: Exit Function 'Exit if position is invalid
            MI = MenuItemDim(m.hwnd, m.hSubMenu, nPos) 'Get dimensinos
            If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit if fails
            p = ToScreen(rOffset + MI.Center.x, MI.Center.y) 'Convert point
            If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then ClickMenu = False: Exit Function 'Move and exit if failure
            For i = 6 To ArrayLength Step 2 ''''''''''Step through the array
                If MoveItemToItem(mnuNames(i), mnuNames(i + 1), TopMost, nPos, rOffset, tOffset, m, MI) = False Then ClickMenu = False: Exit Function
            Next
            If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click final menu item in the chain, and exit if failure
        End If
        If r.rBottom <> 0 Then
            Call apiMoveWindow(m.hwnd, r.rLeft, r.rTop, r.rRight - r.rLeft, r.rBottom - r.rTop, True) 'If there was a rectangle move the window back to where it was
        End If
        Call apiSetCursorPos(poi.x, poi.y) '''''''Return the position of the cursor back to the user
        If ws.IsIconic = True Then Call apiShowWindow(m.hwnd, SW_SHOWMINIMIZED)  'If was minimized then re-minimize
        If ws.IsDisabled = True Then Call apiEnableWindow(m.hwnd, False) 'If was disabled then re-disable
        If ws.IsHidden = True Then Call apiShowWindow(m.hwnd, SW_HIDE)  'If was Hidden then re-Hide
    End If
End Function
Private Function FindMenuItemPos(ByVal hMenu As Long, Optional ByVal iName As String = "", Optional ByVal iIndex As Long = 1) As Long
    On Error Resume Next
    Dim i          As Long
    Dim itemCount  As Long
    Dim indexCount As Long
    Dim retValue   As Long
    Dim mnuCaption As String
    Dim woShortcut As String
    If apiIsMenu(hMenu) = False Then FindMenuItemPos = -1: Exit Function 'Return negative result if it's not a menu handle
    FindMenuItemPos = NEGATIVE '''''''''''''''''''Set a default return value
    itemCount = apiGetMenuItemCount(hMenu) '''''''Count the number of menu items
    For i = 0 To itemCount - 1 '''''''''''''''''''Loop through all menu items
        mnuCaption = "" ''''''''''''''''''''''''''Initialize
        mnuCaption = Space(1024) '''''''''''''''''Pad with a buffer
        retValue = apiGetMenuString(hMenu, i, mnuCaption, Len(mnuCaption), 1024) 'Get menu caption
        mnuCaption = Left(mnuCaption, retValue) ''mnuCaption.Substring(0, retValue) 'Strip off buffer
        woShortcut = "" ''''''''''''''''''''''''''Initialize
        If InStr(mnuCaption, "&") = True Then woShortcut = Replace(mnuCaption, "&", "") 'If the & character exists, then remove it, so the developer doesn't have to specify
        apiCharLower (iName) '''''''''''''''''''''Convert to lower case
        apiCharLower (woShortcut)
        apiCharLower (mnuCaption)
        If iName = woShortcut Or iName = mnuCaption Then 'if specified name matches menu name, as non-case sensitive
            FindMenuItemPos = i ''''''''''''''''''Set return value as that position
            indexCount = indexCount + 1 ''''''''''Increment index by one
            If indexCount = iIndex Then Exit For 'If index matches the specification then exit loop
        End If
    Next
End Function
Private Function GetSetZOrder(ByVal hwnd As Long, Optional ByVal sPosition As Long = NEGATIVE) As Long
    On Error Resume Next
    Dim z    As Long
    Dim swnd As Long
    swnd = apiGetTopWindow(0)  'apiGetWindow(hwnd, GW_HWNDFIRST) ''''''Get top or topmost window in context
    Do
        If sPosition = NEGATIVE Then '''''''''''''If not setting the z-order
            If swnd = hwnd Then GetSetZOrder = (z + 1): Exit Function 'If handle specified matches sibling window, then return the position in the z-order
        Else '''''''''''''''''''''''''''''''''''''Then this function sets the z-order to the specified position
            If z = sPosition - 1 Then GetSetZOrder = apiSetWindowPos(hwnd, swnd, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE): Exit Function
        End If
        swnd = apiGetWindow(swnd, GW_HWNDNEXT) '''Get the next sibling window handle, in the loop
        If swnd = 0 Then Exit Do '''''''''''''''''If there are no more sibling windows, then exit loop with default return
        z = z + 1 ''''''''''''''''''''''''''''''''Increment i by one
    Loop
    GetSetZOrder = swnd
End Function
Private Function MenuItemDim(ByVal hwnd As Long, ByVal hMenu As Long, ByVal nPos As Long) As ITEMINFO
    On Error Resume Next
    Dim m As ITEMINFO
    Dim r As RECT
    If apiGetMenuItemRect(hwnd, hMenu, nPos, r) = 0 Then 'If rectangle not found then set negative returns
        r.rTop = NEGATIVE ''''''''''''''''''''''''Fail
        r.rBottom = NEGATIVE '''''''''''''''''''''Fail
        r.rLeft = NEGATIVE '''''''''''''''''''''''Fail
        r.rRight = NEGATIVE ''''''''''''''''''''''Fail
    Else '''''''''''''''''''''''''''''''''''''''''Or set dimensions of menu item
        m.Width = (r.rRight - r.rLeft) '''''''''''Set width
        m.Height = (r.rBottom - r.rTop) ''''''''''Set height
        m.Center.x = CInt(r.rLeft + (m.Width / 2)) 'Set center point x
        m.Center.y = CInt(r.rTop + (m.Height / 2)) 'Set center point y
    End If
    m.Left = r.rLeft '''''''''''''''''''''''''''''Set left coordinate
    m.Right = r.rRight '''''''''''''''''''''''''''Set right coordinate
    m.Top = r.rTop '''''''''''''''''''''''''''''''Set top coordinate
    m.Bottom = r.rBottom '''''''''''''''''''''''''Set bottom coordinate
    MenuItemDim = m
End Function
Private Function MouseAbort(ByRef WSTATE As WINSTATE, ByVal rCursor As Boolean, ByVal hwnd As Long, ByVal cwnd As Long, ByVal zOrder As Long, ByVal zOrderChild As Long, ByRef p As POINTAPI) As Boolean
    On Error Resume Next
    If rCursor = True Then Call apiSetCursorPos(p.x, p.y) 'If it was a click then return cursor to user position
    If WSTATE.IsIconic = True Then Call apiShowWindow(hwnd, SW_SHOWMINIMIZED)  'If main window was minimized before, then re-minimize it
    If WSTATE.IsDisabled = True Then Call apiEnableWindow(hwnd, False)  'If main window was disabled before, then re-disable it
    If WSTATE.IsChildDisabled = True Then Call apiEnableWindow(cwnd, False)  'If child window was disabled before, then re-disable it
    If WSTATE.IsHidden = True Then Call apiShowWindow(hwnd, SW_HIDE)  'If main window was hidden before, then re-hide it
    If WSTATE.IsChildHidden = True Then Call apiShowWindow(cwnd, SW_HIDE)  'If main window was hidden before, then re-hide it
    Call apiSetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE) 'Set the window position to not topmost window.  TODO Get topmost status first, so original state can be restored.
    If zOrderChild > 0 Then Call GetSetZOrder(cwnd, zOrderChild)  'If a z-order for the child window was obtained, reset the z-order of the child window
    If zOrder > 0 Then Call GetSetZOrder(hwnd, zOrder)  'If a z-order for the main window was obtained, reset the z-order of the main window
    MouseAbort = True ''''''''''''''''''''''''''''Return when finished
End Function
Public Function MouseEvent(Optional ByVal mEvents As Long = 0, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0) As Boolean
    On Error Resume Next
    MouseEvent = apimouse_event(mEvents, x, y, 0, apiGetMessageExtraInfo) 'Return results
End Function
Private Function MoveItemToItem(ByVal wName As String, ByVal wName2 As String, ByVal tMost As Long, ByRef nPos As Long, ByRef rOffset As Long, ByRef tOffset As Long, ByRef m As MENUINFO, ByRef MI As ITEMINFO) As Boolean
    On Error Resume Next
    Dim p As POINTAPI
    tOffset = tOffset + MI.Top - tMost '''''''''''Keep offset from top most
    rOffset = rOffset + MI.Width '''''''''''''''''Keep offset from left most
    If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then MoveItemToItem = False: Exit Function 'click sub menu, and return false if failure
    m.hSubMenu = apiGetSubMenu(m.hSubMenu, nPos) 'Get handle of submenu
    If wName = "" Then '''''''''''''''''''''''''''If name not specified
        nPos = CInt(wName2) - 1 ''''''''''''''''''Then the search is by index
    Else '''''''''''''''''''''''''''''''''''''''''Then the search is by name
        nPos = FindMenuItemPos(m.hSubMenu, wName, CInt(wName2)) 'Get position from handle and name
    End If
    If nPos = -1 Then MoveItemToItem = False: Exit Function '''''''''''''''If return is negative then exit and return false
    MI = MenuItemDim(m.hwnd, m.hSubMenu, 0) ''''''Get menu item dimensions
    If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit if there is a negative return
    p = ToScreen(rOffset + MI.Center.x, tOffset + MI.Center.y) 'Covert to screen coordinates
    If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then MoveItemToItem = False: Exit Function 'Move to new screen location
    MI = MenuItemDim(m.hwnd, m.hSubMenu, nPos) ''Get menu item dimensions
    If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit upon negative result
    p = ToScreen(rOffset + MI.Center.x, tOffset + MI.Center.y) 'Convert to screen
    If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then MoveItemToItem = False: Exit Function 'Move to point
    MoveItemToItem = True
End Function
Friend Function ToScreen(ByVal x As Long, ByVal y As Long) As POINTAPI
    On Error Resume Next
    ToScreen.x = CInt(x * SM_FULLSCREEN / apiGetSystemMetrics(SM_CXSCREEN)) 'Set the return value for x.
    ToScreen.y = CInt(y * SM_FULLSCREEN / apiGetSystemMetrics(SM_CYSCREEN)) 'Set the return value for y.
End Function
Friend Function InvokeElement(ByVal hwnd As Long, ByVal sText As String) As Boolean
    On Error GoTo poop
    InvokeElement = False
    If hwnd = 0 Then Exit Function
    Dim uAuto   As IUIAutomation
    Dim el      As IUIAutomationElement
    Dim uCond   As IUIAutomationCondition
    Dim elName  As IUIAutomationElement
    Dim pInvoke As IUIAutomationInvokePattern
    Set uAuto = New CUIAutomation
    Set uCond = uAuto.CreatePropertyCondition(UIA_NamePropertyId, sText)
    If uCond Is Nothing Then Exit Function
    Set el = uAuto.ElementFromHandle(ByVal hwnd)
    If el Is Nothing Then Exit Function
    Set elName = el.FindFirst(TreeScope_Children, uCond)
    If elName Is Nothing Then Exit Function
    Set pInvoke = elName.GetCurrentPattern(UIA_InvokePatternId)
    If pInvoke Is Nothing Then Exit Function
    pInvoke.Invoke
    InvokeElement = True
poop:
End Function
Friend Function GetUIADesktopElements() As IUIAutomationElementArray
    On Error GoTo poop
    Dim allChilds As IUIAutomationElementArray
    Set GetUIADesktopElements = allChilds
    Dim oUIAutomation As New CUIAutomation
    Dim oUIADesktop   As IUIAutomationElement
    Set oUIADesktop = oUIAutomation.GetRootElement
    If oUIADesktop Is Nothing Then Exit Function
    Set allChilds = oUIADesktop.FindAll(TreeScope_Children, oUIAutomation.CreateTrueCondition)
    Set GetUIADesktopElements = allChilds
poop:
End Function
Friend Function WalkUIADesktopElements() As IUIAutomationElement()
    On Error GoTo poop
    Dim el()          As IUIAutomationElement
    Dim oUIAutomation As New CUIAutomation
    Dim oUIADesktop   As IUIAutomationElement
    Dim allChilds     As IUIAutomationElementArray
    Dim oUIElement    As IUIAutomationElement
    Dim oTW           As IUIAutomationTreeWalker
    Set oUIADesktop = oUIAutomation.GetRootElement
    If oUIADesktop Is Nothing Then Exit Function
    Set oTW = oUIAutomation.ControlViewWalker
    Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
    Dim i As Long
    Do
        If oUIElement Is Nothing Then Exit Do ' exit loop
        On Error GoTo skip
        ReDim Preserve el(i)
        Set el(i) = oUIElement
skip:
        Set oUIElement = oTW.GetNextSiblingElement(oUIElement)
        i = i + 1
    Loop
poop:
    WalkUIADesktopElements = el
End Function
Friend Function ProcessIdToHandleUIA(ByVal pid As Long) As Long
    ProcessIdToHandleUIA = 0
    On Error GoTo poop
    Dim el()          As IUIAutomationElement
    Dim oUIAutomation As New CUIAutomation
    Dim oUIADesktop   As IUIAutomationElement
    Dim allChilds     As IUIAutomationElementArray
    Dim oUIElement    As IUIAutomationElement
    Dim oTW           As IUIAutomationTreeWalker
    Set oUIADesktop = oUIAutomation.GetRootElement
    If oUIADesktop Is Nothing Then Exit Function
    Set oTW = oUIAutomation.ControlViewWalker
    Set oUIElement = oTW.GetFirstChildElement(oUIADesktop)
    ' auwlk = Windows.Automation.TreeWalker.ControlViewWalker.GetFirstChild(AutomationElement.RootElement)
    Dim i As Long
    For i = 1 To 1000
        On Error GoTo skip
        If oUIElement Is Nothing Then Exit For
        Dim hwnd As Long
        ' hwnd = GetUIAHandle(oUIElement)
        '        If hwnd <> 0 And pid = GetWindowProcessId(hwnd) Then
        '            ProcessIdToHandleUIA = hwnd: Exit Function
        '        End If
        '        oUIElement = oTW.GetNextSibling(oUIElement)
skip:
    Next
poop:
End Function
'Friend Function GetUIAHandle(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAHandle = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_NativeWindowHandlePropertyId)
'skip:
'    GetUIAHandle = hwnd
'End Function
'Friend Function GetUIAProcessId(ByVal el As IUIAutomationElement) As Long
'    Dim pid As Long: pid = 0
'    GetUIAProcessId = pid
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    pid = el.GetCurrentPropertyValue(UIA_ProcessIdPropertyId)
'skip:
'    GetUIAProcessId = pid
'End Function
''Friend Function GetUIAControlType(ByVal el As IUIAutomationElement) As Long
''    Dim hwnd As Long: hwnd = 0
''    GetUIAControlType = hwnd
''    If el Is Nothing Then Exit Function
''    On Error GoTo skip
''    hwnd = el.GetCurrentPropertyValue(UIA_ControlTypePropertyId)
''skip:
''    GetUIAControlType = hwnd
''End Function
'Friend Function GetUIALocalizedControlType(ByVal el As IUIAutomationElement) As String
'    Dim lct As String: lct = ""
'    GetUIALocalizedControlType = lct
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    lct = el.GetCurrentPropertyValue(UIA_LocalizedControlTypePropertyId)
'skip:
'    GetUIALocalizedControlType = lct
'End Function
'Friend Function GetUIANameProperty(ByVal el As IUIAutomationElement) As String
'    Dim name As String: name = ""
'    GetUIANameProperty = name
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    name = el.GetCurrentPropertyValue(UIA_NamePropertyId)
'skip:
'    GetUIANameProperty = name
'End Function
'Friend Function GetUIAClassName(ByVal el As IUIAutomationElement) As String
'    Dim cname As String: cname = ""
'    GetUIAClassName = cname
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    cname = el.GetCurrentPropertyValue(UIA_ClassNamePropertyId)
'skip:
'    GetUIAClassName = cname
'End Function
'Friend Function GetUIAFullDescription(ByVal el As IUIAutomationElement) As String
'    Dim desc As String: desc = ""
'    GetUIAFullDescription = desc
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    desc = el.GetCurrentPropertyValue(UIA_FullDescriptionPropertyId)
'skip:
'    GetUIAFullDescription = desc
'End Function
'Friend Function GetUIAHelpText(ByVal el As IUIAutomationElement) As String
'    Dim help As String: help = ""
'    GetUIAHelpText = help
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    help = el.GetCurrentPropertyValue(UIA_HelpTextPropertyId)
'skip:
'    GetUIAHelpText = help
'End Function
'
'
'Friend Function GetUIAHasKeyboardFocus(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAHasKeyboardFocus = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_HasKeyboardFocusPropertyId)
'skip:
'    GetUIAHasKeyboardFocus = b
'End Function
'Friend Function GetUIAIsEnabled(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsEnabled = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsEnabledPropertyId)
'skip:
'    GetUIAIsEnabled = b
'End Function
'Friend Function GetUIAIsKeyboardFocusable(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsKeyboardFocusable = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsKeyboardFocusablePropertyId)
'skip:
'    GetUIAIsKeyboardFocusable = b
'End Function
'Friend Function GetUIAIsControlElement(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsControlElement = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsControlElementPropertyId)
'skip:
'    GetUIAIsControlElement = b
'End Function
'Friend Function GetUIAIsContentElement(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsContentElement = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsContentElementPropertyId)
'skip:
'    GetUIAIsContentElement = b
'End Function
'Friend Function GetUIAIsOffscreen(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsOffscreen = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsOffscreenPropertyId)
'skip:
'    GetUIAIsOffscreen = hwnd
'End Function
'Friend Function GetUIAIsPassword(ByVal el As IUIAutomationElement) As Boolean
'    Dim b As Boolean: b = False
'    GetUIAIsPassword = b
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    b = el.GetCurrentPropertyValue(UIA_IsPasswordPropertyId)
'skip:
'    GetUIAIsPassword = b
'End Function
'Friend Function GetUIABoundingRectangle(ByVal el As IUIAutomationElement) As Variant
'    Dim hwnd As Variant
'    GetUIABoundingRectangle = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId)
'skip:
'    GetUIABoundingRectangle = hwnd
'End Function
'Friend Function GetUIASize(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIASize = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_SizePropertyId)
'skip:
'    GetUIASize = hwnd
'End Function
'Friend Function GetUIAClickablePoint(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAClickablePoint = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_ClickablePointPropertyId)
'skip:
'    GetUIAClickablePoint = hwnd
'End Function
'Friend Function GetUIACenterPoint(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIACenterPoint = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_CenterPointPropertyId)
'skip:
'    GetUIACenterPoint = hwnd
'End Function
'Friend Function GetUIAItemStatus(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAItemStatus = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_ItemStatusPropertyId)
'skip:
'    GetUIAItemStatus = hwnd
'End Function
'Friend Function GetUIAItemType(ByVal el As IUIAutomationElement) As Long
'    Dim hwnd As Long: hwnd = 0
'    GetUIAItemType = hwnd
'    If el Is Nothing Then Exit Function
'    On Error GoTo skip
'    hwnd = el.GetCurrentPropertyValue(UIA_ItemTypePropertyId)
'skip:
'    GetUIAItemType = hwnd
'End Function
'
'
'
'
'
'
Friend Function GetElementFromPoint() As IUIAutomationElement
    '    On Error GoTo poop
    '    Dim UI As New UIAutomationClient.CUIAutomation
    '    Dim El As UIAutomationClient.IUIAutomationElement
    '    Dim tag As UIAutomationClient.tagPOINT
    '    Set GetElementFromPoint = El
    '    Set UI = New CUIAutomation
    '    Dim p As POINTAPI
    '    If apiGetCursorPos(p) = 0 Then Exit Function
    '    tag.x = p.x
    '    tag.y = p.y
    '    Set El = UI.ElementFromPoint(tag)
    '    Set GetElementFromPoint = El
    'poop:
End Function
Friend Function MainWindowHandles(ByVal pid As Long) As Long()
    Dim hwnds() As Long
    '    Dim c() As Long
    '    c = ChildWindows(0)
    '    Dim n As Long: n = 0
    '    Dim i As Long: i = 0
    '    Do
    '        On Error GoTo skip
    '        Dim id As Long
    '        id = GetPIDFromHWND(c(i))
    '        If id = pid Then
    '            MsgBox id & " " & c(i)
    '
    '            ReDim Preserve hwnds(n)
    '            hwnds(n) = c(i)
    '            n = n + 1
    '        End If
    'skip:
    '        If i = UBound(c) Then Exit Do
    '        i = i + 1
    '    Loop
    MainWindowHandles = hwnds
End Function
Friend Function MainWindowTitles(ByVal pid As Long) As String()
    Dim titles() As String
    '    Dim c() As Long
    '    c = ChildWindows(0)
    '    Dim n As Long: n = 0
    '    Dim i As Long: i = 0
    '    For i = 0 To UBound(c)
    '        On Error GoTo skip
    '        Dim id As Long
    '        id = GetPIDFromHWND(c(i))
    '        If id = pid Then
    '            ReDim Preserve titles(n)
    '            ' titles(n) = c(i)
    '            n = n + 1
    '        End If
    'skip:
    '    Next
    MainWindowTitles = titles
End Function
Friend Sub CloseMainWindows(ByVal pid As Long)
    '    Dim c() As Long
    '    c = ChildWindows(0)
    '    Dim n As Long: n = 0
    '    Dim i As Long: i = 0
    '    For i = 0 To UBound(c)
    '        On Error GoTo skip
    '        Dim id As Long
    '        id = GetPIDFromHWND(c(i))
    '        If id = pid Then
    '            'apiPostMessage(c(i), WM_CLOSE, 0,0)
    '            n = n + 1
    '        End If
    'skip:
    '    Next
End Sub

'Friend Function Message(ByVal vKey As Keys, ByRef WFOCUS As WINFOCUS, Optional ByVal kDown As Boolean = True, Optional ByVal kUp As Boolean = True, Optional ByVal bPost As Boolean = False) As Boolean
'    On Error Resume Next
'    Message = True '''''''''''''''''''''''''''''''Set default return value
'    If bPost = True Then '''''''''''''''''''''''''If post specified
'        If kDown = True Then
'            If apiPostMessage(WFOCUS.Focus, WM_KEYDOWN, vKey, vbNullString) = False Then Message = False 'If key is to be pressed down
'        End If
'        If kUp = True Then
'            If apiPostMessage(WFOCUS.Focus, WM_KEYUP, vKey, vbNullString) = False Then Message = False 'If key is to be lifted up
'        End If
'    Else '''''''''''''''''''''''''''''''''''''''''If Sending message
'        If kDown = True Then
'            If apiSendMessage(WFOCUS.Focus, WM_KEYDOWN, vKey, vbNullString) = False Then Message = False 'If key is to be pressed down
'        End If
'        If kUp = True Then
'            If apiSendMessage(WFOCUS.Focus, WM_KEYUP, vKey, vbNullString) = False Then Message = False 'If key is to be lifted up
'        End If
'    End If
'End Function
Friend Function MakeWindowVisible(ByVal hwnd As Long)
    If apiIsWindow(hwnd) = False Then MakeWindowVisible = 0: Exit Function ''Make sure it's a valid window, if not abort
    '
    '    If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL) 'If window is minimized then show it, and remember
    '    If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True) 'If window is disabled then enable it and remember
    '    If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True) 'If child window is disabled then enable it and remember
    '    If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL) 'If window is hidden then show it, and remember
    '    If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL) 'If child window is hidden then show it, and remember
End Function

