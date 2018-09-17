VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Public Enum Buttons
'    LeftClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
'    LeftDoubleClick = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP + MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP
'    MiddleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
'    MiddleDoubleClick = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP + MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP
'    Move = MOUSEEVENTF_MOVE
'    MoveAbsolute = MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE
'    RightClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
'    RightDoubleClick = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP + MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP
'    VirtualDesk = MOUSEEVENTF_VIRTUALDESK
'    Wheel = MOUSEEVENTF_WHEEL
'    xClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
'    xDoubleClick = MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP + MOUSEEVENTF_XDOWN + MOUSEEVENTF_XUP
'End Enum
'' This function sends window messages directly to a window.
'' If the asMessage parameter is false then, this function simulates mouse events directly to a window.
'' Returns false if the specified window/s cannot be found.
'' If the asMessage parameter is false then, this function returns false if a handle cannot be established or
'' the top window could not be set, or a window rectangle could not be found, or the cursor could not be set,
'' or the mouse event could not be simulated.
'' "mButtons" The button to click.  Use the Buttons enumeration for this parameter.
''  "wFocus" The structure of the window to focus on.
'' Use GetWinHandles or GetWinFocus to get a structure.
''  "mDown" (Optional) Press mouse button down only.
''  "mUp" (Optional) Lift mouse button up only.
''  "asMessage" (Optional) Send as window message, or as a mouse event.
'' It's recommended that you use True in this parameter.
'' If false then the x and y parameters may apply.
''  "x" (Optional) The x coordinate only applies if asMessage is false, and mButtons is
'' Buttons.Move or Buttons.MoveAbsolute.
''  "y" (Optional) The y coordinate only applies if asMessage is false, and mButtons is
'' Buttons.Move or Buttons.MoveAbsolute.
'Friend Function Click(ByVal mButtons As Long, ByRef WFOCUS As WINFOCUS, Optional ByVal mDown As Boolean = True, Optional ByVal mUp As Boolean = True, Optional ByVal asMessage As Boolean = True, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1) As Boolean
'    On Error Resume Next
'    Dim hwnd        As Long
'    Dim cwnd        As Long
'    Dim zOrder      As Long
'    Dim zOrderChild As Long
'    Dim p           As POINTAPI
'    Dim ws          As WINSTATE
'    Click = True
'    If WFOCUS.Foreground = -1 Then Click = False: Exit Function
'    If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False)
'    If asMessage = True Then  ''''''''''''''''''''If sending as message
'        Dim WM_DOWN As Long
'        Dim WM_UP   As Long
'        If mButtons = Buttons.Move Then ''''''''''If moving from current position
'            Dim cP As POINTAPI
'            Call apiGetCursorPos(cP)  '''''''''''''Get cursor position
'            Call apiSetCursorPos(cP.x + x, cP.y + y) 'Add to coordinates, and set cursor there
'            Click = True: Exit Function ''''''''''Return success
'        ElseIf mButtons = Buttons.MoveAbsolute Then
'            Call apiSetCursorPos(x, y) '''''''''''Set to absolute coordinates
'            Click = False: Exit Function '''''''''Return success
'        ElseIf mButtons = Buttons.LeftClick Or mButtons = Buttons.LeftDoubleClick Then 'If button specified is left
'            WM_DOWN = WM_LBUTTONDOWN: WM_UP = WM_LBUTTONUP 'Then set left messages
'        ElseIf mButtons = Buttons.RightClick Or mButtons = Buttons.RightDoubleClick Then 'If button specified is right
'            WM_DOWN = WM_RBUTTONDOWN: WM_UP = WM_RBUTTONUP 'Then set right messages
'        ElseIf mButtons = Buttons.MiddleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If button specified is middle
'            WM_DOWN = WM_MBUTTONDOWN: WM_UP = WM_MBUTTONUP ' Then set middle messages
'        End If
'        hwnd = WFOCUS.Foreground '''''''''''''''''Set main handle to something smaller
'        cwnd = WFOCUS.Focus ''''''''''''''''''''''Set child handle(if any) to something smaller
'        If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
'        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'if minimized then show normal
'        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enable
'        If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True): Sleep (25) 'If child disabled then enable
'        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show
'        If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL): Sleep (25) 'If child hidden then show
'        zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
'        zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
'        If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
'        If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
'        If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'If button down is specified, then press down
'        If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'If button up is specified, then lift up
'        If mButtons = Buttons.LeftDoubleClick Or mButtons = Buttons.RightDoubleClick Or mButtons = Buttons.MiddleDoubleClick Then 'If it's a double click
'            If mDown = True Then Call apiSendMessage(cwnd, WM_DOWN, 0, vbNullString) 'Down again
'            If mUp = True Then Call apiSendMessage(cwnd, WM_UP, 0, vbNullString) 'Up again
'        End If
'        Call MouseAbort(ws, False, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
'    Else '''''''''''''''''''''''''''''''''''''''''Else it's an event to send
'        Dim repeat  As Long
'        Dim rCursor As Boolean
'        Dim pts     As POINTAPI
'        If mButtons = Buttons.Move Then  '''''''''Move cursor to click point
'            Call MouseEvent(mButtons, x, y) ''''''Move cursor to click point
'            Exit Function '''''''''''''''''''''''''''''''''Exit thread
'        ElseIf mButtons = Buttons.MoveAbsolute Then
'            pts = ToScreen(x, y)   ''''''''''''''Convert to screen coordinates
'            If pts.x <> 0 Then x = pts.x  ''''''''If x point found then set
'            If pts.y <> 0 Then y = pts.y  ''''''''If y point found then set
'            Call MouseEvent(mButtons, x, y) ''''''Move cursor to click point
'            Exit Function ''''''''''''''''''''''''Exit thread
'        End If
'        If WFOCUS.Foreground = -1 Then Exit Function  'Exit if return from GetWinHandles is negative
'        If WFOCUS.Foreground = 0 Then WFOCUS = GetWinFocus(False, False) 'Get current focus
'        hwnd = WFOCUS.Foreground  ''''''''''''''''Set main handle to something smaller
'        cwnd = WFOCUS.Focus  '''''''''''''''''''''Set child handle(if any) to something smaller
'        If cwnd = 0 Then cwnd = hwnd '''''''''''''If no child specified, then set it to the main window
'        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'if minimized then show normal
'        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enable
'        If apiIsWindowEnabled(cwnd) = False Then ws.IsChildDisabled = apiEnableWindow(cwnd, True): Sleep (25) 'If child disabled then enable
'        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show
'        If apiIsWindowVisible(cwnd) = False Then ws.IsChildHidden = Not apiShowWindow(cwnd, SW_SHOWNORMAL): Sleep (25) 'If child hidden then show
'        zOrder = GetSetZOrder(hwnd) ''''''''''''''Remember main window's place in the z-order
'        zOrderChild = GetSetZOrder(cwnd) '''''''''Remember child window's place in the z-order
'        If hwnd <> apiGetTopWindow(HWND_DESKTOP) Then Call apiSetWindowPos(hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
'        If cwnd <> apiGetTopWindow(apiGetParent(cwnd)) Then Call apiSetWindowPos(cwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE)
'        If mButtons = 12 Or mButtons = 48 Or mButtons = 768 Then repeat = 1    'If double click
'        If mDown = True And mUp = False Then '''''If mouse down
'            If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
'                mButtons = MOUSEEVENTF_LEFTDOWN 'Set as left down
'            ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
'                mButtons = MOUSEEVENTF_RIGHTDOWN 'Set as right down
'            ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
'                mButtons = MOUSEEVENTF_MIDDLEDOWN 'Set as middle down
'            End If
'        ElseIf mDown = False And mUp = True Then 'If mouse up
'            If mButtons = MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP Then  'If left click
'                mButtons = MOUSEEVENTF_LEFTUP '''Set as left up
'            ElseIf mButtons = MOUSEEVENTF_RIGHTDOWN + MOUSEEVENTF_RIGHTUP Then  'If right click
'                mButtons = MOUSEEVENTF_RIGHTUP ''Set as right up
'            ElseIf mButtons = MOUSEEVENTF_MIDDLEDOWN + MOUSEEVENTF_MIDDLEUP Then  'If middle click
'                mButtons = MOUSEEVENTF_MIDDLEUP 'Set as middle up
'            End If
'        End If
'        If mButtons <> Buttons.Wheel And mButtons <> Buttons.VirtualDesk Then   'If it's a click
'            Dim r As RECT
'            If apiGetWindowRect(cwnd, r) = False Then 'If no rectangle is found
'                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'If rectangle not found, then exit failure
'                Exit Function ''''''''''''''''''''Exit
'            End If
'            pts.x = CInt(r.rLeft + ((r.rRight - r.rLeft) / 2)) 'Set to the center of the horizon
'            pts.y = CInt(r.rTop + ((r.rBottom - r.rTop) / 2)) 'Set to the center of the vertical
'            pts = ToScreen(pts.x, pts.y) '''''''''Convert to screen coordinates
'            rCursor = True '''''''''''''''''''''''Cursor position changed, remember to return it later
'            x = 0 '''''''''''''''''''''''''''''''null for click
'            y = 0 '''''''''''''''''''''''''''''''null for click
'            Call apiGetCursorPos(p)  '''''''''''''Get the current cursor position, to be returned later
'            If MouseEvent(Buttons.MoveAbsolute, pts.x, pts.y) = False Then 'Move cursor to click point
'                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if fails to move
'                Exit Function ''''''''''''''''''''Exit thread
'            End If
'        End If
'        Dim i As Long
'        For i = 1 To repeat + 1  '''''''''''''''''Loop the number of repeats
'            If MouseEvent(mButtons, x, y) = False Then  'Do mouse event repeated number of times
'                Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Abort if failure
'                Exit Function ''''''''''''''''''''Exit thread
'            End If
'        Next
'        Call MouseAbort(ws, rCursor, hwnd, cwnd, zOrder, zOrderChild, p) 'Final abort return without failure
'    End If
'End Function
'' Clicks on the specified menu item, by sending it a command message.
'' If asMessage is false then this function clicks on the specified menu item,
'' by simulating an entire chain of events.
'' Returns false if window or menu cannot be found.
'' If asMessage is false then this function returns false if a window or menu or rectangle cannot be found,
'' or an event fails to be simulated.
''  "asMessage"Send as window message, or as a mouse event.
'' It's recommended that you use True in this parameter.
''  "wName"Title or class name of the main window.
''  "wIndex"Index of the main window.
''  "mnuName1"Main menu title.
''  "mnuIndex1"Main menu index.
''  "mnuName2"(Optional) Sub menu title.
''  "mnuIndex2"(Optional) Sub menu index.
''  "mnuName3"(Optional) Sub menu title.
''  "mnuIndex3"(Optional) Sub menu index.
''  "mnuName4"(Optional) Sub menu title.
''  "mnuIndex4"(Optional) Sub menu index.
''  "mnuName5"(Optional) Sub menu title.
''  "mnuIndex5"(Optional) Sub menu index.
''  "mnuName6"(Optional) Sub menu title.
''  "mnuIndex6"(Optional) Sub menu index.
''  "mnuName7"(Optional) Sub menu title.
''  "mnuIndex7"(Optional) Sub menu index.
''  "mnuName8"(Optional) Sub menu title.
''  "mnuIndex8"(Optional) Sub menu index.
''  "mnuName9"(Optional) Sub menu title.
''  "mnuIndex9"(Optional) Sub menu index.
''  "mnuName10"Sub menu title.
''  "mnuIndex10"(Optional) Sub menu index.</param>
'Friend Function ClickMenu(ByVal asMessage As Boolean, ByVal wName As String, ByVal wIndex As Long, ByVal mnuName1 As String, Optional ByVal mnuIndex1 As Long = 1, Optional ByVal mnuName2 As String = " ", Optional ByVal mnuIndex2 As Long = 1, Optional ByVal mnuName3 As String = " ", Optional ByVal mnuIndex3 As Long = 1, Optional ByVal mnuName4 As String = " ", Optional ByVal mnuIndex4 As Long = 1, Optional ByVal mnuName5 As String = " ", Optional ByVal mnuIndex5 As Long = 1, Optional ByVal mnuName6 As String = " ", Optional ByVal mnuIndex6 As Long = 1, Optional ByVal mnuName7 As String = " ", Optional ByVal mnuIndex7 As Long = 1, Optional ByVal mnuName8 As String = " ", Optional ByVal mnuIndex8 As Long = 1, Optional ByVal mnuName9 As String = " ", Optional ByVal mnuIndex9 As Long = 1, Optional ByVal mnuName10 As String = " ", Optional ByVal mnuIndex10 As Long = 1) As Boolean
'    On Error Resume Next
'    Dim n          As Long
'    Dim mnuNames() As String
'    Dim i          As Long
'    Dim nPos       As Long
'    Dim ws         As WINSTATE
'    If asMessage = True Then
'        Dim hwnd     As Long
'        Dim hMenu    As Long
'        Dim hSubMenu As Long
'        Dim hId      As Long
'        hwnd = GetWinHandles(wName, wIndex).Foreground 'Set handle to specification
'        If hwnd = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
'        If apiIsIconic(hwnd) = True Then ws.IsIconic = apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If minimized then show it
'        If apiIsWindowEnabled(hwnd) = False Then ws.IsDisabled = apiEnableWindow(hwnd, True): Sleep (25) 'If disabled then enabled it
'        If apiIsWindowVisible(hwnd) = False Then ws.IsHidden = Not apiShowWindow(hwnd, SW_SHOWNORMAL): Sleep (25) 'If Hidden then show it
'        hMenu = apiGetMenu(hwnd) '''''''''''''''''Set handle of main menu
'        If hMenu = 0 Then ClickMenu = False: Exit Function 'If handle not found then exit failure
'        hSubMenu = hMenu '''''''''''''''''''''''''Initialize sub menu to main menu
'        If mnuName1 <> " " Then n = 1: ReDim mnuNames(1) 'Re-dimension mnuNames array
'        If mnuName2 <> " " Then n = 3: ReDim mnuNames(3)
'        If mnuName3 <> " " Then n = 5: ReDim mnuNames(5)
'        If mnuName4 <> " " Then n = 7: ReDim mnuNames(7)
'        If mnuName5 <> " " Then n = 9: ReDim mnuNames(9)
'        If mnuName6 <> " " Then n = 11: ReDim mnuNames(11)
'        If mnuName7 <> " " Then n = 13: ReDim mnuNames(13)
'        If mnuName8 <> " " Then n = 15: ReDim mnuNames(15)
'        If mnuName9 <> " " Then n = 17: ReDim mnuNames(17)
'        If mnuName10 <> " " Then n = 19: ReDim mnuNames(19)
'        If n > 0 Then mnuNames(0) = mnuName1: mnuNames(1) = CStr(mnuIndex1) 'Set elements of array to specification
'        If n > 2 Then mnuNames(2) = mnuName2: mnuNames(3) = CStr(mnuIndex2)
'        If n > 4 Then mnuNames(4) = mnuName3: mnuNames(5) = CStr(mnuIndex3)
'        If n > 6 Then mnuNames(6) = mnuName4: mnuNames(7) = CStr(mnuIndex4)
'        If n > 8 Then mnuNames(8) = mnuName5: mnuNames(9) = CStr(mnuIndex5)
'        If n > 10 Then mnuNames(10) = mnuName6: mnuNames(11) = CStr(mnuIndex6)
'        If n > 12 Then mnuNames(12) = mnuName7: mnuNames(13) = CStr(mnuIndex7)
'        If n > 14 Then mnuNames(14) = mnuName8: mnuNames(15) = CStr(mnuIndex8)
'        If n > 16 Then mnuNames(16) = mnuName9: mnuNames(17) = CStr(mnuIndex9)
'        If n > 18 Then mnuNames(18) = mnuName10: mnuNames(19) = CStr(mnuIndex10)
'        For i = 0 To (n - 1) Step 2  '''''''''''''Loop through menu tree
'            If mnuNames(i) = "" Then '''''''''''''If menu name not specified then
'                nPos = CInt(mnuNames(i + 1)) - 1 'Set position as an index
'            Else '''''''''''''''''''''''''''''''''Then name was set
'                nPos = FindMenuItemPos(hSubMenu, mnuNames(i), CInt(mnuNames(i + 1))) 'Set position to name and index of specified item.
'            End If
'            If nPos <> -1 Then
'                If apiGetSubMenu(hSubMenu, nPos) <> 0 Then hSubMenu = apiGetSubMenu(hSubMenu, nPos) 'If sub menu exits, and is specified then get the handle
'            End If
'        Next
'        If nPos <> -1 Then '''''''''''''''''''''''If final item has a valid position
'            hId = apiGetMenuItemID(hSubMenu, nPos) 'Get menu id
'            If hId <> -1 Then ''''''''''''''''''''If item has no sub menus
'                ClickMenu = Not CBool(apiSendMessage(hwnd, WM_COMMAND, hId, vbNullString)) 'Send command message
'            Else
'                ClickMenu = False ''''''''''''''''Return failure
'            End If
'        End If
'        If ws.IsIconic = True Then Sleep (25): Call apiShowWindow(hwnd, SW_SHOWMINIMIZED) 'If window was minimized then re-minimize it
'        If ws.IsDisabled = True Then Sleep (25): Call apiEnableWindow(hwnd, False) 'If window was disabled then re-disable it
'        If ws.IsHidden = True Then Sleep (25): Call apiShowWindow(hwnd, SW_HIDE) 'If window was hidden then re-hide it
'    Else
'        Dim rOffset     As Long
'        Dim tOffset     As Long
'        Dim LeftMost    As Long
'        Dim TopMost     As Long
'        Dim ArrayLength As Long
'        Dim m           As MENUINFO
'        Dim MI          As ITEMINFO
'        Dim p           As POINTAPI
'        Dim poi         As POINTAPI
'        Dim r           As RECT
'        ClickMenu = True '''''''''''''''''''''''''Signal true indicating that the thread was started
'        If mnuName1 <> " " Then n = 3: ReDim mnuNames(3)  'Re-dimension mnuNames array
'        If mnuName2 <> " " Then n = 5: ReDim mnuNames(5)
'        If mnuName3 <> " " Then n = 7: ReDim mnuNames(7)
'        If mnuName4 <> " " Then n = 9: ReDim mnuNames(9)
'        If mnuName5 <> " " Then n = 11: ReDim mnuNames(11)
'        If mnuName6 <> " " Then n = 13: ReDim mnuNames(13)
'        If mnuName7 <> " " Then n = 15: ReDim mnuNames(15)
'        If mnuName8 <> " " Then n = 17: ReDim mnuNames(17)
'        If mnuName9 <> " " Then n = 19: ReDim mnuNames(19)
'        If mnuName10 <> " " Then n = 21: ReDim mnuNames(21)
'        If n > 0 Then mnuNames(0) = wName: mnuNames(1) = wIndex  'Set elements of array to specification
'        If n > 2 Then mnuNames(2) = mnuName1: mnuNames(3) = CStr(mnuIndex1)
'        If n > 4 Then mnuNames(4) = mnuName2: mnuNames(5) = CStr(mnuIndex2)
'        If n > 6 Then mnuNames(6) = mnuName3: mnuNames(7) = CStr(mnuIndex3)
'        If n > 8 Then mnuNames(8) = mnuName4: mnuNames(9) = CStr(mnuIndex4)
'        If n > 10 Then mnuNames(10) = mnuName5: mnuNames(11) = CStr(mnuIndex5)
'        If n > 12 Then mnuNames(12) = mnuName6: mnuNames(13) = CStr(mnuIndex6)
'        If n > 14 Then mnuNames(14) = mnuName7: mnuNames(15) = CStr(mnuIndex7)
'        If n > 16 Then mnuNames(16) = mnuName8: mnuNames(17) = CStr(mnuIndex8)
'        If n > 18 Then mnuNames(18) = mnuName9: mnuNames(19) = CStr(mnuIndex9)
'        If n > 20 Then mnuNames(20) = mnuName10: mnuNames(21) = CStr(mnuIndex10)
'        ArrayLength = n ''''''''''''''''''''''''''Set length of array
'        m.hwnd = GetWinHandles(wName).Foreground 'Get the handle of the specified window
'        If m.hwnd = 0 Then ClickMenu = False: Exit Function 'If window not found, then exit failure
'        If apiIsIconic(m.hwnd) = True Then ws.IsIconic = apiShowWindow(m.hwnd, SW_SHOWNORMAL): Sleep (25) 'If minimized then show it
'        If apiIsWindowEnabled(m.hwnd) = False Then ws.IsDisabled = apiEnableWindow(m.hwnd, True): Sleep (25) 'If disabled then enable it
'        If apiIsWindowVisible(m.hwnd) = False Then ws.IsHidden = Not apiShowWindow(m.hwnd, SW_SHOWNORMAL): Sleep (25) 'If hidden then show it
'        Call apiGetWindowRect(m.hwnd, r) '''''''''Set confirmation of window rectangle
'        If r.rBottom <> 0 Then
'            Call apiMoveWindow(m.hwnd, 0, 0, r.rRight - r.rLeft, r.rBottom - r.rTop, True): Sleep (25) 'If rectangle found then move window with coordinates
'        End If
'        Call apiGetCursorPos(poi)  '''''''''''''''Get the current position of the user's cursor, so that it can be returned
'        m.hMenu = apiGetMenu(m.hwnd) '''''''''''''Set handle of the main menu
'        If m.hMenu = 0 Then ClickMenu = False: Exit Function '''''''''''''If no handle found then exit sub with failure
'        m.hSubMenu = apiGetSubMenu(m.hMenu, 0) '''Set handle of the first sub menu if any
'        MI = MenuItemDim(m.hwnd, m.hMenu, 0) '''''Get the dimensions of the menu item
'        If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon negative results
'        LeftMost = MI.Left '''''''''''''''''''''''Initialize the left most coordinate
'        If mnuNames(2) = "" Then '''''''''''''''''If no name specified
'            nPos = CInt(mnuNames(3)) - 1 '''''''''Set position by index only
'        Else '''''''''''''''''''''''''''''''''''''Otherwise set position by name and index
'            nPos = FindMenuItemPos(m.hMenu, mnuNames(2), CInt(mnuNames(3))) 'Find position of the menu item
'        End If
'        If nPos = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
'        MI = MenuItemDim(m.hwnd, m.hMenu, nPos) ''Get item  dimensions
'        If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit upon failure
'        p = ToScreen(MI.Center.x, MI.Center.y) '''Convert point to screen coordinates
'        If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then ClickMenu = False: Exit Function 'Move mouse, and exit if failure
'        rOffset = MI.Left - LeftMost '''''''''''''Initialize offset from the left
'        TopMost = MI.Bottom ''''''''''''''''''''''Initialize offset from the top
'        If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click mouse, and exit upon failure
'        m.hSubMenu = apiGetSubMenu(m.hMenu, nPos) 'Set handle of submenu
'        If m.hSubMenu <> 0 Then ''''''''''''''''''If handle found
'            If mnuNames(4) = "" Then '''''''''''''''''If no name specified
'                nPos = CInt(mnuNames(5)) - 1 '''''''''Set by index only
'            Else '''''''''''''''''''''''''''''''''''''Name and index specified
'                nPos = FindMenuItemPos(m.hSubMenu, mnuNames(4), CInt(mnuNames(5))) 'Find by name and index
'            End If
'            If nPos = -1 Then ClickMenu = False: Exit Function 'Exit if position is invalid
'            MI = MenuItemDim(m.hwnd, m.hSubMenu, nPos) 'Get dimensinos
'            If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then ClickMenu = False: Exit Function 'Exit if fails
'            p = ToScreen(rOffset + MI.Center.x, MI.Center.y) 'Convert point
'            If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then ClickMenu = False: Exit Function 'Move and exit if failure
'            For i = 6 To ArrayLength Step 2 ''''''''''Step through the array
'                If MoveItemToItem(mnuNames(i), mnuNames(i + 1), TopMost, nPos, rOffset, tOffset, m, MI) = False Then ClickMenu = False: Exit Function
'            Next
'            If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then ClickMenu = False: Exit Function 'Click final menu item in the chain, and exit if failure
'        End If
'        If r.rBottom <> 0 Then
'            Call apiMoveWindow(m.hwnd, r.rLeft, r.rTop, r.rRight - r.rLeft, r.rBottom - r.rTop, True) 'If there was a rectangle move the window back to where it was
'        End If
'        Call apiSetCursorPos(poi.x, poi.y) '''''''Return the position of the cursor back to the user
'        If ws.IsIconic = True Then Sleep (25): Call apiShowWindow(m.hwnd, SW_SHOWMINIMIZED) 'If was minimized then re-minimize
'        If ws.IsDisabled = True Then Sleep (25): Call apiEnableWindow(m.hwnd, False) 'If was disabled then re-disable
'        If ws.IsHidden = True Then Sleep (25): Call apiShowWindow(m.hwnd, SW_HIDE) 'If was Hidden then re-Hide
'    End If
'End Function
'Private Function FindMenuItemPos(ByVal hMenu As Long, Optional ByVal iName As String = "", Optional ByVal iIndex As Long = 1) As Long
'    On Error Resume Next
'    Dim i          As Long
'    Dim itemCount  As Long
'    Dim indexCount As Long
'    Dim retValue   As Long
'    Dim mnuCaption As String
'    Dim woShortcut As String
'    If apiIsMenu(hMenu) = False Then FindMenuItemPos = -1: Exit Function 'Return negative result if it's not a menu handle
'    FindMenuItemPos = NEGATIVE '''''''''''''''''''Set a default return value
'    itemCount = apiGetMenuItemCount(hMenu) '''''''Count the number of menu items
'    For i = 0 To itemCount - 1 '''''''''''''''''''Loop through all menu items
'        mnuCaption = "" ''''''''''''''''''''''''''Initialize
'        mnuCaption = Space(1024) '''''''''''''''''Pad with a buffer
'        retValue = apiGetMenuString(hMenu, i, mnuCaption, Len(mnuCaption), 1024) 'Get menu caption
'        mnuCaption = Left(mnuCaption, retValue) ''mnuCaption.Substring(0, retValue) 'Strip off buffer
'        woShortcut = "" ''''''''''''''''''''''''''Initialize
'        If InStr(mnuCaption, "&") = True Then woShortcut = Replace(mnuCaption, "&", "") 'If the & character exists, then remove it, so the developer doesn't have to specify
'        apiCharLower (iName) '''''''''''''''''''''Convert to lower case
'        apiCharLower (woShortcut)
'        apiCharLower (mnuCaption)
'        If iName = woShortcut Or iName = mnuCaption Then 'if specified name matches menu name, as non-case sensitive
'            FindMenuItemPos = i ''''''''''''''''''Set return value as that position
'            indexCount = indexCount + 1 ''''''''''Increment index by one
'            If indexCount = iIndex Then Exit For 'If index matches the specification then exit loop
'        End If
'    Next
'End Function
'Private Function GetSetZOrder(ByVal hwnd As Long, Optional ByVal sPosition As Long = NEGATIVE) As Long
'    On Error Resume Next
'    Dim z    As Long
'    Dim swnd As Long
'    swnd = apiGetTopWindow(0)  'apiGetWindow(hwnd, GW_HWNDFIRST) ''''''Get top or topmost window in context
'    Do
'        If sPosition = NEGATIVE Then '''''''''''''If not setting the z-order
'            If swnd = hwnd Then GetSetZOrder = (z + 1): Exit Function 'If handle specified matches sibling window, then return the position in the z-order
'        Else '''''''''''''''''''''''''''''''''''''Then this function sets the z-order to the specified position
'            If z = sPosition - 1 Then GetSetZOrder = apiSetWindowPos(hwnd, swnd, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE): Exit Function
'        End If
'        swnd = apiGetWindow(swnd, GW_HWNDNEXT) '''Get the next sibling window handle, in the loop
'        If swnd = 0 Then Exit Do '''''''''''''''''If there are no more sibling windows, then exit loop with default return
'        z = z + 1 ''''''''''''''''''''''''''''''''Increment i by one
'    Loop
'    GetSetZOrder = swnd
'End Function
'Private Function MenuItemDim(ByVal hwnd As Long, ByVal hMenu As Long, ByVal nPos As Long) As ITEMINFO
'    On Error Resume Next
'    Dim m As ITEMINFO
'    Dim r As RECT
'    If apiGetMenuItemRect(hwnd, hMenu, nPos, r) = 0 Then 'If rectangle not found then set negative returns
'        r.rTop = NEGATIVE ''''''''''''''''''''''''Fail
'        r.rBottom = NEGATIVE '''''''''''''''''''''Fail
'        r.rLeft = NEGATIVE '''''''''''''''''''''''Fail
'        r.rRight = NEGATIVE ''''''''''''''''''''''Fail
'    Else '''''''''''''''''''''''''''''''''''''''''Or set dimensions of menu item
'        m.Width = (r.rRight - r.rLeft) '''''''''''Set width
'        m.Height = (r.rBottom - r.rTop) ''''''''''Set height
'        m.Center.x = CInt(r.rLeft + (m.Width / 2)) 'Set center point x
'        m.Center.y = CInt(r.rTop + (m.Height / 2)) 'Set center point y
'    End If
'    m.Left = r.rLeft '''''''''''''''''''''''''''''Set left coordinate
'    m.Right = r.rRight '''''''''''''''''''''''''''Set right coordinate
'    m.Top = r.rTop '''''''''''''''''''''''''''''''Set top coordinate
'    m.Bottom = r.rBottom '''''''''''''''''''''''''Set bottom coordinate
'    MenuItemDim = m
'End Function
'Private Function MouseAbort(ByRef WSTATE As WINSTATE, ByVal rCursor As Boolean, ByVal hwnd As Long, ByVal cwnd As Long, ByVal zOrder As Long, ByVal zOrderChild As Long, ByRef p As POINTAPI) As Boolean
'    On Error Resume Next
'    If rCursor = True Then Call apiSetCursorPos(p.x, p.y) 'If it was a click then return cursor to user position
'    If WSTATE.IsIconic = True Then Sleep (25): Call apiShowWindow(hwnd, SW_SHOWMINIMIZED) 'If main window was minimized before, then re-minimize it
'    If WSTATE.IsDisabled = True Then Sleep (25): Call apiEnableWindow(hwnd, False) 'If main window was disabled before, then re-disable it
'    If WSTATE.IsChildDisabled = True Then Sleep (25): Call apiEnableWindow(cwnd, False) 'If child window was disabled before, then re-disable it
'    If WSTATE.IsHidden = True Then Sleep (25): Call apiShowWindow(hwnd, SW_HIDE) 'If main window was hidden before, then re-hide it
'    If WSTATE.IsChildHidden = True Then Sleep (25): Call apiShowWindow(cwnd, SW_HIDE) 'If main window was hidden before, then re-hide it
'    Call apiSetWindowPos(hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE + SWP_SHOWWINDOW + SWP_NOMOVE + SWP_NOSIZE): Sleep (25) 'Set the window position to not topmost window.  TODO Get topmost status first, so original state can be restored.
'    If zOrderChild > 0 Then Sleep (25): Call GetSetZOrder(cwnd, zOrderChild) 'If a z-order for the child window was obtained, reset the z-order of the child window
'    If zOrder > 0 Then Sleep (25): Call GetSetZOrder(hwnd, zOrder) 'If a z-order for the main window was obtained, reset the z-order of the main window
'    MouseAbort = True ''''''''''''''''''''''''''''Return when finished
'End Function
'Public Function MouseEvent(Optional ByVal mEvents As Long = 0, Optional ByVal x As Long = 0, Optional ByVal y As Long = 0) As Boolean
'    On Error Resume Next
'    MouseEvent = apimouse_event(mEvents, x, y, 0, apiGetMessageExtraInfo) 'Return results
'End Function
'Private Function MoveItemToItem(ByVal wName As String, ByVal wName2 As String, ByVal tMost As Long, ByRef nPos As Long, ByRef rOffset As Long, ByRef tOffset As Long, ByRef m As MENUINFO, ByRef MI As ITEMINFO) As Boolean
'    On Error Resume Next
'    Dim p As POINTAPI
'    tOffset = tOffset + MI.Top - tMost '''''''''''Keep offset from top most
'    rOffset = rOffset + MI.Width '''''''''''''''''Keep offset from left most
'    If MouseEvent(MOUSEEVENTF_LEFTDOWN + MOUSEEVENTF_LEFTUP) = False Then MoveItemToItem = False: Exit Function 'click sub menu, and return false if failure
'    m.hSubMenu = apiGetSubMenu(m.hSubMenu, nPos) 'Get handle of submenu
'    If wName = "" Then '''''''''''''''''''''''''''If name not specified
'        nPos = CInt(wName2) - 1 ''''''''''''''''''Then the search is by index
'    Else '''''''''''''''''''''''''''''''''''''''''Then the search is by name
'        nPos = FindMenuItemPos(m.hSubMenu, wName, CInt(wName2)) 'Get position from handle and name
'    End If
'    If nPos = -1 Then MoveItemToItem = False: Exit Function '''''''''''''''If return is negative then exit and return false
'    MI = MenuItemDim(m.hwnd, m.hSubMenu, 0) ''''''Get menu item dimensions
'    If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit if there is a negative return
'    p = ToScreen(rOffset + MI.Center.x, tOffset + MI.Center.y) 'Covert to screen coordinates
'    If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then MoveItemToItem = False: Exit Function 'Move to new screen location
'    MI = MenuItemDim(m.hwnd, m.hSubMenu, nPos) ''Get menu item dimensions
'    If MI.Top = -1 And MI.Bottom = -1 And MI.Left = -1 And MI.Right = -1 Then MoveItemToItem = False: Exit Function 'Exit upon negative result
'    p = ToScreen(rOffset + MI.Center.x, tOffset + MI.Center.y) 'Convert to screen
'    If MouseEvent(MOUSEEVENTF_ABSOLUTE + MOUSEEVENTF_MOVE, p.x, p.y) = False Then MoveItemToItem = False: Exit Function 'Move to point
'    MoveItemToItem = True
'End Function
'Friend Function ToScreen(ByVal x As Long, ByVal y As Long) As POINTAPI
'    On Error Resume Next
'    ToScreen.x = CInt(x * SM_FULLSCREEN / apiGetSystemMetrics(SM_CXSCREEN)) 'Set the return value for x.
'    ToScreen.y = CInt(y * SM_FULLSCREEN / apiGetSystemMetrics(SM_CYSCREEN)) 'Set the return value for y.
'End Function
'
'
