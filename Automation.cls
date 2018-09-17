VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Automation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'ThisWorkbook.VBProject.References.AddFromFile "C:\Windows\SysWOW64\UIAutomationCore.dll"
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type tagPOINT
    x As Long
    y As Long
End Type
Private Declare Function apiGetCursorPos Lib "user32" Alias "GetCursorPos" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function apiElementFromPoint Lib "UIAutomationCore" Alias "ElementFromPoint" (ByRef tp As tagPOINT) As IUIAutomationElement
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
'' Waits for a window to exist, and waits for it to be ready for keyboard and mouse input.
'' Returns true if a window exists. Returns false if it does not, or the specified time ends.
'' wTitle The title or class name of the window.  Specify an application title or class name.
'' dwMilliseconds The number of milliseconds to wait.  Specify an integer.
'Friend Function WaitForWindow(ByVal wTitle As String, Optional ByVal dwMilliseconds As Long = 5000) As Boolean
'    On Error Resume Next
'    Dim tick As Long
'    Dim hwnd As Long
'    tick = apiGetTickCount '''''''''''''''''''''''Get the number of millisecons since last log on
'    hwnd = 0 '''''''''''''''''''''''''''''''''''''Initialize handle
'    Do '''''''''''''''''''''''''''''''''''''''''''Loop this thread until it's time
'        hwnd = apiFindWindow(vbNullString, wTitle) 'Find window by title
'        If hwnd = 0 Then hwnd = apiFindWindow(wTitle, vbNullString) 'If not found by title then find by class name
'        If hwnd <> 0 Then WaitForWindow = WaitForWindowIdle(hwnd): Exit Do  'If found, then wait for input idle and return
'        If apiGetTickCount >= (tick + dwMilliseconds) Then WaitForWindow = False: Exit Do  'If time is up then return that value
'        apiSleep (1) '''''''''''''''''''''''''''''Sleep thread for one millisecond
'        Flush ''''''''''''''''''''''''''''''''''''Process messages in the queue
'    Loop
'End Function
'Private Function ProcessIdToHandle(ByVal pid As Long) As Long
'    On Error Resume Next
'    Dim hwnd      As Long
'    Dim ProcessID As Long
'    hwnd = apiGetTopWindow(0) ''''''''''''''''''''Get the top window in the z-order
'    Do
'        If apiGetParent(hwnd) = 0 Then '''''''''''If window has no parents
'            Call apiGetWindowThreadProcessId(hwnd, ProcessID)  'Get the window's process id
'            If ProcessID = pid Then ProcessIdToHandle = hwnd: Exit Do 'If pid matches then return
'        End If
'        hwnd = apiGetWindow(hwnd, GW_HWNDNEXT) '''Retrieve the next window
'        If hwnd = 0 Then Exit Do '''''''''''''''''Exit if there are no more windows
'    Loop
'End Function
'Private Function ProcessNameToId(ByVal ProcessName As String, Optional ByVal pid As Long = 0) As Long
'    On Error Resume Next
'    Dim pLength            As Long
'    Dim cbSizeReturned     As Long
'    Dim numElements        As Long
'    Dim pIDs()             As Long
'    Dim cbSize             As Long
'    Dim cbSize2            As Long
'    Dim ret                As Long
'    Dim pSize              As Long
'    Dim hProcess           As Long
'    Dim pLoop              As Long
'    Dim pname              As String
'    Dim mName              As String
'    Dim prName             As String
'    Dim pModules(1 To 200) As Long
'    ProcessName = UCase(Trim(ProcessName))
'    pLength = Len(ProcessName)
'    cbSize = 8
'    cbSizeReturned = 96
'    Do While cbSize <= cbSizeReturned
'        cbSize = cbSize * 2 ''''''''''''''''''Increment Size
'        ReDim pIDs(cbSize / 4) As Long '''''''Allocate Memory for Array
'        ret = apiEnumProcesses(pIDs(1), cbSize, cbSizeReturned) 'Get Process ID's
'    Loop
'    numElements = cbSizeReturned / 4  ''''''''Count number of processes returned
'    For pLoop = 1 To numElements '''''''''''''Loop thru each process
'        hProcess = apiOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0, pIDs(pLoop))  'Get a handle to the Process and Open
'        If hProcess <> 0 Then
'            ret = apiEnumProcessModules(hProcess, pModules(1), 200, cbSize2) 'Get an array of the module handles for the specified process
'            If ret <> 0 Then '''''''''''''''''If the Module Array is retrieved, Get the ModuleFileName
'                mName = Space(MAX_PATH) ''''''Buffer with spaces first to allocate memory for byte array
'                pSize = 500  '''''''''''''''''Must be set prior to calling API
'                ret = apiGetModuleFileNameExA(hProcess, pModules(1), mName, pSize) 'Get Process Name
'                pname = Left(mName, ret) '''''Remove trailing spaces
'                pname = UCase(Trim(pname)) '''Check for Matching Upper case result
'                prName = GetElement(Trim(Replace(pname, Chr(0), "")), "\", 0, 0, GetNumElements(Trim(Replace(pname, Chr(0), "")), "\") - 1)
'                If pid = 0 Then ''''''''''''''If getting process id
'                    If prName = ProcessName Then 'If process name matches specification
'                        ProcessNameToId = pIDs(pLoop) 'Set pid return
'                        Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
'                    End If
'                Else '''''''''''''''''''''''''If getting process handle from id
'                    If pid = pIDs(pLoop) Then 'If process id matches specification
'                        ProcessNameToId = hProcess 'Set process handle return
'                        Call apiWaitForInputIdle(hProcess, 5000) 'Wait for input idle
'                    End If
'                End If
'            End If
'        End If
'        ret = apiCloseHandle(hProcess) '''''''Close the handle to this process
'    Next
'End Function
'Private Function WaitForWindowIdle(ByVal hwnd As Long) As Boolean
'    On Error Resume Next
'    Dim pid As Long
'    pid = HandleToProcessId(hwnd) ''''''''''''''''Get process id from handle
'    If pid <> 0 Then WaitForWindowIdle = CBool(ProcessNameToId("", pid)) 'If process id found then set return value of specified pId, which is a process handle
'End Function
'Private Function HandleToProcessId(ByVal hwnd As Long) As Long
'    On Error Resume Next
'    Dim wnd As Long
'    Dim pid As Long
'    If apiGetParent(hwnd) <> 0 Then hwnd = apiGetAncestor(hwnd, GA_ROOT)
'    wnd = apiGetTopWindow(0) '''''''''''''''''''''Get the top window in the z-order
'    Do
'        If wnd = 0 Then Exit Do
'        If wnd = hwnd Then
'            Call apiGetWindowThreadProcessId(wnd, pid)  'Get the window's process id
'            HandleToProcessId = pid: Exit Do '''''Set process id and return
'        End If
'        wnd = apiGetWindow(wnd, GW_HWNDNEXT) '''''Retrieve the next window
'    Loop
'End Function
'Private Function AttachInput(ByVal hwnd As Long, Optional ByVal bAttach As Boolean = True) As Boolean
'    On Error Resume Next
'    AttachInput = CBool(apiAttachThreadInput(apiGetWindowThreadProcessId(hwnd, 0), apiGetCurrentThreadId, CInt(bAttach)))
'End Function
'Private Function ForceForeground(ByVal hwnd As Long) As Boolean
'    On Error Resume Next
'    apiLockSetForegroundWindow (LSFW_UNLOCK) '''''Unlock setforegroundwindow calls
'    apiAllowSetForegroundWindow (ASFW_ANY) '''''''Allow setforeground window calls
'    Call KeyEvent(VK_Menu, False, True) ''''''''''Lift menu key if pressed, and it also allows the foreground window to be set
'    ForceForeground = CBool(apiSetForegroundWindow(hwnd)) 'Set foreground window
'    apiLockSetForegroundWindow (LSFW_LOCK) '''''''Lock other apps from using setforegroundwindow
'End Function
'Friend Function GetChildWindow(ByVal hwnd As Long, Optional ByVal wName As String = "", Optional ByVal wIndex As Long = 1) As Long
'    On Error Resume Next
'    Dim cwnd       As Long
'    Dim TextCount  As Long
'    Dim ClassCount As Long
'    Dim NullCount  As Long
'    Dim w          As WINNAME
'    Do '''''''''''''''''''''''''''''''''''''''''''Loop through sibling windows
'        cwnd = apiFindWindowEx(hwnd, cwnd, vbNullString, vbNullString) 'Set child handle
'        If cwnd = 0 Then GetChildWindow = cwnd: Exit Do  'If no more sibling children then return
'        w = GetWinName(cwnd, True, True) '''''''''Get the text, and class from that window
'        If w.lpText = wName Then TextCount = TextCount + 1 'If text matches the specified wName then increment the text count
'        If TextCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If text count is equal to the specified wIndex then return
'        If w.lpClass = wName Then ClassCount = ClassCount + 1 'If class name matches the specified wName then increment the class count
'        If ClassCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If class count is equal to the specified wIndex then return
'        If w.lpText = "" And w.lpClass = "" And wName = "" Then NullCount = NullCount + 1 'If text and class name not specified, then increment by index only
'        If NullCount = wIndex Then GetChildWindow = cwnd: Exit Do 'If null count is equal to wIndex then return
'    Loop
'End Function
'Friend Function GetChildWindow2(ByVal hwnd As Long, Optional ByVal wName As String = "", Optional ByVal wIndex As Long = 1) As Long
'    On Error Resume Next
'    Dim cwnd       As Long
'    Dim ClassCount As Long
'    Dim NullCount  As Long
'    Dim w          As WINNAME
'    Do '''''''''''''''''''''''''''''''''''''''''''Loop through sibling windows
'        cwnd = apiFindWindowEx(hwnd, cwnd, vbNullString, vbNullString) 'Set child handle
'        If cwnd = 0 Then GetChildWindow2 = cwnd: Exit Do  'If no more sibling children then return
'        w = GetWinName(cwnd, False, True) '''''''''Get the text, and class from that window
'        If InStr(0, w.lpClass, wName) <> 0 Then
'            GetChildWindow2 = cwnd
'            Exit Do
'        End If
'    Loop
'End Function
'Private Function GetElement(ByVal strList As String, ByVal strDelimiter As String, ByVal lngNumColumns As Long, ByVal lngRow As Long, ByVal lngColumn As Long) As String
'    On Error Resume Next
'    Dim lngCounter As Long
'    strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
'    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn) ' Calculate the offset for the item required based on the number of columns the list strList' has i.e. 'lngNumColumns' and from which row the element is to be  selected i.e. 'lngRow'.
'    For lngCounter = 0 To lngColumn - 1 ''''''''''Search for the 'lngColumn' item from the list 'strList'.
'        strList = Mid(strList, InStr(strList, strDelimiter) + Len(strDelimiter), Len(strList))  ' Remove each item from the list.
'        If Len(strList) = 0 Then GetElement = "": Exit Function 'If list becomes empty before 'lngColumn' is found then just return an empty string.
'    Next lngCounter
'    GetElement = Left(strList, InStr(strList, strDelimiter) - 1) 'Return the sought list element.
'End Function
'Private Function GetNumElements(ByVal strList As String, ByVal strDelimiter As String) As Integer
'    On Error Resume Next
'    Dim intElementCount As Integer
'    If Len(strList) = 0 Then GetNumElements = 0: Exit Function 'If no elements in the list 'strList' then just return 0.
'    strList = strList & strDelimiter '''''''''''''Append delimiter text to the end of the list as a terminator.
'    While InStr(strList, strDelimiter) > 0 '''''''Count the number of elements in 'strlist'
'        intElementCount = intElementCount + 1
'        strList = Mid(strList, InStr(strList, strDelimiter) + 1, Len(strList))
'    Wend
'    GetNumElements = intElementCount '''''''''''''Return the number of elements in 'strList'.
'End Function
'Private Function GetWinAncestory(ByVal cwnd As Long, ByVal gName As Boolean) As Long()
'    On Error Resume Next
'    Dim i       As Long
'    Dim z       As Long
'    Dim cWnd2   As Long
'    Dim hwnds() As Long
'    Dim w       As WINNAME
'    cWnd2 = cwnd '''''''''''''''''''''''''''''''''Remember handle for seconds loop
'    w = GetWinName(cwnd, True, True) '''''''''''''Get text and class name of the window
'    If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Focus window:  Title | Class name") 'Display to developer if specified
'    Do '''''''''''''''''''''''''''''''''''''''''''Loop while counting ancestors
'        cwnd = apiGetParent(cwnd) ''''''''''''''''Get new parent handle if any found
'        If cwnd = 0 Then Exit Do '''''''''''''''''If there are no more parents then abort without setting anymore to the array
'        i = i + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
'    Loop
'    ReDim hwnds(i) '''''''''''''''''''''''''''''''Re-Dimension array to number of ancestors including the handle specified
'    Do '''''''''''''''''''''''''''''''''''''''''''Do look for parents of the window specified.
'        hwnds(z) = cWnd2 '''''''''''''''''''''''''Set the indexed array element to the specified handle or it's ancestors
'        cWnd2 = apiGetParent(cWnd2) ''''''''''''''Get new parent handle if any found
'        If cWnd2 = 0 Then Exit Do ''''''''''''''''If there are no more parents then abort
'        z = z + 1 ''''''''''''''''''''''''''''''''Increment the index by 1
'        w = GetWinName(cWnd2, True, True) ''''''''Get text and class name of the window
'        If gName = True Then Call MsgBox(CStr(z) & ":  " & w.lpText & "  |  " & w.lpClass, vbInformation, "Parent window:  Title | Class name")  'Display to developer if specified
'    Loop
'    GetWinAncestory = hwnds ''''''''''''''''''''''Return array of integer handles, or handle.
'End Function
'Friend Function GetWinName(ByVal hwnd As Long, Optional ByVal gText As Boolean = True, Optional ByVal gClass As Boolean = True) As WINNAME
'    On Error Resume Next
'    Dim tLength As Long
'    Dim rValue  As Long
'    GetWinName.lpText = "" '''''''''''''''''''''''Initialize string for text name
'    GetWinName.lpClass = "" ''''''''''''''''''''''Initialize string for class name
'    If gText = True Then '''''''''''''''''''''''''If text is to be retrieved
'        tLength = apiGetWindowTextLength(hwnd) + 4 'Get length
'        GetWinName.lpText = Strings.Space(tLength) 'Pad with buffer
'        rValue = apiGetWindowText(hwnd, GetWinName.lpText, tLength) 'Get text
'        GetWinName.lpText = Left(GetWinName.lpText, rValue) 'Strip buffer
'    End If
'    If gClass = True Then ''''''''''''''''''''''''If class name is to be retrieved
'        GetWinName.lpClass = Strings.Space(260) ''Pad with buffer
'        rValue = apiGetClassName(hwnd, GetWinName.lpClass, 260) 'Get classname
'        GetWinName.lpClass = Left(GetWinName.lpClass, rValue) 'Strip buffer
'    End If
'End Function
' Gets the handle of the current foreground window and the handle of the child window with keyboard focus.
' If keyFocus is False then the handles retrieved are the windows under the cursor with mouse focus.
' Returns the WINFOCUS structure containing the current focus handles. Depends on the keyFocus parameter.
' "showNames"(Optional) Shows the title and class name of the windows,
' to be used with the GetWinHandles function.
' "keyFocus"(Optional) Returns the handle of the main window, and child window with keyboard focus.
' If false then the function returns the handles of the windows currently under the cursor.</param>
'Friend Function GetWinFocus(Optional ByVal showNames As Boolean = False, Optional ByVal keyFocus As Boolean = True) As WINFOCUS
'    On Error Resume Next
'    If keyFocus = True Then ''''''''''''''''''''''Keyboard focus
'        Dim g As GUITHREADINFO '''''''''''''''''''Dimension a thread input structure
'        g.cbSize = Len(g) ''''''''''''''''''''''''Initialize structure
'        Call apiGetGUIThreadInfo(apiGetWindowThreadProcessId(0, 0), g) 'Retrieve information about the active window
'        GetWinFocus.Foreground = g.hWndActive ''''Set handle of active foreground
'        GetWinFocus.Focus = g.hWndFocus ''''''''''Set handle of focus window
'    Else '''''''''''''''''''''''''''''''''''''''''Mouse focus instead
'        Dim p As POINTAPI  '''''''''''''''''''''''Dimension a point for the mouse position
'        Call apiGetCursorPos(p)  '''''''''''''''''Get current cursor position
'        GetWinFocus.Focus = apiWindowFromPoint(p.x, p.y) 'Get handle of window under cursor
'        GetWinFocus.Foreground = apiGetAncestor(GetWinFocus.Focus, GA_ROOT) 'Try to get it's ancestor
'        If GetWinFocus.Foreground = 0 Then GetWinFocus.Foreground = GetWinFocus.Focus 'If no ancestor then set main focus to child focus
'    End If
'    If showNames = True Then Call GetWinAncestory(GetWinFocus.Focus, True) 'Set ancestory using focus window
'End Function

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
'Friend Function GetWinHandles(ByVal wName1 As String, Optional ByVal wIndex1 As Long = 1, Optional ByVal wName2 As String = " ", Optional ByVal wIndex2 As Long = 1, Optional ByVal wName3 As String = " ", Optional ByVal wIndex3 As Long = 1, Optional ByVal wName4 As String = " ", Optional ByVal wIndex4 As Long = 1, Optional ByVal wName5 As String = " ", Optional ByVal wIndex5 As Long = 1, Optional ByVal wName6 As String = " ", Optional ByVal wIndex6 As Long = 1, Optional ByVal wName7 As String = " ", Optional ByVal wIndex7 As Long = 1, Optional ByVal wName8 As String = " ", Optional ByVal wIndex8 As Long = 1, Optional ByVal wName9 As String = " ", Optional ByVal wIndex9 As Long = 1, Optional ByVal wName10 As String = " ", Optional ByVal wIndex10 As Long = 1) As WINFOCUS
'    On Error Resume Next
'    Dim hwnd     As Long
'    Dim cwnd     As Long
'    Dim i        As Long
'    Dim wn       As String
'    Dim wf       As WINFOCUS
'    Dim wNames() As String '''''''''''''''''''''''Dimension an array string the size of max index
'    If wName1 <> " " Then i = 1: ReDim wNames(1) 'Set i to index of parameters
'    If wName2 <> " " Then i = 3: ReDim wNames(3)
'    If wName3 <> " " Then i = 5: ReDim wNames(5)
'    If wName4 <> " " Then i = 7: ReDim wNames(7)
'    If wName5 <> " " Then i = 9: ReDim wNames(9)
'    If wName6 <> " " Then i = 11: ReDim wNames(11)
'    If wName7 <> " " Then i = 13: ReDim wNames(13)
'    If wName8 <> " " Then i = 15: ReDim wNames(15)
'    If wName9 <> " " Then i = 17: ReDim wNames(17)
'    If wName10 <> " " Then i = 19: ReDim wNames(19)
'    If wName1 <> " " Then wNames(0) = wName1: wNames(1) = CStr(wIndex1) 'Set array elements
'    If wName2 <> " " Then wNames(2) = wName2: wNames(3) = CStr(wIndex2)
'    If wName3 <> " " Then wNames(4) = wName3: wNames(5) = CStr(wIndex3)
'    If wName4 <> " " Then wNames(6) = wName4: wNames(7) = CStr(wIndex4)
'    If wName5 <> " " Then wNames(8) = wName5: wNames(9) = CStr(wIndex5)
'    If wName6 <> " " Then wNames(10) = wName6: wNames(11) = CStr(wIndex6)
'    If wName7 <> " " Then wNames(12) = wName7: wNames(13) = CStr(wIndex7)
'    If wName8 <> " " Then wNames(14) = wName8: wNames(15) = CStr(wIndex8)
'    If wName9 <> " " Then wNames(16) = wName9: wNames(17) = CStr(wIndex9)
'    If wName10 <> " " Then wNames(18) = wName10: wNames(19) = CStr(wIndex10)
'    hwnd = apiFindWindow(vbNullString, wName1) '''Look for handle from title
'    If hwnd = 0 Then hwnd = apiFindWindow(wName1, vbNullString) 'If not found by title, then look for handle from class name
'    If hwnd <> 0 And CInt(wNames(1)) > 1 Then  '''If searching for window by index, ie more than 1
'        Dim nxtwnd As Long
'        nxtwnd = hwnd ''''''''''''''''''''''''''''Initialize handle containing the next window in the top level z-order
'        Dim Index As Long
'        Index = 1 ''''''''''''''''''''''''''''''''Initialize index to 1, since handle was found.
'        Dim n As WINNAME  ''''''''''''''''''''''''Create structure for window names
'        Do '''''''''''''''''''''''''''''''''''''''Do loop to look for handles matching paramaters
'            nxtwnd = apiGetWindow(hwnd, GW_HWNDNEXT) 'Set next window
'            If nxtwnd = 0 Then Exit Do '''''''''''Eject if there are no more top level windows
'            If wNames(0) <> "" Then ''''''''''''''If title or class name was given
'                n = GetWinName(nxtwnd, True, True) 'Get title and classname of the window
'                If n.lpText = wNames(0) Or n.lpClass = wNames(0) Then 'If title or class name matches first wNames parameter
'                    Index = Index + 1 ''''''''''''Increment Index for the matching window
'                    hwnd = nxtwnd ''''''''''''''''Set handle to the handle of the matching window
'                    If Index >= CInt(wNames(1)) Then Exit Do 'If the specified index has been reached then exit with last matching handle
'                End If
'            Else '''''''''''''''''''''''''''''''''Then "" indicates the search is by index only
'                Index = Index + 1 ''''''''''''''''Increment Index for the matching window
'                hwnd = nxtwnd ''''''''''''''''''''Set handle to the handle of the matching window
'                If Index >= CInt(wNames(1)) Then Exit Do 'If index specified has been reached then exit with last matching handle
'            End If
'        Loop
'    End If
'    If hwnd = 0 Then '''''''''''''''''''''''''''''If not found by title or class name, then look for special commands
'        Dim pid As Long ''''''''''''''''''''''''''Dimension process identification
'        Dim p   As POINTAPI  '''''''''''''''''''''''Dimension a point structure
'        wn = wName1 ''''''''''''''''''''''''''''''Set string to name for lower case conversion
'        Call apiCharLower(wn)  '''''''''''''''''''Does not need to be case sensitive
'        If wn = "{focus}" Then '''''''''''''''''''If focus window is specified
'            hwnd = GetWinFocus.Focus '''''''''''''Get the handle of the focus window
'        ElseIf wn = "{foreground}" Then ''''''''''If foreground window is specified
'            hwnd = apiGetForegroundWindow ''''''''Get the handle of the foreground window
'        ElseIf wn = "{active}" Then ''''''''''''''If active window is specified
'            hwnd = apiGetActiveWindow ''''''''''''Get the handle of the active window
'        ElseIf wn = "{desktop}" Then '''''''''''''If desktop window is specified
'            hwnd = apiGetDesktopWindow() '''''''''Get the handle of the desktop window
'        ElseIf wn = "{top}" Then '''''''''''''''''If top window is specified
'            hwnd = apiGetTopWindow(hwnd) '''''''''Get the handle of the top window
'        ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
'            hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
'        ElseIf wn = "{parent}" Then ''''''''''''''If parent window is specified
'            hwnd = apiGetParent(wIndex1) '''''''''Get the handle of the parent window
'        ElseIf wn = "{ancestor}" Then ''''''''''''If ancestor window is specified
'            hwnd = apiGetAncestor(wIndex1, GA_ROOT) 'Get the handle of the ancestor window
'        ElseIf wn = "{frompoint}" Then '''''''''''If window under the cursor is specified
'            If wIndex1 <> 1 And wIndex2 <> 1 Then 'If point specified
'                p.x = wIndex1 ''''''''''''''''''''If x specified
'                p.y = wIndex2 ''''''''''''''''''''If y specified
'            Else
'                Call apiGetCursorPos(p)  ''''''''''Get cursor position as POINTAPI
'            End If
'            hwnd = apiWindowFromPoint(p.x, p.y) ''Get the handle of the window under the cursor
'        ElseIf wn = "{childfrompoint}" Then ''''''If child window under the cursor is specified
'            If wIndex2 <> 1 And wIndex3 <> 1 Then 'If point specified
'                p.x = wIndex2 ''''''''''''''''''''If x specified
'                p.y = wIndex3 ''''''''''''''''''''If y specified
'            Else
'                Call apiGetCursorPos(p)  '''''''''Get cursor position as POINTAPI
'            End If
'            hwnd = apiChildWindowFromPointEx(wIndex1, p.x, p.y, wIndex2) 'Get the handle of the child window under the cursor
'        ElseIf InStr(1, wName1, ".exe") <> 0 Then 'If process name specified
'            pid = ProcessNameToId(wName1) ''''''''Get process Id and wait for input idle
'            If pid <> 0 Then hwnd = ProcessIdToHandle(pid) 'If process id found then set handle
'        End If
'    End If
'    If hwnd = 0 Then '''''''''''''''''''''''''''''If no handle found
'        pid = Shell(wName1, wIndex1) '''''''''''''Shell(start) the process
'        If pid <> 0 Then '''''''''''''''''''''''''If process id found
'            hwnd = ProcessIdToHandle(pid) ''''''''Get handle from process identification
'            Call ProcessNameToId("", pid) ''''''''Wait for process id to be idle by specifying it in the second parameter
'        End If
'    End If
'    If hwnd = 0 Then hwnd = -1 '''''''''''''''''''If main window not found, then set failure return
'    wf.Foreground = hwnd '''''''''''''''''''''''''Set structure return
'    wf.Focus = hwnd ''''''''''''''''''''''''''''''Set focus to main handle for now
'    If hwnd = -1 Then GetWinHandles = wf: Exit Function ''''''''''''''''''If no handle found then return now
'    If i = 1 Then ''''''''''''''''''''''''''''''''If only the main window was specified
'        Dim prvswnd As Long
'        prvswnd = apiGetForegroundWindow '''''''''Remember the current foreground window
'        ForceForeground (hwnd) '''''''''''''''''''Force foreground onto specified main window
'        Sleep (0) ''''''''''''''''''''''''''''''''Sleep a moment
'        wf = GetWinFocus() '''''''''''''''''''''''Set structure
'        ForceForeground (prvswnd) ''''''''''''''''Force foreground back to where it was
'    ElseIf i > 1 Then ''''''''''''''''''''''''''''If there are child windows specified
'        cwnd = GetChildWindow(hwnd, wNames(2), CInt(wNames(3))) 'Set first child handle
'        wf.Focus = cwnd ''''''''''''''''''''''''''Set focus handle to first child
'        If (i - 1) > 1 Then ''''''''''''''''''''''If more than one child specified
'            Dim q As Long
'            For q = 4 To i - 1 Step 2  '''''''''''Step through array looking for grandchildren.
'                cwnd = GetChildWindow(cwnd, wNames(q), CInt(wNames(q + 1))) 'Set new child window
'                wf.Focus = cwnd ''''''''''''''''''Set focus handle to youngest grandchild so far
'            Next '''''''''''''''''''''''''''''''''Next in array
'        End If
'    End If
'    GetWinHandles = wf ''''''''''''''''''''''''''''''''''''Return WINFOCUS structure
'End Function
