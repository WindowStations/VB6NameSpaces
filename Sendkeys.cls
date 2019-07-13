VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Sendkeys"
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
'Attribute VB_Name = "Sendkeys"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Const KEYEVENTF_KEYDOWN As Long = 0
Const KEYEVENTF_KEYUP As Long = 2
Const NEGATIVE As Long = -1
Private Declare Function apiCharLower Lib "user32" Alias "CharLowerA" (ByVal lpsz As String) As Long
Private Declare Function apiGetInputState Lib "user32" Alias "GetInputState" () As Long
Private Declare Function apiGetMessageExtraInfo Lib "user32" Alias "GetMessageExtraInfo" () As Long
Private Declare Function apikeybd_event Lib "user32" Alias "keybd_event" (ByVal vKey As Long, ByVal bScan As Long, ByVal dwFlags As Long, ByVal dwExtraInfo As Long) As Boolean
Private Declare Function apiVkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar2 As Long) As Long
Friend Function Send(ByVal cText As String) As Long
    On Error Resume Next
    Dim UserCapsLock   As Boolean
    Dim TextLength     As Long
    Dim vk             As Keys
    Dim TextRemaining  As String
    Dim CurrentLetter  As String
    Dim PreviousLetter As String '''''''''''''''''''''Dimension some strings
    Dim LowerCase      As String
    Dim i              As Long
    i = 1
    TextRemaining = cText ''''''''''''''''''''''''''''Initialize remaining text
    TextLength = Len(TextRemaining) '''''''''''''''''''''''Initialize Text
    CurrentLetter = "" ''''''''''''''''''''''''''''''''''Initialize the first letter to nothing
    PreviousLetter = ""
    LowerCase = ""
    vk = Keys.None_
    Dim ret As Long
    Dim b As Boolean
    Do
        PreviousLetter = CurrentLetter ''''''''''''''''''''''Set previous letter to the current before setting the current
        CurrentLetter = Left(TextRemaining, 1) ''''''''''''''Set current letter to the left most
        LowerCase = CurrentLetter '''''''''''''''''''''''Initialize lower case letter
        TextRemaining = Mid(TextRemaining, 2) ''''''''''''Cut string to get remaining text
        vk = apiVkKeyScan(Asc(CurrentLetter)) And 255 '''Set keyscan of that letter
        Call apiCharLower(LowerCase)  '''''''''''''''''Convert in place to lowercase
        If CurrentLetter <> LowerCase Or CurrentLetter = "!" Or CurrentLetter = "@" Or CurrentLetter = "&" Or CurrentLetter = "*" Or CurrentLetter = "_" Or CurrentLetter = "|" Or CurrentLetter = ":" Or CurrentLetter = "<" Or CurrentLetter = ">" Or CurrentLetter = "?" Then
            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYDOWN, -11)
        End If
        ret = apikeybd_event(vk, 0, KEYEVENTF_KEYDOWN, -11)
        Call apikeybd_event(vk, 0, KEYEVENTF_KEYUP, -11)
        If CurrentLetter <> LowerCase Or CurrentLetter = "!" Or CurrentLetter = "@" Or CurrentLetter = "&" Or CurrentLetter = "*" Or CurrentLetter = "_" Or CurrentLetter = "|" Or CurrentLetter = ":" Or CurrentLetter = "<" Or CurrentLetter = ">" Or CurrentLetter = "?" Then
            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
        End If
        If ret <> 0 Then kSent = kSent + 1 '''''''''''Count the sent keys send down
        If i >= TextLength Then Exit Do
        i = i + 1
    Loop
    Send = kSent
End Function
Friend Sub Flush()
    On Error Resume Next
    If apiGetInputState <> 0 Then DoEvents
End Sub
Friend Function SendKey(ByVal vk As Keys, Optional ByVal down As Boolean = True, Optional ByVal up As Boolean = True) As Boolean
    On Error Resume Next
    If down = True Then SendKey = CBool(apikeybd_event(vk, 0, KEYEVENTF_KEYDOWN, -11))
    If up = True Then SendKey = CBool(apikeybd_event(vk, 0, KEYEVENTF_KEYUP, -11))
End Function
'        If CurrentLetter <> LowerCase Then
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYDOWN, -11)
'        ElseIf CurrentLetter = "!" Or CurrentLetter = "@" Or CurrentLetter = "&" Or CurrentLetter = "*" Or CurrentLetter = "_" Or CurrentLetter = "|" Or CurrentLetter = ":" Or CurrentLetter = "<" Or CurrentLetter = ">" Or CurrentLetter = "?" Then
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYDOWN, -11) 'Or CurrentLetter = " "
'        ElseIf CurrentLetter = "~" Then '''''''''''''''''If letter is tilde
'            vk = 13 ''''''''''''''''''''''''''''''Set to return(enter) key
'        ElseIf CurrentLetter = "+" Then '''''''''''''''''If letter is plus
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYDOWN, -11)
'        ElseIf CurrentLetter = "^" Then '''''''''''''''''If letter is caret
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'            Call apikeybd_event(Keys.ControlKey_, 0, KEYEVENTF_KEYDOWN, -11)
'        ElseIf CurrentLetter = "#" Then '''''''''''''''''If letter is number signifier
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'            Call apikeybd_event(Keys.LWin_, 0, KEYEVENTF_KEYDOWN, -11)
'        ElseIf CurrentLetter = "%" Then '''''''''''''''''If letter is percent
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'            Call apikeybd_event(Keys.Menu_, 0, KEYEVENTF_KEYDOWN, -11)
'        ElseIf CurrentLetter = "(" Then '''''''''''''''''If letter is left parenthesis
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'        ElseIf CurrentLetter = ")" Then '''''''''''''''''If letter is right parenthesis
'            vk = 0 '''''''''''''''''''''''''''''''Do not send a regular key
'            Call Lift ''''''''''''''''''''''''''''Lift extented keys
'        ElseIf CurrentLetter = "{" Then '''''''''''''''''''''If letter is a command bracket
'            vk = 0 '''''''''''''''''''''''''''''''Set virtual key to zero(none)
'            cmdKey = Mid(TextRemaining, 1, InStr(1, TextRemaining, "}") - 1) 'Get text within command brackets
'            frmNameSpacesVB6.txtOutput.Text = frmNameSpacesVB6.txtOutput.Text & " " & cmdKey & vbCrLf
'            '
'            If Len(TextRemaining) - (Len(cmdKey) + 1) > 0 Then
'                TextRemaining = Mid(TextRemaining, InStr(1, TextRemaining, "}") + 1) ' 'Set remaining text, as the text(if any) to the right of the command bracket
'            End If
'            frmNameSpacesVB6.txtOutput.Text = frmNameSpacesVB6.txtOutput.Text & " " & TextRemaining & vbCrLf
'            '
'            i2 = i2 + Len(cmdKey) + 1  ''''''''''''''Increment i the number of characters within command brackets since they are not going to be processed individually
'            '
'            If InStr(1, cmdKey, " ") <> 0 And InStr(1, cmdKey, ", ") = 0 Then 'If there is a space and no comma preceeding it then the command is to be repeated
'                If IsNumeric(Mid(cmdKey, InStr(1, cmdKey, " ") + 1)) Then
'                    repeat = CInt(Mid(cmdKey, InStr(1, cmdKey, " ") + 1)) 'If a number can be identified, then set it otherwise it stays 1._+{}|:"<>?
'                End If
'                cmdKey = Left(cmdKey, InStr(1, cmdKey, " ") - 1) 'Strip off left side, which is the actual command
'            End If
'            frmNameSpacesVB6.txtOutput.Text = frmNameSpacesVB6.txtOutput.Text & " " & cmdKey & vbCrLf
'            If Len(cmdKey) = 1 Then ''''''''''''''If command is a single character
'                If cmdKey = "#" Or cmdKey = "+" Or cmdKey = "^" Or cmdKey = "%" Or cmdKey = "~" Or cmdKey = "(" Or cmdKey = ")" Or cmdKey = "[" Or cmdKey = "]" Then
'                    Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYDOWN, -11)
'                End If
'                vk = apiVkKeyScan(Asc(cmdKey)) And 255 'Set keyscan for that letter
'            Else '''''''''''''''''''''''''''''''''If command is a key name
'                vk = NEGATIVE ''''''''''''''''''''Set to negative in case this is not a valid key command, then it's a simple string
'                apiCharLower (cmdKey)
'                vk = GetKeyCodeFromName(cmdKey)
'            End If
'            Call apikeybd_event(vk, 0, KEYEVENTF_KEYDOWN, -11)
'            Call apikeybd_event(vk, 0, KEYEVENTF_KEYUP, -11)
'            frmNameSpacesVB6.txtOutput.Text = frmNameSpacesVB6.txtOutput.Text & " " & vk & " " & repeat & vbCrLf
'        Else
'            '
'        End If
'        'Lift extended keys
'        If CurrentLetter <> LowerCase Then
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'        ElseIf CurrentLetter = "!" Or CurrentLetter = "@" Or CurrentLetter = " " Or CurrentLetter = "&" Or CurrentLetter = "*" Or CurrentLetter = "_" Or CurrentLetter = "|" Or CurrentLetter = ":" Or CurrentLetter = "<" Or CurrentLetter = ">" Or CurrentLetter = "?" Then
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'        ElseIf upp = 1 And (cmdKey = "#" Or cmdKey = "+" Or cmdKey = "^" Or cmdKey = "%" Or cmdKey = "~" Or cmdKey = "(" Or cmdKey = ")" Or cmdKey = "[" Or cmdKey = "]") Then
'            Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'        ElseIf CurrentLetter <> "(" Then ''''''''''''''''''''If character is not parenthesis
'            If PreviousLetter = "#" Then '''''''''''''If previous letter was numeric signifier
'                Call apikeybd_event(Keys.LWin_, 0, KEYEVENTF_KEYUP, -11)
'            ElseIf PreviousLetter = "+" Then '''''''''If previous letter was shift
'                Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'            ElseIf PreviousLetter = "%" Then '''''''''If previous letter was percent
'                Call apikeybd_event(Keys.Menu_, 0, KEYEVENTF_KEYUP, -11)
'            ElseIf PreviousLetter = "^" Then '''''''''If previous letter was caret
'                Call apikeybd_event(Keys.ControlKey_, 0, KEYEVENTF_KEYUP, -11)
'            End If
'        End If

'Private Function GetKeyCodeFromName(ByVal cmdKey As String) As Keys
'    Dim vk As Keys
'    If cmdKey = "none" Then ''''''''''Set virtual key to specification
'        vk = None_
'    ElseIf cmdKey = "lbutton" Then
'        vk = LButton_
'    ElseIf cmdKey = "rbutton" Then
'        vk = RButton_
'    ElseIf cmdKey = "cancel" Then
'        vk = Cancel_
'    ElseIf cmdKey = "mbutton" Then
'        vk = MButton_
'    ElseIf cmdKey = "xbutton1" Then
'        vk = XButton1_
'    ElseIf cmdKey = "xbutton2" Then
'        vk = XButton2_
'    ElseIf cmdKey = "lbutton, xbutton2" Then
'        vk = LButton_XButton2_
'    ElseIf cmdKey = "back" Or cmdKey = "backspace" Or cmdKey = "bs" Or cmdKey = "bksp" Then
'        vk = Back_
'    ElseIf cmdKey = "tab" Then
'        vk = Tab_
'    ElseIf cmdKey = "linefeed" Then
'        vk = LineFeed_
'    ElseIf cmdKey = "lbutton, linefeed" Then
'        vk = LButton_LineFeed_
'    ElseIf cmdKey = "clear" Then
'        vk = Clear_
'    ElseIf cmdKey = "return" Or cmdKey = "enter" Then
'        vk = Return_
'    ElseIf cmdKey = "rbutton, clear" Then
'        vk = RButton_Clear_
'    ElseIf cmdKey = "rbutton, return" Then
'        vk = RButton_Return_
'    ElseIf cmdKey = "shiftkey" Or cmdKey = "shift" Then
'        vk = ShiftKey_
'    ElseIf cmdKey = "controlkey" Or cmdKey = "control" Then
'        vk = ControlKey_
'    ElseIf cmdKey = "menu" Or cmdKey = "alt" Then
'        vk = Menu_
'    ElseIf cmdKey = "pause" Or cmdKey = "break" Then
'        vk = Pause_
'    ElseIf cmdKey = "capital" Or cmdKey = "capslock" Then
'        vk = Capital_
'    ElseIf cmdKey = "kanamode" Then
'        vk = KanaMode_
'    ElseIf cmdKey = "rbutton, capital" Then
'        vk = RButton_Capital_
'    ElseIf cmdKey = "junjamode" Then
'        vk = JunjaMode_
'    ElseIf cmdKey = "finalmode" Then
'        vk = FinalMode_
'    ElseIf cmdKey = "hanjamode" Then
'        vk = HanjaMode_
'    ElseIf cmdKey = "rbutton, finalmode" Then
'        vk = RButton_FinalMode_
'    ElseIf cmdKey = "escape" Or cmdKey = "esc" Then
'        vk = Escape_
'    ElseIf cmdKey = "imeconvert" Then
'        vk = IMEConvert_
'    ElseIf cmdKey = "imenonconvert" Then
'        vk = IMENonconvert_
'    ElseIf cmdKey = "imeaceept" Then
'        vk = IMEAceept_
'    ElseIf cmdKey = "imemodechange" Then
'        vk = IMEModeChange_
'    ElseIf cmdKey = "space" Then
'        vk = Space_
'    ElseIf cmdKey = "pageup" Or cmdKey = "pgup" Then
'        vk = PageUp_
'    ElseIf cmdKey = "next" Then
'        vk = Next_
'    ElseIf cmdKey = "end" Then
'        vk = End_
'    ElseIf cmdKey = "home" Then
'        vk = Home_
'    ElseIf cmdKey = "left" Then
'        vk = Left_
'    ElseIf cmdKey = "up" Then
'        vk = Up_
'    ElseIf cmdKey = "right" Then
'        vk = Right_
'    ElseIf cmdKey = "down" Then
'        vk = Down_
'    ElseIf cmdKey = "select" Then
'        vk = Select_
'    ElseIf cmdKey = "print" Then
'        vk = Print_
'    ElseIf cmdKey = "execute" Then
'        vk = Execute_
'    ElseIf cmdKey = "printscreen" Or cmdKey = "prtsc" Or cmdKey = "snapshot" Then
'        vk = PrintScreen_
'    ElseIf cmdKey = "Insert" Or cmdKey = "ins" Then
'        vk = Insert_
'    ElseIf cmdKey = "delete" Or cmdKey = "del" Then
'        vk = Delete_
'    ElseIf cmdKey = "help" Then
'        vk = Help_
'    ElseIf cmdKey = "d0" Then
'        vk = D0_
'    ElseIf cmdKey = "d1" Then
'        vk = D1_
'    ElseIf cmdKey = "d2" Then
'        vk = D2_
'    ElseIf cmdKey = "d3" Then
'        vk = D3_
'    ElseIf cmdKey = "d4" Then
'        vk = D4_
'    ElseIf cmdKey = "d5" Then
'        vk = D5_
'    ElseIf cmdKey = "d6" Then
'        vk = D6_
'    ElseIf cmdKey = "d7" Then
'        vk = D7_
'    ElseIf cmdKey = "d8" Then
'        vk = D8_
'    ElseIf cmdKey = "d9" Then
'        vk = D9_
'    ElseIf cmdKey = "rbutton, d8" Then
'        vk = RButton_D8_
'    ElseIf cmdKey = "rbutton, d9" Then
'        vk = RButton_D9_
'    ElseIf cmdKey = "mbutton, d8" Then
'        vk = MButton_D8_
'    ElseIf cmdKey = "mbutton, d9" Then
'        vk = MButton_D9_
'    ElseIf cmdKey = "xbutton2, d8" Then
'        vk = XButton2_D8_
'    ElseIf cmdKey = "xbutton2, d9" Then
'        vk = XButton2_D9_
'    ElseIf cmdKey = "64" Then
'        vk = VK_64_
'    ElseIf cmdKey = "a" Then
'        vk = a_
'    ElseIf cmdKey = "b" Then
'        vk = b_
'    ElseIf cmdKey = "c" Then
'        vk = c_
'    ElseIf cmdKey = "d" Then
'        vk = d_
'    ElseIf cmdKey = "e" Then
'        vk = e_
'    ElseIf cmdKey = "f" Then
'        vk = f_
'    ElseIf cmdKey = "g" Then
'        vk = g_
'    ElseIf cmdKey = "h" Then
'        vk = h_
'    ElseIf cmdKey = "i" Then
'        vk = i_
'    ElseIf cmdKey = "j" Then
'        vk = J_
'    ElseIf cmdKey = "k" Then
'        vk = k_
'    ElseIf cmdKey = "l" Then
'        vk = l_
'    ElseIf cmdKey = "m" Then
'        vk = m_
'    ElseIf cmdKey = "n" Then
'        vk = n_
'    ElseIf cmdKey = "o" Then
'        vk = O_
'    ElseIf cmdKey = "p" Then
'        vk = p_
'    ElseIf cmdKey = "q" Then
'        vk = q_
'    ElseIf cmdKey = "r" Then
'        vk = r_
'    ElseIf cmdKey = "s" Then
'        vk = s_
'    ElseIf cmdKey = "t" Then
'        vk = t_
'    ElseIf cmdKey = "u" Then
'        vk = u_
'    ElseIf cmdKey = "v" Then
'        vk = v_
'    ElseIf cmdKey = "w" Then
'        vk = w_
'    ElseIf cmdKey = "X" Then
'        vk = x_
'    ElseIf cmdKey = "Y" Then
'        vk = y_
'    ElseIf cmdKey = "z" Then
'        vk = z_
'    ElseIf cmdKey = "lwin" Then
'        vk = LWin_
'    ElseIf cmdKey = "rwin" Then
'        vk = RWin_
'    ElseIf cmdKey = "apps" Then
'        vk = Apps_
'    ElseIf cmdKey = "rbutton, rwin" Then
'        vk = RButton_RWin_
'    ElseIf cmdKey = "Sleep" Then
'        vk = Sleep_
'    ElseIf cmdKey = "numpad0" Then
'        vk = NumPad0_
'    ElseIf cmdKey = "numpad1" Then
'        vk = NumPad1_
'    ElseIf cmdKey = "numpad2" Then
'        vk = NumPad2_
'    ElseIf cmdKey = "numpad3" Then
'        vk = NumPad3_
'    ElseIf cmdKey = "numpad4" Then
'        vk = NumPad4_
'    ElseIf cmdKey = "numpad5" Then
'        vk = NumPad5_
'    ElseIf cmdKey = "numpad6" Then
'        vk = NumPad6_
'    ElseIf cmdKey = "numpad7" Then
'        vk = NumPad7_
'    ElseIf cmdKey = "numpad8" Then
'        vk = NumPad8_
'    ElseIf cmdKey = "numpad9" Then
'        vk = NumPad9_
'    ElseIf cmdKey = "multiply" Then
'        vk = Multiply_
'    ElseIf cmdKey = "add" Then
'        vk = Add_
'    ElseIf cmdKey = "separator" Then
'        vk = Separator_
'    ElseIf cmdKey = "subtract" Then
'        vk = Subtract_
'    ElseIf cmdKey = "decimal" Then
'        vk = Decimal_
'    ElseIf cmdKey = "divide" Then
'        vk = Divide_
'    ElseIf cmdKey = "f1" Then
'        vk = F1_
'    ElseIf cmdKey = "f2" Then
'        vk = F2_
'    ElseIf cmdKey = "f3" Then
'        vk = F3_
'    ElseIf cmdKey = "f4" Then
'        vk = F4_
'    ElseIf cmdKey = "f5" Then
'        vk = F5_
'    ElseIf cmdKey = "f6" Then
'        vk = F6_
'    ElseIf cmdKey = "f7" Then
'        vk = F7_
'    ElseIf cmdKey = "f8" Then
'        vk = F8_
'    ElseIf cmdKey = "f9" Then
'        vk = F9_
'    ElseIf cmdKey = "f10" Then
'        vk = F10_
'    ElseIf cmdKey = "f11" Then
'        vk = F11_
'    ElseIf cmdKey = "f12" Then
'        vk = F12_
'    ElseIf cmdKey = "f13" Then
'        vk = F13_
'    ElseIf cmdKey = "f14" Then
'        vk = F14_
'    ElseIf cmdKey = "f15" Then
'        vk = F15_
'    ElseIf cmdKey = "f16" Then
'        vk = F16_
'    ElseIf cmdKey = "f17" Then
'        vk = F17_
'    ElseIf cmdKey = "f18" Then
'        vk = F18_
'    ElseIf cmdKey = "f19" Then
'        vk = F19_
'    ElseIf cmdKey = "f20" Then
'        vk = F20_
'    ElseIf cmdKey = "f21" Then
'        vk = F21_
'    ElseIf cmdKey = "f22" Then
'        vk = F22_
'    ElseIf cmdKey = "f23" Then
'        vk = F23_
'    ElseIf cmdKey = "f24" Then
'        vk = F24_
'    ElseIf cmdKey = "back, f17" Then
'        vk = Back_F17_
'    ElseIf cmdKey = "back, f18" Then
'        vk = Back_F18_
'    ElseIf cmdKey = "back, f19" Then
'        vk = Back_F19_
'    ElseIf cmdKey = "back, f20" Then
'        vk = Back_F20_
'    ElseIf cmdKey = "back, f21" Then
'        vk = Back_F21_
'    ElseIf cmdKey = "back, f22" Then
'        vk = Back_F22_
'    ElseIf cmdKey = "back, f23" Then
'        vk = Back_F23_
'    ElseIf cmdKey = "back, f24" Then
'        vk = Back_F24_
'    ElseIf cmdKey = "numlock" Then
'        vk = NumLock_
'    ElseIf cmdKey = "scroll" Or cmdKey = "scrolllock" Then
'        vk = Scroll_
'    ElseIf cmdKey = "rbutton, numlock" Then
'        vk = RButton_NumLock_
'    ElseIf cmdKey = "rbutton, scroll" Then
'        vk = RButton_Scroll_
'    ElseIf cmdKey = "mbutton, numlock" Then
'        vk = MButton_NumLock_
'    ElseIf cmdKey = "mbutton, scroll" Then
'        vk = MButton_Scroll_
'    ElseIf cmdKey = "xbutton2, numlock" Then
'        vk = XButton2_NumLock_
'    ElseIf cmdKey = "xbutton2, scroll" Then
'        vk = XButton2_Scroll_
'    ElseIf cmdKey = "back, numlock" Then
'        vk = Back_NumLock_
'    ElseIf cmdKey = "back, scroll" Then
'        vk = Back_Scroll_
'    ElseIf cmdKey = "linefeed, numlock" Then
'        vk = LineFeed_NumLock_
'    ElseIf cmdKey = "linefeed, scroll" Then
'        vk = LineFeed_Scroll_
'    ElseIf cmdKey = "clear, numlock" Then
'        vk = Clear_NumLock_
'    ElseIf cmdKey = "clear, scroll" Then
'        vk = Clear_Scroll_
'    ElseIf cmdKey = "rbutton, clear, numlock" Then
'        vk = RButton_Clear_NumLock_
'    ElseIf cmdKey = "rbutton, clear, scroll" Then
'        vk = RButton_Clear_Scroll_
'    ElseIf cmdKey = "lshiftkey" Or cmdKey = "lshift" Then
'        vk = LShiftKey_
'    ElseIf cmdKey = "rshiftkey" Or cmdKey = "rshift" Then
'        vk = RShiftKey_
'    ElseIf cmdKey = "lcontrolkey" Or cmdKey = "lcontrol" Then
'        vk = LControlKey_
'    ElseIf cmdKey = "rcontrolkey" Or cmdKey = "rcontrol" Then
'        vk = RControlKey_
'    ElseIf cmdKey = "lmenu" Then
'        vk = LMenu_
'    ElseIf cmdKey = "rmenu" Then
'        vk = RMenu_
'    ElseIf cmdKey = "browserback" Then
'        vk = BrowserBack_
'    ElseIf cmdKey = "browserforward" Then
'        vk = BrowserForward_
'    ElseIf cmdKey = "browserrefresh" Then
'        vk = BrowserRefresh_
'    ElseIf cmdKey = "browserstop" Then
'        vk = BrowserStop_
'    ElseIf cmdKey = "browsersearch" Then
'        vk = BrowserSearch_
'    ElseIf cmdKey = "browserfavorites" Then
'        vk = BrowserFavorites_
'    ElseIf cmdKey = "browserhome" Then
'        vk = BrowserHome_
'    ElseIf cmdKey = "volumemute" Then
'        vk = VolumeMute_
'    ElseIf cmdKey = "volumedown" Then
'        vk = VolumeDown_
'    ElseIf cmdKey = "volumeup" Then
'        vk = VolumeUp_
'    ElseIf cmdKey = "medianexttrack" Then
'        vk = MediaNextTrack_
'    ElseIf cmdKey = "mediaprevioustrack" Then
'        vk = MediaPreviousTrack_
'    ElseIf cmdKey = "mediastop" Then
'        vk = MediaStop_
'    ElseIf cmdKey = "mediaplaypause" Then
'        vk = MediaPlayPause_
'    ElseIf cmdKey = "launchmail" Then
'        vk = LaunchMail_
'    ElseIf cmdKey = "selectmedia" Then
'        vk = SelectMedia_
'    ElseIf cmdKey = "launchapplication1" Then
'        vk = LaunchApplication1_
'    ElseIf cmdKey = "launchapplication2" Then
'        vk = LaunchApplication2_
'    ElseIf cmdKey = "back, medianexttrack" Then
'        vk = Back_MediaNextTrack_
'    ElseIf cmdKey = "back, mediaprevioustrack" Then
'        vk = Back_MediaPreviousTrack_
'    ElseIf cmdKey = "oem1" Then
'        vk = Oem1_
'    ElseIf cmdKey = "oemplus" Then
'        vk = Oemplus_
'    ElseIf cmdKey = "oemcomma" Then
'        vk = Oemcomma_
'    ElseIf cmdKey = "oemminus" Then
'        vk = OemMinus_
'    ElseIf cmdKey = "oemperiod" Then
'        vk = OemPeriod_
'    ElseIf cmdKey = "oemquestion" Then
'        vk = OemQuestion_
'    ElseIf cmdKey = "oemtilde" Then
'        vk = Oemtilde_
'    ElseIf cmdKey = "lbutton, oemtilde" Then
'        vk = LButton_Oemtilde_
'    ElseIf cmdKey = "rbutton, oemtilde" Then
'        vk = RButton_Oemtilde_
'    ElseIf cmdKey = "cancel, oemtilde" Then
'        vk = Cancel_Oemtilde_
'    ElseIf cmdKey = "mbutton, oemtilde" Then
'        vk = MButton_Oemtilde_
'    ElseIf cmdKey = "xbutton1, oemtilde" Then
'        vk = XButton1_Oemtilde_
'    ElseIf cmdKey = "xbutton2, oemtilde" Then
'        vk = XButton2_Oemtilde_
'    ElseIf cmdKey = "lbutton, xbutton2, oemtilde" Then
'        vk = LButton_XButton2_Oemtilde_
'    ElseIf cmdKey = "back, oemtilde" Then
'        vk = Back_Oemtilde_
'    ElseIf cmdKey = "tab, oemtilde" Then
'        vk = Tab_Oemtilde_
'    ElseIf cmdKey = "linefeed, oemtilde" Then
'        vk = LineFeed_Oemtilde_
'    ElseIf cmdKey = "lbutton, linefeed, oemtilde" Then
'        vk = LButton_LineFeed_Oemtilde_
'    ElseIf cmdKey = "clear, oemtilde" Then
'        vk = Clear_Oemtilde_
'    ElseIf cmdKey = "return, oemtilde" Then
'        vk = Return_Oemtilde_
'    ElseIf cmdKey = "rbutton, clear, oemtilde" Then
'        vk = RButton_Clear_Oemtilde_
'    ElseIf cmdKey = "rbutton, return, oemtilde" Then
'        vk = RButton_Return_Oemtilde_
'    ElseIf cmdKey = "shiftkey, oemtilde" Then
'        vk = ShiftKey_Oemtilde_
'    ElseIf cmdKey = "controlkey, oemtilde" Then
'        vk = ControlKey_Oemtilde_
'    ElseIf cmdKey = "menu, oemtilde" Then
'        vk = Menu_Oemtilde_
'    ElseIf cmdKey = "pause, oemtilde" Then
'        vk = Pause_Oemtilde_
'    ElseIf cmdKey = "capital, oemtilde" Then
'        vk = Capital_Oemtilde_
'    ElseIf cmdKey = "kanamode, oemtilde" Then
'        vk = KanaMode_Oemtilde_
'    ElseIf cmdKey = "rbutton, capital, oemtilde" Then
'        vk = RButton_Capital_Oemtilde_
'    ElseIf cmdKey = "junjamode, oemtilde" Then
'        vk = JunjaMode_Oemtilde_
'    ElseIf cmdKey = "finalmode, oemtilde" Then
'        vk = FinalMode_Oemtilde_
'    ElseIf cmdKey = "hanjamode, oemtilde" Then
'        vk = HanjaMode_Oemtilde_
'    ElseIf cmdKey = "rbutton, finalmode, oemtilde" Then
'        vk = RButton_FinalMode_Oemtilde_
'    ElseIf cmdKey = "oemopenbrackets" Then
'        vk = OemOpenBrackets_
'    ElseIf cmdKey = "oem5" Then
'        vk = Oem5_
'    ElseIf cmdKey = "oem6" Then
'        vk = Oem6_
'    ElseIf cmdKey = "oem7" Then
'        vk = Oem7_
'    ElseIf cmdKey = "oem8" Then
'        vk = Oem8_
'    ElseIf cmdKey = "space, oemtilde" Then
'        vk = Space_Oemtilde_
'    ElseIf cmdKey = "pageup, oemtilde" Then
'        vk = PageUp_Oemtilde_
'    ElseIf cmdKey = "oembackslash" Then
'        vk = OemBackslash_
'    ElseIf cmdKey = "lbutton, oembackslash" Then
'        vk = LButton_OemBackslash_
'    ElseIf cmdKey = "home, oemtilde" Then
'        vk = Home_Oemtilde_
'    ElseIf cmdKey = "processkey" Then
'        vk = ProcessKey_
'    ElseIf cmdKey = "mbutton, oembackslash" Then
'        vk = MButton_OemBackslash_
'    ElseIf cmdKey = "packet" Then
'        vk = Packet_
'    ElseIf cmdKey = "down, oemtilde" Then
'        vk = Down_Oemtilde_
'    ElseIf cmdKey = "select, oemtilde" Then
'        vk = Select_Oemtilde_
'    ElseIf cmdKey = "back, oembackslash" Then
'        vk = Back_OemBackslash_
'    ElseIf cmdKey = "tab, oembackslash" Then
'        vk = Tab_OemBackslash_
'    ElseIf cmdKey = "printscreen, oemtilde" Then
'        vk = PrintScreen_Oemtilde_
'    ElseIf cmdKey = "back, processkey" Then
'        vk = Back_ProcessKey_
'    ElseIf cmdKey = "clear, oembackslash" Then
'        vk = Clear_OemBackslash_
'    ElseIf cmdKey = "back, packet" Then
'        vk = Back_Packet_
'    ElseIf cmdKey = "d0, oemtilde" Then
'        vk = D0_Oemtilde_
'    ElseIf cmdKey = "d1, oemtilde" Then
'        vk = D1_Oemtilde_
'    ElseIf cmdKey = "shiftkey, oembackslash" Then
'        vk = ShiftKey_OemBackslash_
'    ElseIf cmdKey = "controlkey, oembackslash" Then
'        vk = ControlKey_OemBackslash_
'    ElseIf cmdKey = "d4, oemtilde" Then
'        vk = D4_Oemtilde_
'    ElseIf cmdKey = "shiftkey, processkey" Then
'        vk = ShiftKey_ProcessKey_
'    ElseIf cmdKey = "attn" Then
'        vk = Attn_
'    ElseIf cmdKey = "crsel" Then
'        vk = Crsel_
'    ElseIf cmdKey = "exsel" Then
'        vk = Exsel_
'    ElseIf cmdKey = "eraseeof" Then
'        vk = EraseEof_
'    ElseIf cmdKey = "play" Then
'        vk = Play_
'    ElseIf cmdKey = "zoom" Then
'        vk = Zoom_
'    ElseIf cmdKey = "noname" Then
'        vk = NoName_
'    ElseIf cmdKey = "pa1" Then
'        vk = Pa1_
'    ElseIf cmdKey = "oemclear" Then
'        vk = OemClear_
'    ElseIf cmdKey = "lbutton, oemclear" Then
'        vk = LButton_OemClear_
'    End If
'    GetKeyCodeFromName = vk
''End Function
'Private Function Lift() As Boolean
'    On Error Resume Next
'    Call apikeybd_event(Keys.ControlKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.LControlKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.RControlKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.ShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.LShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.RShiftKey_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.Menu_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.LMenu_, 0, KEYEVENTF_KEYUP, -11)
'    Call apikeybd_event(Keys.RMenu_, 0, KEYEVENTF_KEYUP, -11)
'    Lift = True
'End Function

