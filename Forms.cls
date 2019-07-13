VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Forms"
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
'Attribute VB_Name = "Forms"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"SendKeys"
'Attribute VB_Ext_KEY = "Member1" ,"Messagebox"
'Attribute VB_Ext_KEY = "Member2" ,"Application"
'Attribute VB_Ext_KEY = "Member3" ,"SaveFileDialog"
'Attribute VB_Ext_KEY = "Member4" ,"PrintDialog"
'Attribute VB_Ext_KEY = "Member5" ,"OpenFileDialog"
'Attribute VB_Ext_KEY = "Member6" ,"FontDialog"
'Attribute VB_Ext_KEY = "Member7" ,"ColorDialog"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private mvarApplication    As Application
Private mvarMessagebox     As MessageBox
Private mvarSendKeys       As Sendkeys
Private mvarColorDialog    As ColorDialog
Private mvarFontDialog     As FontDialog
Private mvarOpenFileDialog As OpenFileDialog
Private mvarPrintDialog    As PrintDialog
Private mvarSaveFileDialog As SaveFileDialog
Public Enum Keys
    None_ = 0
    LButton_ = 1
    RButton_ = 2
    Cancel_ = 3
    MButton_ = 4
    XButton1_ = 5
    XButton2_ = 6
    LButton_XButton2_ = 7
    Back_ = 8
    Tab_ = 9
    LineFeed_ = 10
    LButton_LineFeed_ = 11
    Clear_ = 12
    Return_ = 13
    RButton_Clear_ = 14
    RButton_Return_ = 15
    ShiftKey_ = 16
    ControlKey_ = 17
    Menu_ = 18
    Pause_ = 19
    Capital_ = 20
    KanaMode_ = 21
    RButton_Capital_ = 22
    JunjaMode_ = 23
    FinalMode_ = 24
    HanjaMode_ = 25
    RButton_FinalMode_ = 26
    Escape_ = 27
    IMEConvert_ = 28
    IMENonconvert_ = 29
    IMEAceept_ = 30
    IMEModeChange_ = 31
    Space_ = 32
    PageUp_ = 33
    Next_ = 34
    End_ = 35
    Home_ = 36
    Left_ = 37
    Up_ = 38
    Right_ = 39
    Down_ = 40
    Select_ = 41
    Print_ = 42
    Execute_ = 43
    PrintScreen_ = 44
    Insert_ = 45
    Delete_ = 46
    Help_ = 47
    D0_ = 48
    D1_ = 49
    D2_ = 50
    D3_ = 51
    D4_ = 52
    D5_ = 53
    D6_ = 54
    D7_ = 55
    D8_ = 56
    D9_ = 57
    RButton_D8_ = 58
    RButton_D9_ = 59
    MButton_D8_ = 60
    MButton_D9_ = 61
    XButton2_D8_ = 62
    XButton2_D9_ = 63
    VK_64_ = 64
    a_ = 65
    b_ = 66
    c_ = 67
    d_ = 68
    e_ = 69
    f_ = 70
    g_ = 71
    h_ = 72
    i_ = 73
    J_ = 74
    k_ = 75
    l_ = 76
    m_ = 77
    n_ = 78
    O_ = 79
    p_ = 80
    q_ = 81
    r_ = 82
    s_ = 83
    t_ = 84
    u_ = 85
    v_ = 86
    w_ = 87
    x_ = 88
    y_ = 89
    z_ = 90
    LWin_ = 91
    RWin_ = 92
    Apps_ = 93
    RButton_RWin_ = 94
    Sleep_ = 95
    NumPad0_ = 96
    NumPad1_ = 97
    NumPad2_ = 98
    NumPad3_ = 99
    NumPad4_ = 100
    NumPad5_ = 101
    NumPad6_ = 102
    NumPad7_ = 103
    NumPad8_ = 104
    NumPad9_ = 105
    Multiply_ = 106
    Add_ = 107
    Separator_ = 108
    Subtract_ = 109
    Decimal_ = 110
    Divide_ = 111
    F1_ = 112
    F2_ = 113
    F3_ = 114
    F4_ = 115
    F5_ = 116
    F6_ = 117
    F7_ = 118
    F8_ = 119
    F9_ = 120
    F10_ = 121
    F11_ = 122
    F12_ = 123
    F13_ = 124
    F14_ = 125
    F15_ = 126
    F16_ = 127
    F17_ = 128
    F18_ = 129
    F19_ = 130
    F20_ = 131
    F21_ = 132
    F22_ = 133
    F23_ = 134
    F24_ = 135
    Back_F17_ = 136
    Back_F18_ = 137
    Back_F19_ = 138
    Back_F20_ = 139
    Back_F21_ = 140
    Back_F22_ = 141
    Back_F23_ = 142
    Back_F24_ = 143
    NumLock_ = 144
    Scroll_ = 145
    RButton_NumLock_ = 146
    RButton_Scroll_ = 147
    MButton_NumLock_ = 148
    MButton_Scroll_ = 149
    XButton2_NumLock_ = 150
    XButton2_Scroll_ = 151
    Back_NumLock_ = 152
    Back_Scroll_ = 153
    LineFeed_NumLock_ = 154
    LineFeed_Scroll_ = 155
    Clear_NumLock_ = 156
    Clear_Scroll_ = 157
    RButton_Clear_NumLock_ = 158
    RButton_Clear_Scroll_ = 159
    LShiftKey_ = 160
    RShiftKey_ = 161
    LControlKey_ = 162
    RControlKey_ = 163
    LMenu_ = 164
    RMenu_ = 165
    BrowserBack_ = 166
    BrowserForward_ = 167
    BrowserRefresh_ = 168
    BrowserStop_ = 169
    BrowserSearch_ = 170
    BrowserFavorites_ = 171
    BrowserHome_ = 172
    VolumeMute_ = 173
    VolumeDown_ = 174
    VolumeUp_ = 175
    MediaNextTrack_ = 176
    MediaPreviousTrack_ = 177
    MediaStop_ = 178
    MediaPlayPause_ = 179
    LaunchMail_ = 180
    SelectMedia_ = 181
    LaunchApplication1_ = 182
    LaunchApplication2_ = 183
    Back_MediaNextTrack_ = 184
    Back_MediaPreviousTrack_ = 185
    Oem1_ = 186
    Oemplus_ = 187
    Oemcomma_ = 188
    OemMinus_ = 189
    OemPeriod_ = 190
    OemQuestion_ = 191
    Oemtilde_ = 192
    LButton_Oemtilde_ = 193
    RButton_Oemtilde_ = 194
    Cancel_Oemtilde_ = 195
    MButton_Oemtilde_ = 196
    XButton1_Oemtilde_ = 197
    XButton2_Oemtilde_ = 198
    LButton_XButton2_Oemtilde_ = 199
    Back_Oemtilde_ = 200
    Tab_Oemtilde_ = 201
    LineFeed_Oemtilde_ = 202
    LButton_LineFeed_Oemtilde_ = 203
    Clear_Oemtilde_ = 204
    Return_Oemtilde_ = 205
    RButton_Clear_Oemtilde_ = 206
    RButton_Return_Oemtilde_ = 207
    ShiftKey_Oemtilde_ = 208
    ControlKey_Oemtilde_ = 209
    Menu_Oemtilde_ = 210
    Pause_Oemtilde_ = 211
    Capital_Oemtilde_ = 212
    KanaMode_Oemtilde_ = 213
    RButton_Capital_Oemtilde_ = 214
    JunjaMode_Oemtilde_ = 215
    FinalMode_Oemtilde_ = 216
    HanjaMode_Oemtilde_ = 217
    RButton_FinalMode_Oemtilde_ = 218
    OemOpenBrackets_ = 219
    Oem5_ = 220
    Oem6_ = 221
    Oem7_ = 222
    Oem8_ = 223
    Space_Oemtilde_ = 224
    PageUp_Oemtilde_ = 225
    OemBackslash_ = 226
    LButton_OemBackslash_ = 227
    Home_Oemtilde_ = 228
    ProcessKey_ = 229
    MButton_OemBackslash_ = 230
    Packet_ = 231
    Down_Oemtilde_ = 232
    Select_Oemtilde_ = 233
    Back_OemBackslash_ = 234
    Tab_OemBackslash_ = 235
    PrintScreen_Oemtilde_ = 236
    Back_ProcessKey_ = 237
    Clear_OemBackslash_ = 238
    Back_Packet_ = 239
    D0_Oemtilde_ = 240
    D1_Oemtilde_ = 241
    ShiftKey_OemBackslash_ = 242
    ControlKey_OemBackslash_ = 243
    D4_Oemtilde_ = 244
    ShiftKey_ProcessKey_ = 245
    Attn_ = 246
    Crsel_ = 247
    Exsel_ = 248
    EraseEof_ = 249
    Play_ = 250
    Zoom_ = 251
    NoName_ = 252
    Pa1_ = 253
    OemClear_ = 254
    LButton_OemClear_ = 255
End Enum
Private Sub Class_Initialize()
    Set mvarApplication = New Application
    Set mvarMessagebox = New MessageBox
    Set mvarSendKeys = New Sendkeys
    Set mvarColorDialog = New ColorDialog
    Set mvarFontDialog = New FontDialog
    Set mvarOpenFileDialog = New OpenFileDialog
    Set mvarPrintDialog = New PrintDialog
    Set mvarSaveFileDialog = New SaveFileDialog
End Sub
Private Sub Class_Terminate()
    Set mvarApplication = Nothing
    Set mvarMessagebox = Nothing
    Set mvarSendKeys = Nothing
    Set mvarColorDialog = Nothing
    Set mvarFontDialog = Nothing
    Set mvarOpenFileDialog = Nothing
    Set mvarPrintDialog = Nothing
    Set mvarSaveFileDialog = Nothing
End Sub

Friend Property Get ColorDialog() As ColorDialog
    Set ColorDialog = mvarColorDialog
End Property
Friend Property Get FontDialog() As FontDialog
    Set FontDialog = mvarFontDialog
End Property
Friend Property Get OpenFileDialog() As OpenFileDialog
    Set OpenFileDialog = mvarOpenFileDialog
End Property
Friend Property Get PrintDialog() As PrintDialog
    Set PrintDialog = mvarPrintDialog
End Property
Friend Property Get SaveFileDialog() As SaveFileDialog
    Set SaveFileDialog = mvarSaveFileDialog
End Property
Friend Property Get Application() As Application
    Set Application = mvarApplication
End Property
Friend Property Get MessageBox() As MessageBox
    Set MessageBox = mvarMessagebox
End Property
Friend Property Get Sendkeys() As Sendkeys
    Set Sendkeys = mvarSendKeys
End Property

