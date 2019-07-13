VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Color"
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
'Attribute VB_Name = "Color"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Declare Function apiGetSysColor Lib "user32" Alias "GetSysColor" (ByVal nIndex As Long) As Long
Private Enum ESysColor
    eSysColor_COLOR_SCROLLBAR = 0
    eSysColor_COLOR_BACKGROUND = 1
    eSysColor_COLOR_ACTIVECAPTION = 2
    eSysColor_COLOR_INACTIVECAPTION = 3
    eSysColor_COLOR_MENU = 4
    eSysColor_COLOR_WINDOW = 5
    eSysColor_COLOR_WINDOWFRAME = 6
    eSysColor_COLOR_MENUTEXT = 7
    eSysColor_COLOR_WINDOWTEXT = 8
    eSysColor_COLOR_CAPTIONTEXT = 9
    eSysColor_COLOR_ACTIVEBORDER = 10
    eSysColor_COLOR_INACTIVEBORDER = 11
    eSysColor_COLOR_APPWORKSPACE = 12
    eSysColor_COLOR_HIGHLIGHT = 13
    eSysColor_COLOR_HIGHLIGHTTEXT = 14
    eSysColor_COLOR_BTNFACE = 15
    eSysColor_COLOR_BTNSHADOW = 16
    eSysColor_COLOR_GRAYTEXT = 17
    eSysColor_COLOR_BTNTEXT = 18
    eSysColor_COLOR_INACTIVECAPTIONTEXT = 19
    eSysColor_COLOR_BTNHIGHLIGHT = 20
    eSysColor_COLOR_3DDKSHADOW = 21
    eSysColor_COLOR_3DLIGHT = 22
    eSysColor_COLOR_INFOTEXT = 23
    eSysColor_COLOR_INFOBK = 24
    eSysColor_COLOR_DESKTOP = eSysColor_COLOR_BACKGROUND
    eSysColor_COLOR_3DFACE = eSysColor_COLOR_BTNFACE
    eSysColor_COLOR_3DSHADOW = eSysColor_COLOR_BTNSHADOW
    eSysColor_COLOR_3DHIGHLIGHT = eSysColor_COLOR_BTNHIGHLIGHT
End Enum
' Private variables to hold property values
Private m_lColor                 As Long
Private m_iRed                   As Byte
Private m_iGreen                 As Byte
Private m_iBlue                  As Byte
Private m_fHue                   As Single
Private m_fSaturation            As Single
Private m_fValue                 As Single
Private mvarTransparent          As Color
Private mvarAliceBlue            As Color
Private mvarAntiqueWhite         As Color
Private mvarAqua                 As Color
Private mvarAquamarine           As Color
Private mvarAzure                As Color
Private mvarBeige                As Color
Private mvarBisque               As Color
Private mvarBlack                As Color
Private mvarBlanchedAlmond       As Color
Private mvarBlue                 As Color
Private mvarBlueViolet           As Color
Private mvarBrown                As Color
Private mvarBurlyWood            As Color
Private mvarCadetBlue            As Color
Private mvarChartreuse           As Color
Private mvarChocolate            As Color
Private mvarCoral                As Color
Private mvarCornflowerBlue       As Color
Private mvarCornsilk             As Color
Private mvarCrimson              As Color
Private mvarCyan                 As Color
Private mvarDarkBlue             As Color
Private mvarDarkCyan             As Color
Private mvarDarkGoldenrod        As Color
Private mvarDarkGray             As Color
Private mvarDarkGreen            As Color
Private mvarDarkKhaki            As Color
Private mvarDarkMagenta          As Color
Private mvarDarkOliveGreen       As Color
Private mvarDarkOrange           As Color
Private mvarDarkOrchid           As Color
Private mvarDarkRed              As Color
Private mvarDarkSalmon           As Color
Private mvarDarkSeaGreen         As Color
Private mvarDarkSlateBlue        As Color
Private mvarDarkSlateGray        As Color
Private mvarDarkTurquoise        As Color
Private mvarDarkViolet           As Color
Private mvarDeepPink             As Color
Private mvarDeepSkyBlue          As Color
Private mvarDimGray              As Color
Private mvarDodgerBlue           As Color
Private mvarFirebrick            As Color
Private mvarFloralWhite          As Color
Private mvarForestGreen          As Color
Private mvarFuchsia              As Color
Private mvarGainsboro            As Color
Private mvarGhostWhite           As Color
Private mvarGold                 As Color
Private mvarGoldenrod            As Color
Private mvarGray                 As Color
Private mvarGreen                As Color
Private mvarGreenYellow          As Color
Private mvarHoneydew             As Color
Private mvarHotPink              As Color
Private mvarIndianRed            As Color
Private mvarIndigo               As Color
Private mvarIvory                As Color
Private mvarKhaki                As Color
Private mvarLavender             As Color
Private mvarLavenderBlush        As Color
Private mvarLawnGreen            As Color
Private mvarLemonChiffon         As Color
Private mvarLightBlue            As Color
Private mvarLightCoral           As Color
Private mvarLightCyan            As Color
Private mvarLightGoldenrodYellow As Color
Private mvarLightGreen           As Color
Private mvarLightGray            As Color
Private mvarLightPink            As Color
Private mvarLightSalmon          As Color
Private mvarLightSeaGreen        As Color
Private mvarLightSkyBlue         As Color
Private mvarLightSlateGray       As Color
Private mvarLightSteelBlue       As Color
Private mvarLightYellow          As Color
Private mvarLime                 As Color
Private mvarLimeGreen            As Color
Private mvarLinen                As Color
Private mvarMagenta              As Color
Private mvarMaroon               As Color
Private mvarMediumAquamarine     As Color
Private mvarMediumBlue           As Color
Private mvarMediumOrchid         As Color
Private mvarMediumPurple         As Color
Private mvarMediumSeaGreen       As Color
Private mvarMediumSlateBlue      As Color
Private mvarMediumSpringGreen    As Color
Private mvarMediumTurquoise      As Color
Private mvarMediumVioletRed      As Color
Private mvarMidnightBlue         As Color
Private mvarMintCream            As Color
Private mvarMistyRose            As Color
Private mvarMoccasin             As Color
Private mvarNavajoWhite          As Color
Private mvarNavy                 As Color
Private mvarOldLace              As Color
Private mvarOlive                As Color
Private mvarOliveDrab            As Color
Private mvarOrange               As Color
Private mvarOrangeRed            As Color
Private mvarOrchid               As Color
Private mvarPaleGoldenrod        As Color
Private mvarPaleGreen            As Color
Private mvarPaleTurquoise        As Color
Private mvarPaleVioletRed        As Color
Private mvarPapayaWhip           As Color
Private mvarPeachPuff            As Color
Private mvarPeru                 As Color
Private mvarPink                 As Color
Private mvarPlum                 As Color
Private mvarPowderBlue           As Color
Private mvarPurple               As Color
Private mvarRed                  As Color
Private mvarRosyBrown            As Color
Private mvarRoyalBlue            As Color
Private mvarSaddleBrown          As Color
Private mvarSalmon               As Color
Private mvarSandyBrown           As Color
Private mvarSeaGreen             As Color
Private mvarSeaShell             As Color
Private mvarSienna               As Color
Private mvarSilver               As Color
Private mvarSkyBlue              As Color
Private mvarSlateBlue            As Color
Private mvarSlateGray            As Color
Private mvarSnow                 As Color
Private mvarSpringGreen          As Color
Private mvarSteelBlue            As Color
Private mvarTan                  As Color
Private mvarTeal                 As Color
Private mvarThistle              As Color
Private mvarTomato               As Color
Private mvarTurquoise            As Color
Private mvarViolet               As Color
Private mvarWheat                As Color
Private mvarWhite                As Color
Private mvarWhiteSmoke           As Color
Private mvarYellow               As Color
Private mvarYellowGreen          As Color
Friend Property Get Transparent() As Color
    Set Transparent = mvarTransparent
End Property
Friend Property Get AliceBlue() As Color
    Set AliceBlue = mvarAliceBlue
End Property
Friend Property Get AntiqueWhite() As Color
    Set AntiqueWhite = mvarAntiqueWhite
End Property
Friend Property Get Aqua() As Color
    Set Aqua = mvarAqua
End Property
Friend Property Get Aquamarine() As Color
    Set Aquamarine = mvarAquamarine
End Property
Friend Property Get Azure() As Color
    Set Azure = mvarAzure
End Property
Friend Property Get Beige() As Color
    Set Beige = mvarBeige
End Property
Friend Property Get Bisque() As Color
    Set Bisque = mvarBisque
End Property
Friend Property Get Black() As Color
    Set Black = mvarBlack
End Property
Friend Property Get BlanchedAlmond() As Color
    Set BlanchedAlmond = mvarBlanchedAlmond
End Property
Friend Property Get Blue() As Color
    Set Blue = mvarBlue
End Property
Friend Property Get BlueViolet() As Color
    Set BlueViolet = mvarBlueViolet
End Property
Friend Property Get Brown() As Color
    Set Brown = mvarBrown
End Property
Friend Property Get BurlyWood() As Color
    Set BurlyWood = mvarBurlyWood
End Property
Friend Property Get CadetBlue() As Color
    Set CadetBlue = mvarCadetBlue
End Property
Friend Property Get Chartreuse() As Color
    Set Chartreuse = mvarChartreuse
End Property
Friend Property Get Chocolate() As Color
    Set Chocolate = mvarChocolate
End Property
Friend Property Get Coral() As Color
    Set Coral = mvarCoral
End Property
Friend Property Get CornflowerBlue() As Color
    Set CornflowerBlue = mvarCornflowerBlue
End Property
Friend Property Get Cornsilk() As Color
    Set Cornsilk = mvarCornsilk
End Property
Friend Property Get Crimson() As Color
    Set Crimson = mvarCrimson
End Property
Friend Property Get Cyan() As Color
    Set Cyan = mvarCyan
End Property
Friend Property Get DarkBlue() As Color
    Set DarkBlue = mvarDarkBlue
End Property
Friend Property Get DarkCyan() As Color
    Set DarkCyan = mvarDarkCyan
End Property
Friend Property Get DarkGoldenrod() As Color
    Set DarkGoldenrod = mvarDarkGoldenrod
End Property
Friend Property Get DarkGray() As Color
    Set DarkGray = mvarDarkGray
End Property
Friend Property Get DarkGreen() As Color
    Set DarkGreen = mvarDarkGreen
End Property
Friend Property Get DarkKhaki() As Color
    Set DarkKhaki = mvarDarkKhaki
End Property
Friend Property Get DarkMagenta() As Color
    Set DarkMagenta = mvarDarkMagenta
End Property
Friend Property Get DarkOliveGreen() As Color
    Set DarkOliveGreen = mvarDarkOliveGreen
End Property
Friend Property Get DarkOrange() As Color
    Set DarkOrange = mvarDarkOrange
End Property
Friend Property Get DarkOrchid() As Color
    Set DarkOrchid = mvarDarkOrchid
End Property
Friend Property Get DarkRed() As Color
    Set DarkRed = mvarDarkRed
End Property
Friend Property Get DarkSalmon() As Color
    Set DarkSalmon = mvarDarkSalmon
End Property
Friend Property Get DarkSeaGreen() As Color
    Set DarkSeaGreen = mvarDarkSeaGreen
End Property
Friend Property Get DarkSlateBlue() As Color
    Set DarkSlateBlue = mvarDarkSlateBlue
End Property
Friend Property Get DarkSlateGray() As Color
    Set DarkSlateGray = mvarDarkSlateGray
End Property
Friend Property Get DarkTurquoise() As Color
    Set DarkTurquoise = mvarDarkTurquoise
End Property
Friend Property Get DarkViolet() As Color
    Set DarkViolet = mvarDarkViolet
End Property
Friend Property Get DeepPink() As Color
    Set DeepPink = mvarDeepPink
End Property
Friend Property Get DeepSkyBlue() As Color
    Set DeepSkyBlue = mvarDeepSkyBlue
End Property
Friend Property Get DimGray() As Color
    Set DimGray = mvarDimGray
End Property
Friend Property Get DodgerBlue() As Color
    Set DodgerBlue = mvarDodgerBlue
End Property
Friend Property Get Firebrick() As Color
    Set Firebrick = mvarFirebrick
End Property
Friend Property Get FloralWhite() As Color
    Set FloralWhite = mvarFloralWhite
End Property
Friend Property Get ForestGreen() As Color
    Set ForestGreen = mvarForestGreen
End Property
Friend Property Get Fuchsia() As Color
    Set Fuchsia = mvarFuchsia
End Property
Friend Property Get Gainsboro() As Color
    Set Gainsboro = mvarGainsboro
End Property
Friend Property Get GhostWhite() As Color
    Set GhostWhite = mvarGhostWhite
End Property
Friend Property Get Gold() As Color
    Set Gold = mvarGold
End Property
Friend Property Get Goldenrod() As Color
    Set Goldenrod = mvarGoldenrod
End Property
Friend Property Get Gray() As Color
    Set Gray = mvarGray
End Property
Friend Property Get Green() As Color
    Set Green = mvarGreen
End Property
Friend Property Get GreenYellow() As Color
    Set GreenYellow = mvarGreenYellow
End Property
Friend Property Get Honeydew() As Color
    Set Honeydew = mvarHoneydew
End Property
Friend Property Get HotPink() As Color
    Set HotPink = mvarHotPink
End Property
Friend Property Get IndianRed() As Color
    Set IndianRed = mvarIndianRed
End Property
Friend Property Get Indigo() As Color
    Set Indigo = mvarIndigo
End Property
Friend Property Get Ivory() As Color
    Set Ivory = mvarIvory
End Property
Friend Property Get Khaki() As Color
    Set Khaki = mvarKhaki
End Property
Friend Property Get Lavender() As Color
    Set Lavender = mvarLavender
End Property
Friend Property Get LavenderBlush() As Color
    Set LavenderBlush = mvarLavenderBlush
End Property
Friend Property Get LawnGreen() As Color
    Set LawnGreen = mvarLawnGreen
End Property
Friend Property Get LemonChiffon() As Color
    Set LemonChiffon = mvarLemonChiffon
End Property
Friend Property Get LightBlue() As Color
    Set LightBlue = mvarLightBlue
End Property
Friend Property Get LightCoral() As Color
    Set LightCoral = mvarLightCoral
End Property
Friend Property Get LightCyan() As Color
    Set LightCyan = mvarLightCyan
End Property
Friend Property Get LightGoldenrodYellow() As Color
    Set LightGoldenrodYellow = mvarLightGoldenrodYellow
End Property
Friend Property Get LightGreen() As Color
    Set LightGreen = mvarLightGreen
End Property
Friend Property Get LightGray() As Color
    Set LightGray = mvarLightGray
End Property
Friend Property Get LightPink() As Color
    Set LightPink = mvarLightPink
End Property
Friend Property Get LightSalmon() As Color
    Set LightSalmon = mvarLightSalmon
End Property
Friend Property Get LightSeaGreen() As Color
    Set LightSeaGreen = mvarLightSeaGreen
End Property
Friend Property Get LightSkyBlue() As Color
    Set LightSkyBlue = mvarLightSkyBlue
End Property
Friend Property Get LightSlateGray() As Color
    Set LightSlateGray = mvarLightSlateGray
End Property
Friend Property Get LightSteelBlue() As Color
    Set LightSteelBlue = mvarLightSteelBlue
End Property
Friend Property Get LightYellow() As Color
    Set LightYellow = mvarLightYellow
End Property
Friend Property Get Lime() As Color
    Set Lime = mvarLime
End Property
Friend Property Get LimeGreen() As Color
    Set LimeGreen = mvarLimeGreen
End Property
Friend Property Get Linen() As Color
    Set Linen = mvarLinen
End Property
Friend Property Get Magenta() As Color
    Set Magenta = mvarMagenta
End Property
Friend Property Get Maroon() As Color
    Set Maroon = mvarMaroon
End Property
Friend Property Get MediumAquamarine() As Color
    Set MediumAquamarine = mvarMediumAquamarine
End Property
Friend Property Get MediumBlue() As Color
    Set MediumBlue = mvarMediumBlue
End Property
Friend Property Get MediumOrchid() As Color
    Set MediumOrchid = mvarMediumOrchid
End Property
Friend Property Get MediumPurple() As Color
    Set MediumPurple = mvarMediumPurple
End Property
Friend Property Get MediumSeaGreen() As Color
    Set MediumSeaGreen = mvarMediumSeaGreen
End Property
Friend Property Get MediumSlateBlue() As Color
    Set MediumSlateBlue = mvarMediumSlateBlue
End Property
Friend Property Get MediumSpringGreen() As Color
    Set MediumSpringGreen = mvarMediumSpringGreen
End Property
Friend Property Get MediumTurquoise() As Color
    Set MediumTurquoise = mvarMediumTurquoise
End Property
Friend Property Get MediumVioletRed() As Color
    Set MediumVioletRed = mvarMediumVioletRed
End Property
Friend Property Get MidnightBlue() As Color
    Set MidnightBlue = mvarMidnightBlue
End Property
Friend Property Get MintCream() As Color
    Set MintCream = mvarMintCream
End Property
Friend Property Get MistyRose() As Color
    Set MistyRose = mvarMistyRose
End Property
Friend Property Get Moccasin() As Color
    Set Moccasin = mvarMoccasin
End Property
Friend Property Get NavajoWhite() As Color
    Set NavajoWhite = mvarNavajoWhite
End Property
Friend Property Get Navy() As Color
    Set Navy = mvarNavy
End Property
Friend Property Get OldLace() As Color
    Set OldLace = mvarOldLace
End Property
Friend Property Get Olive() As Color
    Set Olive = mvarOlive
End Property
Friend Property Get OliveDrab() As Color
    Set OliveDrab = mvarOliveDrab
End Property
Friend Property Get Orange() As Color
    Set Orange = mvarOrange
End Property
Friend Property Get OrangeRed() As Color
    Set OrangeRed = mvarOrangeRed
End Property
Friend Property Get Orchid() As Color
    Set Orchid = mvarOrchid
End Property
Friend Property Get PaleGoldenrod() As Color
    Set PaleGoldenrod = mvarPaleGoldenrod
End Property
Friend Property Get PaleGreen() As Color
    Set PaleGreen = mvarPaleGreen
End Property
Friend Property Get PaleTurquoise() As Color
    Set PaleTurquoise = mvarPaleTurquoise
End Property
Friend Property Get PaleVioletRed() As Color
    Set PaleVioletRed = mvarPaleVioletRed
End Property
Friend Property Get PapayaWhip() As Color
    Set PapayaWhip = mvarPapayaWhip
End Property
Friend Property Get PeachPuff() As Color
    Set PeachPuff = mvarPeachPuff
End Property
Friend Property Get Peru() As Color
    Set Peru = mvarPeru
End Property
Friend Property Get Pink() As Color
    Set Pink = mvarPink
End Property
Friend Property Get Plum() As Color
    Set Plum = mvarPlum
End Property
Friend Property Get PowderBlue() As Color
    Set PowderBlue = mvarPowderBlue
End Property
Friend Property Get Purple() As Color
    Set Purple = mvarPurple
End Property
Friend Property Get Red() As Color
    Set Red = mvarRed
End Property
Friend Property Get RosyBrown() As Color
    Set RosyBrown = mvarRosyBrown
End Property
Friend Property Get RoyalBlue() As Color
    Set RoyalBlue = mvarRoyalBlue
End Property
Friend Property Get SaddleBrown() As Color
    Set SaddleBrown = mvarSaddleBrown
End Property
Friend Property Get Salmon() As Color
    Set Salmon = mvarSalmon
End Property
Friend Property Get SandyBrown() As Color
    Set SandyBrown = mvarSandyBrown
End Property
Friend Property Get SeaGreen() As Color
    Set SeaGreen = mvarSeaGreen
End Property
Friend Property Get SeaShell() As Color
    Set SeaShell = mvarSeaShell
End Property
Friend Property Get Sienna() As Color
    Set Sienna = mvarSienna
End Property
Friend Property Get Silver() As Color
    Set Silver = mvarSilver
End Property
Friend Property Get SkyBlue() As Color
    Set SkyBlue = mvarSkyBlue
End Property
Friend Property Get SlateBlue() As Color
    Set SlateBlue = mvarSlateBlue
End Property
Friend Property Get SlateGray() As Color
    Set SlateGray = mvarSlateGray
End Property
Friend Property Get Snow() As Color
    Set Snow = mvarSnow
End Property
Friend Property Get SpringGreen() As Color
    Set SpringGreen = mvarSpringGreen
End Property
Friend Property Get SteelBlue() As Color
    Set SteelBlue = mvarSteelBlue
End Property
Friend Property Get Tan() As Color
    Set Tan = mvarTan
End Property
Friend Property Get Teal() As Color
    Set Teal = mvarTeal
End Property
Friend Property Get Thistle() As Color
    Set Thistle = mvarThistle
End Property
Friend Property Get Tomato() As Color
    Set Tomato = mvarTomato
End Property
Friend Property Get Turquoise() As Color
    Set Turquoise = mvarTurquoise
End Property
Friend Property Get Violet() As Color
    Set Violet = mvarViolet
End Property
Friend Property Get Wheat() As Color
    Set Wheat = mvarWheat
End Property
Friend Property Get White() As Color
    Set White = mvarWhite
End Property
Friend Property Get WhiteSmoke() As Color
    Set WhiteSmoke = mvarWhiteSmoke
End Property
Friend Property Get Yellow() As Color
    Set Yellow = mvarYellow
End Property
Friend Property Get YellowGreen() As Color
    Set YellowGreen = mvarYellowGreen
End Property
'Dim lngR As Long
'Dim lngG As Long
'Dim lngB As Long
'
'For lngR = 0 To 255
'   For lngG = 0 To 255
'      For lngB = 0 To 255
'         lblClock.BackColor = RGB(lngR, lngG, lngB)
'         Sleep 30 'change this value to change speed
'      Next lngB
'   Next lngG
'Next lngR
Friend Function FromArgb(ByVal argb As Long) As Color
End Function
Friend Function FromKnownColor(ByVal kColor As Color) As Color
End Function
Friend Function FromName(ByVal cName As String) As Color
End Function
Friend Function GetBrightness() As Integer
End Function
Friend Function GetHue() As Integer
End Function
Friend Function GetSaturation() As Color
End Function
Friend Function ToArgb() As Long
End Function
Friend Function ToKnownColor() As Color
End Function
Private Sub Class_Initialize()
    '  Set mvarTransparent = 16777215
    '  Set mvarAliceBlue = -984833
    '  Set mvarAntiqueWhite = -332841
    '  Set mvarAqua = -16711681
    '  Set mvarAquamarine = -8388652
    '  Set mvarAzure = -983041
    '  Set mvarBeige = -657956
    '  Set mvarBisque = -6972
    '  Set mvarBlack = -16777216
    '  Set mvarBlanchedAlmond = -5171
    '  Set mvarBlue = -16776961
    '  Set mvarBlueViolet = -7722014
    '  Set mvarBrown = -5952982
    '  Set mvarBurlyWood = -2180985
    '  Set mvarCadetBlue = -10510688
    '  Set mvarChartreuse = -8388864
    '  Set mvarChocolate = -2987746
    '  Set mvarCoral = -32944
    '  Set mvarCornflowerBlue = -10185235
    '  Set mvarCornsilk = -1828
    '  Set mvarCrimson = -2354116
    '  Set mvarCyan = -16711681
    '  Set mvarDarkBlue = -16777077
    '  Set mvarDarkCyan = -16741493
    '  Set mvarDarkGoldenrod = -4684277
    '  Set mvarDarkGray = -5658199
    '  Set mvarDarkGreen = -16751616
    '  Set mvarDarkKhaki = -4343957
    '  Set mvarDarkMagenta = -7667573
    '  Set mvarDarkOliveGreen = -11179217
    '  Set mvarDarkOrange = -29696
    '  Set mvarDarkOrchid = -6737204
    '  Set mvarDarkRed = -7667712
    '  Set mvarDarkSalmon = -1468806
    '  Set mvarDarkSeaGreen = -7357301
    '  Set mvarDarkSlateBlue = -12042869
    '  Set mvarDarkSlateGray = -13676721
    '  Set mvarDarkTurquoise = -16724271
    '  Set mvarDarkViolet = -7077677
    '  Set mvarDeepPink = -60269
    '  Set mvarDeepSkyBlue = -16728065
    '  Set mvarDimGray = -9868951
    '  Set mvarDodgerBlue = -14774017
    '  Set mvarFirebrick = -5103070
    '  Set mvarFloralWhite = -1296
    '  Set mvarForestGreen = -14513374
    '  Set mvarFuchsia = -65281
    '  Set mvarGainsboro = -2302756
    '  Set mvarGhostWhite = -460545
    '  Set mvarGold = -10496
    '  Set mvarGoldenrod = -2448096
    '  Set mvarGray = -8355712
    '  Set mvarGreen = -16744448
    '  Set mvarGreenYellow = -5374161
    '  Set mvarHoneydew = -983056
    '  Set mvarHotPink = -38476
    '  Set mvarIndianRed = -3318692
    '  Set mvarIndigo = -11861886
    '  Set mvarIvory = -16
    '  Set mvarKhaki = -989556
    '  Set mvarLavender = -1644806
    '  Set mvarLavenderBlush = -3851
    '  Set mvarLawnGreen = -8586240
    '  Set mvarLemonChiffon = -1331
    '  Set mvarLightBlue = -5383962
    '  Set mvarLightCoral = -1015680
    '  Set mvarLightCyan = -2031617
    '  Set mvarLightGoldenrodYellow = -329006
    '  Set mvarLightGreen = -7278960
    '  Set mvarLightGray = -2894893
    '  Set mvarLightPink = -18751
    '  Set mvarLightSalmon = -24454
    '  Set mvarLightSeaGreen = -14634326
    '  Set mvarLightSkyBlue = -7876870
    '  Set mvarLightSlateGray = -8943463
    '  Set mvarLightSteelBlue = -5192482
    '  Set mvarLightYellow = -32
    '  Set mvarLime = -16711936
    '  Set mvarLimeGreen = -13447886
    '  Set mvarLinen = -331546
    '  Set mvarMagenta = -65281
    '  Set mvarMaroon = -8388608
    '  Set mvarMediumAquamarine = -10039894
    '  Set mvarMediumBlue = -16777011
    '  Set mvarMediumOrchid = -4565549
    '  Set mvarMediumPurple = -7114533
    '  Set mvarMediumSeaGreen = -12799119
    '  Set mvarMediumSlateBlue = -8689426
    '  Set mvarMediumSpringGreen = -16713062
    '  Set mvarMediumTurquoise = -12004916
    '  Set mvarMediumVioletRed = -3730043
    '  Set mvarMidnightBlue = -15132304
    '  Set mvarMintCream = -655366
    '  Set mvarMistyRose = -6943
    '  Set mvarMoccasin = -6987
    '  Set mvarNavajoWhite = -8531
    '  Set mvarNavy = -16777088
    '  Set mvarOldLace = -133658
    '  Set mvarOlive = -8355840
    '  Set mvarOliveDrab = -9728477
    '  Set mvarOrange = -23296
    '  Set mvarOrangeRed = -47872
    '  Set mvarOrchid = -2461482
    '  Set mvarPaleGoldenrod = -1120086
    '  Set mvarPaleGreen = -6751336
    '  Set mvarPaleTurquoise = -5247250
    '  Set mvarPaleVioletRed = -2396013
    '  Set mvarPapayaWhip = -4139
    '  Set mvarPeachPuff = -9543
    '  Set mvarPeru = -3308225
    '  Set mvarPink = -16181
    '  Set mvarPlum = -2252579
    '  Set mvarPowderBlue = -5185306
    '  Set mvarPurple = -8388480
    '  Set mvarRed = -65536
    '  Set mvarRosyBrown = -4419697
    '  Set mvarRoyalBlue = -12490271
    '  Set mvarSaddleBrown = -7650029
    '  Set mvarSalmon = -360334
    '  Set mvarSandyBrown = -744352
    '  Set mvarSeaGreen = -13726889
    '  Set mvarSeaShell = -2578
    '  Set mvarSienna = -6270419
    '  Set mvarSilver = -4144960
    '  Set mvarSkyBlue = -7876885
    '  Set mvarSlateBlue = -9807155
    '  Set mvarSlateGray = -9404272
    '  Set mvarSnow = -1286
    '  Set mvarSpringGreen = -16711809
    '  Set mvarSteelBlue = -12156236
    '  Set mvarTan = -2968436
    '  Set mvarTeal = -16744320
    '  Set mvarThistle = -2572328
    '  Set mvarTomato = -40121
    '  Set mvarTurquoise = -12525360
    '  Set mvarViolet = -1146130
    '  Set mvarWheat = -663885
    '  Set mvarWhite = -1
    '  Set mvarWhiteSmoke = -657931
    '  Set mvarYellow = -256
    '  Set mvarYellowGreen = -6632142
End Sub
'*****************************************************************************************
'* Sub         : InitFromLongColor
'* Notes       : Use this method to initialize the object from a long (absolute) color.
'*****************************************************************************************
Public Sub InitFromLongColor(LongColor As Long)
    On Error GoTo hComponentFailure
    m_lColor = LongColor
    m_iRed = m_lColor And &HFF
    m_iGreen = (m_lColor \ &H100) And &HFF
    m_iBlue = (m_lColor \ &H10000) And &HFF
    RGBToHSV
    Exit Sub
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Sub
Public Sub InitFromRGB(Red As Byte, Green As Byte, Blue As Byte)
    On Error GoTo hComponentFailure
    m_lColor = RGB(Red, Green, Blue)
    m_iRed = Red
    m_iGreen = Green
    m_iBlue = Blue
    RGBToHSV
    Exit Sub
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Sub
Public Sub InitFromHSV(Hue As Single, Saturation As Single, Value As Single)
    On Error GoTo hComponentFailure
    If Not (0 <= Hue And Hue <= 359) Then
        On Error GoTo 0
        'Err.Raise eErrColor_InvalidHue, "VB65" & ".CColor", S_ERR_InvalidHue
    End If
    If Not (0 <= Saturation And Saturation <= 100) Then
        On Error GoTo 0
        'Err.Raise eErrColor_InvalidSaturation, "VB65" & ".CColor", S_ERR_InvalidSaturation
    End If
    If Not (0 <= Value And Value <= 100) Then
        On Error GoTo 0
        'Err.Raise eErrColor_InvalidValue, "VB65" & ".CColor", S_ERR_InvalidValue
    End If
    m_fHue = Hue
    m_fSaturation = Saturation
    m_fValue = Value
    m_lColor = ColorFromHSV(Hue, Saturation, Value)
    m_iRed = m_lColor And &HFF
    m_iGreen = (m_lColor \ &H100) And &HFF
    m_iBlue = (m_lColor \ &H10000) And &HFF
    Exit Sub
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Sub
'*****************************************************************************************
'Public Sub InitFromSysColor(Colr As SystemColorConstants)
'    On Error GoTo hComponentFailure
'    Dim lRet As Long
'    Select Case Colr
'        Case vb3DDKShadow
'            lRet = apiGetSysColor(eSysColor_COLOR_3DDKSHADOW)
'        Case vb3DFace
'            lRet = apiGetSysColor(eSysColor_COLOR_3DFACE)
'        Case vb3DHighlight
'            lRet = apiGetSysColor(eSysColor_COLOR_3DHIGHLIGHT)
'        Case vb3DLight
'            lRet = apiGetSysColor(eSysColor_COLOR_3DLIGHT)
'        Case vb3DShadow
'            lRet = apiGetSysColor(eSysColor_COLOR_3DSHADOW)
'        Case vbActiveBorder
'            lRet = apiGetSysColor(eSysColor_COLOR_ACTIVEBORDER)
'        Case vbActiveTitleBar
'            lRet = apiGetSysColor(eSysColor_COLOR_ACTIVECAPTION)
'        Case vbActiveTitleBarText
'            lRet = apiGetSysColor(eSysColor_COLOR_CAPTIONTEXT)
'        Case vbApplicationWorkspace
'            lRet = apiGetSysColor(eSysColor_COLOR_APPWORKSPACE)
'        Case vbButtonFace
'            lRet = apiGetSysColor(eSysColor_COLOR_BTNFACE)
'        Case vbButtonShadow
'            lRet = apiGetSysColor(eSysColor_COLOR_BTNSHADOW)
'        Case vbButtonText
'            lRet = apiGetSysColor(eSysColor_COLOR_BTNTEXT)
'        Case vbDesktop
'            lRet = apiGetSysColor(eSysColor_COLOR_DESKTOP)
'        Case vbGrayText
'            lRet = apiGetSysColor(eSysColor_COLOR_GRAYTEXT)
'        Case vbHighlight
'            lRet = apiGetSysColor(eSysColor_COLOR_HIGHLIGHT)
'        Case vbHighlightText
'            lRet = apiGetSysColor(eSysColor_COLOR_HIGHLIGHTTEXT)
'        Case vbInactiveBorder
'            lRet = apiGetSysColor(eSysColor_COLOR_INACTIVEBORDER)
'        Case vbInactiveCaptionText
'            lRet = apiGetSysColor(eSysColor_COLOR_INACTIVECAPTION)
'        Case vbInactiveTitleBar
'            lRet = apiGetSysColor(eSysColor_COLOR_INACTIVECAPTION)
'        Case vbInactiveTitleBarText
'            lRet = apiGetSysColor(eSysColor_COLOR_INACTIVECAPTIONTEXT)
'        Case vbInfoBackground
'            lRet = apiGetSysColor(eSysColor_COLOR_INFOBK)
'        Case vbInfoText
'            lRet = apiGetSysColor(eSysColor_COLOR_INFOTEXT)
'        Case vbMenuBar
'            lRet = apiGetSysColor(eSysColor_COLOR_MENU)
'        Case vbMenuText
'            lRet = apiGetSysColor(eSysColor_COLOR_MENUTEXT)
'        Case vbScrollBars
'            lRet = apiGetSysColor(eSysColor_COLOR_SCROLLBAR)
'        Case vbTitleBarText
'            lRet = apiGetSysColor(eSysColor_COLOR_CAPTIONTEXT)
'        Case vbWindowBackground
'            lRet = apiGetSysColor(eSysColor_COLOR_WINDOW)
'        Case vbWindowFrame
'            lRet = apiGetSysColor(eSysColor_COLOR_WINDOWFRAME)
'        Case vbWindowText
'            lRet = apiGetSysColor(eSysColor_COLOR_WINDOWTEXT)
'        Case Else
'            On Error GoTo 0
'            'Err.Raise eErrColor_UnknownSystemColor, "VB65" & ".CColor", S_ERR_UnknownSystemColor
'    End Select
'    If lRet = 0 Then
'        On Error GoTo 0
'        'Err.Raise eErrColor_CannotGetSystemColor, "VB65" & ".CColor", S_ERR_CannotGetSystemColor
'    End If
'    Exit Sub
'hComponentFailure:
'    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
'End Sub
Friend Property Get LongColor()
    On Error GoTo hComponentFailure
    LongColor = m_lColor
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get RgbRed() As Byte
    On Error GoTo hComponentFailure
    RgbRed = m_iRed
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get RgbGreen() As Byte
    On Error GoTo hComponentFailure
    RgbGreen = m_iGreen
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get RgbBlue() As Byte
    On Error GoTo hComponentFailure
    RgbBlue = m_iBlue
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get HsvHue() As Single
    On Error GoTo hComponentFailure
    HsvHue = m_fHue
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get HsvSaturation() As Single
    On Error GoTo hComponentFailure
    HsvSaturation = m_fSaturation
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Friend Property Get HsvValue() As Single
    On Error GoTo hComponentFailure
    HsvValue = m_fValue
    Exit Property
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Property
Private Function ColorFromHSV(Hue As Single, Saturation As Single, Value As Single) As Long
    On Error GoTo hComponentFailure
    Dim fSaturation As Single
    Dim fValue      As Single
    Dim fHue        As Single
    Dim nI          As Integer
    Dim fF          As Single
    Dim fP          As Single
    Dim fQ          As Single
    Dim ft          As Single
    Dim fRed        As Single
    Dim fGreen      As Single
    Dim fBlue       As Single
    fSaturation = Saturation / 100
    fValue = Value / 100
    If Saturation = 0 Then
        fRed = fValue
        fGreen = fValue
        fBlue = fValue
    Else
        fHue = Hue / 60
        If fHue = 6 Then fHue = 0
        nI = Int(fHue)
        fF = fHue - nI
        fP = fValue * (1 - fSaturation)
        fQ = fValue * (1 - (fSaturation * fF))
        ft = fValue * (1 - (fSaturation * (1 - fF)))
        Select Case nI
            Case 0
                fRed = fValue
                fGreen = ft
                fBlue = fP
            Case 1
                fRed = fQ
                fGreen = fValue
                fBlue = fP
            Case 2
                fRed = fP
                fGreen = fValue
                fBlue = ft
            Case 3
                fRed = fP
                fGreen = fQ
                fBlue = fValue
            Case 4
                fRed = ft
                fGreen = fP
                fBlue = fValue
            Case 5
                fRed = fValue
                fGreen = fP
                fBlue = fQ
        End Select
    End If
    ColorFromHSV = RGB(Int(255.9999 * fRed), Int(255.9999 * fGreen), Int(255.9999 * fBlue))
    Exit Function
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Function
Private Sub RGBToHSV()
    On Error GoTo hComponentFailure
    Dim fRed   As Single
    Dim fGreen As Single
    Dim fBlue  As Single
    Dim fMx    As Single
    Dim fMn    As Single
    Dim fVa    As Single
    Dim fSa    As Single
    Dim fRc    As Single
    Dim fGc    As Single
    Dim fBc    As Single
    fRed = m_iRed / 255
    fGreen = m_iGreen / 255
    fBlue = m_iBlue / 255
    fMx = fRed
    If fGreen > fMx Then fMx = fGreen
    If fBlue > fMx Then fMx = fBlue
    fMn = fRed
    If fGreen < fMn Then fMn = fGreen
    If fBlue < fMn Then fMn = fBlue
    fVa = fMx
    If fMx Then
        fSa = (fMx - fMn) / fMx
    Else
        fSa = 0
    End If
    If fSa = 0 Then
        m_fHue = 0
    Else
        fRc = (fMx - fRed) / (fMx - fMn)
        fGc = (fMx - fGreen) / (fMx - fMn)
        fBc = (fMx - fBlue) / (fMx - fMn)
        Select Case fMx
            Case fRed
                m_fHue = fBc - fGc
            Case fGreen
                m_fHue = 2 + fRc - fBc
            Case fBlue
                m_fHue = 4 + fGc - fRc
        End Select
        m_fHue = m_fHue * 60
        If m_fHue < 0 Then m_fHue = m_fHue + 360
    End If
    m_fSaturation = fSa * 100
    m_fValue = fVa * 100
    Exit Sub
hComponentFailure:
    'Err.Raise eErrColor_ComponentFailure, "VB65" & ".CColor", S_ERR_ComponentFailure
End Sub

