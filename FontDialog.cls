VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "FontDialog"
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
'Attribute VB_Name = "FontDialog"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const CF_ANSIONLY             As Long = &H400
Private Const CF_APPLY                As Long = &H200
Private Const CF_BOTH                 As Long = &H3
Private Const CF_EFFECTS              As Long = &H100
Private Const CF_ENABLEHOOK           As Long = &H8
Private Const CF_ENABLETEMPLATE       As Long = &H10
Private Const CF_ENABLETEMPLATEHANDLE As Long = &H20
Private Const CF_FIXEDPITCHONLY       As Long = &H4000
Private Const CF_FORCEFONTEXIST       As Long = &H10000
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40
Private Const CF_LIMITSIZE            As Long = &H2000
Private Const CF_NOOEMFONTS           As Long = &H800
Private Const CF_NOFACESEL            As Long = &H80000
Private Const CF_NOSCRIPTSEL          As Long = &H800000
Private Const CF_NOSIZESEL            As Long = &H200000
Private Const CF_NOSIMULATIONS        As Long = &H1000
Private Const CF_NOSTYLESEL           As Long = &H100000
Private Const CF_NOVECTORFONTS        As Long = &H800
Private Const CF_NOVERTFONTS          As Long = &H1000000
Private Const CF_PRINTERFONTS         As Long = &H2
Private Const CF_SCALABLEONLY         As Long = &H20000
Private Const CF_SCREENFONTS          As Long = &H1
Private Const CF_SCRIPTSONLY          As Long = &H400
Private Const CF_SELECTSCRIPT         As Long = &H400000
Private Const CF_SHOWHELP             As Long = &H4
Private Const CF_TTONLY               As Long = &H40000
Private Const CF_USESTYLE             As Long = &H80
Private Const CF_WYSIWYG              As Long = &H8000
Private Const BOLD_FONTTYPE           As Long = &H100
Private Const ITALIC_FONTTYPE         As Long = &H200
Private Const PRINTER_FONTTYPE        As Long = &H4000
Private Const REGULAR_FONTTYPE        As Long = &H400
Private Const SCREEN_FONTTYPE         As Long = &H2000
Private Const SIMULATED_FONTTYPE      As Long = &H8000
Private Const FW_DONTCARE             As Long = 0
Private Const FW_THIN                 As Long = 100
Private Const FW_EXTRALIGHT           As Long = 200
Private Const FW_ULTRALIGHT           As Long = 200
Private Const FW_LIGHT                As Long = 300
Private Const FW_NORMAL               As Long = 400
Private Const FW_REGULAR              As Long = 400
Private Const FW_MEDIUM               As Long = 500
Private Const FW_SEMIBOLD             As Long = 600
Private Const FW_DEMIBOLD             As Long = 600
Private Const FW_BOLD                 As Long = 700
Private Const FW_EXTRABOLD            As Long = 800
Private Const FW_ULTRABOLD            As Long = 800
Private Const FW_HEAVY                As Long = 900
Private Const FW_BLACK                As Long = 900
Private Const ANSI_CHARSET            As Long = 0
Private Const ARABIC_CHARSET          As Long = 178
Private Const BALTIC_CHARSET          As Long = 186
Private Const CHINESEBIG5_CHARSET     As Long = 136
Private Const DEFAULT_CHARSET         As Long = 1
Private Const EASTEUROPE_CHARSET      As Long = 238
Private Const GB2312_CHARSET          As Long = 134
Private Const GREEK_CHARSET           As Long = 161
Private Const HANGEUL_CHARSET         As Long = 129
Private Const HEBREW_CHARSET          As Long = 177
Private Const JOHAB_CHARSET           As Long = 130
Private Const MAC_CHARSET             As Long = 77
Private Const OEM_CHARSET             As Long = 255
Private Const RUSSIAN_CHARSET         As Long = 204
Private Const SHIFTJIS_CHARSET        As Long = 128
Private Const SYMBOL_CHARSET          As Long = 2
Private Const THAI_CHARSET            As Long = 222
Private Const TURKISH_CHARSET         As Long = 162
Private Const OUT_DEFAULT_PRECIS      As Long = 0
Private Const OUT_DEVICE_PRECIS       As Long = 5
Private Const OUT_OUTLINE_PRECIS      As Long = 8
Private Const OUT_RASTER_PRECIS       As Long = 6
Private Const OUT_STRING_PRECIS       As Long = 1
Private Const OUT_STROKE_PRECIS       As Long = 3
Private Const OUT_TT_ONLY_PRECIS      As Long = 7
Private Const OUT_TT_PRECIS           As Long = 4
Private Const CLIP_DEFAULT_PRECIS     As Long = 0
Private Const CLIP_EMBEDDED           As Long = 128
Private Const CLIP_LH_ANGLES          As Long = 16
Private Const CLIP_STROKE_PRECIS      As Long = 2
Private Const ANTIALIASED_QUALITY     As Long = 4
Private Const DEFAULT_QUALITY         As Long = 0
Private Const DRAFT_QUALITY           As Long = 1
Private Const NONANTIALIASED_QUALITY  As Long = 3
Private Const PROOF_QUALITY           As Long = 2
Private Const DEFAULT_PITCH           As Long = 0
Private Const FIXED_PITCH             As Long = 1
Private Const VARIABLE_PITCH          As Long = 2
Private Const FF_DECORATIVE           As Long = 80
Private Const FF_DONTCARE             As Long = 0
Private Const FF_MODERN               As Long = 48
Private Const FF_ROMAN                As Long = 16
Private Const FF_SCRIPT               As Long = 64
Private Const FF_SWISS                As Long = 32
Private Const GHND                    As Long = &H40 'Same as combining GMEM_MOVEABLE with GMEM_ZEROINIT.
Private Const GMEM_DDESHARE           As Long = &H2000 'Optimize the allocated memory for use in DDE conversations.
Private Const GMEM_DISCARDABLE        As Long = &H100 'Allocate discardable memory. (Cannot be combined with GMEM_FIXED.)
Private Const GMEM_FIXED              As Long = &H0 'Allocate fixed memory. The function's return value is a pointer to the beginning of the memory block. (Cannot be combined with GMEM_DISCARDABLE or GMEM_MOVEABLE.)
Private Const GMEM_MOVEABLE           As Long = &H2 'Allocate moveable memory. The memory block's lock count is initialized at 0 (unlocked). The function's return value is a handle to the beginning of the memory block. (Cannot be combined with GMEM_FIXED.)
Private Const GMEM_NOCOMPACT          As Long = &H10 'Do not compact any memory or discard any discardable memory to allocate the requested block.
Private Const GMEM_NODISCARD          As Long = &H20 'Do not discard any discardable memory to allocate the requested block.
Private Const GMEM_SHARE              As Long = &H2000 'Sa'me as GMEM_DDESHARE.
Private Const GMEM_ZEROINIT           As Long = &H40 'Initialize the contents of the memory block to 0.
Private Const GPTR                    As Long = &H42 'Same as combining GMEM_FIXED with GMEM_ZEROINIT.
Private Type CHOOSEFONT_TYPE
    lStructSize As Long
    hwndOwner As Long
    HDC As Long
    lpLogFont As Long
    iPointSize As Long
    Flags As Long
    rgbColors As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
    hInstance As Long
    lpszStyle As String
    nFontType As Integer
    MISSING_ALIGNMENT As Integer
    nSizeMin As Long
    nSizeMax As Long
End Type
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type
Private Declare Function apiChooseFont Lib "comdlg32" Alias "ChooseFontA" (ByRef lpcf As CHOOSEFONT_TYPE) As Long
Private Declare Function apiGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function apiGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Sub apiCopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal Length As Long)
Private Declare Function apiGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare Function apiGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Friend Function ChooseFontDialog(ByVal hwnd As Long) As Long
    ' Display a Choose Font dialog box.  Print out the typeface name, point size,
    ' and style of the selected font.  More detail about topics in this example can be found in
    ' the pages for CHOOSEFONT_TYPE and LOGFONT.
    Dim cf       As CHOOSEFONT_TYPE  ' data structure needed for function
    Dim lfont    As LOGFONT  ' receives information about the chosen font
    Dim hMem     As Long, pMem As Long  ' handle and pointer to memory buffer
    Dim fontname As String  ' receives name of font selected
    Dim ret      As Long  ' return value
    ' Initialize the default selected font: Times New Roman, regular, black, 12 point.
    ' (Note that some of that information is in the CHOOSEFONT_TYPE structure instead.)
    With lfont
        .lfHeight = 0  ' determine default height
        .lfWidth = 0  ' determine default width
        .lfEscapement = 0  ' angle between baseline and escapement vector
        .lfOrientation = 0  ' angle between baseline and orientation vector
        .lfWeight = FW_NORMAL  ' normal weight i.e. not bold
        .lfItalic = 0  ' not italic
        .lfUnderline = 0  ' not underline
        .lfStrikeOut = 0  ' not strikeout
        .lfCharSet = DEFAULT_CHARSET  ' use default character set
        .lfOutPrecision = OUT_DEFAULT_PRECIS  ' default precision mapping
        .lfClipPrecision = CLIP_DEFAULT_PRECIS  ' default clipping precision
        .lfQuality = DEFAULT_QUALITY  ' default quality setting
        .lfPitchAndFamily = DEFAULT_PITCH Or FF_ROMAN  ' default pitch, proportional with serifs
        .lfFaceName = "Times New Roman" & vbNullChar  ' string must be null-terminated
    End With
    ' Create the memory block which will act as the LOGFONT structure buffer.
    hMem = apiGlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(lfont))
    pMem = apiGlobalLock(hMem)  ' lock and get pointer
    Call apiCopyMemory(ByVal pMem, lfont, Len(lfont))   ' copy structure's contents into block
    ' Initialize dialog box: Screen and printer fonts, point size between 10 and 72.
    cf.lStructSize = Len(cf)  ' size of structure
    With cf
        .hwndOwner = hwnd  ' window Form1 is opening this dialog box
       ' .HDC = Printer.HDC  ' device context of default printer (using VB's mechanism)
        .lpLogFont = pMem   ' pointer to LOGFONT memory block buffer
        .iPointSize = 120  ' 12 point font (in units of 1/10 point)
        .Flags = CF_BOTH Or CF_EFFECTS Or CF_FORCEFONTEXIST Or CF_INITTOLOGFONTSTRUCT Or CF_LIMITSIZE
        .rgbColors = RGB(0, 0, 0)  ' black
        .lCustData = 0  ' we don't use this here...
        .lpfnHook = 0  ' ...or this...
        .lpTemplateName = ""  ' ...or this...
        .hInstance = 0  ' ...or this...
        .lpszStyle = ""  ' ...or this
        .nFontType = REGULAR_FONTTYPE  ' regular font type i.e. not bold or anything
        .nSizeMin = 10  ' minimum point size
        .nSizeMax = 72  ' maximum point size
    End With
    ' Now, call the function.  If successful, copy the LOGFONT structure back into the structure
    ' and then print out the attributes we mentioned earlier that the user selected.
    ret = apiChooseFont(cf)  ' open the dialog box
    ChooseFontDialog = ret
    If ret <> 0 Then  ' success
        Call apiCopyMemory(lfont, ByVal pMem, Len(lfont))   ' copy memory back
        ' Now make the fixed-length string holding the font name into a "normal" string.
        ' fontname = Left(lfont.lfFaceName, InStr(lfont.lfFaceName, vbNullChar) - 1)
        ' Display font name and a few attributes.
        '  Debug.Print "FONT NAME: "; fontname
        '  Debug.Print "FONT SIZE (points):"; cf.iPointSize / 10  ' in units of 1/10 point!
        '  Debug.Print "FONT STYLE(S): ";
        '  If lfont.lfWeight >= FW_BOLD Then Debug.Print "Bold ";
        '  If lfont.lfItalic <> 0 Then Debug.Print "Italic ";
        '  If lfont.lfUnderline <> 0 Then Debug.Print "Underline ";
        '  If lfont.lfStrikeOut <> 0 Then Debug.Print "Strikeout";
    End If
    ' Deallocate the memory block we created earlier.  Note that this must
    ' be done whether the function succeeded or not.
    Call apiGlobalUnlock(hMem)  ' destroy pointer, unlock block
    Call apiGlobalFree(hMem)  ' free the allocated memory
End Function

