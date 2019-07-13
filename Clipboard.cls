VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Clipboard"
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
'Attribute VB_Name = "Clipboard"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const IMAGE_BITMAP            As Long = 0
Private Const LR_COPYRETURNORG        As Long = &H4
Private Const LR_CREATEDIBSECTION     As Long = &H2000
Private Const GMEM_DDESHARE           As Long = &H2000
Private Const GMEM_MOVEABLE           As Long = &H2
Private Const NO_CB_OPEN_ERROR        As Long = 0
Private Const NO_CB_OPENED            As Long = 0
Private Const NO_CB_FORMAT_AVAILABLE  As Long = 0
Private Const NO_CB_VIWER             As Long = 0
Private Const CF_ACCEPT               As Long = &H0
Private Const CF_ANSIONLY             As Long = &H400
Private Const CF_APPLY                As Long = &H200
Private Const CF_BITMAP               As Long = 2
Private Const CF_BOTH                 As Long = (&H1 Or &H2)
Private Const CF_BOTTOMUP_DIB         As Long = 8
Private Const CF_CONVERTONLY          As Long = &H100
Private Const CF_DEFER                As Long = &H2
Private Const CF_DIB                  As Long = 8
Private Const CF_DIBV5                As Long = 17
Private Const CF_DIF                  As Long = 5
Private Const CF_DISABLEACTIVATEAS    As Long = &H40
Private Const CF_DISABLEDISPLAYASICON As Long = &H20
Private Const CF_DSPBITMAP            As Long = &H82
Private Const CF_DSPENHMETAFILE       As Long = &H8E
Private Const CF_DSPMETAFILEPICT      As Long = &H83
Private Const CF_DSPTEXT              As Long = &H81
Private Const CF_EFFECTS              As Long = &H100
Private Const CF_ENABLEHOOK           As Long = &H8
Private Const CF_ENABLETEMPLATE       As Long = &H10
Private Const CF_ENABLETEMPLATEHANDLE As Long = &H20
Private Const CF_ENHMETAFILE          As Long = 14
Private Const CF_FIXEDPITCHONLY       As Long = &H4000
Private Const CF_FORCEFONTEXIST       As Long = &H10000
Private Const CF_GDIOBJFIRST          As Long = &H300
Private Const CF_GDIOBJLAST           As Long = &H3FF
Private Const CF_HDROP                As Long = 15
Private Const CF_HIDECHANGEICON       As Long = &H80
Private Const CF_INITTOLOGFONTSTRUCT  As Long = &H40
Private Const CF_JPEG                 As Long = 19
Private Const CF_LIMITSIZE            As Long = &H2000
Private Const CF_LOCALE               As Long = 16
Private Const CF_MAX                  As Long = 17
Private Const CF_METAFILEPICT         As Long = 3
Private Const CF_MULTI_TIFF           As Long = 22
Private Const CF_NOFACESEL            As Long = &H80000
Private Const CF_NOOEMFONTS           As Long = &H800
Private Const CF_NOSCRIPTSEL          As Long = &H800000
Private Const CF_NOSIMULATIONS        As Long = &H1000
Private Const CF_NOSIZESEL            As Long = &H200000
Private Const CF_NOSTYLESEL           As Long = &H100000
Private Const CF_NOVECTORFONTS        As Long = &H800
Private Const CF_NOVERTFONTS          As Long = &H1000000
Private Const CF_NULL                 As Long = 0
Private Const CF_OEMTEXT              As Long = 7
Private Const CF_OWNERDISPLAY         As Long = &H80
Private Const CF_PALETTE              As Long = 9
Private Const CF_PENDATA              As Long = 10
Private Const CF_PRINTERFONTS         As Long = &H2
Private Const CF_PRIVATEFIRST         As Long = &H200
Private Const CF_PRIVATELAST          As Long = &H2FF
Private Const CF_REJECT               As Long = &H1
Private Const CF_RETEXTOBJ            As String = "RichEdit Text and Objects"
Private Const CF_RIFF                 As Long = 11
Private Const CF_RTF                  As String = "Rich Text Format"
Private Const CF_RTFNOOBJS            As String = "Rich Text Format Without Objects"
Private Const CF_SCALABLEONLY         As Long = &H20000
Private Const CF_SCREENFONTS          As Long = &H1
Private Const CF_SCRIPTSONLY          As Long = CF_ANSIONLY
Private Const CF_SELECTACTIVATEAS     As Long = &H10
Private Const CF_SELECTCONVERTTO      As Long = &H8
Private Const CF_SELECTSCRIPT         As Long = &H400000
Private Const CF_SETACTIVATEDEFAULT   As Long = &H4
Private Const CF_SETCONVERTDEFAULT    As Long = &H2
Private Const CF_SHOWHELP             As Long = &H4
Private Const CF_SHOWHELPBUTTON       As Long = &H1
Private Const CF_SYLK                 As Long = 4
Private Const CF_TEXT                 As Long = 1
Private Const CF_TIFF                 As Long = 6
Private Const CF_TOPDOWN_DIB          As Long = 20
Private Const CF_TTONLY               As Long = &H40000
Private Const CF_UNICODETEXT          As Long = 13
Private Const CF_USESTYLE             As Long = &H80
Private Const CF_WAVE                 As Long = 12
Private Const CF_WYSIWYG              As Long = &H8000
Private Type OLEPIC
   Size As Long
   tType As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type
Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type
Private Declare Function apiGetOpenClipboardWindow Lib "user32" Alias "GetOpenClipboardWindow" () As Long
Private Declare Function apiOpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hwnd As Long) As Long
Private Declare Function apiCloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare Function apiEmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare Function apiSetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function apiGetClipboardData Lib "user32" Alias "GetClipboardData" (ByVal wFormat As Long) As Long
Private Declare Function apiEnumClipboardFormats Lib "user32" Alias "EnumClipboardFormats" (ByVal wFormat As Long) As Long
Private Declare Function apiGetClipboardFormatName Lib "user32" Alias "GetClipboardFormatNameA" (ByVal wFormat As Long, ByRef lpString As String, ByVal nMaxCount As Long) As Long
Private Declare Function apiRegisterClipboardFormat Lib "user32" Alias "RegisterClipboardFormatA" (ByVal lpString As String) As Long
Private Declare Function apiCountClipboardFormats Lib "user32" Alias "CountClipboardFormats" () As Long
Private Declare Function apiGetClipboardOwner Lib "user32" Alias "GetClipboardOwner" () As Long
Private Declare Function apiSetClipboardViewer Lib "user32" Alias "SetClipboardViewer" (ByVal hwnd As Long) As Long
Private Declare Function apiChangeClipboardChain Lib "user32" Alias "ChangeClipboardChain" (ByVal hwnd As Long, ByVal hWndNext As Long) As Long
Private Declare Function apiGetClipboardViewer Lib "user32" Alias "GetClipboardViewer" () As Long
Private Declare Function apiGetPriorityClipboardFormat Lib "user32" Alias "GetPriorityClipboardFormat" (ByRef lpPriorityList As Long, ByVal nCount As Long) As Long
Private Declare Function apiIsClipboardFormatAvailable Lib "user32" Alias "IsClipboardFormatAvailable" (ByVal wFormat As Long) As Long
Private Declare Function apiCopyImage Lib "user32" Alias "CopyImage" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function apiOleCreatePictureIndirect Lib "OlePro32" Alias "OleCreatePictureIndirect" (ByRef PicDesc As OLEPIC, ByRef RefIID As GUID, ByVal fPictureOwnsHandle As Long, ByRef IPic As Any) As Long
Private Declare Function apiGlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function apiGlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As Long) As Long
Private Declare Function apiGlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As Long) As Long
Private Declare Function apiGlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As Long) As Long
Private Declare Function apiGlobalSize Lib "kernel32" Alias "GlobalSize" (ByVal hMem As Long) As Long
Private Declare Sub apiMoveMemoryStringLong Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As String, ByRef source As Long, ByVal Length As Long)
Private Declare Sub apiMoveMemoryLongString Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Long, ByRef source As String, ByVal Length As Long)
Private Const vbPicTypeBitmap As Long = 1
Private Sub EnumClipboardFormats()
   On Error Resume Next
   Dim cfs  As Long
   Dim ncfs As Long
   Dim s    As String * 260
   If apiGetOpenClipboardWindow() <> NO_CB_OPENED Then Exit Sub
   If apiOpenClipboard(0) = NO_CB_OPEN_ERROR Then Exit Sub
   cfs = apiEnumClipboardFormats(&H0)
   Do
      ncfs = apiEnumClipboardFormats(cfs)
      If ncfs = NO_CB_OPEN_ERROR Then Exit Do
      cfs = ncfs
      s = String(260, 0)
      Call apiGetClipboardFormatName(cfs, ByVal s, 260)
   Loop
   Call apiCloseClipboard
End Sub
Friend Sub SetText(ByVal Text As String)
   On Error Resume Next
   If apiGetOpenClipboardWindow() <> NO_CB_OPENED Then Exit Sub
   If apiOpenClipboard(0) = NO_CB_OPEN_ERROR Then Exit Sub
   Dim ga As Long
   Dim hPtr As Long
   Dim buff As Long
   Dim s    As String
   Call apiEmptyClipboard
   buff = Len(Text) + 1
   s = String(buff, 0)
   Mid(s, 1, buff - 1) = Text
   ga = apiGlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, buff)
   hPtr = apiGlobalLock(ga)
   Call apiMoveMemoryLongString(hPtr, s, buff)
   Call apiGlobalUnlock(ga)
   Call apiSetClipboardData(CF_TEXT, ga)
   Call apiCloseClipboard
End Sub
Friend Function GetText() As String
   On Error Resume Next
   If apiGetOpenClipboardWindow() <> NO_CB_OPENED Then Exit Function
   If apiOpenClipboard(0) = NO_CB_OPEN_ERROR Then Exit Function
   If apiIsClipboardFormatAvailable(CF_TEXT) = NO_CB_FORMAT_AVAILABLE Then Call apiCloseClipboard: Exit Function
   Dim ga As Long
   Dim buff As Long
   Dim s    As String
   ga = apiGetClipboardData(CF_TEXT)
   buff = apiGlobalSize(ga)
   s = String(buff, 0)
   Call apiMoveMemoryStringLong(s, ga, buff)
   Call apiCloseClipboard
   If s = "" Then GetText = "": Exit Function
   GetText = Left(s, InStr(s, Chr(0)) - 1)
End Function
Friend Sub SetPicture(ByVal hPic As Long)
   On Error Resume Next
   If apiGetOpenClipboardWindow() <> NO_CB_OPENED Then Exit Sub
   If apiOpenClipboard(0) = NO_CB_OPEN_ERROR Then Exit Sub
   Dim ga As Long
   Call apiEmptyClipboard
   ga = apiCopyImage(hPic, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG Or LR_CREATEDIBSECTION)
   Call apiSetClipboardData(CF_BITMAP, ga)
   Call apiCloseClipboard
End Sub
Friend Function GetPicture() As Object
   On Error Resume Next
   If apiGetOpenClipboardWindow() <> NO_CB_OPENED Then Exit Function
   If apiOpenClipboard(0) = NO_CB_OPEN_ERROR Then Exit Function
   If apiIsClipboardFormatAvailable(CF_BITMAP) = NO_CB_FORMAT_AVAILABLE Then Call apiCloseClipboard: Exit Function
   Dim ga          As Long
   Dim r             As Long
   Dim IID_IDispatch As GUID
   Dim OPic          As OLEPIC
   Dim IObj          As Object
   ga = apiGetClipboardData(CF_BITMAP)
   With IID_IDispatch
      .Data1 = &H20400
      .Data4(0) = &HC0
      .Data4(7) = &H46
   End With
   With OPic
      .Size = Len(OPic)
      .tType = vbPicTypeBitmap
      
      .hBmp = ga
   End With
   r = apiOleCreatePictureIndirect(OPic, IID_IDispatch, 0, IObj)
   Call apiCloseClipboard
   Set GetPicture = IObj
   Set IObj = Nothing
End Function
Friend Function Clear()
   On Error Resume Next
   Call apiEmptyClipboard
End Function

