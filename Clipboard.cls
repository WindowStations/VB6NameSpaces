VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clipboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const IMAGE_BITMAP        As Long = 0
Private Const LR_COPYRETURNORG    As Long = &H4
Private Const LR_CREATEDIBSECTION As Long = &H2000
Private Const GMEM_DDESHARE = &H2000
Private Const GMEM_MOVEABLE = &H2
Private Const CF_TEXT = 1
Private Const CF_BITMAP = 2
Private Const NO_CB_OPEN_ERROR = 0
Private Const NO_CB_OPENED = 0
Private Const NO_CB_FORMAT_AVAILABLE = 0
Private Const NO_CB_VIWER = 0
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
Private Declare Sub apiMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef source As Any, ByVal length As Long)
Private Enum eCBERRORMSG
    CB_OPEN_ERROR = 0
    CB_NO_BITMAP_FORMAT_AVAILABLE = 1
    CB_NO_TEXT_FORMAT_AVAILABLE = 2
    CB_ALREADY_OPEN = 3
End Enum
Private sME(3) As String
Private Sub CBEnumerateFormats()
    Dim fmt  As Long
    Dim nfmt As Long
    Dim s    As String * 256
    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
            fmt = apiEnumClipboardFormats(&H0)
            Do
                nfmt = apiEnumClipboardFormats(fmt)
                If nfmt = NO_CB_OPEN_ERROR Then
                    Exit Do
                Else
                    fmt = nfmt
                    s = String(256, 0)
                    Call apiGetClipboardFormatName(fmt, ByVal s, 256)
                    'List1.AddItem CStr(nfmt) + " = " + s + vbCrLf
                End If
            Loop
            Call apiCloseClipboard
            Exit Sub
        Else
            Messagebox.Show CB_OPEN_ERROR
        End If
    Else
        Messagebox.Show CB_ALREADY_OPEN
    End If
End Sub
Friend Sub SetText(ByVal Text As String)
    Dim hMem       As Long
    Dim hPtr       As Long
    Dim lLenBuffer As Long
    Dim s          As String
    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
            Call apiEmptyClipboard
            lLenBuffer = Len(Text) + 1
            s = String(lLenBuffer, 0)
            Mid(s, 1, lLenBuffer - 1) = Text
            hMem = apiGlobalAlloc(GMEM_MOVEABLE Or GMEM_DDESHARE, lLenBuffer)
            hPtr = apiGlobalLock(hMem)
            Call apiMoveMemory(ByVal hPtr, ByVal s, lLenBuffer)
            Call apiGlobalUnlock(hMem)
            Call apiSetClipboardData(CF_TEXT, hMem)
            Call apiCloseClipboard
        Else
            Messagebox.Show CB_OPEN_ERROR
        End If
    Else
        Messagebox.Show CB_ALREADY_OPEN
    End If
    '    GlobalUnlock hMem ' Never leave hMem Locked with SetClipboardData
    '    GlobalFree hMem   ' Never Free hMem with SetClipboardData
    ' Uses GMEM_MOVEABLE Or GMEM_DDESHARE
End Sub
Friend Function GetText() As String
    Dim hMem       As Long
    Dim lLenBuffer As Long
    Dim s          As String
    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
            If apiIsClipboardFormatAvailable(CF_TEXT) <> NO_CB_FORMAT_AVAILABLE Then
                hMem = apiGetClipboardData(CF_TEXT)
                lLenBuffer = apiGlobalSize(hMem)
                s = String(lLenBuffer, 0)
                Call apiMoveMemory(ByVal s, ByVal hMem, lLenBuffer)
                Call apiCloseClipboard
                If s = "" Then: Exit Function
                GetText = Left(s, InStr(s, Chr(0)) - 1)
                Exit Function
            Else
                Call apiCloseClipboard
                Messagebox.Show CB_NO_TEXT_FORMAT_AVAILABLE
            End If
        Else
            Messagebox.Show CB_OPEN_ERROR
        End If
    Else
        Messagebox.Show CB_ALREADY_OPEN
    End If
End Function
Friend Sub SetPicture(ByVal hPic As Long)
    Dim hMem As Long
    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
            Call apiEmptyClipboard
            hMem = apiCopyImage(hPic, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG Or LR_CREATEDIBSECTION)
            Call apiSetClipboardData(CF_BITMAP, hMem)
            Call apiCloseClipboard
        Else
            Messagebox.Show CB_OPEN_ERROR
        End If
    Else
        Messagebox.Show CB_ALREADY_OPEN
    End If
End Sub
Friend Function GetPicture() As Object
    Dim hMem          As Long
    Dim r             As Long
    Dim IID_IDispatch As GUID
    Dim OPic          As OLEPIC
    Dim IObj          As Object
    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
            If apiIsClipboardFormatAvailable(CF_BITMAP) <> NO_CB_FORMAT_AVAILABLE Then
                hMem = apiGetClipboardData(CF_BITMAP)
                With IID_IDispatch
                    .Data1 = &H20400
                    .Data4(0) = &HC0
                    .Data4(7) = &H46
                End With
                With OPic
                    .Size = Len(OPic)        'Lunghezza della struttura.
                    .tType = vbPicTypeBitmap 'Tipo dell'immagine (bitmap).
                    .hBmp = hMem             'L'handle dell'immagine.
                    '                    .hPal = hMem + 40       ' 40 Len BITMAP structure before palette
                End With
                r = apiOleCreatePictureIndirect(OPic, IID_IDispatch, 0, IObj)
                Call apiCloseClipboard
                Set GetPicture = IObj
                Set IObj = Nothing
                Exit Function
            Else
                Call apiCloseClipboard
                Messagebox.Show CB_NO_BITMAP_FORMAT_AVAILABLE
            End If
        Else
            Messagebox.Show CB_OPEN_ERROR
        End If
    Else
        Messagebox.Show CB_ALREADY_OPEN
    End If
    Set IObj = Nothing
    Set GetPicture = Nothing
End Function
Friend Function Clear()
    Call apiEmptyClipboard  ' Clipboard.Clear
End Function
'Private Enum ClipBoardFormat
'    CF_ACCEPT = &H0
'    CF_ANSIONLY = &H400&
'    CF_APPLY = &H200&
'    CF_BITMAP = 2
'    CF_BOTH = (CF_SCREENFONTS Or CF_PRINTERFONTS)
'    CF_BOTTOMUP_DIB = CF_DIB
'    CF_CONVERTONLY = &H100&
'    CF_DEFER = &H2
'    CF_DIB = 8
'    CF_DIBV5 = 17
'    CF_DIF = 5
'    CF_DISABLEACTIVATEAS = &H40&
'    CF_DISABLEDISPLAYASICON = &H20&
'    CF_DSPBITMAP = &H82
'    CF_DSPENHMETAFILE = &H8E
'    CF_DSPMETAFILEPICT = &H83
'    CF_DSPTEXT = &H81
'    CF_EFFECTS = &H100&
'    CF_ENABLEHOOK = &H8&
'    CF_ENABLETEMPLATE = &H10&
'    CF_ENABLETEMPLATEHANDLE = &H20&
'    CF_ENHMETAFILE = 14
'    CF_FIXEDPITCHONLY = &H4000&
'    CF_FORCEFONTEXIST = &H10000
'    CF_GDIOBJFIRST = &H300
'    CF_GDIOBJLAST = &H3FF
'    CF_HDROP = 15
'    CF_HIDECHANGEICON = &H80&
'    CF_INITTOLOGFONTSTRUCT = &H40&
'    CF_JPEG = 19
'    CF_LIMITSIZE = &H2000&
'    CF_LOCALE = 16
'    CF_MAX = 17
'    CF_METAFILEPICT = 3
'    CF_MULTI_TIFF = 22
'    CF_NOFACESEL = &H80000
'    CF_NOOEMFONTS = CF_NOVECTORFONTS
'    CF_NOSCRIPTSEL = &H800000
'    CF_NOSIMULATIONS = &H1000&
'    CF_NOSIZESEL = &H200000
'    CF_NOSTYLESEL = &H100000
'    CF_NOVECTORFONTS = &H800&
'    CF_NOVERTFONTS = &H1000000
'    CF_NULL = 0
'    CF_OEMTEXT = 7
'    CF_OWNERDISPLAY = &H80
'    CF_PALETTE = 9
'    CF_PENDATA = 10
'    CF_PRINTERFONTS = &H2
'    CF_PRIVATEFIRST = &H200
'    CF_PRIVATELAST = &H2FF
'    CF_REJECT = &H1
'    CF_RETEXTOBJ = ("RichEdit Text and Objects")
'    CF_RIFF = 11
'    CF_RTF = ("Rich Text Format")
'    CF_RTFNOOBJS = ("Rich Text Format Without Objects")
'    CF_SCALABLEONLY = &H20000
'    CF_SCREENFONTS = &H1
'    CF_SCRIPTSONLY = CF_ANSIONLY
'    CF_SELECTACTIVATEAS = &H10&
'    CF_SELECTCONVERTTO = &H8&
'    CF_SELECTSCRIPT = &H400000
'    CF_SETACTIVATEDEFAULT = &H4&
'    CF_SETCONVERTDEFAULT = &H2&
'    CF_SHOWHELP = &H4&
'    CF_SHOWHELPBUTTON = &H1&
'    CF_SYLK = 4
'    CF_TEXT = 1
'    CF_TIFF = 6
'    CF_TOPDOWN_DIB = 20
'    CF_TTONLY = &H40000
'    CF_UNICODETEXT = 13
'    CF_USESTYLE = &H80&
'    CF_WAVE = 12
'    CF_WYSIWYG = &H8000
'End Enum
'Private Sub Command1_Click()
'    CBSetText "This is a ClipBoard Test!!!"
'End Sub
'Private Sub Command2_Click()
'    List1.AddItem CBGetText()
'    List1.ListIndex = List1.ListCount - 1
'End Sub
'Private Sub Command3_Click()
'    If GetOpenClipboardWindow() = NO_CB_OPENED Then
'        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
'            EmptyClipboard
'            CloseClipboard
'        End If
'    End If
'End Sub
'Private Sub Command4_Click()
'    List1.Clear
'End Sub
'Private Sub Command5_Click()
'    If GetOpenClipboardWindow() = NO_CB_OPENED Then
'        If OpenClipboard(Me.hwnd) <> NO_CB_OPEN_ERROR Then
'            EmptyClipboard
'            CloseClipboard
'        End If
'    End If
'End Sub
'Private Sub Command6_Click()
'    Set Image2.Picture = CBGetPicture()
'    Image2.Refresh
'End Sub
'Private Sub Command7_Click()
'    CBSetPicture Image1.Picture.handle
'End Sub
'Private Sub Command8_Click()
'    Set Image2.Picture = Nothing
'    Image2.Refresh
'End Sub
'Private Sub Form_Load()
'    '    Debug.Print RegisterClipboardFormat("GioRock Clipboard Format" + Chr$(0))
'    '    CloseClipboard
'    '    CBEnumerateFormats
'    If GetOpenClipboardWindow() <> NO_CB_OPENED Then
'        CloseClipboard
'        SetClipboardViewer Me.hwnd
'    End If
'    sME(0) = "Clipboard open error!!!"
'    sME(1) = "Not Clipboard BITMAP format available!!!"
'    sME(2) = "Not Clipboard TEXT format available!!!"
'    sME(3) = "Clipboard already opened by other application!!!"
'End Sub
'Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    If GetOpenClipboardWindow() = Me.hwnd Then
'        '        EmptyClipboard
'        CloseClipboard
'    End If
'End Sub
'Private Sub Form_Unload(Cancel As Integer)
'    End
'    Set Form1 = Nothing
'End Sub

