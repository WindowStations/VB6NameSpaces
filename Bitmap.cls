VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Bitmap"
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
'Attribute VB_Name = "Bitmap"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const E_ERR_BASE              As Long = 17660 + vbObjectError
Private Const S_ERR_NotInitialized    As String = "Object not properly initialized"
Private Const S_ERR_InvalidFileName   As String = "Invalid file name"
Private Const S_ERR_FileNotFound      As String = "File not found"
Private Const S_ERR_UnknownFileFormat As String = "Unknown file format"
Private Const S_ERR_ComponentFailure  As String = "CImageInfo component failure"
Private Const INVALID_HANDLE_VALUE    As Long = -1
Private Const MAX_LENGTH              As Long = 512
Private Const CF_BITMAP               As Long = 2
Private Const IMAGE_BITMAP            As Long = 0
Private Const LR_COPYRETURNORG        As Long = &H4
Private Const LR_CREATEDIBSECTION     As Long = &H2000
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
Private Type BITMAPFILEHEADER
    bfType       As Integer  ' The string "BM" reversed (hex value &H4D42) - 2bytes
    bfSize       As Long     ' The size of the file, measured in bytes - 4bytes
    bfReserved1  As Integer  ' Not used - 2 bytes
    bfReserved2  As Integer  ' Not Used - 2 bytes
    bfOffBits    As Long     ' The start Offset of the bitmap data in the file - 4bytes
End Type
Private Type BITMAPINFOHEADER        ' 40 Bytes
    biSize               As Long     ' It has 40 in it - the size of this structure
    bm_Width             As Long     ' The Width of the bitmap in pixels
    bm_Height            As Long     ' The Height of the bitmap in pixels
    biPlanes             As Integer  ' 1 (DIBs always have one plane)
    biBitCount           As Integer  ' 1 for monochrome; 4 for 16 colors; 8 for 256 colors; 24 for 24bit RGB color.
    biCompression        As Long     ' Compression method - BI_RGB (&H00), BI_RLE8 (&H1), BI_RLE4 (&H2)
    biSizeImage          As Long     ' Size in bytes. May be zero if biCompression is BI_RGB
    biXPixelsPerMeter    As Long     ' Number of horizontal pixels per meter for which this DIB was designed
    biYPixelsPerMeter    As Long     ' Same but for vertical
    biCtlrUsed           As Long     ' Number of entries in color table that as actually used. May be zero
    biCtlrImportant      As Long     ' Number of entries in color table that are important. Zero for all.
End Type
Private Type BITMAPINFO
    bmiHeader    As BITMAPINFOHEADER
    bmiColor     As String * 64
End Type
Private Type FILETIME
    dwLowDateTime    As Long
    dwHighDateTime   As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes     As Long
    ftCreationTime       As FILETIME
    ftLastAccessTime     As FILETIME
    ftLastWriteTime      As FILETIME
    nFileSizeHigh        As Long
    nFileSizeLow         As Long
    dwReserved0          As Long
    dwReserved1          As Long
    cFileName            As String * MAX_LENGTH
    cAlternate           As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function apiCopyImage Lib "user32" Alias "CopyImage" (ByVal Handle As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private m_sFileName      As String
Private m_itImageType    As EImageType
Private m_lHeight        As Long
Private m_lWidth         As Long
Private m_bInfoExtracted As Boolean
Public Enum EErrImageInfo
    eErrImageInfo_NotInitialized = E_ERR_BASE + 1
    eErrImageInfo_InvalidFileName
    eErrImageInfo_FileNotFound
    eErrImageInfo_UnknownFileFormat
    eErrImageInfo_ComponentFailure
End Enum
Public Enum EImageType
    eImageType_Unknown = -1
    eImageType_Bitmap = 1
    eImageType_Gif = 2
    eImageType_Jpeg = 3
End Enum
Private Sub Class_Initialize()
    On Error GoTo hComponentFailure
    m_sFileName = ""
    m_itImageType = eImageType_Unknown
    m_lWidth = 0
    m_lHeight = 0
    Exit Sub
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Sub
Public Sub Init(FileName As String)
    On Error GoTo hComponentFailure
    If Len(FileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_InvalidFileName, "VB65" & ".CImageInfo", S_ERR_InvalidFileName
    End If
    If Not FileExists(FileName) Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_FileNotFound, "VB65" & ".CImageInfo", S_ERR_FileNotFound
    End If
    m_sFileName = FileName
    m_bInfoExtracted = False
    Exit Sub
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Sub
Friend Property Let FileName(Value As String)
    On Error GoTo hComponentFailure
    If Len(Value) = 0 Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_InvalidFileName, "VB65" & ".CImageInfo", S_ERR_InvalidFileName
    End If
    If Not FileExists(Value) Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_FileNotFound, "VB65" & ".CImageInfo", S_ERR_FileNotFound
    End If
    m_sFileName = Value
    m_bInfoExtracted = False
    Exit Property
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Property
Friend Property Get FileName() As String
    On Error GoTo hComponentFailure
    FileName = m_sFileName
    Exit Property
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Property
Friend Property Get Height() As Single
    On Error GoTo hComponentFailure
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_NotInitialized, "VB65" & ".CImageInfo", S_ERR_NotInitialized
    End If
    If Not m_bInfoExtracted Then
        Call ExtractInfo
    End If
    Height = m_lHeight
    Exit Property
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Property
Friend Property Get Width() As Single
    On Error GoTo hComponentFailure
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_NotInitialized, "VB65" & ".CImageInfo", S_ERR_NotInitialized
    End If
    If Not m_bInfoExtracted Then
        Call ExtractInfo
    End If
    Width = m_lWidth
    Exit Property
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Property
Friend Property Get ImageType() As EImageType
    On Error GoTo hComponentFailure
    If Len(m_sFileName) = 0 Then
        On Error GoTo 0
        Err.Raise eErrImageInfo_NotInitialized, "VB65" & ".CImageInfo", S_ERR_NotInitialized
    End If
    If Not m_bInfoExtracted Then
        Call ExtractInfo
    End If
    ImageType = m_itImageType
    Exit Property
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Property
Private Sub ExtractInfo()
    On Error GoTo hComponentFailure
    Dim fileHandle As Integer
    Dim buffer     As String
    Dim Header     As BITMAPFILEHEADER
    Dim Info       As BITMAPINFO
    Dim SizeMark   As String
    Dim Marker     As Long
    On Error GoTo hComponentFailure
    m_itImageType = eImageType_Unknown
    fileHandle = FreeFile
    Open m_sFileName For Binary As #fileHandle
    Seek #fileHandle, 1
    Get #fileHandle, , Header
    If Header.bfType = &H4D42 Then m_itImageType = eImageType_Bitmap
    If m_itImageType = eImageType_Unknown Then
        Seek #fileHandle, 1
        buffer = Space$(6)
        Get #fileHandle, , buffer
        If Left$(buffer, 3) = "GIF" Then m_itImageType = eImageType_Gif
    End If
    If m_itImageType = eImageType_Unknown Then
        Seek #fileHandle, 1
        buffer = Space$(100)
        Get fileHandle, , buffer
        If InStr(buffer, "JFIF" + Chr$(0)) <> 0 Then m_itImageType = eImageType_Jpeg
    End If
    If m_itImageType = eImageType_Unknown Then
        Close #fileHandle
        On Error GoTo 0
        Err.Raise eErrImageInfo_UnknownFileFormat, "VB65" & ".CImageInfo", S_ERR_UnknownFileFormat
    End If
    Select Case m_itImageType
        Case eImageType_Bitmap
            Seek #fileHandle, 1
            Get #fileHandle, , Header
            Get #fileHandle, , Info
            m_lHeight = Info.bmiHeader.bm_Height
            m_lWidth = Info.bmiHeader.bm_Width
        Case eImageType_Gif
            Seek #fileHandle, 7
            buffer = Space$(2)
            Get #fileHandle, , buffer
            m_lWidth = Asc(Left$(buffer, 1)) + 256! * Asc(Right$(buffer, 1))
            Seek #fileHandle, 9
            buffer = Space$(2)
            Get #fileHandle, , buffer
            m_lHeight = Asc(Left$(buffer, 1)) + 256! * Asc(Right$(buffer, 1))
        Case eImageType_Jpeg
            SizeMark = Chr$(255) + Chr$(192)
            Seek #1, 1
            buffer = Space$(1000)
            Get #fileHandle, , buffer
            Marker = InStr(buffer, SizeMark)
            If Marker <> 0 Then
                Seek #1, Marker + 5
                buffer = Space$(2)
                Get #fileHandle, , buffer
                m_lHeight = Asc(Right$(buffer, 1)) + 256! * Asc(Left$(buffer, 1))
                Seek #1, Marker + 7
                buffer = Space$(2)
                Get #fileHandle, , buffer
                m_lWidth = Asc(Right$(buffer, 1)) + 256! * Asc(Left$(buffer, 1))
            End If
    End Select
    Close #fileHandle
    m_bInfoExtracted = True
    Exit Sub
hComponentFailure:
    If fileHandle Then Close #fileHandle
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Sub
Private Function FileExists(FileName As String) As Boolean
    On Error GoTo hComponentFailure
    Dim wfd   As WIN32_FIND_DATA
    Dim hFile As Long
    FileExists = False
    hFile = FindFirstFile(FileName, wfd)
    If hFile <> INVALID_HANDLE_VALUE Then
        FileExists = True
        If FindClose(hFile) = 0 Then
            On Error GoTo 0
            Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
        End If
    End If
    Exit Function
hComponentFailure:
    Err.Raise eErrImageInfo_ComponentFailure, "VB65" & ".CImageInfo", S_ERR_ComponentFailure
End Function
Friend Sub SetPicture(ByVal hPic As Long)
    Dim hMem As Long
    hMem = apiCopyImage(hPic, IMAGE_BITMAP, 0, 0, LR_COPYRETURNORG Or LR_CREATEDIBSECTION)
    'Call apiSetClipboardData(CF_BITMAP, hMem)
End Sub
Friend Function GetPicture() As Object
    '    Dim hMem          As Long
    '    Dim r             As Long
    '    Dim IID_IDispatch As GUID
    '    Dim OPic          As OLEPIC
    '    Dim IObj          As Object
    '    If apiGetOpenClipboardWindow() = NO_CB_OPENED Then
    '        If apiOpenClipboard(0) <> NO_CB_OPEN_ERROR Then
    '            If apiIsClipboardFormatAvailable(CF_BITMAP) <> NO_CB_FORMAT_AVAILABLE Then
    '                hMem = apiGetClipboardData(CF_BITMAP)
    '                With IID_IDispatch
    '                    .Data1 = &H20400
    '                    .Data4(0) = &HC0
    '                    .Data4(7) = &H46
    '                End With
    '                With OPic
    '                    .Size = Len(OPic)        'Lunghezza della struttura.
    '                    .tType = vbPicTypeBitmap 'Tipo dell'immagine (bitmap).
    '                    .hBmp = hMem             'L'handle dell'immagine.
    '                    '                    .hPal = hMem + 40       ' 40 Len BITMAP structure before palette
    '                End With
    '                r = apiOleCreatePictureIndirect(OPic, IID_IDispatch, 0, IObj)
    '                Call apiCloseClipboard
    '                Set GetPicture = IObj
    '                Set IObj = Nothing
    '                Exit Function
    '            Else
    '                Call apiCloseClipboard
    '                Messagebox.Show CB_NO_BITMAP_FORMAT_AVAILABLE
    '            End If
    '        Else
    '            Messagebox.Show CB_OPEN_ERROR
    '        End If
    '    Else
    '        Messagebox.Show CB_ALREADY_OPEN
    '    End If
    '    Set IObj = Nothing
    '    Set GetPicture = Nothing
End Function

