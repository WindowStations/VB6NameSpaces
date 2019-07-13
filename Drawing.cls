VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Drawing"
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
'Attribute VB_Name = "Drawing"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Image"
'Attribute VB_Ext_KEY = "Member1" ,"Icon"
'Attribute VB_Ext_KEY = "Member2" ,"Graphics"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Attribute VB_Ext_KEY = "Member3" ,"Brushes"
'Attribute VB_Ext_KEY = "Member4" ,"Brush"
'Attribute VB_Ext_KEY = "Member5" ,"FontFamily"
'Attribute VB_Ext_KEY = "Member6" ,"Font"
'Attribute VB_Ext_KEY = "Member7" ,"Pens"
'Attribute VB_Ext_KEY = "Member8" ,"Pen"
'Attribute VB_Ext_KEY = "Member9" ,"Color"
'Attribute VB_Ext_KEY = "Member10" ,"Size"
'Attribute VB_Ext_KEY = "Member11" ,"Rectangle"
'Attribute VB_Ext_KEY = "Member12" ,"Point"
'Option Explicit
Private Const BITSPIXEL  As Long = 12       ' Number of adjacent color bits for each pixel.
Private Const PLANES     As Long = 14           ' Number of color planes.
Private Const NUMBRUSHES As Long = 16       ' Number of device-specific brushes.
Private Const NUMPENS    As Long = 18          ' Number of device-specific pens.
Private Const NUMMARKERS As Long = 20       ' Number of markers the device has.
Private Const NUMCOLORS  As Long = 24        ' Number of colors the device supports.
Private Declare Function apiGetDeviceCaps Lib "gdi32" Alias "GetDeviceCaps" (ByVal HDC As Long, ByVal nIndex As Long) As Long
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetDC Lib "user32" Alias "GetDC" (ByVal hwnd As Long) As Long
Private m_lhdc         As Long
Private mvarPoint      As Point
Private mvarRectangle  As Rectangle
Private mvarSize       As Size
Private mvarColor      As Color
Private mvarPen        As Pen
Private mvarPens       As Pens
Private mvarFont       As Font
Private mvarFontFamily As FontFamily
Private mvarBrush      As Brush
Private mvarBrushes    As Brushes
Private mvarGraphics   As Graphics
Private mvarIcon       As Icon
Private mvarImage      As Image
Private Sub Class_Initialize()
    Set mvarPoint = New Point
    Set mvarRectangle = New Rectangle
    Set mvarSize = New Size
    Set mvarColor = New Color
    Set mvarPen = New Pen
    Set mvarPens = New Pens
    Set mvarFont = New Font
    Set mvarFontFamily = New FontFamily
    Set mvarBrush = New Brush
    Set mvarBrushes = New Brushes
    Set mvarGraphics = New Graphics
    Set mvarIcon = New Icon
    Set mvarImage = New Image
End Sub
Private Sub Class_Terminate()
    Set mvarPoint = Nothing
    Set mvarRectangle = Nothing
    Set mvarSize = Nothing
    Set mvarColor = Nothing
    Set mvarPen = Nothing
    Set mvarPens = Nothing
    Set mvarFont = Nothing
    Set mvarFontFamily = Nothing
    Set mvarBrush = Nothing
    Set mvarBrushes = Nothing
    Set mvarGraphics = Nothing
    Set mvarIcon = Nothing
    Set mvarImage = Nothing
End Sub
Friend Property Get Point() As Point
    Set Point = mvarPoint
End Property
Friend Property Get Rectangle() As Rectangle
    Set Rectangle = mvarRectangle
End Property
Friend Property Get Size() As Size
    Set Size = mvarSize
End Property
Friend Property Get Color() As Color
    Set Color = mvarColor
End Property
Friend Property Get Pen() As Pen
    Set Pen = mvarPen
End Property
Friend Property Get Pens() As Pens
    Set Pens = mvarPens
End Property
Friend Property Get Font() As Font
    Set Font = mvarFont
End Property
Friend Property Get FontFamily() As FontFamily
    Set FontFamily = mvarFontFamily
End Property
Friend Property Get Brush() As Brush
    Set Brush = mvarBrush
End Property
Friend Property Get Brushes() As Brushes
    Set Brushes = mvarBrushes
End Property
Friend Property Get Graphics() As Graphics
    Set Graphics = mvarGraphics
End Property
Friend Property Get Icon() As Icon
    Set Icon = mvarIcon
End Property
Friend Property Get Image() As Image
    Set Image = mvarImage
End Property
'
'Public Sub Init(Optional HDC As Long = 0)
'    On Error GoTo hComponentFailure
'    If HDC = 0 Then
'        HDC = apiGetDC(apiGetDesktopWindow)
'    End If
'    If HDC = 0 Then
'        On Error GoTo 0
'       ' Err.Raise eErrSysPalette_InvalidDeviceContextHandle, "VB65" & ".CSysPalette", S_ERR_InvalidDeviceContextHandle
'    End If
'    m_lhdc = HDC
'    Exit Sub
'hComponentFailure:
'  '  Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Sub
'Friend Property Get Brushes() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'      '  Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    Brushes = apiGetDeviceCaps(m_lhdc, NUMBRUSHES)
'    Exit Property
'hComponentFailure:
'   ' Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get ColorDepth() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'      '  Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    ColorDepth = apiGetDeviceCaps(m_lhdc, BITSPIXEL)
'    Exit Property
'hComponentFailure:
'  '  Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get ColorPlanes() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'      '  Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    ColorPlanes = apiGetDeviceCaps(m_lhdc, PLANES)
'    Exit Property
'hComponentFailure:
'   ' Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get Colors() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'        'Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    Colors = apiGetDeviceCaps(m_lhdc, NUMCOLORS)
'    Exit Property
'hComponentFailure:
'   ' Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Let HDC(Value As Long)
'    On Error GoTo hComponentFailure
'    If Value = 0 Then
'        On Error GoTo 0
'       ' Err.Raise eErrSysPalette_InvalidDeviceContextHandle, "VB65" & ".CSysPalette", S_ERR_InvalidDeviceContextHandle
'    End If
'    m_lhdc = Value
'    Exit Property
'hComponentFailure:
'   ' Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get HDC() As Long
'    On Error GoTo hComponentFailure
'    HDC = m_lhdc
'    Exit Property
'hComponentFailure:
'    'Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get Markers() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'      '  Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    Markers = apiGetDeviceCaps(m_lhdc, NUMMARKERS)
'    Exit Property
'hComponentFailure:
'   ' Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property
'Friend Property Get Pens() As Long
'    On Error GoTo hComponentFailure
'    If m_lhdc = 0 Then
'        On Error GoTo 0
'        'Err.Raise eErrSysPalette_NotInitialized, "VB65" & ".CSysPalette", S_ERR_NotInitialized
'    End If
'    Pens = apiGetDeviceCaps(m_lhdc, NUMPENS)
'    Exit Property
'hComponentFailure:
'    'Err.Raise eErrSysPalette_ComponentFailure, "VB65" & ".CSysPalette", S_ERR_ComponentFailure
'End Property

