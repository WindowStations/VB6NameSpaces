VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "Rectangle"
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
'Attribute VB_Name = "Rectangle"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Member0" ,"Size"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type
Private Declare Function apiCopyRect Lib "user32" Alias "CopyRect" (ByRef lpDestRect As RECT, ByRef lpSourceRect As RECT) As Long
Private Declare Function apiEqualRect Lib "user32" Alias "EqualRect" (ByRef lpRect1 As RECT, ByRef lpRect2 As RECT) As Long
Private Declare Function apiGetClientRect Lib "user32" Alias "GetClientRect" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiGetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function apiInflateRect Lib "user32" Alias "InflateRect" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiIntersectRect Lib "user32" Alias "IntersectRect" (ByRef lpDestRect As RECT, ByRef lpSrc1Rect As RECT, ByRef lpSrc2Rect As RECT) As Long
Private Declare Function apiIsRectEmpty Lib "user32" Alias "IsRectEmpty" (ByRef lpRect As RECT) As Long
Private Declare Function apiOffsetRect Lib "user32" Alias "OffsetRect" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiPtInRect Lib "user32" Alias "PtInRect" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function apiSetRect Lib "user32" Alias "SetRect" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function apiSetRectEmpty Lib "user32" Alias "SetRectEmpty" (ByRef lpRect As RECT) As Long
Private Declare Function apiSubtractRect Lib "user32" Alias "SubtractRect" (ByRef lprcDst As RECT, ByRef lprcSrc1 As RECT, ByRef lprcSrc2 As RECT) As Long
Private Declare Function apiUnionRect Lib "user32" Alias "UnionRect" (ByRef lprcDst As RECT, ByRef lprcSrc1 As RECT, ByRef lprcSrc2 As RECT) As Long
Private mvarLocation As New Point
Private mvarSize     As New Size
Private mvarTop      As Long
Private mvarLeft     As Long
Private mvarBottom   As Long
Private mvarRight    As Long
Private mvarHeight   As Long
Private mvarWidth_   As Long
Private mvarX        As Long
Private mvarY        As Long
Friend Property Let y(ByVal vData As Long)
    mvarY = vData
End Property
Friend Property Get y() As Long
    y = mvarY
End Property
Friend Property Let x(ByVal vData As Long)
    mvarX = vData
End Property
Friend Property Get x() As Long
    x = mvarX
End Property
Friend Property Let Width_(ByVal vData As Long)
    mvarWidth_ = vData
End Property
Friend Property Get Width_() As Long
    Width_ = mvarWidth_
End Property
Friend Property Let Height(ByVal vData As Long)
    mvarHeight = vData
End Property
Friend Property Get Height() As Long
    Height = mvarHeight
End Property
Friend Property Set Location(ByVal vData As Point)
    Set mvarLocation = vData
End Property
Friend Property Get Location() As Point
    Set Location = mvarLocation
End Property
Friend Property Set Size(ByVal vData As Size)
    Set mvarSize = vData
End Property
Friend Property Get Size() As Size
    Set Size = mvarSize
End Property
Friend Property Let Top(ByVal vData As Long)
    mvarTop = vData
End Property
Friend Property Get Top() As Long
    Top = mvarTop
End Property
Friend Property Let Left(ByVal vData As Long)
    mvarLeft = vData
End Property
Friend Property Get Left() As Long
    Left = mvarLeft
End Property
Friend Property Let Bottom(ByVal vData As Long)
    mvarBottom = vData
End Property
Friend Property Get Bottom() As Long
    Bottom = mvarBottom
End Property
Friend Property Let Right(ByVal vData As Long)
    mvarRight = vData
End Property
Friend Property Get Right() As Long
    Right = mvarRight
End Property
Private Sub FromRectangle(ByRef rDestination As RECT, ByRef rSource As Rectangle)
    Call apiSetRect(rDestination, rSource.x, rSource.y, rSource.Right, rSource.Bottom)
End Sub
Friend Sub FromRECT(ByRef r1 As RECT, ByRef r2 As Rectangle)
    On Error Resume Next
    Call FromLTRB(r1.Left, r1.Top, r1.Right, r1.Bottom, r2)
    With r2
        .x = r1.Left
        .y = r1.Top
        .Width_ = (r1.Right - r1.Left)
        .Height = (r1.Bottom - r1.Top)
        .Location.x = r1.Left
        .Location.y = r1.Top
        .Size.Width_ = (r1.Right - r1.Left)
        .Size.Height_ = (r1.Bottom - r1.Top)
    End With
End Sub
Friend Sub FromLTRB(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByRef r2 As Rectangle)
    On Error Resume Next
    Dim r As New Rectangle
    With r
        .x = Left
        .y = Top
        .Left = Left
        .Top = Top
        .Location.x = Left
        .Location.y = Top
        .Bottom = Bottom
        .Right = Right
        .Width_ = (Right - Left)
        .Height = (Bottom - Top)
        .Size.Width_ = (Right - Left)
        .Size.Height_ = (Bottom - Top)
    End With
    Set r2 = r
End Sub
Friend Sub Copy(ByRef rDestination As Rectangle, ByRef rSource As Rectangle)
    On Error Resume Next
    Dim rctdest As RECT
    Dim rctsour As RECT
    Call FromRectangle(rctdest, rDestination)
    Call FromRectangle(rctsour, rSource)
    Call apiCopyRect(rctdest, rctsour)
    Call FromRECT(rctdest, rDestination)
End Sub
Friend Sub offset(ByRef pt As Point, ByRef lpRect As Rectangle)
    On Error Resume Next
    Dim r As RECT
    Call FromRectangle(r, lpRect)
    Call apiOffsetRect(r, pt.x, pt.y)
    Call FromRECT(r, lpRect)
End Sub
Friend Sub EmptyRectangle(ByRef lpRect As Rectangle)
    On Error Resume Next
    Dim r   As New Rectangle
    Dim rct As RECT
    Call FromRectangle(rct, lpRect)
    Call apiSetRectEmpty(rct)
    Call FromRECT(rct, r)
    Set lpRect = r
End Sub
Friend Function PoinInRectangle(ByRef p As Point, ByRef r As Rectangle) As Boolean
    Dim r2 As RECT
    Call FromRectangle(r2, r)
    PoinInRectangle = CBool(apiPtInRect(r2, p.x, p.y))
End Function
Friend Sub Inflate(ByRef lpRect As Rectangle, ByVal x As Long, ByVal y As Long)
    On Error Resume Next
    Dim r   As New Rectangle
    Dim rct As RECT
    Call FromRectangle(rct, lpRect)
    Call apiInflateRect(rct, x, y)
    Call FromRECT(rct, r)
    Set lpRect = r
End Sub
Friend Sub Subtract(ByRef dest As Rectangle, ByRef source As Rectangle, ByRef source2 As Rectangle)
    On Error Resume Next
    Dim r        As New Rectangle
    Dim rdest    As RECT
    Dim rsource1 As RECT
    Dim rsource2 As RECT
    Call FromRectangle(rdest, dest)
    Call FromRectangle(rsource1, source)
    Call FromRectangle(rsource2, source2)
    Call apiSubtractRect(rdest, rsource1, rsource2)
    Call FromRECT(rdest, r)
    Set dest = r
End Sub
Friend Sub Union(ByRef dest As Rectangle, ByRef source As Rectangle, ByRef source2 As Rectangle)
    On Error Resume Next
    Dim r        As New Rectangle
    Dim rdest    As RECT
    Dim rsource1 As RECT
    Dim rsource2 As RECT
    Call FromRectangle(rdest, dest)
    Call FromRectangle(rsource1, source)
    Call FromRectangle(rsource2, source2)
    Call apiUnionRect(rdest, rsource1, rsource2)
    Call FromRECT(rdest, r)
    Set dest = r
End Sub
Friend Sub Intersect(ByRef dest As Rectangle, ByRef source As Rectangle, ByRef source2 As Rectangle)
    On Error Resume Next
    Dim r        As New Rectangle
    Dim rdest    As RECT
    Dim rsource1 As RECT
    Dim rsource2 As RECT
    Call FromRectangle(rdest, dest)
    Call FromRectangle(rsource1, source)
    Call FromRectangle(rsource2, source2)
    Call apiIntersectRect(rdest, rsource1, rsource2)
    Call FromRECT(rdest, r)
    Set dest = r
End Sub
Friend Function IntersectsWith(ByRef dest As Rectangle, ByRef source As Rectangle, ByRef source2 As Rectangle) As Boolean
    On Error Resume Next
    Dim r        As New Rectangle
    Dim rdest    As RECT
    Dim rsource1 As RECT
    Dim rsource2 As RECT
    Call FromRectangle(rdest, dest)
    Call FromRectangle(rsource1, source)
    Call FromRectangle(rsource2, source2)
    Call apiIntersectRect(rdest, rsource1, rsource2)
    Call FromRECT(rdest, r)
    IntersectsWith = Not IsEmpty(r)
End Function
Friend Function Equals(ByRef r1 As Rectangle, ByRef r2 As Rectangle) As Boolean
    On Error Resume Next
    Dim rect1 As RECT
    Dim rect2 As RECT
    Call apiSetRect(rect1, r1.x, r1.y, r1.Right, r1.Bottom)
    Call FromRectangle(rect1, r1)
    Call FromRectangle(rect2, r2)
    Equals = CBool(apiEqualRect(rect1, rect2))
End Function
Friend Function IsEmpty(ByRef lpRect As Rectangle) As Boolean
    On Error Resume Next
    Dim rct As RECT
    Call FromRectangle(rct, lpRect)
    IsEmpty = CBool(apiSetRectEmpty(rct))
End Function

