VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Array"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'Public Type TSAFEARRAYBOUND
'    lElements As Long
'    lLowest As Long
'End Type
'
'Public Type TSAFEARRAY
'    iDims As Integer
'    iFeatures As Integer
'    lElementSize As Long
'    lLocks As Long
'    lData As Long
'    lPointer As Long
'    lVarType As Long
'    lSorted As Long
'    uBounds() As TSAFEARRAYBOUND
'End Type
'
''CONSTANTS
'
'Public Const L_MX As Long = 2147483647
'Public Const L_NG As Long = -1&
'
''VARIABLES
'
'Public PUB_UNICODE As Boolean
'Public m_Comma As Integer
'Private m_Minus As Integer
'
''WINAPI
'
'Private Declare Function GetDesktopWindow Lib "user32" () As Long
'Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
'Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
'Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long
'Private Declare Function SafeArrayCopy Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Long) As Long
'Private Declare Function SafeArrayCopyData Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Long) As Long
'Private Declare Function SafeArrayCreate Lib "oleaut32" (ByVal lType As Integer, ByVal lDims As Long, uBounds As Any) As Long
'Private Declare Function SafeArrayDestroy Lib "oleaut32" (ByVal lArray As Long) As Long
'Private Declare Function SafeArrayGetElement Lib "oleaut32" (ByVal lArray As Long, ByRef lIndices As Long, uValue As Any) As Long
'Private Declare Function SafeArrayPutElement Lib "oleaut32" (ByVal lArray As Long, ByRef lIndices As Long, uValue As Any) As Long
'Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal lArray As Long, uLastBound As TSAFEARRAYBOUND) As Long
'Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal lSource As Long, iTarget As Integer)
'Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal lTarget As Long, ByVal iSource As Integer)
'Private Declare Sub RtlZeroMemory Lib "kernel32" (uDestination As Any, ByVal lLen As Long)
'
'Public Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)
'
''METHODS
'
'Private Sub Main()
'
'    Dim s As String
'    Dim p As Long
'
'    s = Space$(1&)
'    p = StrPtr(s)
'
'    PUB_UNICODE = IsWindowUnicode(GetDesktopWindow)
'
'    If PUB_UNICODE Then
'
'        If GetLocaleInfoW(1024&, 14&, p, 2&) Then m_Comma = AscW(s)
'        If GetLocaleInfoW(1024&, 81&, p, 2&) Then m_Minus = AscW(s)
'
'    Else
'
'        If GetLocaleInfoA(1024&, 14&, s, 2&) Then m_Comma = AscW(s)
'        If GetLocaleInfoA(1024&, 81&, s, 2&) Then m_Minus = AscW(s)
'
'    End If
'
'End Sub
'
'Public Sub ArrayCopy(ByRef SafeArray As TSAFEARRAY, ByRef SourceOrOutput As Variant)
'
'    If VarType(SourceOrOutput) = vbEmpty Then
'
'        VariantInit SourceOrOutput, vbArray + SafeArray.lVarType
'
'        SafeArrayCopy SafeArray.lPointer, VarPtr(SourceOrOutput) + 8&
'
'    Else
'
'        SafeArrayCopyData SourceOrOutput, SafeArray.lPointer
'
'    End If
'
'End Sub
'
'Public Function ArrayCreate(ByRef SafeArray As TSAFEARRAY, ByVal SafeArrayType As VbVarType, ByVal SafeArrayDims As Long, ByRef SafeArrayBounds() As TSAFEARRAYBOUND) As Boolean
'
'    Dim x As Long
'
'    ArrayDestroy SafeArray
'
'    x = SafeArrayCreate(SafeArrayType Xor (vbArray * (((SafeArrayType And vbArray) = vbArray) * L_NG)), SafeArrayDims, SafeArrayBounds(0))
'
'    If x Then
'
'        ArrayPtr SafeArray, x
'
'        ArrayCreate = SafeArray.lPointer
'
'    End If
'
'End Function
'
'Public Function ArrayCreateVector(ByRef SafeArray As TSAFEARRAY, ByVal SafeArrayType As VbVarType, ByVal SafeArrayElements As Long) As Boolean
'
'    Dim b(0) As TSAFEARRAYBOUND
'
'    b(0).lElements = SafeArrayElements
'
'    ArrayCreateVector = ArrayCreate(SafeArray, SafeArrayType, 1&, b)
'
'End Function
'
'Public Sub ArrayDestroy(ByRef SafeArray As TSAFEARRAY)
'
'    SafeArrayDestroy SafeArray.lPointer
'
'    Erase SafeArray.uBounds
'
'    RtlZeroMemory SafeArray, Len(SafeArray)
'
'End Sub
'
'Public Function ArrayElementGet(ByRef SafeArray As TSAFEARRAY, ByRef Indexes As Long, ByRef Value As Variant) As Boolean
'
'    Select Case SafeArray.lVarType
'        Case vbInteger To vbString, vbError, vbBoolean, vbByte: ArrayElementGet = (SafeArrayGetElement(SafeArray.lPointer, Indexes, ByVal VarPtr(Value) + 8&) = 0&)
'        Case vbObject, vbVariant: ArrayElementGet = (SafeArrayGetElement(SafeArray.lPointer, Indexes, Value) = 0&)
'        Case vbDecimal: ArrayElementGet = (SafeArrayGetElement(SafeArray.lPointer, Indexes, ByVal VarPtr(Value)) = 0&)
'    End Select
'
'End Function
'
'Public Function ArrayElementSet(ByRef SafeArray As TSAFEARRAY, ByRef Indexes As Long, ByRef Value As Variant) As Boolean
'
'    On Error Resume Next
'
'    Select Case SafeArray.lVarType
'        Case vbError, vbLong: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CLng(Value)) = 0&)
'        Case vbInteger: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CInt(Value)) = 0&)
'        Case vbDouble: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CDbl(Value)) = 0&)
'        Case vbCurrency: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CCur(Value)) = 0&)
'        Case vbSingle: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CSng(Value)) = 0&)
'        Case vbByte: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CByte(Value)) = 0&)
'        Case vbBoolean: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CBool(Value)) = 0&)
'        Case vbString: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, ByVal StrPtr(Value)) = 0&)
'        Case vbDate: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, CDate(Value)) = 0&)
'        Case vbObject, vbVariant: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, Value) = 0&)
'        Case vbDecimal: ArrayElementSet = (SafeArrayPutElement(SafeArray.lPointer, Indexes, ByVal VarPtr(CDec(Value))) = 0&)
'    End Select
'
'End Function
'
'Public Sub ArrayPtr(ByRef SafeArray As TSAFEARRAY, ByVal SourceArrayPtr As Long, Optional ByVal IsExternal As Boolean)
'
'    If SourceArrayPtr Then
'
'        RtlMoveMemory SafeArray.lVarType, ByVal SourceArrayPtr + (-4& * ((Not IsExternal) * L_NG)), 2& 'get array type
'
'        If IsExternal Then
'
'            RtlMoveMemory SafeArray.lPointer, ByVal SourceArrayPtr + 8&, 4&
'
'            If SafeArray.lVarType And 16384& Then 'if passed by reference (pointer to pointer)
'                RtlMoveMemory SafeArray.lPointer, ByVal SafeArray.lPointer, 4&
'                SafeArray.lVarType = SafeArray.lVarType Xor 16384&
'            End If
'
'            SafeArray.lVarType = SafeArray.lVarType Xor vbArray
'
'        Else
'            SafeArray.lPointer = SourceArrayPtr
'        End If
'
'        RtlMoveMemory SafeArray, ByVal SafeArray.lPointer, 16& 'fill first fixed 16 bytes from pointer
'
'        ReDim SafeArray.uBounds(SafeArray.iDims + L_NG) 'allocate bounds member
'
'        RtlMoveMemory SafeArray.uBounds(0), ByVal SafeArray.lPointer + 16&, SafeArray.iDims * 8& 'get array dimensions info bytes (in descending order) starting from array pointer adress + 16 bytes offset
'
'    End If
'
'End Sub
'
'Public Sub ArrayRedim(ByRef SafeArray As TSAFEARRAY, ByVal NewCount As Long, Optional ByVal HoldSortedFlag As Boolean)
'
'    If Not HoldSortedFlag Then SafeArray.lSorted = 0&
'
'    SafeArray.uBounds(0&).lElements = NewCount
'
'    SafeArrayRedim SafeArray.lPointer, SafeArray.uBounds(0)
'
'End Sub
'
'Public Function ToNumber(ByVal lPointer As Long, ByVal lLength As Long) As String
'
'    Dim b As Integer
'    Dim c As Long
'    Dim i As Long
'    Dim f As Boolean
'    Dim m As Boolean
'    Dim p As Long
'
'    If lLength Then
'
'        ToNumber = Space$(lLength)
'        p = StrPtr(ToNumber)
'
'        For i = 1& To lLength
'
'            GetMem2 lPointer, b
'
'            If b > 47 And b < 58 Then
'
'                PutMem2 p, b
'
'                c = c + 1&
'                p = p + 2&
'
'            ElseIf Not f And b = m_Comma Then
'
'                PutMem2 p, b
'
'                c = c + 1&
'                p = p + 2&
'
'                f = True
'
'            ElseIf c = 0& And Not f And Not m And b = m_Minus Then
'
'                PutMem2 p, b
'
'                c = c + 1&
'                p = p + 2&
'
'                m = True
'
'            End If
'
'            lPointer = lPointer + 2&
'
'        Next i
'
'        If c > 308& Then
'
'            ToNumber = Left$(ToNumber, 308&)
'
'        ElseIf c = 1& Then
'
'            If AscW(ToNumber) = m_Comma Or AscW(ToNumber) = m_Minus Then ToNumber = ChrW$(48&)
'
'        ElseIf c > 0& Then
'
'            ToNumber = Left$(ToNumber, c)
'
'        Else
'
'            ToNumber = ChrW$(48&)
'
'        End If
'
'    Else
'        ToNumber = ChrW$(48&)
'    End If
'
'End Function
'
'Public Sub VariantInit(ByRef uVariable As Variant, ByVal lVariableType As Long)
'
'    RtlMoveMemory uVariable, lVariableType, 2&
'
'End Sub
'
'Public Sub VariantZero(ByRef uVariable As Variant)
'
'    RtlMoveMemory uVariable, 0&, 4&
'
'End Sub
