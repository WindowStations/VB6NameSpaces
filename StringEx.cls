VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "StringEx"
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
'Attribute VB_Name = "StringEx"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
'INTERFACES
Implements TypesInterface
'CONSTANTS
Private Const L_MX                 As Long = 2147483647
Private Const L_NG                 As Long = -1&
Private Const CONST_CRYPT_PROVIDER As String = "Microsoft Base Cryptographic Provider v1.0"
'VARIABLES
Private p_CryptContext             As Long
Private p_CryptHash                As Long
Private p_CryptKey                 As Long
Private p_Value                    As String
Private p_ValueLen                 As Long
Private p_ValueLenReal             As Long
Private p_ValuePtr                 As Long
'WINAPI
Private Type TSAFEARRAYBOUND
    lElements As Long
    lLowest As Long
End Type
Private Type TSAFEARRAY
    iDims As Integer
    iFeatures As Integer
    lElementSize As Long
    lLocks As Long
    lData As Long
    lPointer As Long
    lVarType As Long
    lSorted As Long
    uBounds() As TSAFEARRAYBOUND
End Type
Private Declare Function CharLowerBuffA Lib "user32" (ByVal sStr As String, ByVal lLength As Long) As Long
Private Declare Function CharLowerBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function CharUpperBuffA Lib "user32" (ByVal sStr As String, ByVal lLength As Long) As Long
Private Declare Function CharUpperBuffW Lib "user32" (ByVal lStr As Long, ByVal lLength As Long) As Long
Private Declare Function CompareStringA Lib "kernel32" (ByVal lLocale As Long, ByVal lFlags As Long, ByVal sString1 As String, ByVal lCount1 As Long, ByVal sString2 As String, ByVal lCount2 As Long) As Long
Private Declare Function CompareStringW Lib "kernel32" (ByVal lLocale As Long, ByVal lFlags As Long, ByVal lString1 As Long, ByVal lCount1 As Long, ByVal lString2 As Long, ByVal lCount2 As Long) As Long
Private Declare Function CryptAcquireContextA Lib "advapi32" (lProvider As Long, ByVal sContainer As String, ByVal sProvider As String, ByVal lType As Long, ByVal lFlags As Long) As Long
Private Declare Function CryptAcquireContextW Lib "advapi32" (lProvider As Long, ByVal lContainer As Long, ByVal lProvider As Long, ByVal lType As Long, ByVal lFlags As Long) As Long
Private Declare Function CryptCreateHash Lib "advapi32" (ByVal lProvider As Long, ByVal lAlgoritm As Long, ByVal lKey As Long, ByVal lFlags As Long, lHash As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32" (ByVal lKey As Long, ByVal lHash As Long, ByVal lFinal As Long, ByVal lFlags As Long, uData As Any, lDataLen As Long) As Long
Private Declare Function CryptDeriveKey Lib "advapi32" (ByVal lProvider As Long, ByVal lAlgoritm As Long, ByVal lData As Long, ByVal lFlags As Long, lKey As Long) As Long
Private Declare Function CryptDestroyHash Lib "advapi32" (ByVal lHash As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32" (ByVal lKey As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32" (ByVal lKey As Long, ByVal lHash As Long, ByVal lFinal As Long, ByVal lFlags As Long, uData As Any, lDataLen As Long, ByVal lBufLen As Long) As Long
Private Declare Function CryptHashData Lib "advapi32" (ByVal lHash As Long, ByVal sData As String, ByVal lLen As Long, ByVal lFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32" (ByVal lProvider As Long, ByVal lFlags As Long) As Long
Private Declare Function IsTextUnicode Lib "advapi32" (ByRef uBuffer As Any, ByVal lBufferLen As Long, ByRef lResult As Long) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lString As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal lCodePage As Long, ByVal lFlags As Long, uMultiByte As Any, ByVal lMultiByteLen As Long, uWideCharStr As Any, ByVal lWideCharStrLen As Long) As Long
Private Declare Sub RtlMoveMemory Lib "kernel32" (uTarget As Any, uSource As Any, ByVal lLen As Long)
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetLocaleInfoA Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
Private Declare Function GetLocaleInfoW Lib "kernel32" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Function IsWindowUnicode Lib "user32" (ByVal lhWnd As Long) As Long
Private Declare Function SafeArrayCopy Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Long) As Long
Private Declare Function SafeArrayCopyData Lib "oleaut32" (ByVal lSource As Long, ByVal lTarget As Long) As Long
Private Declare Function SafeArrayCreate Lib "oleaut32" (ByVal lType As Integer, ByVal lDims As Long, uBounds As Any) As Long
Private Declare Function SafeArrayDestroy Lib "oleaut32" (ByVal lArray As Long) As Long
Private Declare Function SafeArrayGetElement Lib "oleaut32" (ByVal lArray As Long, ByRef lIndices As Long, uValue As Any) As Long
Private Declare Function SafeArrayPutElement Lib "oleaut32" (ByVal lArray As Long, ByRef lIndices As Long, uValue As Any) As Long
Private Declare Function SafeArrayRedim Lib "oleaut32" (ByVal lArray As Long, uLastBound As TSAFEARRAYBOUND) As Long
Private Declare Sub GetMem2 Lib "msvbvm60" (ByVal lSource As Long, iTarget As Integer)
Private Declare Sub PutMem2 Lib "msvbvm60" (ByVal lTarget As Long, ByVal iSource As Integer)
Private Declare Sub RtlZeroMemory Lib "kernel32" (uDestination As Any, ByVal lLen As Long)
Private m_Minus     As Integer
Private PUB_UNICODE As Boolean
Private m_Comma     As Integer
'EVENTS
Private Sub Class_Initialize()
    p_ValueLenReal = 256&
End Sub
Private Sub Class_Terminate()
    CryptDestroy
End Sub
'PROPERTIES
Private Property Get TypesInterface_Pointer() As Long
    TypesInterface_Pointer = p_ValuePtr
End Property
Friend Property Get Asc(Optional ByVal Index As Long) As Long
    If p_ValueLen Then
        If Index >= 0& And Index <= p_ValueLen Then RtlMoveMemory Asc, ByVal p_ValuePtr + ((Index + (Index > 0&)) * 2&), 2&
    End If
End Property
Friend Property Get Capacity() As Long
    Capacity = p_ValueLenReal
End Property
Friend Property Get Equals(ByRef CompareString As String, Optional ByVal CompareMethod As VbCompareMethod) As Long
    Dim l As Long
    l = Len(CompareString)
    If p_ValueLen > l Then
        Equals = 1&
    ElseIf p_ValueLen < l Then
        Equals = L_NG
    ElseIf p_ValueLen > 0& And l > 0& Then
        If PUB_UNICODE Then Equals = CompareStringW(0&, (CompareMethod > 0&) * L_NG, p_ValuePtr, p_ValueLen, StrPtr(CompareString), p_ValueLen) - 2& Else Equals = CompareStringA(0&, (CompareMethod > 0&) * L_NG, p_Value, p_ValueLen, CompareString, p_ValueLen) - 2&
    End If
End Property
Friend Property Get HashCode() As Long
    If CryptCreate Then HashCode = p_CryptHash
End Property
Friend Property Get IndexOf(ByRef SearchString As String, Optional ByVal Start As Long, Optional ByVal Reverse As Boolean, Optional ByVal CompareMethod As VbCompareMethod) As Long
    Dim x As Long
    If p_ValueLen > 0& And Len(SearchString) > 0& Then
        If Reverse Then
            If Start > 0& And Start < p_ValueLen Then x = Start Else x = p_ValueLen
            IndexOf = InStrRev(p_Value, SearchString, x, CompareMethod)
        Else
            If Start > 0& Then x = Start Else x = 1&
            x = InStr(x, p_Value, SearchString, CompareMethod)
            If x <= p_ValueLen Then IndexOf = x
        End If
    End If
End Property
Friend Property Get Length() As Long
    Length = p_ValueLen
End Property
Friend Property Get Pointer() As Long
    Pointer = TypesInterface_Pointer
End Property
Friend Property Get Value() As String
    If p_ValueLen Then Value = VBA.Left$(p_Value, p_ValueLen)
End Property
Friend Property Let Value(ByRef NewValue As String)
    Dim l As Long
    l = Len(NewValue)
    If l Then
        BufferRedim l, True
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(NewValue), p_ValueLen * 2&
    Else
        p_ValueLen = 0&
    End If
End Property
'METHODS
Private Function TypesInterface_Clone() As TypesInterface
    Set TypesInterface_Clone = New StringEx
    If p_ValueLen Then TypesInterface_Clone.Parse VBA.Left$(p_Value, p_ValueLen)
End Function
Private Sub ArrayPtr(ByRef SafeArray As TSAFEARRAY, ByVal SourceArrayPtr As Long, Optional ByVal IsExternal As Boolean)
    If SourceArrayPtr Then
        RtlMoveMemory SafeArray.lVarType, ByVal SourceArrayPtr + (-4& * ((Not IsExternal) * L_NG)), 2& 'get array type
        If IsExternal Then
            RtlMoveMemory SafeArray.lPointer, ByVal SourceArrayPtr + 8&, 4&
            If SafeArray.lVarType And 16384& Then 'if passed by reference (pointer to pointer)
                RtlMoveMemory SafeArray.lPointer, ByVal SafeArray.lPointer, 4&
                SafeArray.lVarType = SafeArray.lVarType Xor 16384&
            End If
            SafeArray.lVarType = SafeArray.lVarType Xor vbArray
        Else
            SafeArray.lPointer = SourceArrayPtr
        End If
        RtlMoveMemory SafeArray, ByVal SafeArray.lPointer, 16& 'fill first fixed 16 bytes from pointer
        ReDim SafeArray.uBounds(SafeArray.iDims + L_NG) 'allocate bounds member
        RtlMoveMemory SafeArray.uBounds(0), ByVal SafeArray.lPointer + 16&, SafeArray.iDims * 8& 'get array dimensions info bytes (in descending order) starting from array pointer adress + 16 bytes offset
    End If
End Sub
Private Function TypesInterface_Parse(Value As Variant) As TypesInterface
    Dim a As TSAFEARRAY
    Dim b As Integer
    Dim i As Long
    Dim j As Long
    Dim x As Long
    Set TypesInterface_Parse = Me
    x = VarType(Value)
    Select Case x
        Case vbInteger To vbString, vbError, vbBoolean, vbDecimal, vbByte
            Me.Value = Value
        Case vbArray + vbByte
            ArrayPtr a, VarPtr(Value), True
            If a.iDims = 1 Then
                x = a.uBounds(0).lElements 'get items count
                If x Mod 2& Then 'odd bytes count (ANSI, UTF-8)
                    If x >= 3& Then
                        If Value(0) = 239 And Value(1) = 187 And Value(2) = 191 Then j = 3& 'check for UTF-8 byte order mark
                        i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                        If i Then '7-bit ANSI or UTF-8
                            BufferRedim i - ((j > 0&) * L_NG), True
                            MultiByteToWideChar 65001, 0&, ByVal a.lData + j, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                            Exit Function
                        End If
                    End If
                    BufferRedim x, True
                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.lData, x 'copy bytes as 8-bit ANSI
                ElseIf x > 0& Then 'even bytes count (ANSI, UTF-8, UTF-16 BE/LE)
                    Select Case Value(0) 'check if first byte BOM-like
                        Case 255
                            If Value(1) = 254 Then i = 2& 'UTF-16 Little Endian
                        Case 254
                            If Value(1) = 255 Then 'UTF-16 Big Endian
                                For j = 2& To x + L_NG Step 2& 'convert to Little Endian order
                                    RtlMoveMemory b, ByVal a.lData + j, 2&
                                    b = (((b And &HFF00) \ &H100) And &HFF) Or (b And &H7F) * &H100 Or (b And &H80) * &HFF00 'swap integer low & high words
                                    RtlMoveMemory ByVal a.lData + j, b, 2&
                                Next j
                                i = 2&
                            End If
                        Case 239
                            If x >= 3& Then
                                If Value(1) = 187 And Value(2) = 191 Then 'UTF-8
                                    i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                                    If i Then
                                        BufferRedim i + L_NG, True
                                        MultiByteToWideChar 65001, 0&, ByVal a.lData + 3&, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                                        Exit Function
                                    End If
                                End If
                            End If
                    End Select
                    If i = 0& Then 'no BOM
                        If IsTextUnicode(ByVal a.lData, x, 7&) = 0& Then 'if IS_TEXT_UNICODE_ASCII16, IS_TEXT_UNICODE_STATISTICS and IS_TEXT_UNICODE_CONTROLS tests fail
                            i = MultiByteToWideChar(65001, 8&, ByVal a.lData, x, ByVal 0&, 0&) 'estimate buffer size (with MB_ERR_INVALID_CHARS flag)
                            If i Then '7-bit ANSI or UTF-8
                                BufferRedim i, True
                                MultiByteToWideChar 65001, 0&, ByVal a.lData, x, ByVal p_ValuePtr, p_ValueLen 'convert multi-bytes
                            Else '8-bit ANSI or UTF-16 case (currently difference can not be determined correctly, f.e. Chinese double-byte encoded text is recognized as ANSI)
                                BufferRedim x, True
                                RtlMoveMemory ByVal p_Value, ByVal a.lData, x 'copy bytes as 8-bit ANSI
                            End If
                            Exit Function
                        End If
                    End If
                    BufferRedim (x - i) \ 2&, True
                    RtlMoveMemory ByVal p_ValuePtr, ByVal a.lData + i, x - i 'copy bytes as Unicode
                End If
            End If
        Case Else
            p_ValueLen = 0&
    End Select
End Function
Public Function Clone() As StringEx
    Set Clone = TypesInterface_Clone
End Function
Public Function Concat(ByRef ConcatString As String) As StringEx
    Dim l As Long
    Set Concat = Me
    l = Len(ConcatString)
    If l Then
        BufferRedim l, False
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal StrPtr(ConcatString), l * 2&
    End If
End Function
Public Function ConcatPointer(ByVal ConcatStringPointer As Long) As StringEx
    Dim l As Long
    Set ConcatPointer = Me
    If ConcatStringPointer Then
        l = lstrlenW(ConcatStringPointer)
        If l Then
            BufferRedim l, False
            RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - l) * 2&), ByVal ConcatStringPointer, l * 2&
        End If
    End If
End Function
Public Function Decrypt() As StringEx
    Set Decrypt = Me
    If CryptCreate And p_ValueLen Then CryptDecrypt p_CryptKey, 0&, 1&, 0&, ByVal p_ValuePtr, p_ValueLenReal * 2&
End Function
Public Function Duplicate(ByVal Count As Long) As StringEx
    Dim i As Long
    Dim l As Long
    Set Duplicate = Me
    If p_ValueLen > 0& And Count > 0& Then
        l = p_ValueLen
        BufferRedim l * Count, False
        For i = 1& To Count
            RtlMoveMemory ByVal p_ValuePtr + ((l * i) * 2&), ByVal p_ValuePtr, l * 2&
        Next i
    End If
End Function
Public Function Encrypt() As StringEx
    Set Encrypt = Me
    If CryptCreate And p_ValueLen Then CryptEncrypt p_CryptKey, 0&, 1&, 0&, ByVal p_ValuePtr, p_ValueLen * 2&, p_ValueLenReal * 2&
End Function
Public Function Insert(ByVal Index As Long, ByRef InsertString As String) As StringEx
    Dim l As Long
    Set Insert = Me
    l = Len(InsertString)
    If l > 0& And Index >= 0& And Index <= p_ValueLen Then
        BufferRedim l, False
        RtlMoveMemory ByVal p_ValuePtr + (Index * 2&) + l + l, ByVal p_ValuePtr + (Index * 2&), (p_ValueLen - (Index + l)) * 2&
        RtlMoveMemory ByVal p_ValuePtr + (Index * 2&), ByVal StrPtr(InsertString), l + l
    End If
End Function
Public Function Left(ByVal Length As Long) As StringEx
    Set Left = Me
    If Length >= 0& And Length < p_ValueLen Then p_ValueLen = Length
End Function
Public Function Lower() As StringEx
    Set Lower = Me
    If p_ValueLen Then
        If PUB_UNICODE Then CharLowerBuffW p_ValuePtr, p_ValueLen Else CharLowerBuffA p_Value, p_ValueLen
    End If
End Function
Public Function Mid(ByVal Start As Long, Optional ByVal Length As Long) As StringEx
    Set Mid = Me
    If p_ValueLen > 0& And Start > 0& And Start <= p_ValueLen Then
        If Length > 0& And Length <= (p_ValueLen - Start) Then p_ValueLen = Length Else p_ValueLen = p_ValueLen - Start + 1&
        RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + ((Start + L_NG) * 2&), p_ValueLen * 2&
    End If
End Function
Public Function Numeric() As StringEx
    Set Numeric = Me
    Value = ToNumber(p_ValuePtr, p_ValueLen)
End Function
Public Function PadLeft(ByVal Width As Long) As StringEx
    Dim x As Long
    Set PadLeft = Me
    If Width > p_ValueLen Then
        x = Width - p_ValueLen
        BufferRedim x, False
        RtlMoveMemory ByVal p_ValuePtr + x + x, ByVal p_ValuePtr, (p_ValueLen - x) * 2&
        RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(Space$(x)), x + x
    End If
End Function
Public Function PadRight(ByVal Width As Long) As StringEx
    Dim x As Long
    Set PadRight = Me
    If Width > p_ValueLen Then
        x = Width - p_ValueLen
        BufferRedim x, False
        RtlMoveMemory ByVal p_ValuePtr + ((p_ValueLen - x) * 2&), ByVal StrPtr(Space$(x)), x + x
    End If
End Function
Public Function Parse(ByRef Value As Variant) As StringEx
    Set Parse = TypesInterface_Parse(Value)
End Function
Public Function Remove(ByVal Index As Long, ByVal Length As Long) As StringEx
    Dim x As Long
    Set Remove = Me
    If p_ValueLen > 0& And Index > 0& And Index <= p_ValueLen And Length > 0& Then
        x = Length - (Length - (p_ValueLen - Index + 1&))
        If Length <= x Then x = Length
        If x < p_ValueLen Then
            RtlMoveMemory ByVal p_ValuePtr + ((Index + L_NG) * 2&), ByVal p_ValuePtr + ((Index + L_NG) * 2&) + (x * 2&), (p_ValueLen - (Index + x) + 1&) * 2&
            p_ValueLen = p_ValueLen - x
        Else
            p_ValueLen = 0&
        End If
    End If
End Function
Public Function Replace(ByRef SearchString As String, ByRef ReplaceString As String, Optional ByVal Start As Long = 1&, Optional ByVal Count As Long = L_NG, Optional ByVal CompareMethod As VbCompareMethod) As StringEx
    Set Replace = Me
    If p_ValueLen Then Value = VBA.Replace(VBA.Left$(p_Value, p_ValueLen), SearchString, ReplaceString, Start, Count, CompareMethod)
End Function
Public Function Right(ByVal Length As Long) As StringEx
    Set Right = Me
    If Length >= 0& And Length < p_ValueLen Then
        RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + ((p_ValueLen - Length) * 2&), Length + Length
        p_ValueLen = Length
    End If
End Function
'Public Function Split(Optional ByRef Delimeter As String, Optional ByVal Limit As Long = L_NG, Optional ByVal CompareMethod As VbCompareMethod, Optional ByVal ArrayType As VbVarType = vbString) As ListEx
'
'    On Error Resume Next
'
'    Dim i As Long
'    Dim l As Long
'    Dim p As Long
'    Dim s As String
'    Dim x As Long
'    Dim y As Long
'
'    Set Split = New ListEx
'
'    If p_ValueLen > 0& And Len(Delimeter) > 0& Then
'
'        Split.Create ArrayType, p_ValueLen + L_NG
'
'        s = Space$(p_ValueLen)
'        p = StrPtr(s)
'
'        Do
'
'            x = InStr(x + 1&, p_Value, Delimeter, CompareMethod)
'
'            If x = 0& Or x > p_ValueLen Then x = p_ValueLen + 1&
'
'            l = x - y + L_NG
'
'            RtlMoveMemory ByVal p, ByVal p_ValuePtr + y + y, l + l
'
'            Split.Item(i) = VBA.Left$(s, l)
'
'            i = i + 1&
'            y = x
'
'        Loop While y < p_ValueLen
'
'        If Limit > L_NG And Limit < i Then i = Limit Else i = i + L_NG
'
'        Split.Resize i
'
'    Else
'
'        Split.Parse VBA.Left$(p_Value, p_ValueLen)
'
'    End If
'
'End Function
'Public Function ToArray(Optional ByVal AsANSI As Boolean) As ListEx
'
'    Dim x As Long
'
'    Set ToArray = New ListEx
'
'    If p_ValueLen Then
'
'        x = p_ValueLen * (((Not AsANSI) * L_NG) + 1&)
'
'        ToArray.Create vbByte, x
'
'        If AsANSI Then RtlMoveMemory ByVal ToArray.Data, ByVal p_Value, x Else RtlMoveMemory ByVal ToArray.Data, ByVal p_ValuePtr, x
'
'    End If
'
'End Function
'Public Function ToBoolean() As BooleanEx
'
'    Set ToBoolean = New BooleanEx
'
'    If p_ValueLen Then ToBoolean.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToByte() As ByteEx
'
'    Set ToByte = New ByteEx
'
'    If p_ValueLen Then ToByte.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToCurrency() As CurrencyEx
'
'    Set ToCurrency = New CurrencyEx
'
'    If p_ValueLen Then ToCurrency.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToDecimal() As DecimalEx
'
'    Set ToDecimal = New DecimalEx
'
'    If p_ValueLen Then ToDecimal.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToDouble() As DoubleEx
'
'    Set ToDouble = New DoubleEx
'
'    If p_ValueLen Then ToDouble.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToInteger() As IntegerEx
'
'    Set ToInteger = New IntegerEx
'
'    If p_ValueLen Then ToInteger.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToList() As ListEx
'
'    Set ToList = New ListEx
'
'    ToList.Parse Me.Value
'
'End Function
'
'Public Function ToLong() As LongEx
'
'    Set ToLong = New LongEx
'
'    If p_ValueLen Then ToLong.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
'
'Public Function ToSingle() As SingleEx
'
'    Set ToSingle = New SingleEx
'
'    If p_ValueLen Then ToSingle.Parse VBA.Left$(p_Value, p_ValueLen)
'
'End Function
Public Function TrimL() As StringEx
    Dim x As Long
    Set TrimL = Me
    If p_ValueLen Then
        If AscW(p_Value) = 32 Then
            Do
                x = x + 1&
            Loop While InStr(x + 1&, p_Value, ChrW$(32&), vbBinaryCompare) = (x + 1&)
            If x Then
                p_ValueLen = p_ValueLen - x
                RtlMoveMemory ByVal p_ValuePtr, ByVal p_ValuePtr + (x * 2&), p_ValueLen * 2&
            End If
        End If
    End If
End Function
Public Function TrimNull() As StringEx
    Dim x As Long
    Set TrimNull = Me
    If p_ValueLen Then
        x = InStr(1&, p_Value, vbNullChar, vbBinaryCompare)
        If x > 0& And x <= p_ValueLen Then p_ValueLen = x + L_NG
    End If
End Function
Public Function TrimR() As StringEx
    Set TrimR = Me
    If p_ValueLen Then
        If AscW(Strings.Mid$(p_Value, p_ValueLen, 1&)) = 32 Then
            Do
                p_ValueLen = p_ValueLen + L_NG
            Loop While InStrRev(p_Value, ChrW$(32&), p_ValueLen, vbBinaryCompare) = p_ValueLen
        End If
    End If
End Function
Public Function Upper() As StringEx
    Set Upper = Me
    If p_ValueLen Then
        If PUB_UNICODE Then CharUpperBuffW p_ValuePtr, p_ValueLen Else CharUpperBuffA p_Value, p_ValueLen
    End If
End Function
Private Function ToNumber(ByVal lPointer As Long, ByVal lLength As Long) As String
    Dim b As Integer
    Dim C As Long
    Dim i As Long
    Dim f As Boolean
    Dim m As Boolean
    Dim p As Long
    If lLength Then
        ToNumber = Space$(lLength)
        p = StrPtr(ToNumber)
        For i = 1& To lLength
            GetMem2 lPointer, b
            If b > 47 And b < 58 Then
                PutMem2 p, b
                C = C + 1&
                p = p + 2&
            ElseIf Not f And b = m_Comma Then
                PutMem2 p, b
                C = C + 1&
                p = p + 2&
                f = True
            ElseIf C = 0& And Not f And Not m And b = m_Minus Then
                PutMem2 p, b
                C = C + 1&
                p = p + 2&
                m = True
            End If
            lPointer = lPointer + 2&
        Next i
        If C > 308& Then
            ' ToNumber = Left(ToNumber, 308&)
        ElseIf C = 1& Then
            If AscW(ToNumber) = m_Comma Or AscW(ToNumber) = m_Minus Then ToNumber = ChrW$(48&)
        ElseIf C > 0& Then
            '  ToNumber = Left$(ToNumber, c)
        Else
            '  ToNumber = ChrW$(48&)
        End If
    Else
        ToNumber = ChrW$(48&)
    End If
End Function
Private Sub BufferRedim(ByVal NewStringLen As Long, ByVal FromScratch As Boolean)
    Dim b As String
    Dim r As Boolean
    Dim x As Long
    p_ValueLen = (p_ValueLen * ((Not FromScratch) * L_NG)) + NewStringLen
    If p_ValueLen > (p_ValueLenReal \ 2&) Then p_ValueLenReal = p_ValueLenReal + (p_ValueLen * 2&) Else r = p_ValuePtr
    If Not r Then
        If FromScratch Then
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
        Else
            x = p_ValueLen - NewStringLen
            If x Then b = VBA.Left$(p_Value, x)
            p_Value = Space$(p_ValueLenReal)
            p_ValuePtr = StrPtr(p_Value)
            If x Then RtlMoveMemory ByVal p_ValuePtr, ByVal StrPtr(b), x * 2&
        End If
    End If
End Sub
Private Function CryptCreate() As Boolean
    If p_CryptContext Then
        CryptCreate = True
    Else
        If PUB_UNICODE Then
            If CryptAcquireContextW(p_CryptContext, StrPtr(appTitled), StrPtr(CONST_CRYPT_PROVIDER), 1&, 8&) = 0& Then CryptAcquireContextW p_CryptContext, StrPtr(appTitled), StrPtr(CONST_CRYPT_PROVIDER), 1&, 0&
        Else
            If CryptAcquireContextA(p_CryptContext, appTitled, CONST_CRYPT_PROVIDER, 1&, 8&) Then CryptAcquireContextA p_CryptContext, appTitled, CONST_CRYPT_PROVIDER, 1&, 0&
        End If
        If p_CryptContext Then
            If CryptCreateHash(p_CryptContext, 32771, 0&, 0&, p_CryptHash) Then 'ALG_CLASS_HASH + ALG_SID_MD5
                If CryptHashData(p_CryptHash, appTitled, 6&, 0&) Then
                    If CryptDeriveKey(p_CryptContext, 26625&, p_CryptHash, 0&, p_CryptKey) Then CryptCreate = True 'ALG_CLASS_DATA_ENCRYPT + ALG_TYPE_STREAM + ALG_SID_RC4
                End If
            End If
        End If
    End If
End Function
Private Sub CryptDestroy()
    If p_CryptContext Then
        If p_CryptKey Then CryptDestroyKey p_CryptKey
        If p_CryptHash Then CryptDestroyHash p_CryptHash
        CryptReleaseContext p_CryptContext, 0&
    End If
End Sub

