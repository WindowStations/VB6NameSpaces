VERSION 1.0 CLASS
BEGIN
   MultiUse = -1
   Persistable = 0
   DataBindingBehavior = 0
   DataSourceBehavior = 0
   MTSTransactionMode = 0
END
Attribute VB_Name = "ArrayEx"
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
'Attribute VB_Name = "ArrayEx"
'Attribute VB_GlobalNameSpace = False
'Attribute VB_Creatable = True
'Attribute VB_PredeclaredId = False
'Attribute VB_Exposed = False
'Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
'Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'Option Explicit
Private Const L_MX As Long = 2147483647
Private Const L_NG As Long = -1&
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type
Private Declare Function apiGetDesktopWindow Lib "user32" Alias "GetDesktopWindow" () As Long
Private Declare Function apiGetLocaleInfoA Lib "kernel32" Alias "GetLocaleInfoA" (ByVal lLocale As Long, ByVal lType As Long, ByVal sBuffer As String, ByVal lBufferLen As Long) As Long
Private Declare Function apiGetLocaleInfoW Lib "kernel32" Alias "GetLocaleInfoW" (ByVal lLocale As Long, ByVal lType As Long, ByVal lBuffer As Long, ByVal lBufferLen As Long) As Long
Private Declare Function apiIsWindowUnicode Lib "user32" Alias "IsWindowUnicode" (ByVal lhWnd As Long) As Long
Private Declare Function apiSafeArrayAccessData Lib "oleaut32" Alias "SafeArrayAccessData" (ByRef psa As SAFEARRAY2D, ByRef ppvData As Any) As Long
Private Declare Function apiSafeArrayAllocData Lib "oleaut32" Alias "SafeArrayAllocData" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayAllocDescriptor Lib "oleaut32" Alias "SafeArrayAllocDescriptor" (ByVal cDims As Long, ByRef ppsaOut As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayAllocDescriptorEx Lib "oleaut32" Alias "SafeArrayAllocDescriptorEx" (ByVal vt As Integer, ByVal cDims As Long, ByRef ppsaOut As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayCopy Lib "oleaut32" Alias "SafeArrayCopy" (ByRef lSource As SAFEARRAY2D, ByRef lTarget As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayCopyData Lib "oleaut32" Alias "SafeArrayCopyData" (ByRef lSource As SAFEARRAY2D, ByRef lTarget As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayCreate Lib "oleaut32" Alias "SafeArrayCreate" (ByVal lType As Integer, ByVal lDims As Long, ByRef uBounds As TSAFEARRAYBOUND) As Long
Private Declare Function apiSafeArrayCreateEx Lib "oleaut32" Alias "SafeArrayCreateEx" (ByVal vt As Integer, ByVal cDims As Long, ByRef rgsabound As TSAFEARRAYBOUND, ByRef pvExtra As Any) As Long
Private Declare Function apiSafeArrayCreateVector Lib "oleaut32" Alias "SafeArrayCreateVector" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long) As Long
Private Declare Function apiSafeArrayCreateVectorEx Lib "oleaut32" Alias "SafeArrayCreateVectorEx" (ByVal vt As Integer, ByVal lLbound As Long, ByVal cElements As Long, ByRef pvExtra As Any) As Long
Private Declare Function apiSafeArrayDestroy Lib "oleaut32" Alias "SafeArrayDestroy" (ByRef lArray As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayDestroyData Lib "oleaut32" Alias "SafeArrayDestroyData" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayDestroyDescriptor Lib "oleaut32" Alias "SafeArrayDestroyDescriptor" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayGetDim Lib "oleaut32" Alias "SafeArrayGetDim" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayGetElement Lib "oleaut32" Alias "SafeArrayGetElement" (ByRef lArray As SAFEARRAY2D, ByRef lIndices As Long, ByRef uValue As Any) As Long
Private Declare Function apiSafeArrayGetElemsize Lib "oleaut32" Alias "SafeArrayGetElemsize" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayGetIID Lib "oleaut32" Alias "SafeArrayGetIID" (ByRef psa As SAFEARRAY2D, ByRef pguid As GUID) As Long
Private Declare Function apiSafeArrayGetLBound Lib "oleaut32" Alias "SafeArrayGetLBound" (ByRef psa As SAFEARRAY2D, ByVal nDim As Long, ByRef plLbound As Long) As Long
Private Declare Function apiSafeArrayGetRecordInfo Lib "oleaut32" Alias "SafeArrayGetRecordInfo" (ByRef psa As SAFEARRAY2D, ByRef prinfo As Long) As Long
Private Declare Function apiSafeArrayGetUBound Lib "oleaut32" Alias "SafeArrayGetUBound" (ByRef psa As SAFEARRAY2D, ByVal nDim As Long, ByRef plUbound As Long) As Long
Private Declare Function apiSafeArrayGetVartype Lib "oleaut32" Alias "SafeArrayGetVartype" (ByRef psa As SAFEARRAY2D, ByRef pvt As Integer) As Long
Private Declare Function apiSafeArrayLock Lib "oleaut32" Alias "SafeArrayLock" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayPutElement Lib "oleaut32" Alias "SafeArrayPutElement" (ByRef lArray As SAFEARRAY2D, ByRef lIndices As Long, ByRef uValue As Any) As Long
Private Declare Function apiSafeArrayPtrOfIndex Lib "oleaut32" Alias "SafeArrayPtrOfIndex" (ByRef psa As SAFEARRAY2D, ByRef rgIndices As Long, ByRef ppvData As Any) As Long
Private Declare Function apiSafeArrayRedim Lib "oleaut32" Alias "SafeArrayRedim" (ByRef lArray As SAFEARRAY2D, ByRef uLastBound As TSAFEARRAYBOUND) As Long
Private Declare Function apiSafeArraySetIID Lib "oleaut32" Alias "SafeArraySetIID" (ByRef psa As SAFEARRAY2D, ByVal GUID As Long) As Long
Private Declare Function apiSafeArraySetRecordInfo Lib "oleaut32" Alias "SafeArraySetRecordInfo" (ByRef psa As SAFEARRAY2D, ByRef prinfo As Long) As Long
Private Declare Function apiSafeArrayUnaccessData Lib "oleaut32" Alias "SafeArrayUnaccessData" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiSafeArrayUnlock Lib "oleaut32" Alias "SafeArrayUnlock" (ByRef psa As SAFEARRAY2D) As Long
Private Declare Function apiVarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef pt() As Any) As Long
Private Declare Function apiPathFindSuffixArray Lib "shlwapi" Alias "PathFindSuffixArrayA" (ByVal pszPath As String, ByVal apszSuffix As String, ByVal iArraySize As Long) As Long
Private Declare Sub apiGetMem2 Lib "msvbvm60" Alias "GetMem2" (ByVal lSource As Long, ByRef iTarget As Integer)
Private Declare Sub apiPutMem2 Lib "msvbvm60" Alias "PutMem2" (ByVal lTarget As Long, ByVal iSource As Integer)
Private Declare Sub apiRtlZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByRef uDestination As Any, ByVal lLen As Long)
Private Declare Sub apiRtlMoveMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef uTarget As Any, ByRef uSource As Any, ByVal lLen As Long)
Private PUB_UNICODE As Boolean
Private m_Comma As Integer
Private m_Minus As Integer
'Private Sub Class_initialize()
'    Dim s As String
'    Dim p As Long
'    s = Space$(1)
'    p = StrPtr(s)
'    PUB_UNICODE = apiIsWindowUnicode(apiGetDesktopWindow)
'    If PUB_UNICODE <> 0 Then
'        If apiGetLocaleInfoW(1024, 14, p, 2) Then m_Comma = AscW(s)
'        If apiGetLocaleInfoW(1024, 81, p, 2) Then m_Minus = AscW(s)
'    Else
'        If apiGetLocaleInfoA(1024, 14, s, 2) Then m_Comma = AscW(s)
'        If apiGetLocaleInfoA(1024, 81, s, 2) Then m_Minus = AscW(s)
'    End If
'End Sub
'Friend Sub Copy(ByRef SArray As TSAFEARRAY, ByRef SourceOrOutput As Variant)
''    If VarType(SourceOrOutput) = vbEmpty Then
''        VariantInit SourceOrOutput, vbArray + SArray.lVarType
''        apiSafeArrayCopy SArray.lPointer, VarPtr(SourceOrOutput) + 8&
''    Else
''        apiSafeArrayCopyData SourceOrOutput, SafeArray.lPointer
''    End If
'End Sub
'Friend Function Create(ByRef SafeArray As TSAFEARRAY, ByVal SafeArrayType As VbVarType, ByVal SafeArrayDims As Long, ByRef SafeArrayBounds() As TSAFEARRAYBOUND) As Boolean
'    Dim x As Long
'    Destroy SafeArray
'    x = apiSafeArrayCreate(SafeArrayType Xor (vbArray * (((SafeArrayType And vbArray) = vbArray) * L_NG)), SafeArrayDims, SafeArrayBounds(0))
'    If x Then
'        Ptr SafeArray, x
'        Create = SafeArray.lPointer
'    End If
'End Function
'Friend Function CreateVector(ByRef SafeArray As TSAFEARRAY, ByVal SafeArrayType As VbVarType, ByVal SafeArrayElements As Long) As Boolean
'    Dim b(0) As TSAFEARRAYBOUND
'    b(0).lElements = SafeArrayElements
'    CreateVector = Create(SafeArray, SafeArrayType, 1, b)
'End Function
'Friend Sub Destroy(ByRef SafeArray As TSAFEARRAY)
''    Call apiSafeArrayDestroy(SafeArray.lPointer)
''    Erase SafeArray.uBounds
''    apiRtlZeroMemory SafeArray, Len(SafeArray)
'End Sub
'Friend Function ElementGet(ByRef SafeArray As TSAFEARRAY, ByRef Indexes As Long, ByRef Value As Variant) As Boolean
'    Select Case SafeArray.lVarType
'        Case vbInteger To vbString, vbError, vbBoolean, vbByte: ElementGet = (apiSafeArrayGetElement(SafeArray.lPointer, Indexes, ByVal VarPtr(Value) + 8&) = 0&)
'        Case vbObject, vbVariant: ElementGet = (apiSafeArrayGetElement(SafeArray.lPointer, Indexes, Value) = 0)
'        Case vbDecimal: ElementGet = (apiSafeArrayGetElement(SafeArray.lPointer, Indexes, ByVal VarPtr(Value)) = 0)
'    End Select
'End Function
'Friend Function ElementSet(ByRef SafeArray As TSAFEARRAY, ByRef Indexes As Long, ByRef Value As Variant) As Boolean
'    On Error Resume Next
'    Select Case SafeArray.lVarType
'        Case vbError, vbLong: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CLng(Value)) = 0&)
'        Case vbInteger: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CInt(Value)) = 0&)
'        Case vbDouble: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CDbl(Value)) = 0&)
'        Case vbCurrency: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CCur(Value)) = 0&)
'        Case vbSingle: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CSng(Value)) = 0&)
'        Case vbByte: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CByte(Value)) = 0&)
'        Case vbBoolean: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CBool(Value)) = 0&)
'        Case vbString: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, ByVal StrPtr(Value)) = 0&)
'        Case vbDate: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, CDate(Value)) = 0&)
'        Case vbObject, vbVariant: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, Value) = 0&)
'        Case vbDecimal: ElementSet = (apiSafeArrayPutElement(SafeArray.lPointer, Indexes, ByVal VarPtr(CDec(Value))) = 0&)
'    End Select
'End Function
'Friend Sub Ptr(ByRef SafeArray As TSAFEARRAY, ByVal SourceArrayPtr As Long, Optional ByVal IsExternal As Boolean)
'    If SourceArrayPtr Then
'        apiRtlMoveMemory SafeArray.lVarType, ByVal SourceArrayPtr + (-4 * ((Not IsExternal) * L_NG)), 2 'get array type
'        If IsExternal = True Then
'            Call apiRtlMoveMemory(SafeArray.lPointer, ByVal SourceArrayPtr + 8, 4)
'            If SafeArray.lVarType And 16384& Then 'if passed by reference (pointer to pointer)
'                apiRtlMoveMemory SafeArray.lPointer, ByVal SafeArray.lPointer, 4
'                SafeArray.lVarType = SafeArray.lVarType Xor 16384
'            End If
'            SafeArray.lVarType = SafeArray.lVarType Xor vbArray
'        Else
'            SafeArray.lPointer = SourceArrayPtr
'        End If
'        Call apiRtlMoveMemory(SafeArray, ByVal SafeArray.lPointer, 16)  'fill first fixed 16 bytes from pointer
'        ReDim SafeArray.uBounds(SafeArray.iDims + L_NG) 'allocate bounds member
'        Call apiRtlMoveMemory(SafeArray.uBounds(0), ByVal SafeArray.lPointer + 16, SafeArray.iDims * 8)  'get array dimensions info bytes (in descending order) starting from array pointer adress + 16 bytes offset
'    End If
'End Sub
'Friend Sub Resize(ByRef SafeArray As TSAFEARRAY, ByVal NewCount As Long, Optional ByVal HoldSortedFlag As Boolean)
'    If Not HoldSortedFlag Then SafeArray.lSorted = 0
'    SafeArray.uBounds(0).lElements = NewCount
'    Call apiSafeArrayRedim(SafeArray.lPointer, SafeArray.uBounds(0))
'End Sub
'Friend Function ToNumber(ByVal lPointer As Long, ByVal lLength As Long) As String
'    Dim b As Integer
'    Dim c As Long
'    Dim i As Long
'    Dim f As Boolean
'    Dim m As Boolean
'    Dim p As Long
'    If lLength Then
'        ToNumber = Space(lLength)
'        p = StrPtr(ToNumber)
'        For i = 1 To lLength
'            apiGetMem2 lPointer, b
'            If b > 47 And b < 58 Then
'                apiPutMem2 p, b
'                c = c + 1
'                p = p + 2
'            ElseIf Not f And b = m_Comma Then
'                apiPutMem2 p, b
'                c = c + 1
'                p = p + 2
'                f = True
'            ElseIf c = 0& And Not f And Not m And b = m_Minus Then
'                apiPutMem2 p, b
'                c = c + 1
'                p = p + 2
'                m = True
'            End If
'            lPointer = lPointer + 2
'        Next i
'        If c > 308 Then
'            ToNumber = Left$(ToNumber, 308)
'        ElseIf c = 1 Then
'            If AscW(ToNumber) = m_Comma Or AscW(ToNumber) = m_Minus Then ToNumber = ChrW(48)
'        ElseIf c > 0 Then
'            ToNumber = Left(ToNumber, c)
'        Else
'            ToNumber = ChrW(48)
'        End If
'    Else
'        ToNumber = ChrW(48)
'    End If
'End Function
'Public Sub VariantInit(ByRef uVariable As Variant, ByVal lVariableType As Long)
'    Call apiRtlMoveMemory(uVariable, lVariableType, 2)
'End Sub
'Public Sub VariantZero(ByRef uVariable As Variant)
'    Call apiRtlMoveMemory(uVariable, 0, 4)
'End Sub
'This section is purely for your interest.
'
'If you try the following:
'
'Dim arrTemp(1)   As Integer
'Dim ptrVarPtr    As Long
'
'ptrVarPtr = VarPtr(arrTemp)
'you 'll get a nice "Type Mismatch" error. VB absolutely will not let you pass an array to VarPtr (and passing an element just returns the data area). So, how do you ever test the memory storage of an array? Or for that matter, ever hack it? Well, you wrap it in a UDT and then get the address from that... but that doesn't work directly. Try this:
'
'Private Type udtMyType
'   arrbytVar(1)    As Byte
'End Type
'
':
':
'
'Dim udtMyTemp   As udtMyType
'
'udtMyTemp.arrbytVar(0) = 1
'udtMyTemp.arrbytVar(1) = 2
'
'Text1.Text = Text1.Text & PrintOut(ptrToVar, _
'             LenB(udtMyTemp), "ptrToVar Dump For udtMyTemp")
'And hey, presto, it actually prints out the data elements of the array (in other words, each element of the array is an element of the UDT; it isn't stored as a SAFE_ARRAY)! Now, try dimensioning the array only later in code:
'
'Private Type udtMyType
'   arrbytVar()    As Byte
'End Type
'
':
':
'
'Dim udtMyTemp   As udtMyType
'
'ReDim udtMyTemp.arrbytVar(1)
'
'udtMyTemp.arrbytVar(0) = 1
'udtMyTemp.arrbytVar(1) = 2
'
'Text1.Text = Text1.Text & PrintOut(ptrToVar, LenB
'(udtMyTemp), "ptrToVar Dump For udtMyTemp")
'Bingo. There you go. You get a pointer in the printout. Follow the pointer if you like and you'll get the "base" of the array. Here is the complete code for doing this:
'
'Dim udtMyTemp       As udtMyType
'Dim ptrToVar        As Long
'Dim ptrToArrBase    As Long
'Dim ptrToArrData    As Long
'
'ReDim udtMyTemp.arrbytVar(0 To 1, 1 To 2)
'
'udtMyTemp.arrbytVar(0, 1) = 1
'udtMyTemp.arrbytVar(1, 1) = 2
'udtMyTemp.arrbytVar(0, 2) = 3
'udtMyTemp.arrbytVar(1, 2) = 4
'
'ptrToVar = VarPtr(udtMyTemp)
'
'Text1.Text = Text1.Text & "ptrToVar = " & Hex
'(ptrToVar) & "h (" & ptrToVar &")" & vbNewLine
'
'Text1.Text = Text1.Text & PrintOut(ptrToVar, LenB
'(udtMyTemp), "ptrToVar Dump For udtMyTemp")
'
''Follow the safe array "base" pointer
'CopyMemory ByVal VarPtr(ptrToArrBase), ByVal ptrToVar, 4
'
'Text1.Text = Text1.Text & "ptrToArrBase = " & Hex
'(ptrToArrBase) & vbNewLine
'Text1.Text = Text1.Text & PrintOut(ptrToArrBase, 32,
'"ptrToArrBase Dump For Array")"
'
''Follow the data area pointer
'CopyMemory ByVal VarPtr(ptrToArrData), ByVal ptrToArrBase + 12, 4
'
'Text1.Text = Text1.Text & "ptrToArrData = " & Hex
'(ptrToArrData) & vbNewLine
'Text1.Text = Text1.Text & PrintOut(ptrToArrData, 4, "ptrToArrData Dump For Array")


