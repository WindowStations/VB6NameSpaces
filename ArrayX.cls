VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"

Option Explicit
''INTERFACES
'
'Implements TypesInterface
'
''VARIABLES
'
'Private p_Array As TSAFEARRAY
'
''EVENTS
'
'Private Sub Class_Terminate()
'
'    ArrayDestroy p_Array
'
'End Sub
'
''PROPERTIES
'
'Private Property Get TypesInterface_Pointer() As Long
'
'    TypesInterface_Pointer = p_Array.lPointer
'
'End Property
'
'Public Property Get BoundL(Optional ByVal Dimension As Long) As Long
'
'    If p_Array.lPointer Then
'
'        If Dimension > 0& And Dimension <= p_Array.iDims Then BoundL = p_Array.uBounds(p_Array.iDims - Dimension).lLowest Else BoundL = p_Array.uBounds(p_Array.iDims + L_NG).lLowest
'
'    End If
'
'End Property
'
'Public Property Get BoundU(Optional ByVal Dimension As Long) As Long
'
'    Dim x As Long
'
'    If p_Array.lPointer Then
'
'        If Dimension > 0& And Dimension <= p_Array.iDims Then x = p_Array.iDims - Dimension Else x = p_Array.iDims + L_NG
'
'        BoundU = p_Array.uBounds(x).lLowest + p_Array.uBounds(x).lElements + L_NG
'
'    End If
'
'End Property
'
'Public Property Get Data() As Long
'
'    Data = p_Array.lData
'
'End Property
'
'Public Property Get Dimensions() As Long
'
'    Dimensions = p_Array.iDims
'
'End Property
'
'Public Property Get Elements(Optional ByVal Dimension As Long) As Long
'
'    If p_Array.lPointer Then
'
'        If Dimension > 1& And Dimension <= p_Array.iDims Then Elements = p_Array.uBounds(p_Array.iDims - Dimension).lElements Else Elements = p_Array.uBounds(p_Array.iDims + L_NG).lElements
'
'    End If
'
'End Property
'
'Public Property Get ElementSize() As Long
'
'    ElementSize = p_Array.lElementSize
'
'End Property
'
'Public Property Get Item(ParamArray Indexes()) As Variant
'
'    On Error GoTo ErrHandler
'
'    Dim b() As Long
'
'    If p_Array.lPointer Then
'
'        VariantInit Item, p_Array.lVarType
'
'        If p_Array.iDims > 1 Then
'
'            ArrayElementGet p_Array, ArrVariantToIndexes(Indexes, b), Item
'
'            Erase b
'
'        Else
'            ArrayElementGet p_Array, CLng(Indexes(0)), Item
'        End If
'
'    End If
'
'ErrHandler:
'
'End Property
'
'Public Property Let Item(ParamArray Indexes(), ByRef NewValue As Variant)
'
'    On Error GoTo ErrHandler
'
'    Dim b() As Long
'
'    If p_Array.lPointer Then
'
'        If p_Array.iDims > 1 Then
'
'            ArrayElementSet p_Array, ArrVariantToIndexes(Indexes, b), NewValue
'
'            Erase b
'
'        Else
'            ArrayElementSet p_Array, CLng(Indexes(0)), NewValue
'        End If
'
'    End If
'
'ErrHandler:
'
'End Property
'
'Public Property Get Pointer() As Long
'
'    Pointer = TypesInterface_Pointer
'
'End Property
'
'Public Property Get Type_() As VbVarType
'
'    Type_ = p_Array.lVarType
'
'End Property
'
'Public Property Get Value() As Variant
'
'    If p_Array.lPointer Then ArrayCopy p_Array, Value
'
'End Property
'
'Public Property Let Value(ByRef NewValue As Variant)
'
'    Dim a As TSAFEARRAY
'    Dim b() As TSAFEARRAYBOUND
'    Dim i As Long
'
'    ArrayPtr a, VarPtr(NewValue) * (((VarType(NewValue) And vbArray) = vbArray) * L_NG), True
'
'    If a.lPointer Then
'
'        ReDim b(a.iDims + L_NG)
'
'        For i = 0& To a.iDims + L_NG: b(a.iDims - i + L_NG) = a.uBounds(i): Next i
'
'        If ArrayCreate(p_Array, a.lVarType, a.iDims, b) Then ArrayCopy p_Array, a.lPointer
'
'        Erase b
'
'    End If
'
'End Property
'
''METHODS
'
'Private Function TypesInterface_Clone() As TypesInterface
'
'    Set TypesInterface_Clone = New ArrayEx
'
'    If p_Array.lPointer Then TypesInterface_Clone.Parse Me.Value
'
'End Function
'
'Private Function TypesInterface_Parse(Value As Variant) As TypesInterface
'
'    Set TypesInterface_Parse = Me
'
'    If (VarType(Value) And vbArray) = vbArray Then
'
'        Me.Value = Value
'
'    Else
'
'        If ArrayCreateVector(p_Array, VarType(Value), 1&) Then ArrayElementSet p_Array, 0&, Value
'
'    End If
'
'End Function
'
'Public Function Add(ByRef Item As Variant) As ArrayEx
'
'    Dim b() As Long
'    Dim x As Long
'
'    Set Add = Me
'
'    If p_Array.lPointer Then
'
'        ArrayRedim p_Array, p_Array.uBounds(0).lElements + 1&
'
'        If p_Array.iDims > 1& Then
'
'            x = p_Array.iDims + L_NG
'
'            ReDim b(x)
'
'            For x = 0& To x: b(x) = BoundU(x + 1&): Next x
'
'            ArrayElementSet p_Array, b(0), Item
'
'            Erase b
'
'        Else
'            ArrayElementSet p_Array, BoundU, Item
'        End If
'
'    Else
'
'        If ArrayCreateVector(p_Array, VarType(Item), 1&) Then ArrayElementSet p_Array, 0&, Item
'
'    End If
'
'End Function
'
'Public Function AddRange(ByRef Range As Variant) As ArrayEx
'
'    Dim a As TSAFEARRAY
'    Dim b() As Long
'    Dim i As Long
'    Dim x As Long
'
'    Set AddRange = Me
'
'    If p_Array.lPointer Then
'
'        ArrayPtr a, VarPtr(Range) * (((VarType(Range) And vbArray) = vbArray) * L_NG), True
'
'        If a.lPointer <> 0& And a.iDims = 1& Then
'
'            x = p_Array.iDims + L_NG
'
'            ReDim b(x)
'
'            For i = 0& To x: b(i) = BoundU(i + 1&): Next i
'
'            ArrayRedim p_Array, p_Array.uBounds(0).lElements + a.uBounds(0).lElements
'
'            For i = a.uBounds(0).lLowest To a.uBounds(0).lLowest + a.uBounds(0).lElements + L_NG
'                b(x) = b(x) + 1&
'                ArrayElementSet p_Array, b(0), Range(i)
'            Next i
'
'            Erase b
'
'        End If
'
'    Else
'
'        Me.Value = Range
'
'    End If
'
'End Function
'
'Public Function Clone() As ArrayEx
'
'    Set Clone = TypesInterface_Clone
'
'End Function
'
'Public Function Create(ByVal ArrayType As VbVarType, ParamArray Bounds()) As ArrayEx
'
'    On Error GoTo ErrHandler
'
'    Dim b() As TSAFEARRAYBOUND
'    Dim i As Long
'    Dim x As Long
'
'    Set Create = Me
'
'    x = UBound(Bounds)
'
'    ReDim b(x)
'
'    For i = 0& To x
'        b(i).lLowest = Bounds(i)(0)
'        b(i).lElements = Bounds(i)(1) - b(i).lLowest + 1&
'    Next i
'
'    ArrayCreate p_Array, ArrayType, x + 1&, b
'
'    Erase b
'
'ErrHandler:
'
'End Function
'
'Public Function Destroy() As ArrayEx
'
'    Set Destroy = Me
'
'    ArrayDestroy p_Array
'
'End Function
'
'Public Function Parse(ByRef Value As Variant) As ArrayEx
'
'    Set Parse = TypesInterface_Parse(Value)
'
'End Function
'
'Public Function Resize(ByVal NewUpperBound As Long) As ArrayEx
'
'    Set Resize = Me
'
'    If p_Array.lPointer Then
'
'        If NewUpperBound >= p_Array.uBounds(0&).lLowest Then ArrayRedim p_Array, NewUpperBound - p_Array.uBounds(0&).lLowest + 1&
'
'    End If
'
'End Function
'
'Public Function ToList(ByRef StartIndexes As Variant, ByRef EndIndexes As Variant) As ListEx
'
'    On Error GoTo ErrHandler
'
'    Dim b() As Long
'    Dim i As Long
'    Dim x As Long
'    Dim u As Long
'    Dim v As Variant
'
'    Set ToList = New ListEx
'
'    If p_Array.lPointer Then
'
'        x = ArrVariantToIndexes(StartIndexes, b)
'        u = UBound(b)
'
'        VariantInit v, p_Array.lVarType
'
'        For i = b(u) To EndIndexes(UBound(EndIndexes))
'
'            b(u) = i
'
'            If ArrayElementGet(p_Array, x, v) Then
'
'                If ToList.Pointer Then ToList.Add v Else ToList.Create(p_Array.lVarType, 1&).Item(0&) = v
'
'            End If
'
'        Next i
'
'    End If
'
'ErrHandler:
'
'    VariantZero v
'
'    Erase b
'
'End Function
'
'Public Function ToString(ByRef StartIndexes As Variant, ByRef EndIndexes As Variant, Optional ByRef Delimeter As String) As StringEx
'
'    On Error GoTo ErrHandler
'
'    Dim b() As Long
'    Dim i As Long
'    Dim x As Long
'    Dim u As Long
'    Dim v As Variant
'
'    Set ToString = New StringEx
'
'    If p_Array.lPointer Then
'
'        x = ArrVariantToIndexes(StartIndexes, b)
'        u = UBound(b)
'
'        VariantInit v, p_Array.lVarType
'
'        For i = b(u) To EndIndexes(UBound(EndIndexes))
'
'            b(u) = i
'
'            If ArrayElementGet(p_Array, x, v) Then ToString.Concat((v)).Concat Delimeter
'
'        Next i
'
'    End If
'
'ErrHandler:
'
'    VariantZero v
'
'    Erase b
'
'End Function
'
'Private Function ArrVariantToIndexes(ByVal SourceArray As Variant, ByRef OutputArray() As Long) As Long
'
'    Dim i As Long
'    Dim u As Long
'
'    u = UBound(SourceArray)
'
'    ReDim OutputArray(u)
'
'    For i = 0& To u
'        OutputArray(i) = SourceArray(i)
'    Next i
'
'    ArrVariantToIndexes = OutputArray(0)
'
'End Function

