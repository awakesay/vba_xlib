Attribute VB_Name = "console"
'   ■Version
'       v0.5

'   ■Description
'       Print functions.

'   ■Function list
'       Pprint      Display value in a structured way in the Immediate window. (Pretty print)
'       Println     Display value in the immediate window. (Print line)
'       Print2D     Display two dimensional array in the immediate window. (Print two dimension array)

'   ■Aboud display
'       String:     String type is enclose in ".
'                   Examples -> "String", "hell world!"
'
'       Numeric:    Display with thousands separator. If there are decimals, also display the decimals.
'                   Examples -> 1,234, 123.456
'
'       Boolean:    Display as is.
'                   Examples -> True or False
'
'       Date:       Display as is.
'                   Examples -> 2023/05/07 9:30:13, 2023/05/07, 9:30:13
'
'       Object:     Object type is enclose in <>.
'                   Examples -> <Worksheet>, <Collection(2)>, <Dictionary(1)
'
'       Array:      Array is display array type and array size. If empty array then not display array size.
'                   Examples -> Variant(3), Long(), String(2 To 3), Variant(2, 3 To 5)
'
'       Special:    Empty, Null, Nothing is display as is.
'                   Examples -> Empty, Null, Nothing
'


Option Explicit
Option Base 0


' ================================================================================
' Pprint (Pretty print)


Public Sub Pprint(ParamArray Any_() As Variant)
    ' Display Any in a structured string in the Immediate window.

    Dim Element As Variant
    
    Call console.DisplayImmediateWindow
    
    For Each Element In Any_
        Debug.Print console.EncodeArrayToStructString(Element)
    Next Element
    
End Sub


Private Function EncodeArrayToStructString(Any_ As Variant) As String
    ' Returns Any as a structured string.
    
    Const INDENT_SIZE = 4
    Dim EncodeArray As Variant
    Dim i As Long
    Dim Indent As Long
    Dim Values(-1 To 1) As String
    
    ' Get encode array.
    EncodeArray = EncodeAny(Any_)
    
    For i = 0 To UBound(EncodeArray)
        ' Get value and previous value and next value.
        Values(-1) = console.GetDefaultOfArray(EncodeArray, i - 1, vbNullString)
        Values(0) = EncodeArray(i)
        Values(1) = console.GetDefaultOfArray(EncodeArray, i + 1, vbNullString)
        
        If Values(0) = ", " Then
            ' Comma.
            EncodeArray(i) = Values(0) & vbNewLine
            
        ElseIf Values(1) = ": " Then
            ' Dictionary key. (Processed before the array.)
            EncodeArray(i) = String(Indent, " ") & Values(0)
        
        ElseIf Values(0) = ": " Then
            ' Dictionary separeter.
            ' Pass
        
        ElseIf Values(-1) = ": " Then
            ' Dictionary item.
            If IsOpenBracket(Values(0)) Then
                ' If open bracket then add new line.
                EncodeArray(i) = Values(0) & vbNewLine
                Indent = Indent + INDENT_SIZE
            End If
        
        ElseIf IsOpenBracket(Values(0)) Then
            ' Open bracket.
            EncodeArray(i) = String(Indent, " ") & Values(0) & vbNewLine
            Indent = Indent + INDENT_SIZE
        
        
        ElseIf IsCloseBracket(Values(0)) Then
            ' Close bracket.
            Indent = Indent - INDENT_SIZE
            If Not IsOpenBracket(Values(-1)) Then
                ' If previou value is not open bracket then add new line and indent.
                EncodeArray(i) = vbNewLine & String(Indent, " ") & Values(0)
            End If
        
        Else
            ' Value. (Add indent.)
            EncodeArray(i) = String(Indent, " ") & Values(0)
                
        End If
        
    Next i
    
    ' Return. (ArrayList object to string.)
    EncodeArrayToStructString = Join(EncodeArray, vbNullString)
    
End Function


' ================================================================================
' Println (Print line)


Public Sub Println(ParamArray Any_() As Variant)
    ' Displays Any in the Immediate window.
    
    Dim El As Variant
    
    Call console.DisplayImmediateWindow
    
    For Each El In Any_
        Debug.Print Join(console.EncodeAny(El), vbNullString)
    Next El
    
End Sub


' ================================================================================
' "Pprint" and "Println" common functions.


Private Function EncodeAny(Any_ As Variant) As Variant
    ' Returns the array to display the string.
    
    Dim EncodeStrings As Object
    Dim ObjectAddress As Object
    
    Set EncodeStrings = CreateObject("System.Collections.ArrayList")
    Set ObjectAddress = CreateObject("Scripting.Dictionary")
    
    EncodeAny = EncodeAnyRecursive(Any_, EncodeStrings, ObjectAddress, True)
    
End Function


Private Function EncodeAnyRecursive(Any_ As Variant, _
    ByRef EncodeStrings As Object, ByRef ObjectAddress As Object, _
    Optional IsReturnValue As Boolean = False _
) As Variant
    ' Recursive function.
    
    Dim AnyTypeName As String
    
    AnyTypeName = TypeName(Any_)
    
    If console.GetDimensionOfArray(Any_) = 1 Then
        ' One dimension array. (Display elements in the array.)
        Call console.EncodeArray1D(Any_, EncodeStrings, ObjectAddress)
    
    ElseIf AnyTypeName = "Dictionary" Then
        ' Dictionary object. (Display element in the dictionary.)
        Call console.ParseDictionary(Any_, EncodeStrings, ObjectAddress)
        
    ElseIf console.GetCountOfIterableObject(Any_) >= 0 Then
        ' Iterable object. (Display element in the iterable object.)
        Call console.ParseIterableObject(Any_, EncodeStrings, ObjectAddress)
        
    Else
        ' Not iterable value. (End of recursive function.)
        Call EncodeStrings.Add(AnyToString(Any_))

    End If
    
    ' Return. (Not recursive call.)
    If IsReturnValue Then EncodeAnyRecursive = EncodeStrings.ToArray
    
End Function


Private Sub EncodeArray1D(Any_ As Variant, ByRef EncodeStrings As Object, ByRef ObjectAddress As Object)
    ' Encoding one dimension array.
    
    Dim i As Long
    Dim UBnd As Long
    
    UBnd = UBound(Any_)
    
    Call EncodeStrings.Add(console.ArrayToString(Any_) & "[")
    
    For i = LBound(Any_) To UBound(Any_)
        Call console.EncodeAnyRecursive(Any_(i), EncodeStrings, ObjectAddress)
        If Not i = UBnd Then Call EncodeStrings.Add(", ")
    Next i
    
    Call EncodeStrings.Add("]")
    
End Sub


Private Sub ParseDictionary(Any_ As Variant, ByRef EncodeStrings As Object, ByRef ObjectAddress As Object)
    ' Encoding dictionary object.
    
    ' If recursive object then omitted elements of the object.
    If console.GetAndSetObjectAddress(Any_, ObjectAddress) Then
        Call EncodeStrings.Add(console.ObjectToString(Any_) & "{ … ]")
        Exit Sub
    End If
    
    Dim Keys As Variant
    Dim Items As Variant
    Dim i As Long
    
    Keys = Any_.Keys()
    Items = Any_.Items()
    
    Call EncodeStrings.Add(console.ObjectToString(Any_) & "{")
    For i = 0 To Any_.Count - 1
        Call EncodeStrings.Add(console.AnyToString(Keys(i)))
        Call EncodeStrings.Add(": ")
        Call console.EncodeAnyRecursive(Items(i), EncodeStrings, ObjectAddress)
        If Not i = Any_.Count - 1 Then Call EncodeStrings.Add(", ")
    Next i

    Call EncodeStrings.Add("}")

End Sub


Private Sub ParseIterableObject(Any_ As Variant, ByRef EncodeStrings As Object, ByRef ObjectAddress As Object)
    ' Encoding iterable object. (ArrayList, Collection, Workbooks, ... )
    
    ' If recursive object then omitted elements of the object.
    If console.GetAndSetObjectAddress(Any_, ObjectAddress) Then
        Call EncodeStrings.Add(console.AnyToString(Any_) & "[ ... ]")
        Exit Sub
    End If
    
    Dim Element As Variant
    Dim i As Long
    
    i = 1
    
    Call EncodeStrings.Add(console.AnyToString(Any_) & "[")
    
    For Each Element In Any_
        Call console.EncodeAnyRecursive(Element, EncodeStrings, ObjectAddress)
        If Not i = Any_.Count Then Call EncodeStrings.Add(", ")
        i = i + 1
    Next Element
    
    Call EncodeStrings.Add("]")
    
End Sub


Private Function GetAndSetObjectAddress(Any_ As Variant, ByRef ObjectAddress As Object) As Boolean

    If IsObject(Any_) Then
        If ObjectAddress.exists(CDbl(ObjPtr(Any_))) Then
            ' If already exist object then return "True".
            GetAndSetObjectAddress = True
        Else
            ' If not registered then return "False" and register memory address.
            Call ObjectAddress.Add(CDbl(ObjPtr(Any_)), Empty)
        End If
    End If
    
End Function


Private Function IsOpenBracket(Str_ As String) As Boolean

    Select Case Right(Str_, 1)
        Case "[", "{"
            IsOpenBracket = True
    End Select
    
End Function
        

Private Function IsCloseBracket(Str_ As String) As Boolean

    Select Case Str_
        Case "]", "}"
            IsCloseBracket = True
    End Select

End Function


' ================================================================================
' Print2D (Print two dimension array)


Public Sub Print2D(Any_ As Variant, Optional Head As Long = 10, Optional Tail As Long = 10)

    ' Check arguments.
    Debug.Assert IsArray(Any_)                          ' Not array.
    Debug.Assert console.GetDimensionOfArray(Any_) = 2  ' Not two dimension array.
    
    ' Get array bounds.
    Dim RowLBnd As Long: RowLBnd = LBound(Any_, 1)
    Dim RowUBnd As Long: RowUBnd = UBound(Any_, 1)
    Dim ColLBnd As Long: ColLBnd = LBound(Any_, 2)
    Dim ColUBnd As Long: ColUBnd = UBound(Any_, 2)
    
    ' Get display row indexes.
    Dim RowIndexes As New Collection
    Dim ColIndexes As New Collection
    Dim r As Long
    Dim c As Long
    
    For r = RowLBnd To RowUBnd
        If r < RowLBnd + Head Then
            Call RowIndexes.Add(r)
        ElseIf r >= (RowUBnd - Tail) Then
            Call RowIndexes.Add(r)
        Else
            ' Jump index.
            r = RowUBnd - Tail
        End If
    Next
    
    ' And column indexes.
    For c = ColLBnd To ColUBnd
        Call ColIndexes.Add(c)
    Next c
    
    ' Create strings array and align array. Column oriented array.
    Dim ColsArray As Variant
    Dim AlignColsArray As Variant
    Dim RowArray() As String
    Dim AlignRowArray() As Byte
    Dim RowStringIndexes() As String
    Dim AlignRowStringIndexes() As Byte
    Dim RowIndex As Variant
    Dim ColIndex As Variant
    
    ReDim ColsArray(ColIndexes.Count - 1) As Variant
    ReDim AlignColsArray(ColIndexes.Count - 1) As Variant
    
    c = 0
    For Each ColIndex In ColIndexes
        ReDim RowArray(RowIndexes.Count) As String
        ReDim AlignRowArray(RowIndexes.Count) As Byte
        r = 1
        RowArray(0) = console.AnyToString(ColIndex)     ' Index string.
        AlignRowArray(0) = 1                            ' Align center.
        For Each RowIndex In RowIndexes
            RowArray(r) = console.AnyToString(Any_(RowIndex, ColIndex))
            AlignRowArray(r) = console.GetAlignEnum(Any_(RowIndex, ColIndex))
            r = r + 1
        Next RowIndex
        ColsArray(c) = RowArray
        AlignColsArray(c) = AlignRowArray
        c = c + 1
    Next ColIndex
    
    ReDim RowStringIndexes(RowIndexes.Count) As String
    ReDim AlignRowStringIndexes(RowIndexes.Count) As Byte
    r = 1
    RowStringIndexes(0) = "i"
    AlignRowStringIndexes(0) = 2
    For Each RowIndex In RowIndexes
        RowStringIndexes(r) = console.AnyToString(RowIndex)
        AlignRowStringIndexes(r) = 2
        r = r + 1
    Next RowIndex
    
    ' Set align.
    RowStringIndexes = console.SetFillAlign(RowStringIndexes, AlignRowStringIndexes)
    For c = LBound(ColsArray) To UBound(ColsArray)
        ColsArray(c) = console.SetFillAlign(ColsArray(c), AlignColsArray(c))
    Next c
    
    ' Create display string.
    Dim DisplayStrings As Collection
    Dim InfoRowString As String
    Dim LastRowIndex As Long
    
    Set DisplayStrings = New Collection
    
    ReDim InfoRowArray(ColIndexes.Count - 1) As String
    
    InfoRowString = GetDisplayRowString(ColsArray, RowStringIndexes, 0)
    
    ' Add header string.
    Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┏┳┯━┓"))
    Call DisplayStrings.Add(InfoRowString)
    Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┣╋┿━┫"))
    
    r = 1
    LastRowIndex = RowIndexes.Item(1) - 1
    For Each RowIndex In RowIndexes
    
        If Not RowIndex = LastRowIndex + 1 Then
            ' Add middle header string.
            Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┠╂┼─┨"))
            Call DisplayStrings.Add(InfoRowString)
            Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┠╂┼─┨"))
        End If
        
        ' Add value strings.
        Call DisplayStrings.Add(GetDisplayRowString(ColsArray, RowStringIndexes, r))
        
        r = r + 1
        LastRowIndex = RowIndex
        
    Next RowIndex
    
    ' Add fotter string.
    Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┣╋┿━┫"))
    Call DisplayStrings.Add(InfoRowString)
    Call DisplayStrings.Add(console.GetRuledLineString(ColsArray, RowStringIndexes, "┗┻┷━┛"))
    
    ' Display in the immediate window.
    Dim DisplayRowString As Variant
    
    For Each DisplayRowString In DisplayStrings
        Debug.Print DisplayRowString
    Next DisplayRowString
    
    Call console.DisplayImmediateWindow
    
End Sub


Private Function GetRuledLineString( _
    ColsArray As Variant, _
    RowStringIndexes() As String, _
    RuledLineStrings As String _
) As String
    ' RuledLineStrings:
    ' 1: Right
    ' 2: IndexSplit
    ' 3: Split
    ' 4: HorizontalLine
    ' 5: Left
    ' Example: "┏┳┯━┓"
    
    Dim RuledLineArray() As String
    Dim i As Long
    
    ReDim RuledLineArray(UBound(ColsArray)) As String
    
    For i = 0 To UBound(ColsArray)
        RuledLineArray(i) = String(console.LenByte(CStr(ColsArray(i)(0))) / 2, Mid(RuledLineStrings, 4, 1))
    Next i
    
    GetRuledLineString = _
        Mid(RuledLineStrings, 1, 1) & _
        String(console.LenByte(RowStringIndexes(0)) / 2, Mid(RuledLineStrings, 4, 1)) & _
        Mid(RuledLineStrings, 2, 1) & _
        Join(RuledLineArray, Mid(RuledLineStrings, 3, 1)) & _
        Mid(RuledLineStrings, 2, 1) & _
        String(console.LenByte(RowStringIndexes(0)) / 2, Mid(RuledLineStrings, 4, 1)) & _
        Mid(RuledLineStrings, 5, 1)
        
End Function


Private Function GetDisplayRowString( _
    ColsArray As Variant, _
    RowStringIndexes() As String, _
    Index As Long _
) As String
    
    Dim RowStringArray() As String
    Dim IndexString As String
    Dim i As Long
    
    ReDim RowStringArray(UBound(ColsArray))
    
    For i = 0 To UBound(ColsArray)
        RowStringArray(i) = ColsArray(i)(Index)
    Next i
    
    IndexString = "┃" & RowStringIndexes(Index) & "┃"
    
    GetDisplayRowString = IndexString & Join(RowStringArray, "│") & IndexString

End Function


Private Function SetFillAlign(RowArray As Variant, AlignArray As Variant) As String()
    ' Return align string of space filled.
    
    Dim FillArray() As String
    Dim MaxLength As Long
    Dim i As Long
    
    ReDim FillArray(UBound(RowArray)) As String
    
    ' Get max length.
    For i = LBound(RowArray) To UBound(RowArray)
        If console.LenByte(CStr(RowArray(i))) > MaxLength Then MaxLength = console.LenByte(CStr(RowArray(i)))
    Next i
    
    ' Fix for even number.
    If Not MaxLength Mod 2 = 0 Then MaxLength = MaxLength + 1
    
    ' Filling.
    For i = LBound(RowArray) To UBound(RowArray)
        Select Case AlignArray(i)
            Case 0  ' Left
                FillArray(i) = console.FillByteRight(CStr(RowArray(i)), MaxLength)
            Case 1  ' Center
                FillArray(i) = console.FillByteCenter(CStr(RowArray(i)), MaxLength)
            Case 2  ' Right
                FillArray(i) = console.FillByteLeft(CStr(RowArray(i)), MaxLength)
        End Select
    Next i
    
    ' Return.
    SetFillAlign = FillArray
    
End Function


Private Function GetAlignEnum(Any_ As Variant) As Byte
    ' 0: Left
    ' 1: Center
    ' 2: Right
    
    Select Case TypeName(Any_)
        Case "String"
            GetAlignEnum = 0
        Case "Byte", "Integer", "Long", "LongLong", "Currency"
            GetAlignEnum = 2
        Case "Single", "Double", "Decimal"
            GetAlignEnum = 2
        Case "Boolean"
            GetAlignEnum = 1
        Case "Date"
            GetAlignEnum = 0
        Case "Empty", "Null", "Nothing"
            GetAlignEnum = 1
        Case Else
            If IsArray(Any_) Then
                GetAlignEnum = 1
            Else
                GetAlignEnum = 1
            End If
    End Select
    
End Function


Private Function IsInteger(Any_ As Variant) As Boolean
    
    Select Case TypeName(Any_)
        Case "Byte", "Integer", "Long", "LongLong", "Currency"
            IsInteger = True
        Case "Single", "Double", "Decimal"
            If InStr(1, CStr(Any_), ".") = 0 Then
                IsInteger = True
            End If
    End Select
    
End Function


Private Function LenByte(Str_ As String) As Long

    LenByte = LenB(StrConv(Str_, vbFromUnicode))
    
End Function


Private Function FillByteLeft(Str_ As String, ByteLength As Long, Optional FillChar As String = " ") As String

    Dim FillLength As Long
    
    FillLength = ByteLength - console.LenByte(Str_)
    
    If FillLength < 0 Then
        FillByteLeft = Str_
    Else
        FillByteLeft = String(FillLength, FillChar) & Str_
    End If
    
End Function


Private Function FillByteRight(Str_ As String, ByteLength As Long, Optional FillChar As String = " ") As String

    Dim FillLength As Long
    
    FillLength = ByteLength - console.LenByte(Str_)
    
    If FillLength < 0 Then
        FillByteRight = Str_
    Else
        FillByteRight = Str_ & String(FillLength, FillChar)
    End If
    
End Function


Private Function FillByteCenter(Str_ As String, Length As Long, Optional FillChar As String = " ") As String

    Dim LeftFillByteLength As Long
    Dim RightFillByteLength As Long
    
    LeftFillByteLength = Int((Length - console.LenByte(Str_)) / 2)
    RightFillByteLength = LeftFillByteLength
    
    If (Length - console.LenByte(Str_)) Mod 2 Then
        RightFillByteLength = RightFillByteLength + 1
    End If
    
    If LeftFillByteLength < 0 Then
        FillByteCenter = Str_
    Else
        FillByteCenter = String(LeftFillByteLength, FillChar) & Str_ & String(RightFillByteLength, FillChar)
    End If
    
End Function


' ================================================================================
' Common functions.


Private Sub DisplayImmediateWindow()

    Static ImmediateWindow As Object
    
    If ImmediateWindow Is Nothing Then Set ImmediateWindow = console.GetImmediateWindow()
    
    If Not ImmediateWindow.Visible Then ImmediateWindow.Visible = True
    
End Sub


Private Function GetImmediateWindow() As Object
    ' Return immediate window object.
    
    Dim WindowElement As Object
    ' Check "Trust access to the VBA project object model".
    ' JP: [オプション] -> [トラスト センター] -> [マクロの設定] -> [VBA プロジェクト オブジェクト モデルへのアクセスを信頼する] にチェックを入れる。
    ' EN: I don't understand this notation because there is no English version.
    For Each WindowElement In Application.VBE.Windows
        If WindowElement.Type = 5 Then
            ' 5: vbext_wt_Immediate
            Set GetImmediateWindow = WindowElement
            Exit Function
        End If
    Next WindowElement
    
End Function


Private Function AnyToString(Any_ As Variant) As String
    
    Dim AnyTypeName As String
    
    AnyTypeName = TypeName(Any_)
    
    Select Case AnyTypeName
    
        Case "String"
            AnyToString = """" & Any_ & """"
            
        Case "Byte", "Integer", "Long", "LongLong", "Currency"
            AnyToString = FormatNumber(Any_, 0)
        
        Case "Single", "Double", "Decimal"
            Dim DecPointIndex As Long
             ' Get decimal point index.
            DecPointIndex = InStr(Any_, ".")
            If DecPointIndex Then
                ' Exist decimal point.
                AnyToString = FormatNumber(Any_, Len(Any_) - DecPointIndex)
            Else
                ' Not Exist decimal point.
                AnyToString = FormatNumber(Any_, 0)
            End If
            
        Case "Boolean"
            AnyToString = CStr(Any_)
        
        Case "Date"
            AnyToString = FormatDateTime(Any_, vbGeneralDate)
        
        Case "Empty", "Null", "Nothing"
            AnyToString = AnyTypeName
        
        Case Else
            If IsArray(Any_) Then
                ' Array.
                AnyToString = console.ArrayToString(Any_)

            ElseIf IsObject(Any_) Then
                ' Object.
                AnyToString = console.ObjectToString(Any_)
            
            Else
                ' Unknown (Please report issue.)
                Debug.Assert False
                
            End If
            
    End Select
    
End Function


Private Function ArrayToString(Any_ As Variant) As String
    ' Return array string and array size.
    
    Dim BoundArray As Variant
    Dim i As Long
    
    BoundArray = console.GetBoundsOfArray(Any_)     ' Example:
                                                    '   Variant(10)         Base 0.
    For i = 0 To UBound(BoundArray)                 '   String(5 To 15)     Base not 0.
        If BoundArray(i)(0) = 0 Then                '   Long(100, 2 To 130) Multi dimension array.
            BoundArray(i) = BoundArray(i)(1)        '   Object()            Empty array.
        Else
            BoundArray(i) = BoundArray(i)(0) & " To " & BoundArray(i)(1)
        End If
    Next i
    
    ArrayToString = Replace(TypeName(Any_), "()", "") & "(" & Join(BoundArray, ", ") & ")"
    
End Function


Private Function ObjectToString(Any_ As Variant) As String
    
    Dim IterCount As Long
    
    IterCount = console.GetCountOfIterableObject(Any_)
    
    If IterCount = -1 Then
        ' Not iterable object.
        ObjectToString = "<" & TypeName(Any_) & ">"
    Else
        ' Iterable object. Add "Count" property.
        ObjectToString = "<" & TypeName(Any_) & "(" & IterCount & ")>"
    End If
    
End Function


Private Function GetBoundsOfArray(Any_ As Variant) As Variant
    ' Example
    '   Dim Array(1, 2 To 5) As Variant
    '   Return: [
    '       [0, 1],
    '       [2, 5]
    '   ]
    
    Dim DimensionLength As Long
    
    DimensionLength = console.GetDimensionOfArray(Any_)
    
    If DimensionLength < 1 Then
        GetBoundsOfArray = Array()
        Exit Function
    End If
    
    Dim BoundArray As Variant
    Dim i As Long
    
    ReDim BoundArray(DimensionLength - 1)
    For i = 0 To DimensionLength - 1
        BoundArray(i) = Array(LBound(Any_, i + 1), UBound(Any_, i + 1))
    Next i
    
    GetBoundsOfArray = BoundArray
    
End Function


Private Function GetDimensionOfArray(Any_ As Variant) As Long
    ' -1: Not array.
    '  0: Empty array.
    '  1: 1 dimension array.
    '  n: n dimension array.
    
    Dim DimensionLength As Long
    Dim Dummy As Long

    Call Err.Clear
    On Error Resume Next
    Do
        DimensionLength = DimensionLength + 1
        Dummy = UBound(Any_, DimensionLength)
        If Err.Number = 9 Then              ' Not found dimension.
            GetDimensionOfArray = DimensionLength - 1
            Exit Do
        ElseIf Err.Number = 13 Then         ' Not array.
            GetDimensionOfArray = -1
            Exit Do
        ElseIf Dummy = -1 Then
            GetDimensionOfArray = 0         ' Empty array.
            Exit Do
        End If
    Loop
    Call Err.Clear
    
End Function


Private Function GetCountOfIterableObject(Any_ As Variant) As Long
    ' If Iterable object then return "Count" property of object.
    ' Other than that, return -1.
    
    Dim Element As Variant
    Dim AnyCount As Long
    
    Call Err.Clear
    On Error Resume Next
    
    AnyCount = Any_.Count
    For Each Element In Any_
        Exit For
    Next Element
    
    If Err.Number = 0 Then
        GetCountOfIterableObject = AnyCount
    Else
        GetCountOfIterableObject = -1
    End If
    
    Call Err.Clear

End Function


Private Function GetDefaultOfArray(Any_ As Variant, Index As Long, Optional Default As Variant = Empty) As Variant
    ' Returns the element of the specified index of the array.
    ' If index is out of range then return "Default" value.
    
    Call Err.Clear
    On Error Resume Next
    GetDefaultOfArray = Any_(Index)
    If Err.Number = 9 Then GetDefaultOfArray = Default
    Call Err.Clear
    
End Function









