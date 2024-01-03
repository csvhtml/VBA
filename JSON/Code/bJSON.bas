Attribute VB_Name = "bJSON"
Function JSONString(var As Variant, Optional ws As String = "    ") As String
    Dim arr, keys As Variant
    arr = AddQuotes(var)
    keys = Application.Index(arr, 1, 0)
    Dim ret As String: ret = ""
    
    ret = ret + "[" + Chr(10)
    For i = 2 To UBound(var, 1)
        ret = ret + ws + "{" + Chr(10)
        For j = 1 To UBound(var, 2)
            ret = ret + ws + ws + keys(j) + ": " + arr(i, j) + "," + Chr(10)
        Next
        ret = RemoveLastCharacters(ret, 2) + Chr(10)  ' = remove comma
        ret = ret + ws + "}," + Chr(10)
    Next
    ret = RemoveLastCharacters(ret, 2) + Chr(10)  ' = remove comma
    ret = ret + "]" + Chr(10)
    
    JSONString = ret
    
    Dim a As String: a = JSONString_Dict(keys, var, ws)
End Function

Function JSONString_List(values As Variant, ws As String, Optional nthIndent As Integer = 0) As String

    Dim ret As String, wsIndent As String: ret = "": wsIndent = ""
    If JSONString_List_Assert(values) = False Then
        JSONString_List = "": Exit Function: End If
    For i = 1 To nthIndent
        wsIndent = wsIndent + ws
    Next
    
    ret = ret + wsIndent + "[" + Chr(10)
    For i = LBound(values) To UBound(values)
        ret = ret + wsIndent + ws + values(i) + "," + Chr(10)
    Next
    ret = RemoveLastCharacters(ret, 2) + Chr(10)  ' = remove last comma
    ret = ret + wsIndent + "]"

    JSONString_List = ret
    
End Function


Function JSONString_Dict(keys As Variant, values As Variant, ws As String, Optional nthIndent As Integer = 0) As String

    Dim ret As String, wsIndent As String: ret = "": wsIndent = ""
    If JSONString_Dict_Assert(keys, values) = False Then
        JSONString_Dict = "": Exit Function: End If
    
    For i = 1 To nthIndent
        wsIndent = wsIndent + ws
    Next
    
    ret = ret + wsIndent + "{" + Chr(10)
    For i = 1 To UBound(keys)
        ret = ret + wsIndent + ws + keys(i) + ": " + values(i) + "," + Chr(10)
    Next
    ret = RemoveLastCharacters(ret, 2) + Chr(10)  ' = remove last comma
    ret = ret + wsIndent + "}"

    JSONString_Dict = ret
    
End Function


'######################################################################################
' Assert                                                                              #
'######################################################################################


Function JSONString_List_Assert(values As Variant) As Boolean
    JSONString_List_Assert = False
    
    If LBoundX(values, 1) > -1 And LBoundX(values, 2) = -1 Then
        JSONString_List_Assert = True
    End If
End Function

Function JSONString_Dict_Assert(keys As Variant, values As Variant) As Boolean
    JSONString_Dict_Assert = False
    
    If LBoundX(values, 2) > -1 Or LBoundX(keys, 2) > -1 Then
        Exit Function: End If
        
    If LBoundX(keys, 1) <> 1 Then
        Exit Function: End If
    If LBoundX(values, 1) <> 1 Then
        Exit Function: End If
    If UBoundX(values, 1) <> UBoundX(keys, 1) Then
        Exit Function: End If

    JSONString_Dict_Assert = True
End Function


'######################################################################################
' Test                                                                                #
'######################################################################################

Sub Test_JSONString_ListofDicts()
    Dim vals, keys As Variant: keys = test_keys1(): vals = test_vals1()
    Dim str As String, elements As Variant: ReDim elements(LBound(vals) To UBound(vals))
    Dim i As Integer
    
    For i = LBound(vals) To UBound(vals)
        sub1D = SubListFrom2D(vals, i)
        elements(i) = JSONString_Dict(AddQuotes(keys), AddQuotes(sub1D), "    ", 1)
    Next
    
    str = JSONString_List(elements, "    ")

    Call bConfig.Init
    Call bBasis.SaveStringAsTextFile(ByVal str, TARGET_PATH)
    
End Sub

Sub Test_JSONString_Dict()
    Dim vals, keys As Variant: keys = test_keys1(): vals = test_vals1()
    Dim str As String, sub1D As Variant, i As Integer
    
    str = ""
    For i = LBound(vals) To UBound(vals)
        sub1D = SubListFrom2D(vals, i)
        str = str + JSONString_Dict(AddQuotes(keys), AddQuotes(sub1D), "    ")
    Next

    Call bConfig.Init
    Call bBasis.SaveStringAsTextFile(ByVal str, TARGET_PATH)
    
End Sub

Sub Test_DictString()
    Dim arr, keys As Variant: arr = test_arr()
    Dim str, str2 As String: str = test_str()
    
    str2 = JSONString(arr)
    Debug.Print ("Test_DictString: " + CStr(IsEqual(str, str2)))
    
    Call bConfig.Init
    Call bBasis.SaveStringAsTextFile(ByVal str2, TARGET_PATH)
    
    
End Sub

Sub Test_JSONString_Dict_Assert()
    Dim tX, t2, t3 As Boolean

    Dim keys0(5), vals0(5) As Variant
    Dim keysShorter(1 To 3), ThanVals(1 To 5) As Variant
    Dim keys2D(1 To 2, 1 To 3), vals2D(1 To 2, 1 To 3) As Variant
    Dim keys(1 To 5), vals(1 To 5) As Variant

    tX = JSONString_Dict_Assert(keys0, vals): Debug.Print (Not tX)
    tX = JSONString_Dict_Assert(keys, vals0): Debug.Print (Not tX)
    tX = JSONString_Dict_Assert(keysShorter, ThanVals): Debug.Print (Not tX)
    tX = JSONString_Dict_Assert(keys2D, vals): Debug.Print (Not tX)
    tX = JSONString_Dict_Assert(keys, vals2D): Debug.Print (Not tX)
    
    tX = JSONString_Dict_Assert(keys, vals)
    Debug.Print (tX)


End Sub

Private Function test_keys1() As Variant
    Dim ret(1 To 5) As Variant
    ret(1) = "Country"
    ret(2) = "City"
    ret(3) = "River"
    ret(4) = "Person"
    ret(5) = "Food"
    
    test_keys1 = ret
End Function


Private Function test_vals1() As Variant
    Dim ret(2 To 3, 1 To 5) As Variant
    
    ret(2, 1) = "Germany"
    ret(2, 2) = "Berlin"
    ret(2, 3) = "Spree"
    ret(2, 4) = "Peter"
    ret(2, 5) = "Bratwurst"
    
    ret(3, 1) = "France"
    ret(3, 2) = "Paris"
    ret(3, 3) = "Seine"
    ret(3, 4) = "Chanel"
    ret(3, 5) = "Baguette"
    
    test_vals1 = ret
End Function



Private Function test_arr() As Variant
    Dim ret(1 To 3, 1 To 5) As Variant
    ret(1, 1) = "Country"
    ret(1, 2) = "City"
    ret(1, 3) = "River"
    ret(1, 4) = "Person"
    ret(1, 5) = "Food"
    
    ret(2, 1) = "Germany"
    ret(2, 2) = "Berlin"
    ret(2, 3) = "Spree"
    ret(2, 4) = "Peter"
    ret(2, 5) = "Bratwurst"
    
    ret(3, 1) = "France"
    ret(3, 2) = "Paris"
    ret(3, 3) = "Seine"
    ret(3, 4) = "Chanel"
    ret(3, 5) = "Baguette"
    
    test_arr = ret
End Function

Private Function test_str() As String
    Dim ret, ws As String: ret = "": ws = "    "
    
    ret = ret + "[" + Chr(10)
    ret = ret + ws + "{" + Chr(10)
    ret = ret + ws + ws + """Country"": ""Germany""," + Chr(10)
    ret = ret + ws + ws + """City"": ""Berlin""," + Chr(10)
    ret = ret + ws + ws + """River"": ""Spree""," + Chr(10)
    ret = ret + ws + ws + """Person"": ""Peter""," + Chr(10)
    ret = ret + ws + ws + """Food"": ""Bratwurst""" + Chr(10)
    ret = ret + ws + "}," + Chr(10)
    ret = ret + ws + "{" + Chr(10)
    ret = ret + ws + ws + """Country"": ""France""," + Chr(10)
    ret = ret + ws + ws + """City"": ""Paris""," + Chr(10)
    ret = ret + ws + ws + """River"": ""Seine""," + Chr(10)
    ret = ret + ws + ws + """Person"": ""Chanel""," + Chr(10)
    ret = ret + ws + ws + """Food"": ""Baguette""" + Chr(10)
    ret = ret + ws + "}" + Chr(10)
    ret = ret + "]" + Chr(10)
    
    test_str = ret
End Function
