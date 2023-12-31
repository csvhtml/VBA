Attribute VB_Name = "bJSON"
Function DictString(var As Variant, Optional ws As String = "    ") As String
    Dim arr, keys As Variant
    arr = WithQuotes(var)
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
    
    DictString = ret
End Function


Function WithQuotes(myArray As Variant) As Variant
    Dim i, j As Long, ret As Variant
    
    ReDim ret(LBound(myArray, 1) To UBound(myArray, 1), LBound(myArray, 2) To UBound(myArray, 2))

    For i = LBound(myArray, 1) To UBound(myArray, 1)
        For j = LBound(myArray, 2) To UBound(myArray, 2)
            ret(i, j) = """" & myArray(i, j) & """"
        Next j
    Next i
    
    WithQuotes = ret
End Function



'######################################################################################
' Test                                                                                #
'######################################################################################


Sub Test_DictString()
    Dim arr, keys As Variant: arr = arrX()
    Dim str, str2 As String: str = strX()
    
    str2 = DictString(arr)
    Debug.Print ("Test_DictString: " + CStr(IsEqual(str, str2)))
    
    Call bConfig.Init
    Call bBasis.SaveStringAsTextFile(ByVal str2, TARGET_PATH)
    
    
End Sub




Function arrX() As Variant
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
    
    arrX = ret
End Function

Function strX() As String
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
    
    strX = ret
End Function
