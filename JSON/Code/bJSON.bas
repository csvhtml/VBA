Attribute VB_Name = "bJSON"
Function DictString(var As Variant) As String

End Function




'######################################################################################
' Test                                                                                #
'######################################################################################


Sub Test_DictString()
    Dim arr As Variant: arr = arrX()
    Dim str As String: str = strX()
    
    
    Call bConfig.Init
    Call bBasis.SaveStringAsTextFile(str, TARGET_PATH)
    
    
End Sub




Function arrX() As Variant
    Dim ret(1 To 5, 1 To 3) As Variant
    ret(1, 1) = "Country"
    ret(2, 1) = "City"
    ret(3, 1) = "River"
    ret(4, 1) = "Person"
    ret(5, 1) = "Food"
    
    ret(1, 2) = "Germany"
    ret(2, 2) = "Berlin"
    ret(3, 2) = "Spree"
    ret(4, 2) = "Peter"
    ret(5, 2) = "Bratwurt"
    
    ret(1, 3) = "France"
    ret(2, 3) = "Paris"
    ret(3, 3) = "Seine"
    ret(4, 3) = "Chanel"
    ret(5, 3) = "Baguette"
    
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
