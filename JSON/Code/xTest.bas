Attribute VB_Name = "xTest"
Function test_arr() As Variant
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

Function test_arr1() As Variant
    Dim ret(1 To 3) As Variant
    ret(1) = "UK"
    ret(2) = "Germany"
    ret(3) = "France"

    test_arr1 = ret
End Function

