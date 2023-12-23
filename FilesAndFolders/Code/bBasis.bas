Attribute VB_Name = "bBasis"
Function InArray(element As String, arr As Variant) As Integer
    InArray = -1
    
    For i = LBound(arr) To UBound(arr)
        If arr(i) = element Then
            InArray = i
            Exit Function
        End If
    Next

End Function


Function ItemList(n As Integer) As Variant
    Dim Items() As clsItem
    ReDim Items(n)
    Dim i As Integer

    For i = 1 To 3
        Set Items(i) = New clsItem
        Items(i).name = "name " & CStr(i)
        Items(i).url = "http " & CStr(i)
    Next i

    ' Return the array of Person objects
    ItemList = Items
End Function

' via CHATGPT
Function GetRightPart(inputText As String, startingWord As String) As String
    Dim startPos As Long
    Dim result As String
    
    ' Find the starting position of the word
    startPos = InStr(1, inputText, startingWord, vbTextCompare)
    
    ' If the word is found, return the right part of the text
    If startPos > 0 Then
        result = Mid(inputText, startPos + Len(startingWord))
    Else
        result = "Word not found"
    End If
    
    GetRightPart = result
End Function

Function GetLeftPart(inputText As String, endingWord As String) As String
    Dim endPos As Long
    Dim result As String
    
    ' Find the starting position of the word
    endPos = InStr(1, inputText, endingWord, vbTextCompare)
    
    ' If the word is found, return the right part of the text
    If endPos > 0 Then
        result = Left(inputText, endPos - 1)
    Else
        result = "Word not found"
    End If
    
    GetLeftPart = result
End Function