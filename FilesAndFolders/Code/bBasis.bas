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

    ItemList = Items
End Function

Function PushToArr(arr As Variant, item As Variant) As Variant
    Dim lastIndex As Integer
    lastIndex = UBound(arr)
    ReDim Preserve arr(lastIndex + 1)
    arr(lastIndex + 1) = item
    PushToArr = arr
End Function

Function IsExisting(Optional item As Variant) As Boolean
     IsExisting = True
     If IsMissing(item) Then
        IsExisting = False: End If
        
End Function

Function IsCharInString(searchChar As String, searchString As String) As Boolean
    IsCharInString = InStrRev(searchString, searchChar) > 0
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

'via ChatGPT
Sub AddSheetIfNotExists(sheetName As String)
    Dim ws As Worksheet

    ' Check if the sheet already exists
    On Error Resume Next
    Set ws = Worksheets(sheetName)
    On Error GoTo 0

    ' If the sheet doesn't exist, add it
    If ws Is Nothing Then
        Sheets.Add(After:=Sheets(Sheets.Count)).name = sheetName
    End If
End Sub

'via CchatGPT
Function CellPositionA1(rowNum As Long, colNum As Long) As String
    Dim colLetter As String
    colLetter = Split(Cells(1, colNum).Address, "$")(1)
    CellPositionA1 = colLetter & rowNum
End Function


Function maxx(a, b) As Long
    maxx = a
    If a < b Then
        maxx = b: End If
End Function

Function minn(a, b) As Long
    minn = a
    If a > b Then
        minn = b: End If
End Function


