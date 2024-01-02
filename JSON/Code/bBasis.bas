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

Function IsEqual(a As Variant, b As Variant) As Boolean
    IsEqual = True
    
    If IsString(a) And IsString(b) Then
        If Len(a) <> Len(a) Then
            IsEqual = False: Exit Function: End If
        For i = 1 To Len(a)
            If Mid(a, i, 1) <> Mid(b, i, 1) Then
                IsEqual = False: Exit Function: End If: Next
    End If
    
End Function

Function IsString(myVariant As Variant) As Boolean
    IsString = (VarType(myVariant) = vbString)
End Function


Function ItemList(n As Integer) As Variant
    Dim Items() As clsItem
    ReDim Items(n)
    Dim i As Integer

    For i = 1 To 3
        Set Items(i) = New clsItem
        Items(i).Name = "name " & CStr(i)
        Items(i).url = "http " & CStr(i)
    Next i

    ItemList = Items
End Function

Function UBoundX(arr As Variant, n As Integer) As Long
    UBoundX = -1
    On Error Resume Next
    UBoundX = UBound(arr, n)
    On Error GoTo 0
End Function

Function LBoundX(arr As Variant, n As Integer) As Long
    LBoundX = -1
    On Error Resume Next
    LBoundX = LBound(arr, n)
    On Error GoTo 0
End Function

Function SubListFrom2D(arr As Variant, n As Integer) As Variant
    If UBoundX(arr, 2) = -1 Then
        SubListFrom2D = -1: Exit Function: End If

    
    Dim ret As Variant
    ReDim ret(LBound(arr, 2) To UBound(arr, 2))
    
    For i = LBound(arr, 2) To UBound(arr, 2)
        ret(i) = arr(n, i)
    Next
    
    ' Return the result array
    SubListFrom2D = ret
End Function

Function AddQuotes(var As Variant) As Variant
    Dim str As String
    If IsString(var) Then
        str = CStr(var)
        AddQuotes = AddQuotes_ToString(str): Exit Function: End If
    
    If LBoundX(var, 1) > -1 And LBoundX(var, 2) = -1 Then
        AddQuotes = AddQuotes_ToList(var): Exit Function: End If

    If LBoundX(var, 2) > -1 And LBound(var, 3) = -1 Then
        Dim i, j As Integer, ret As Variant
        ReDim ret(LBound(var, 1) To UBound(var, 1), LBound(var, 2) To UBound(var, 2))
        For i = LBound(var, 1) To UBound(var, 1)
            For j = LBound(var, 2) To UBound(var, 2)
                str = CStr(var(i, j))
                ret(i, j) = AddQuotes_ToString(str)
            Next j
        Next i
    End If
    
    AddQuotes = ret
End Function

Private Function AddQuotes_ToString(str As String) As String
    AddQuotes_ToString = """" & str & """"
End Function

Private Function AddQuotes_ToList(arr As Variant) As Variant
    Dim str As String
    Dim ret As Variant: ReDim ret(LBound(arr) To UBound(arr))
    For i = LBound(arr) To UBound(arr)
        str = CStr(arr(i))
        ret(i) = AddQuotes_ToString(str)
    Next
    AddQuotes_ToList = ret
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

Function RemoveLastCharacters(inputString As String, n As Integer) As String
    RemoveLastCharacters = Left(inputString, Len(inputString) - n)
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
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
    End If
End Sub

'via CchatGPT
Function CellPositionA1(rowNum As Long, colNum As Long) As String
    Dim colLetter As String
    colLetter = Split(Cells(1, colNum).Address, "$")(1)
    CellPositionA1 = colLetter & rowNum
End Function

Function maxRange(sht As Worksheet) As Range
   Dim a, ret  As Range, col, row As Long
   
   Set a = sht.UsedRange: col = a.Columns.Count: row = a.Rows.Count
   Set maxRange = Range(Cells(1, 1), Cells(row + 1, col + 1))
End Function

Function SheetValues(sht As Worksheet) As Variant
    ' when the range is assigned to a vairant variabel it becomes an 2D array with the cell values (not Formulass)
    Set SheetValues = maxRange(sht)
End Function

Function SheetFormulas(sht As Worksheet) As Variant
    ' in case of no formula then the cell value is taken
    
    Dim rng As Range, ret As Variant
    
    Set rng = maxRange(sht)
    ret = maxRange(sht)
    
    For i = 1 To UBound(ret, 1)
        For j = 1 To UBound(ret, 2)
            ret(i, j) = rng.Cells(i, j).FormulaR1C1
        Next
    Next

    SheetFormulas = ret
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

'######################################################################################
' File Read and Save                                                                  #
'######################################################################################


Sub SaveStringAsTextFile(ByVal myString As String, filePath As String)
    Dim fileNumber As Integer

    ' Open the file for writing
    fileNumber = FreeFile
    Open filePath For Output As fileNumber

    ' Write the string to the file
    Print #fileNumber, myString

    ' Close the file
    Close fileNumber
End Sub



