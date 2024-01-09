Attribute VB_Name = "bFileReader"
'######################################################################################
' File Read and Save                                                                  #
'######################################################################################

Function ArrayFromFile(path As String, Optional IndexStart = 0) As String()
    Dim fileNumber As Integer
    Dim textLine As String
    Dim linesArray() As String
    Dim i As Integer: i = 0

    ' Open the text file
    fileNumber = FreeFile
    Open path For Input As fileNumber

    ' Loop through each line in the file
    Do Until EOF(fileNumber)
        Line Input #fileNumber, textLine
        ReDim Preserve linesArray(i + IndexStart)
        linesArray(i + IndexStart) = textLine
        i = i + 1
    Loop
    
    If IndexStart > 0 Then
        Dim linesArrayX() As String: ReDim linesArrayX(IndexStart To UBound(linesArray))
        For i = IndexStart To UBound(linesArray)
            linesArrayX(i) = linesArray(i)
        Next
        linesArray = linesArrayX
    End If

    Close fileNumber

    ArrayFromFile = linesArray
End Function


' The size of a 2D array is fixed in all dimension, i. .e there each subarray has the same size
Function Array2DFromFile(path As String, Optional Delimiter As String = "|", Optional IndexStart = 0) As Variant
    Dim ret() As String
    Dim i As Integer, j As Integer
    Dim tmp() As String, arr1D() As String: arr1D = ArrayFromFile(path, IndexStart)

    tmp = Split(arr1D(LBound(arr1D)), Delimiter) ' take the first line to determine the cols size
    If IndexStart > 0 Then
        tmp = Array_ShiftIndex(tmp, IndexStart, "String"): End If
    ReDim ret(LBound(arr1D) To UBound(arr1D), LBound(tmp) To UBound(tmp))
    
    For i = LBound(arr1D) To UBound(arr1D)
        tmp = Split(arr1D(i), Delimiter)
        If IndexStart > 0 Then
            tmp = Array_ShiftIndex(tmp, IndexStart, "String"): End If
        For j = LBound(tmp) To UBound(tmp)
            ret(i, j) = tmp(j)
        Next j
    Next i

    Array2DFromFile = ret
End Function


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

Function StringFromArray(arr As Variant) As String
    Dim ret As String: ret = ""
    Dim dem As String: dem = "|"
    
    For i = LBound(arr, 1) To UBound(arr, 1)
        For j = LBound(arr, 2) To UBound(arr, 2)
            ret = ret & arr(i, j) & dem
        Next
        ret = RemoveLastCharacters(ret, Len(dem)) + NEWLINE
    Next
    ret = RemoveLastCharacters(ret, Len(NEWLINE))
    
    StringFromArray = ret
End Function

Sub SaveSheetsAs(sourcePath As String, targetPath As String, Optional Ending As String = ".csv", Optional Delimiter As String = "|")
    Dim wb As Workbook, wb_name As String: wb_name = GetFileNameFromPath(sourcePath)
    Dim flag As Boolean: flag = False
    
    If IsWorkbookOpen(wb_name) = False Then
        Set wb = Workbooks.Open(sourcePath)
        flag = True: End If
    
    Set wb = Workbooks(wb_name): wb.Activate

    Dim str, newFileName As String, ws As Worksheet
    For Each ws In wb.Sheets
        ws.Activate
        If Ending = ".csv" Then
            str = StringFromArray(SheetFormulas(ws)): End If
        If Ending = ".json" Then
            str = bJSON.JSONString_List(SheetFormulas(ws), "    "): End If
            
        newFileName = targetPath & ws.Name & Ending
        Call SaveStringAsTextFile(str, newFileName)
    Next ws
    
    If flag Then
        wb.Close: End If
End Sub


'######################################################################################
' Test                                                                                #
'######################################################################################


Sub bFileReader_test()
    '!!!!!!!!!!!!!!!!!!!!
    Dim LocalTestFilePath As String
    LocalTestFilePath = ThisWorkbook.path + "\bFileReader_test.txt"
    '!!!!!!!!!!!!!!!!!!!!
    
    Dim arr, arrT As Variant: arr = xTest.test_arr()
    Dim str As String
    
    str = StringFromArray(arr)
    Call SaveStringAsTextFile(str, LocalTestFilePath)
    arrT = Array2DFromFile(LocalTestFilePath, , 1)
    
    Debug.Print (IsEqual(arr, arrT))
    
End Sub
