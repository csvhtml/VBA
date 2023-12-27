Attribute VB_Name = "mOutput"
Sub WriteToSheet(out1 As Variant, out2 As Variant, out3 As Variant)
    Sheets(SHEET_OUT).Activate
    
    Sheets(SHEET_OUT).Range(Cells(2, 1), Cells(OUTPUT_MAX + 1, 5)).value = ""
    Call WriteHeader
    For i = 1 To minn(UBound(out1), OUTPUT_MAX)
        Sheets(SHEET_OUT).Cells(1 + i, 1).value = out1(i)
        Sheets(SHEET_OUT).Cells(1 + i, 2).value = out2(i)
        Sheets(SHEET_OUT).Cells(1 + i, 3).value = out3(i)
    Next
    
End Sub


Private Sub WriteHeader()
    Sheets(SHEET_OUT).Activate
    Sheets(SHEET_OUT).Cells(1, 1).value = "name"
    Sheets(SHEET_OUT).Cells(1, 2).value = "relative path"
    Sheets(SHEET_OUT).Cells(1, 3).value = "full path"
End Sub
