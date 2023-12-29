Attribute VB_Name = "mOutput"
Sub WriteToSheet(out1 As Variant, out2 As Variant, out3 As Variant)
    Sheets(SHEET_OUT).Activate
    
    Sheets(SHEET_OUT).Range(Cells(2, 1), Cells(OUTPUT_MAX + 1, 5)).Value = ""
    Call WriteHeader
    For i = 1 To minn(UBound(out1), OUTPUT_MAX)
        Sheets(SHEET_OUT).Cells(1 + i, 1).Value = out1(i)
        Sheets(SHEET_OUT).Cells(1 + i, 2).Value = out2(i)
        Sheets(SHEET_OUT).Cells(1 + i, 3).Value = out3(i)
    Next
    
End Sub


Private Sub WriteHeader()
    Sheets(SHEET_OUT).Activate
    Sheets(SHEET_OUT).Cells(1, 1).Value = "name"
    Sheets(SHEET_OUT).Cells(1, 2).Value = "relative path"
    Sheets(SHEET_OUT).Cells(1, 3).Value = "full path"
End Sub
