Attribute VB_Name = "xmainTest"
Sub mainSaveSheetsAsCSV_001()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht As Worksheet: Set sht = ActiveSheet
    
    Call Init("001")
    Call bFileReader.SaveSheetsAs(SOURCE_FILENAME, TARGET_PATH)
    
    wb.Activate
    sht.Activate
End Sub
