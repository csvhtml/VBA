Attribute VB_Name = "xmainTest"
Sub mainSaveSheetsAsJSON_004()
    'Prepare test
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim vals() As String: vals = RememberValues
    
    Cells(ROW_PARA, COL_PARA).Value = "..\Test\004\test-004.xlsx"
    Cells(ROW_PARA + 2, COL_PARA).Value = "..\Test\004\"
    Call Init
    
    'Call Function to be tested
    Call bFileReader.SaveSheetsFormat(SOURCE_FILENAME, TARGET_PATH, True)
    Call bFileReader.SaveSheetsFormat(SOURCE_FILENAME, TARGET_PATH, False)
    
    ' bring back old condition
    wb.Activate
    sht.Activate
    Call SetValues(vals)
End Sub

Sub mainSaveSheetsAsJSON_003()
    'Prepare test
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim vals() As String: vals = RememberValues
    
    Cells(ROW_PARA, COL_PARA).Value = "..\Test\003\test-003.xlsx"
    Cells(ROW_PARA + 2, COL_PARA).Value = "..\Test\003\"
    Call Init
    
    'Call Function to be tested
    Call bFileReader.SaveSheetsAs(SOURCE_FILENAME, TARGET_PATH, ".json", , True)
    
    ' bring back old condition
    wb.Activate
    sht.Activate
    Call SetValues(vals)
End Sub

Sub mainSaveSheetsAsJSON_002()
    'Prepare test
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim vals() As String: vals = RememberValues
    
    Cells(ROW_PARA, COL_PARA).Value = "..\Test\002\test-002.xlsx"
    Cells(ROW_PARA + 2, COL_PARA).Value = "..\Test\002\"
    Call Init
    
    'Call Function to be tested
    Call bFileReader.SaveSheetsAs(SOURCE_FILENAME, TARGET_PATH, ".json")
    
    ' bring back old condition
    wb.Activate
    sht.Activate
    Call SetValues(vals)
End Sub

Sub mainSaveSheetsAsCSV_001()
    'Prepare test
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sht As Worksheet: Set sht = ActiveSheet
    Dim vals() As String: vals = RememberValues
    
    Cells(ROW_PARA, COL_PARA).Value = "..\Test\001\test-001.xlsx"
    Cells(ROW_PARA + 2, COL_PARA).Value = "..\Test\001\"
    Call Init
    
    'Call Function to be tested
    Call bFileReader.SaveSheetsAs(SOURCE_FILENAME, TARGET_PATH)
    
    ' bring back old condition
    wb.Activate
    sht.Activate
    Call SetValues(vals)
End Sub

Private Function RememberValues() As String()
    Dim vals() As String: ReDim vals(0 To 100)
    
    For i = LBound(vals) To UBound(vals)
        vals(i) = Cells(ROW_PARA + i, COL_PARA).Value
    Next
    RememberValues = vals
    
End Function

Private Sub SetValues(vals() As String)
    For i = LBound(vals) To UBound(vals)
        Cells(ROW_PARA + i, COL_PARA).Value = vals(i)
    Next
    
End Sub
