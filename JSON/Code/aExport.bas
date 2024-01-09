Attribute VB_Name = "aExport"
Private Const EXPORT_PATHVBA = "C:\git\VBA\JSON\Code\"

Sub VBA_Export()
    Dim clnVBAModules_BAS As Collection: Set clnVBAModules_BAS = New Collection
    Dim clnVBAModules_CLS As Collection: Set clnVBAModules_CLS = New Collection
    
    clnVBAModules_BAS.Add ("aExport")
    clnVBAModules_BAS.Add ("bBasis")
    clnVBAModules_BAS.Add ("bConfig")
    clnVBAModules_BAS.Add ("bFileReader")
    clnVBAModules_BAS.Add ("bJSON")
    clnVBAModules_BAS.Add ("mMain")
    clnVBAModules_BAS.Add ("xmainTest")
    clnVBAModules_BAS.Add ("xTest")
    clnVBAModules_BAS.Add ("zBuildSheets")
    clnVBAModules_CLS.Add ("clsFSO")
    
    For i = 1 To clnVBAModules_BAS.Count
        With ThisWorkbook.VBProject.VBComponents(clnVBAModules_BAS(i))
            .Export EXPORT_PATHVBA & .Name & ".bas"
            DoEvents
        End With: Next
    For i = 1 To clnVBAModules_CLS.Count
        With ThisWorkbook.VBProject.VBComponents(clnVBAModules_CLS(i))
            .Export EXPORT_PATHVBA & .Name & ".cls"
            DoEvents
        End With: Next
        
    Set clnVBAModules_BAS = Nothing
    Set clnVBAModules_CLS = Nothing
End Sub


Sub SaveSheetsAsXLCSV()
    Dim ws As Worksheet
    Dim newFileName As String
    Dim filePath As String

    ' Set the path to the current workbook's path
    filePath = ThisWorkbook.path & "\"

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        'Simple Case ###################
        newFileName = filePath & ws.Name & ".csv"
        ws.SaveAs newFileName, xlCSV
        
    Next ws

End Sub


Sub ImportCSVFiles()
    Dim FolderPath As String
    Dim Filename As String
    Dim ws As Worksheet

    ' Set the path to the current workbook's path
    FolderPath = ThisWorkbook.path & "\"

    ' Disable alerts to prevent prompts during file import
    Application.DisplayAlerts = False

    ' Loop through all CSV files in the folder
    Filename = Dir(FolderPath & "*.csv")
    Do While Filename <> ""
        ' Create a new worksheet with the file name (without extension)
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = Left(Filename, Len(Filename) - 4)

        ' Import the CSV file into the new worksheet
        With ws.QueryTables.Add(Connection:="TEXT;" & FolderPath & Filename, Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileOtherDelimiter = "|"
            .TextFileTextQualifier = xlTextQualifierDoubleQuote
            .Refresh
        End With

        ' Get the next CSV file in the folder
        Filename = Dir
    Loop

    ' Enable alerts again
    Application.DisplayAlerts = True
End Sub


