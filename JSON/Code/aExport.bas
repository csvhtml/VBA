Attribute VB_Name = "aExport"
Private Const EXPORT_PATHVBA = "C:\git\VBA\JSON\Code\"

Sub VBA_Export()
    Dim clnVBAModules_BAS As Collection: Set clnVBAModules_BAS = New Collection
    Dim clnVBAModules_CLS As Collection: Set clnVBAModules_CLS = New Collection
    
    clnVBAModules_BAS.Add ("aExport")
    clnVBAModules_BAS.Add ("bBasis")
    clnVBAModules_BAS.Add ("bConfig")
    clnVBAModules_BAS.Add ("bJSON")
    clnVBAModules_BAS.Add ("mMain")

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


Sub SaveSheetsAsCSV()
    Dim ws As Worksheet
    Dim newFileName As String
    Dim filePath As String

    ' Set the path to the current workbook's path
    filePath = ThisWorkbook.path & "\"

    ' Loop through all sheets in the workbook
    For Each ws In ThisWorkbook.Sheets
        ' Create a unique file name for each sheet
        newFileName = filePath & ws.Name & ".csv"

        ' Save the sheet as a CSV file
        ws.SaveAs newFileName, xlCSV
    Next ws

End Sub

Sub ImportCSVFiles()
    Dim folderPath As String
    Dim fileName As String
    Dim ws As Worksheet

    ' Set the path to the current workbook's path
    folderPath = ThisWorkbook.path & "\"

    ' Disable alerts to prevent prompts during file import
    Application.DisplayAlerts = False

    ' Loop through all CSV files in the folder
    fileName = Dir(folderPath & "*.csv")
    Do While fileName <> ""
        ' Create a new worksheet with the file name (without extension)
        Set ws = Sheets.Add(After:=Sheets(Sheets.Count))
        ws.Name = Left(fileName, Len(fileName) - 4)

        ' Import the CSV file into the new worksheet
        With ws.QueryTables.Add(Connection:="TEXT;" & folderPath & fileName, Destination:=ws.Range("A1"))
            .TextFileParseType = xlDelimited
            .TextFileConsecutiveDelimiter = False
            .Refresh
        End With

        ' Get the next CSV file in the folder
        fileName = Dir
    Loop

    ' Enable alerts again
    Application.DisplayAlerts = True
End Sub


