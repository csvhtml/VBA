Attribute VB_Name = "aExport"
Private Const EXPORT_PATHVBA = "C:\git\VBA\JSON\Code\"

Sub VBA_Export()
    Dim clnVBAModules_BAS As Collection: Set clnVBAModules_BAS = New Collection
    Dim clnVBAModules_CLS As Collection: Set clnVBAModules_CLS = New Collection
    
    clnVBAModules_BAS.Add ("aExport")
    clnVBAModules_BAS.Add ("bBasis")
    clnVBAModules_BAS.Add ("bConfig")
    clnVBAModules_BAS.Add ("mMain")

    clnVBAModules_CLS.Add ("clsFSO")
    
    For i = 1 To clnVBAModules_BAS.Count
        With ThisWorkbook.VBProject.VBComponents(clnVBAModules_BAS(i))
            .Export EXPORT_PATHVBA & .name & ".bas"
            DoEvents
        End With: Next
    For i = 1 To clnVBAModules_CLS.Count
        With ThisWorkbook.VBProject.VBComponents(clnVBAModules_CLS(i))
            .Export EXPORT_PATHVBA & .name & ".cls"
            DoEvents
        End With: Next
        
    Set clnVBAModules_BAS = Nothing
    Set clnVBAModules_CLS = Nothing
End Sub


