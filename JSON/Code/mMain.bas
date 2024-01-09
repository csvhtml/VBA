Attribute VB_Name = "mMain"
Sub runMain()
Attribute runMain.VB_ProcData.VB_Invoke_Func = " \n14"
   Dim a, b As Variant
   Dim wb As Workbook, sht As Worksheet
   Call bConfig.Init
   
   Set wb = Workbooks(SOURCE_FILENAME): wb.Activate
   Set sht = wb.Worksheets(SOURCE_SHEETNAME)
   a = bBasis.SheetValues(sht)
   b = bBasis.SheetFormulas(sht)
   
   WB_EGO.Activate

End Sub


