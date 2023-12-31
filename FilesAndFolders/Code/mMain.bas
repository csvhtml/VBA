Attribute VB_Name = "mMain"
Sub runMain()
Attribute runMain.VB_ProcData.VB_Invoke_Func = " \n14"
    mFolderList.FolderList

End Sub

'-----------------------------------------------------------------------------------
'Code for new sheet 'run'. 'Alternative is to run 'runMain' manually within the VBA Editor
'
'Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
'    If Target.name = "List Folders/Files" Then
'        Call mFolderList.FolderList: End If
'
'End Sub
'-----------------------------------------------------------------------------------
Private Sub BuildMainSheet()

    AddSheetIfNotExists (SHEET_RUN)
    Worksheets(SHEET_RUN).Activate
    Sheets(SHEET_RUN).Range(Cells(1, 1), Cells(100, 100)).Value = ""
    Sheets(SHEET_RUN).Range("B1").Select
    ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:="", SubAddress:= _
        "run!B1", TextToDisplay:="List Folders/Files"
    
    Dim sheetContent(1 To 4, 1 To 3) As Variant
    sheetContent(1, 1) = ""
    sheetContent(2, 1) = "[Folders] or [Files]"
    sheetContent(3, 1) = "Depth [1..n]"
    sheetContent(4, 1) = "Target Path"

    sheetContent(1, 2) = "List Folders/Files" ' keep Hyperlink
    sheetContent(2, 2) = "Folders"
    sheetContent(3, 2) = "1"
    sheetContent(4, 2) = ""

    sheetContent(1, 3) = "Ego Path (for information only. Not used by script)"
    sheetContent(2, 3) = "=CELL(""dateiname"")"
    sheetContent(3, 3) = "=FIND(""["",R[-1]C[0])"
    sheetContent(4, 3) = "=LEFT(R[-2]C[0],R[-1]C[0]-1)"

    For i = 1 To 4
        For j = 1 To 3
            If InStrRev(sheetContent(i, j), "=") Then
                Sheets(SHEET_RUN).Cells(i, j).FormulaR1C1 = sheetContent(i, j)
            Else
                Sheets(SHEET_RUN).Cells(i, j).Value = sheetContent(i, j)
            End If
        Next
    Next
    
    Call ApplyFormat
End Sub

Private Sub ApplyFormat()
    Range("A1:C4").Interior.Color = RGB(200, 200, 200)
    Range("A1:C4").Rows.RowHeight = 40
    Range("A1").Columns.ColumnWidth = 20
    Range("B1:C1").Columns.ColumnWidth = 60
    Range("B1:B1").Interior.Color = RGB(150, 180, 215)
    Range("C1:C4").Font.Color = RGB(150, 150, 150)
    Range("B2:B4").Interior.ColorIndex = -4142 ' ColorIndex for no fill
    
    Range("A1:C4").VerticalAlignment = xlCenter
    
    Call SetBorder(Range("A1:C4"))
End Sub

Private Sub SetBorder(rng As Range)
    Dim arr() As Variant: ReDim arr(1 To 6)
    
    arr(1) = xlEdgeLeft
    arr(2) = xlEdgeRight
    arr(3) = xlEdgeTop
    arr(4) = xlEdgeBottom
    arr(5) = xlInsideVertical
    arr(6) = xlInsideHorizontal
    
    For i = 1 To 6
        rng.Borders(arr(i)).LineStyle = xlContinuous
        rng.Borders(arr(i)).Weight = xlThin
    Next
    
End Sub
