Attribute VB_Name = "bConfig"
Public Const EGO_WB_NAME = "JSON.xlsm"
Public Const EGO_SHT_NAME = "JSON"
Public Const COL_PARA = 2
Public Const ROW_PARA = 2
Public Const ROW_PARA_PATH = 4

Public SOURCE_FILENAME As String
Public SOURCE_SHEETNAME As String
Public TARGET_PATH As String
Public HEADERS_JS As String
Public TYPE_JS As String

Public FSO As clsFSO

Sub Init()
    Debug.Assert ActiveWorkbook.Name = EGO_WB_NAME
    SOURCE_FILENAME = Init_Path(Cells(ROW_PARA, COL_PARA).Value)
    SOURCE_SHEETNAME = Cells(ROW_PARA + 1, COL_PARA).Value
    TARGET_PATH = Init_Path(Cells(ROW_PARA + 2, COL_PARA).Value)
    HEADERS_JS = TypeJS()
    TYPE_JS = HeadersJS()
    Set FSO = New clsFSO
End Sub

Function Init_Path(path As String) As String
    Dim ret As String: ret = path
    
    If Left(path, 3) = "..\" Then
        ret = ThisWorkbook.path & "\" & Mid(path, 4): End If
    
    Init_Path = ret
End Function


Private Function HeadersJS() As Long
    HeadersJS = "0"
    Dim ret, val As String
    
    val = Cells(ROW_PARA + 3, COL_PARA).Value
    If Left(val, 4) = "Row " Then
        ret = bBasis.GetRightPart(val, "Row ")
        HeadersJS = CInt(ret)
    End If

End Function

Private Function TypeJS() As String
    TypeJS = "List"

End Function

Private Function IterationsDetph() As Integer
    IterationsDetph = 1
    Dim ret As String: ret = CStr(Cells(ROW_PARA + 1, COL_PARA).Value)
    
    If Len(ret) = 1 And IsCharInString(ret, "123456789") Then
        IterationsDetph = CInt(ret): End If

End Function



