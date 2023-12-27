Attribute VB_Name = "bCONFIG"
Public Const SHEET_RUN = "run"
Public Const SHEET_OUT = "out"
Public Const COL_PARA = 2
Public Const ROW_PARA = 2
Public Const ROW_PARA_PATH = 4

Public Const OUTPUT_MAX = 9000

Public Const HTML_PREFIX_PATH = "Foto Album privat"  ' Source Path for relative html references of fiels/images in the html file.

Public TYPE_OUTPUT As String
Public RECURSIONS As Integer
Public TARGET_PATH As String      ' Target Path to which the Scripts shall apply


Public FORMATS As Variant
Public FSO As clsFSO

Sub Init()

    
    TYPE_OUTPUT = TypeFunction()
    RECURSIONS = IterationsDetph()
    TARGET_PATH = path(Cells(ROW_PARA_PATH, COL_PARA).value)
    Set FSO = New clsFSO
    
    'FORMATS
    Dim a() As String
    ReDim a(3)
    a(0) = "jpg"
    a(1) = "JPG"
    a(2) = "png"
    a(3) = "PNG"
    
    FORMATS = a
    
    AddSheetIfNotExists (SHEET_OUT)
End Sub

Private Function path(pathh As String) As String
    path = pathh
    If Right(pathh, 1) = "\" Or Right(pathh, 1) = "/" Then
        Exit Function: End If
    If InStr(pathh, "/") > 0 Then
        path = pathh + "/"
        Exit Function: End If
    If InStr(pathh, "\") > 0 Then
        path = pathh + "\"
        Exit Function: End If
End Function

Private Function TypeFunction() As String
    TypeFunction = "Folders"
    Dim ret As String: ret = Replace(GetLeftPart(Cells(ROW_PARA, COL_PARA).value, Chr(10)), " ", "")
    If ret = "Files" Then
        TypeFunction = "Files": End If
    If ret = "Folders and Files" Then
        TypeFunction = "Folders and Files": End If

End Function

Private Function IterationsDetph() As Integer
    IterationsDetph = 1
    Dim ret As String: ret = CStr(Cells(ROW_PARA + 1, COL_PARA).value)
    
    If Len(ret) = 1 And IsCharInString(ret, "123456789") Then
        IterationsDetph = CInt(ret): End If

End Function



