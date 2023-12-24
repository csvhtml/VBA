Attribute VB_Name = "bCONFIG"
Public Const ROW_OUT_START = 11  ' Row start in which the Script Output is written
Public Const ROW_OUT_END = 9000   ' Row end in which the Script Output is written
Public Const COL_OUT = 2    ' Col in which the Script Output is written
Public Const COL_PARA = 2       ' Col where Script parameters are defined
Public Const ROW_PARA_PATH = 7  ' Row of target Path Parameter
Public Const HTML_PREFIX_PATH = "Foto Album privat"  ' Source Path for relative html references of fiels/images in the html file.

Public TYPE_OUTPUT As String
Public RECURSIONS As Integer
Public TARGET_PATH As String      ' Target Path to which the Scripts shall apply


Public FORMATS As Variant
Public FSO As clsFSO

Sub Init()

    TYPE_OUTPUT = TypeFunction()
    RECURSIONS = Cells(8, COL_PARA).value
    TARGET_PATH = path(Cells(9, COL_PARA).value)
    Set FSO = New clsFSO
    
    'FORMATS
    Dim a() As String
    ReDim a(3)
    a(0) = "jpg"
    a(1) = "JPG"
    a(2) = "png"
    a(3) = "PNG"
    
    FORMATS = a
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
    Dim ret As String: ret = Replace(GetLeftPart(Cells(7, COL_PARA).value, Chr(10)), " ", "")
    If ret = "Files" Then
        TypeFunction = "Files": End If
    If ret = "Folders and Files" Then
        TypeFunction = "Folders and Files": End If

End Function
