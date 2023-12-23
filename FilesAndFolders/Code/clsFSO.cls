VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objFSO As Scripting.FileSystemObject
Private objFolder As Scripting.Folder
Private objSubFolder As Scripting.Folder
Private objFile As Scripting.File

Private counter As Integer


Private Sub Init(path)
    Dim pathM As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Right(path, 1) = "\" Or Right(path, 1) = "/" Then
        pathM = Left(path, Len(path) - 1)
    Else
        pathM = path
    End If
    Set objFolder = objFSO.GetFolder(pathM)
    counter = 1
End Sub

Function FolderList(path As String, Optional Iterations As Integer = 1, Optional ByRef recursiveArr As Variant, Optional sourcePath As Variant) As Variant
    Call Init(path)
    Dim ret() As String
    If objFolder.SubFolders.Count = 0 Then
         ReDim ret(0): ret(0) = "": Exit Function: End If
    
    ReDim ret(1 To objFolder.SubFolders.Count)
Debug.Print ("FolderList 1")
    For Each objSubFolder In objFolder.SubFolders
        If IsMissing(sourcePath) Then
            ret(counter) = objSubFolder.name
        Else
            ret(counter) = Replace(objSubFolder.path, sourcePath, "")
        End If
        counter = counter + 1
    Next objSubFolder
Debug.Print ("FolderList 2")
    If Iterations > 1 Then
        For i = LBound(ret) To UBound(ret)
            ret = FolderList(path + ret(i) + "\", Iterations - 1, ret, path)
        Next
    End If
Debug.Print ("FolderList 3")
    If IsMissing(recursiveArr) Then
        FolderList = ret
    Else
        Dim rmb As Integer: rmb = UBound(recursiveArr)
        ReDim Preserve recursiveArr(1 To UBound(recursiveArr) + UBound(ret))
        For i = LBound(ret) To UBound(ret)
            recursiveArr(i + rmb) = ret(i)
        Next
        FolderList = recursiveArr
    End If

End Function

'Returns a list of all files of a given path
Function FileList(path As String) As Variant
    Dim ret() As String
    
    'variables
    counter = 1
    Call Init(path)
    
    ReDim ret(1 To objFolder.Files.Count)

    'MAIN
    For Each objFile In objFolder.Files
        ret(counter) = objFile.name
        counter = counter + 1
    Next objFile
    
    FileList = ret

End Function
