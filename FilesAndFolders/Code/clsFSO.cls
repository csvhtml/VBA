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
Debug.Print ("FolderList 1. Iteration: " & CStr(Iterations))
    
    Dim retName() As Variant: ReDim retName(0)
    If IsExisting(recursiveArr) Then
        retName = recursiveArr: End If
    If IsMissing(sourcePath) Then
        sourcePath = path: End If
    Call Init(path)
    
    Dim tmp As String
    For Each objSubFolder In objFolder.SubFolders
        tmp = Replace(objSubFolder.path, sourcePath, "")
        Debug.Print ("FolderList 2. Iteration: " & CStr(Iterations) & CStr(tmp))
        retName = bBasis.PushToArr(retName, tmp)
        If Iterations > 1 Then
            retName = FolderList(sourcePath + tmp + "\", Iterations - 1, retName, sourcePath): End If
    Next objSubFolder
    
    FolderList = retName

End Function


'Returns a list of all files of a given path
Function FileList(path As String, Optional Iterations As Integer = 1, Optional ByRef recursiveArr As Variant, Optional sourcePath As Variant) As Variant
    Dim folders As Variant: folders = FolderList(path, Iterations)
    Dim retName() As Variant: ReDim retName(0)
    
    'variables
    counter = 1
    Dim item As Variant
    For i = LBound(folders) To UBound(folders)
        Set objFolder = objFSO.GetFolder(path + folders(i))
        For Each objFile In objFolder.Files
            retName = bBasis.PushToArr(retName, objFile.name): Next objFile
    Next
    
'    Call Init(path)
'
'    ReDim ret(1 To objFolder.Files.Count)
'
'    'MAIN
'    For Each objFile In objFolder.Files
'        ret(counter) = objFile.name
'        counter = counter + 1
'    Next objFile
    
    FileList = retName

End Function
