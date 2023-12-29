VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFSO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objFSO As Scripting.FileSystemObject  ' Activate in VBA Menu: 'Extras' -> 'Verweise/references'-> 'Microsoft Scripting Runtime' to make this line work
Private objFolder As Scripting.Folder
Private objSubFolder As Scripting.Folder
Private objFile As Scripting.File


Private Sub Init(path)
    Dim pathM As String
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Right(path, 1) = "\" Or Right(path, 1) = "/" Then
        pathM = Left(path, Len(path) - 1)
    Else
        pathM = path
    End If
    Set objFolder = objFSO.GetFolder(pathM)

End Sub

Function FolderList(path As String, Optional Iterations As Integer = 1, Optional ByRef recursiveArr As Variant, Optional sourcePath As Variant, _
    Optional NameRelativeFullKey As String) As Variant
    
    Dim retName() As Variant: ReDim retName(0)
    If IsExisting(recursiveArr) Then
        retName = recursiveArr: End If
    If IsMissing(sourcePath) Then
        sourcePath = path: End If
    If IsMissing(NameRelativeFullKey) Then
        NameRelativeFullKey = "Path": End If
    Call Init(path)
    
    Dim tmp, nextPath As String
    For Each objSubFolder In objFolder.SubFolders
        tmp = NameRelativeFullX(objSubFolder, NameRelativeFullKey, sourcePath)
        retName = bBasis.PushToArr(retName, tmp)
        If Iterations > 1 Then
            nextPath = NameRelativeFullX(objSubFolder, "Full")
            retName = FolderList(nextPath + "\", Iterations - 1, retName, sourcePath, NameRelativeFullKey): End If
    Next objSubFolder
    
    FolderList = retName

End Function


Function FileList(path As String, Optional Iterations As Integer = 1, Optional ByRef recursiveArr As Variant, Optional sourcePath As Variant, _
    Optional NameRelativeFullKey As String) As Variant
    Dim folders As Variant: folders = FolderList(path, Iterations)
    Dim retName() As Variant: ReDim retName(0)
    If IsMissing(sourcePath) Then
        sourcePath = path: End If
        
    Dim tmp As String
    For i = 1 To UBound(folders)
        Set objFolder = objFSO.GetFolder(folders(i))
        Debug.Print (objFolder.path)
        For Each objFile In objFolder.Files
            tmp = NameRelativeFullX(objFile, NameRelativeFullKey, sourcePath)
            retName = bBasis.PushToArr(retName, tmp): Next objFile
    Next
    
    FileList = retName

End Function

Function NameRelativeFullX(obj As Variant, key As String, Optional sourcePath As Variant) As String
    NameRelativeFullX = obj.path
    If key = "Relative" Then
        NameRelativeFullX = Replace(obj.path, sourcePath, ""): End If
    If key = "Name" Then
        NameRelativeFullX = obj.name: End If

End Function
