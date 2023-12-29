Attribute VB_Name = "mFolderList"
'Set a reference to Microsoft Scripting Runtime by using
'Tools > References in the Visual Basic Editor (Alt+F11)

' This Script creates a list of all (Sub-) Folders in a given folder
    
' Force the explicit delcaration of variables
Option Explicit

Sub FolderList()
Debug.Print (Chr(10) & "-----------------------------------------------------" & Chr(10))
    Call bConfig.Init
    
    Dim i As Integer
    Dim outName, outRelative, outFull As Variant
    If bConfig.TYPE_OUTPUT = "Files" Then
        outName = FSO.FileList(TARGET_PATH, RECURSIONS, , , "Name")
        outRelative = FSO.FileList(TARGET_PATH, RECURSIONS, , , "Relative")
        outFull = FSO.FileList(TARGET_PATH, RECURSIONS, , , "Full")
    Else
        outName = FSO.FolderList(TARGET_PATH, RECURSIONS, , , "Name")
        outRelative = FSO.FolderList(TARGET_PATH, RECURSIONS, , , "Relative")
        outFull = FSO.FolderList(TARGET_PATH, RECURSIONS, , , "Full")
    End If
    
    Call mOutput.WriteToSheet(outName, outRelative, outFull)
    
End Sub

Sub recursiveList(liste As Variant, n As Integer)


'Returns a list of all files of a given path
Function ItemList_FilesInPath(path As String) As Variant
    Dim Items() As clsItem
    
    'File System Object Libs
    Dim objFSO As Scripting.FileSystemObject
    Dim objFolder As Scripting.Folder
    Dim objFile As Scripting.File
    
    'variables
    Dim counter As Integer: counter = 1
    
    ' FSO instances
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFolder = objFSO.GetFolder(path)
    
    ReDim Items(1 To objFolder.Files.Count)

    'MAIN
    For Each objFile In objFolder.Files
        Set Items(counter) = New clsItem
        Items(counter).name = objFile.name
        Items(counter).path = objFile.ParentFolder.path & "\"
        counter = counter + 1
    Next objFile
    
    ItemList_FilesInPath = Items

End Function
