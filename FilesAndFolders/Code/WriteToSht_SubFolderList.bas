Attribute VB_Name = "WriteToSht_SubFolderList"
'Set a reference to Microsoft Scripting Runtime by using
'Tools > References in the Visual Basic Editor (Alt+F11)

' This Script creates a list of all (Sub-) Folders in a given folder
    
' Force the explicit delcaration of variables
Option Explicit

Sub SubFolderList()
    Call aCONFIG.Init
    
    Dim i As Integer
    Dim Folders As Variant
    Folders = FSO.FolderList(TARGET_PATH, 2)
    If aCONFIG.RECURSIONS > 1 Then
        
    End If
    Dim Files As Variant: Files = FSO.FileList(TARGET_PATH)
    Dim Items As Variant: Items = ItemList_FilesInPath(TARGET_PATH)
    
    Sheets("Tabelle1").Range(Cells(ROW_OUT_START, COL_OUT), Cells(ROW_OUT_END, COL_OUT)).Value = ""
    For i = LBound(Folders) To UBound(Folders)
        If i > ROW_OUT_END Then: Exit Sub
        Sheets("Tabelle1").Cells(ROW_OUT_START + i - 1, COL_OUT).Value = Folders(i)
    Next
    
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
