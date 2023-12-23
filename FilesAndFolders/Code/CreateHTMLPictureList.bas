Attribute VB_Name = "CreateHTMLPictureList"
'Replace "2016" with correct year/subfolder

'Force the explicit delcaration of variables
Option Explicit
Public Count As Integer

Public sXML As String


Sub FilesInFolder()
    PARA_OFFSET = 1
    Call InitConfig
    
    Dim i As Integer
    Dim Files As Variant: Files = FileList(TARGET_PATH)
    Dim Items As Variant: Items = ItemList_FilesInPath(TARGET_PATH)
    
    Sheets("Tabelle1").Range(Cells(ROW_OUT_START, COL_OUT + PARA_OFFSET), Cells(ROW_OUT_END, COL_OUT + PARA_OFFSET)).Value = ""
    For i = LBound(Items) To UBound(Items)
        If i > ROW_OUT_END * 10 Then: Exit Sub
        Sheets("Tabelle1").Cells(ROW_OUT_START + i - 1, COL_OUT + PARA_OFFSET).Value = Items(i).name
        Sheets("Tabelle1").Cells(ROW_OUT_START + i - 1, COL_OUT + 1 + PARA_OFFSET).Value = HTML_Path_Relative(Items(i).path + Items(i).name)
    Next

End Sub



Sub PicturesToHTML()
    Call InitConfig
    
    'File System Object Libs
    Dim objFSO As Scripting.FileSystemObject
    Dim objFolder, objSubFolder As Scripting.Folder
    

    
    'DIM
    Dim Q, E, name, ThisPath As String
    Dim i As Integer
    
    'SET
    Count = 0
    Q = Chr(34)   ' = "
    E = Chr(10)    ' = Enter
    sXML = "var daten = " & E & "[" & E
'    ThisPath = Application.ActiveWorkbook.Path
    ThisPath = TARGET_PATH
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    
    'MAIN
    For i = ROW_OUT_START To ROW_OUT_END
    name = Sheets("Tabelle1").Cells(i, COL_OUT)
        If name <> "" Then
            '1 Write Folder Header <h6>
            'sXML = sXML & E & "<h6 id=" & Q & i - 1 & Q & " >" & Name & "</h6>" '& E & "<p>" & E
            
            '2 Prepare for 2
            ThisPath = ThisPath & "\" & name
            Set objFolder = objFSO.GetFolder(ThisPath)
            
            '2 Read images in Folder and Write img links
            Call RecursiveFolder(ByVal objFolder, ThisPath)
            
            '3 Write End of table
            sXML = Left(sXML, Len(sXML) - 1) & E & "]"
            
            'Reset ThisPath
            ThisPath = Application.ActiveWorkbook.path
            
            
        End If
    Next
    
    'Write Total Number of Fotos
    'sXML = sXML & E & E & "Summe: " & Count & " Fotos"
    
    'Write to HTML via FileSystemObjects
    Dim FSO, Fileout As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set Fileout = FSO.CreateTextFile(Application.ActiveWorkbook.path & "\daten.json", True, True)
    Fileout.Write sXML
    Fileout.Close
    
End Sub

Sub RecursiveFolder(ByVal objFolder As Scripting.Folder, ThisPath As String)

    'DIM
    Dim objFile As Scripting.File
    Dim objSubFolder As Scripting.Folder
    Dim counter, IndexofSlash As Long
    Dim name, path, tempPath, Q, E, PathforSmall, PathFull, tempXML As String
    
    'SET
    Q = Chr$(34) ' = "
    E = Chr(10) ' = Enter
    tempXML = ""
    counter = 0
    
    'Loop through each file in the folder
    For Each objFile In objFolder.Files
        name = objFile.name
        If InArray(Right(name, 3), FORMATS) > -1 Then
            counter = counter + 1
            tempPath = objFile.path
            path = getParentFolder(objFile.path)
            path = Replace(tempPath, path & "\", "")
            
            PathforSmall = Replace(path, "\", "/small-")
            
            'ADDON when super resizer is used
            IndexofSlash = InStr(1, PathforSmall, "/")
            PathforSmall = Right(PathforSmall, Len(PathforSmall) - IndexofSlash + 1)
            
            path = Replace(path, "\", "/")
            'MODIFY YEAR (Must be changed via repalce manually)
            PathFull = "Foto Album privat/2018/" & path
            'PathFull = "XXFolderXX" & Path
            
            'CHANGE when super resizer is used
            'tempXML = tempXML & E & "<li class=" & Q & "li-img" & Q & "><a href=" & Q & PathFull & Q & "><img src=" & Q & "html/pics/" & PathforSmall & Q & "></a></li>"
            'tempXML = tempXML & E & "<li class=" & Q & "li-img" & Q & "><a href=" & Q & PathFull & Q & "><img src=" & Q & "html/pics/2015" & PathforSmall & Q & "></a></li>"
            tempXML = tempXML & "{" & E & Q & "url" & Q & ": " & Q & PathFull & Q & E & "},"
        End If
    Next objFile
    
    Count = Count + counter
    'Write Number Photos
    'sXML = sXML & counter & " Fotos"
    sXML = sXML & tempXML
    
'EXPANSION FOR SUBFOLDERS
'    'Loop through files in the subfolders
'    If IncludeSubFolders Then 'Variable "IncludeSubFolders" must be defined in funcion makro
'        For Each objSubFolder In objFolder.SubFolders
'            Call RecursiveFolder(objSubFolder, True, ThisPath, sXML)
'        Next objSubFolder
'    End If
    
End Sub

Function HTML_Path_Relative(Absolutpath) As String
    Dim path As String
    path = Replace(Absolutpath, "\", "/")
    path = HTML_PREFIX_PATH + "/..//" + HTML_PREFIX_PATH + GetRightPart(path, HTML_PREFIX_PATH)
'    "Foto Album privat/..//Foto Album privat/"

    HTML_Path_Relative = path
End Function

Function getParentFolder(ByVal strFolder0)
  Dim strFolder
  strFolder = Left(strFolder0, InStrRev(strFolder0, "\") - 1)
  getParentFolder = Left(strFolder, InStrRev(strFolder, "\") - 1)
End Function
