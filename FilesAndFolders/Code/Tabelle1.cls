VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)

    
If Target.name = "FolderList" Then
    Call mFolderList.FolderList
    
ElseIf Target.name = "Pictures To HTML" Then
    Call PicturesToHTML

ElseIf Target.name = "FilesInFolder" Then
    Call FilesInFolder

End If


'Call ListFiles
'ThisWorkbook.Close savechanges:=False
End Sub


