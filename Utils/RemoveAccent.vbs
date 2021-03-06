Set Fso = CreateObject("Scripting.FileSystemObject")

RemoveFolderRecursive "."
MsgBox "Complete"

Sub RemoveFolderRecursive(ByVal FolderName)
  For Each F in Fso.GetFolder(FolderName).Files
    G = RemoveAccent(F.Path)
    If G <> F.Path Then F.Name = Fso.GetFileName(G)
  Next
  For Each F in Fso.GetFolder(FolderName).SubFolders
    RemoveFolderRecursive F.Path
  Next
End Sub

Function RemoveAccent(ByVal Text)
    Dim i, s1, s2
    s1 = "ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜàáâãäåçèéêëìíîïòóôõöùúûüøćčřśšșžß"
    s2 = "AAAAAACEEEEIIIIOOOOOUUUUaaaaaaceeeeiiiiooooouuuuoccrssszs"
    If Len(Text) <> 0 Then
        For i = 1 To Len(s1)
            Text = Replace(Text, Mid(s1,i,1), Mid(s2,i,1))
        Next
    End If
    RemoveAccent = Text
End Function