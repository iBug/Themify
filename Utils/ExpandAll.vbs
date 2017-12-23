Option Explicit

Dim Shell, Fso

Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim WD, File
Set WD = Fso.GetFolder(".")

If Not Fso.FolderExists("Out") Then Fso.CreateFolder "Out"

For Each File In WD.Files
  Dim Extension, Folder : Extension = Fso.GetExtensionName(File.Path)
  If Extension = "themepack" _
    Or Extension = "deskthemepack" Then
    Folder = Fso.GetBaseName(File.Path)
    Fso.CreateFolder Folder
    Shell.Run "expand """ & File.Path & """ -F:* """ & Folder & """",, True
    Fso.MoveFolder Folder & "\DesktopBackground", "Out\" & Folder
    Fso.DeleteFolder Folder
  End If
Next
