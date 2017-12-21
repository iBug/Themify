' File: Collect.vbs
Option Explicit

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim ValidExtensions
Dim CollectDir, OutputDir, Dir

CollectDir = "Collect"
OutputDir = "DesktopBackgound"

' Parse command line arguments
If WScript.Arguments.Count >= 1 Then
  CollectDir = WScript.Arguments(0)
End If
If WScript.Arguments.Count >= 2 Then
  OutputDir = WScript.Arguments(1)
End If

' TODO: Validate arguments

If Not Fso.FolderExists(OutputDir) Then
  Fso.CreateFolder OutputDir
End If

Set Dir = Fso.GetFolder(CollectDir)
CollectDir = Dir.Path
Set Dir = Fso.GetFolder(OutputDir)
OutputDir = Dir.Path
Set Dir = Nothing

ValidExtensions = Array("jpg", "png", "bmp")
CopyFileRecursive CollectDir

Sub CopyFileRecursive(ByVal DirName)
  Dim Dir, File, Folder
  Set Dir = Fso.GetFolder(DirName)
  
  For Each File In Dir.Files
    Dim Extension, FileExtension, Valid
    FileExtension = Fso.GetExtensionName(File.Path)
    Valid = False
    
    For Each Extension In ValidExtensions
      If FileExtension = Extension Then
        Valid = True
        Exit For
      End If
    Next
    
    If Valid Then
      Fso.CopyFile File.Path, OutputDir & "\" & File.Name, True
    End If
  Next
  For Each Folder In Dir.SubFolders
    CopyFileRecursive Folder.Path
  Next
End Sub

'###################
'#  End of Script  #
'###################

' Library Functions

Sub WriteLog(ByVal LogMessage)
  WScript.Echo LogMessage
End Sub