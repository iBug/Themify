' File: Procedure.vbs
Option Explicit

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim ConfigFileName
Dim ConfigFile, Dir, File
Dim C_SourceDir, C_Name, C_Cover, C_OutputFile, C_DeleteWD
Dim S_WD, S_ThemeConfigFileName

C_SourceDir = "Src_"
C_Name = "Test"
C_Cover = "default.jpg"
C_OutputFile = "Generating.deskthemepack"
C_DeleteWD = True
S_WD = Mid(C_OutputFile, 1, Len(C_OutputFile)-1) & "_"
S_ThemeConfigFileName = "config.vbs"

' TODO: Read configuration file
If WScript.Arguments.Count < 2 Then
  WScript.Echo "Usage: " & vbCrLf & Fso.GetFileName(WScript.ScriptFullName) & " <Source> <Name> [OutputFile]"
  WScript.Quit 1
End If

C_SourceDir = WScript.Arguments(0)
C_Name = WScript.Arguments(1)
If WScript.Arguments.Count >= 3 Then
  C_OutputFile = WScript.Arguments(2)
End If

If Fso.FolderExists(S_WD) Then
  Fso.DeleteFolder S_WD
End If
Fso.CreateFolder S_WD

' 1. Collect all images from the given directory
Shell.Run "Collect.vbs """ & C_SourceDir & """ """ & S_WD & "\DesktopBackground""",, True

' 2. Create the theme file
Set ConfigFile = Fso.OpenTextFile(S_ThemeConfigFileName, 2, True)
ConfigFile.WriteLine "C_DefaultBackground = """ & C_Cover & """"
ConfigFile.WriteLine "C_DisplayName = """ & C_Name & """"
ConfigFile.Close
Set ConfigFile = Nothing
Shell.Run "CreateTheme.vbs " & S_WD & "\Generating.theme " & S_ThemeConfigFileName,, True
Fso.DeleteFile S_ThemeConfigFileName

' Rename a file to "default.jpg"
Set Dir = Fso.GetFolder(S_WD & "\DesktopBackground")
If Not Fso.FileExists(Fso.GetParentFolderName(Dir.Path) & "\" & C_Cover) Then
  For Each File In Dir.Files
    If Fso.GetExtensionName(File.Name) = Fso.GetExtensionName(C_Cover) Then
      Fso.MoveFile File.Path, Fso.GetParentFolderName(File.Path) & "\" & C_Cover
      Set File = Nothing
      Exit For
    End If
  Next
End If
Set Dir = Nothing

' 3. Pack them up
Shell.Run "PackTheme.vbs " & S_WD & " " & C_OutputFile,, True

' Cleanup
If C_DeleteWD Then
  Fso.DeleteFolder S_WD
End If

'###################
'#  End of Script  #
'###################

' Library Functions

Sub WriteLog(ByVal LogMessage)
  WScript.Echo LogMessage
End Sub
