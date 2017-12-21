' File: PackTheme.vbs
Option Explicit

Dim Shell, Fso, ConfigFile, File
Dim ConfigFileName, ThemeFolderName, OutputThemeName

Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

ConfigFileName = "Packing.ddf"
ThemeFolderName = "Packing"
OutputThemeName = "Packing.deskthemepack"

' Parse command line arguments
If WScript.Arguments.Count >= 1 Then
  ThemeFolderName = WScript.Arguments(0)
End If
If WScript.Arguments.Count >= 2 Then
  OutputThemeName = WScript.Arguments(1)
End If

' TODO: Validate arguments

' 1. Create the config file for makecab.exe
Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 2, True) 'For Writing
ConfigFile.WriteLine ".OPTION EXPLICIT"
ConfigFile.WriteLine ".Set CabinetNameTemplate=.\" & OutputThemeName
ConfigFile.WriteLine ".Set DiskDirectory1=."
ConfigFile.WriteLine ".Set CompressionType=MSZIP"
ConfigFile.WriteLine ".Set Cabinet=on"
ConfigFile.WriteLine ".Set Compress=on"
ConfigFile.WriteLine ".Set CabinetFileCountThreshold=0"
ConfigFile.WriteLine ".Set FolderFileCountThreshold=0"
ConfigFile.WriteLine ".Set FolderSizeThreshold=0"
ConfigFile.WriteLine ".Set MaxCabinetSize=0"
ConfigFile.WriteLine ".Set MaxDiskFileCount=0"
ConfigFile.WriteLine ".Set MaxDiskSize=0"
WriteDirTree ConfigFile, ThemeFolderName, ""
ConfigFile.Close

' 2. Pack the stuff
Shell.Run "makecab.exe /F """ & ConfigFileName & """",, True

' 3. Cleanup
Fso.DeleteFile "setup.inf"
Fso.DeleteFile "setup.rpt"
Fso.DeleteFile ConfigFileName

'###################
'#  End of Script  #
'###################

' Library Functions

Sub WriteDirTree(ByRef OutFile, ByVal DirPath, ByVal RootPath)
  Dim Dir, File, FilePath, Folder
  Set Dir = Fso.GetFolder(DirPath)
  
  If Len(RootPath) = 0 Then
    RootPath = Dir.Path
  End If
  
  For Each File In Dir.Files
    FilePath = Dir.Path & "\" & File.Name
    OutFile.WriteLine """" & FilePath & """ """ & Mid(FilePath, 2+Len(RootPath)) & """"
  Next
  For Each Folder In Dir.SubFolders
    WriteDirTree OutFile, Fso.GetFolder(Folder.Path), RootPath
  Next
End Sub

Sub WriteLog(ByVal LogMessage)
  WScript.Echo LogMessage
End Sub