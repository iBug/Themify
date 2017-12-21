' File: Main.vbs
Option Explicit

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim ConfigFileName, DefaultConfig, ConfigLine, Config
Dim ConfigFile
Dim C_Source, C_Name, C_Output

ConfigFileName = "config.ini"
DefaultConfig = "; Creation Time: " & Now() & vbCrlf & vbCrLf & _
  "; Path containing images to be made into the theme pack" & vbCrLf & _
  "SourcePath=C:\WINDOWS\Web\Wallpaper\Theme1" & vbCrLf & vbCrLf & _
  "; Name of the theme" & vbCrLf & _
  "ThemeName=My Theme" & vbCrLf & vbCrLf & _
  "; The generated file (extension will be automatically appended)" & vbCrLf & _
  "OutputFile=Output" & vbCrLf & vbCrLf

' Generate the default config
Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 2, True)
ConfigFile.WriteLine DefaultConfig
ConfigFile.Close
Set ConfigFile = Nothing

' Prompt the user to change the config
Shell.Run ConfigFileName,, True

' Read and apply the config
Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 1, False)
Do Until ConfigFile.AtEndOfStream
  Config = Split(ConfigFile.ReadLine(), "=")
  If UBound(Config) = 1 Then
    Select Case Trim(Config(0))
      Case "SourcePath"
        C_Source = Trim(Config(1))
      Case "ThemeName"
        C_Name = Trim(Config(1))
      Case "OutputFile"
        C_Output = Trim(Config(1)) & ".deskthemepack"
      Case Else
        MsgBox "Unknown option """ & Config(0) & """", vbExclamation, "Theme Creator"
    End Select
  End If
Loop
ConfigFile.Close
Set ConfigFile = Nothing
Fso.DeleteFile ConfigFileName

' Run the procedure
Dim Command
Command = "Procedure.vbs """ & C_Source & """ """ & C_Name & """ """ & C_Output & """"
'If MsgBox(Command, vbYesNo, "Theme Creator") = vbYes Then
  Shell.Run Command,, False
'End If
