' iBug Themify VBScript
Option Explicit

Const ProgramName = "iBug Themify"
Const Version = "1.0"
Const VersionNumber = 4
Dim Title : Title = ProgramName & " v" & Version & "." & VersionNumber

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim C_Time, C_Source, C_Name, C_Output
Dim CT_DisplayName, CT_GUID, CT_DefaultBackground, CT_AutoColor, CT_Interval, CT_Shuffle
Dim I_ConfigFileName, I_ThemeFileName, I_DirectiveFileName, I_ValidExtensions

C_Time = Now()
I_ConfigFileName = Fso.GetSpecialFolder(2) & "\themify.ini"
I_ThemeFileName = Fso.GetSpecialFolder(2) & "\Themify.theme"
I_DirectiveFileName = Fso.GetSpecialFolder(2) & "\themify.ddf"
I_ValidExtensions = Array("bmp", "jpg", "jpeg", "png")

GetConfig I_ConfigFileName
GenerateThemeConfig
CreateTheme I_ThemeFileName
WriteDirectiveFile I_DirectiveFileName
CreatePackage I_DirectiveFileName
Fso.DeleteFile I_DirectiveFileName
Fso.DeleteFile I_ThemeFileName

' *********************
' Config File Functions
' *********************

Sub CreateDefaultConfig(ByVal ConfigFileName)
  Dim DefaultConfig, ConfigFile
  DefaultConfig = "; " & ProgramName & " running at: " & C_Time & vbNewline & vbNewline & _
    "; Path containing images to be made into the theme pack" & vbNewline & _
    "SourcePath=C:\WINDOWS\Web\Wallpaper\Theme1" & vbNewline & vbNewline & _
    "; Name of the theme" & vbNewline & _
    "ThemeName=My Theme" & vbNewline & vbNewline & _
    "; Path to the generated theme pack (extension will be automatically appended)" & vbNewline & _
    "OutputFile=Output" & vbNewline & vbNewline
  Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 2, True)
  ConfigFile.WriteLine DefaultConfig
  ConfigFile.Close
End Sub

Sub ReadConfig(ByVal ConfigFileName)
  Dim ConfigFile, Config
  Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 1, False)
  Do Until ConfigFile.AtEndOfStream
    Config = Split(ConfigFile.ReadLine(), "=", 2)
    If UBound(Config) <> 1 Then
      ' Do nothing
    ElseIf Mid(LTrim(Config(0)), 1, 1) <> ";" Then
      Select Case LCase(Trim(Config(0)))
        Case "sourcepath"
          C_Source = Trim(Config(1))
        Case "themename"
          C_Name = Trim(Config(1))
        Case "outputfile"
          C_Output = Trim(Config(1)) & ".deskthemepack"
        Case Else
          MsgBox "Unknown option """ & Config(0) & """", vbExclamation, Title
          ' And... Do nothing
      End Select
    End If
  Loop
  ConfigFile.Close
  
  If Fso.GetExtensionName(C_Output) <> "themepack" _
    And Fso.GetExtensionName(C_Output) <> "deskthemepack" Then
    C_Output = C_Output & ".deskthemepack"
  End If
End Sub

Sub GetConfig(ByVal ConfigFileName)
  ' Notify user
  MsgBox "Please set the parameters in the configuration file.", 0, Title
  CreateDefaultConfig ConfigFileName
  
  ' Prompt the user to change the config
  Shell.Run ConfigFileName,, True

  ' Read and apply the config
  ReadConfig ConfigFileName
  Fso.DeleteFile ConfigFileName

  ' Validate config
  If Not Fso.FolderExists(C_Source) Then
    MsgBox "The source folder """ & C_Source & """ does not exist!", 16, Title
    WScript.Quit 1
  End If
End Sub

' *****************
' Library Functions
' *****************

Function CreateGUID
  CreateGUID = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Sub WriteLog(ByVal LogMessage)
  WScript.Echo LogMessage
End Sub

' ********************
' Theme File Functions
' ********************

Sub GenerateThemeConfig
  CT_DisplayName = C_Name
  CT_GUID = CreateGUID()
  CT_DefaultBackground = "1.jpg"
  CT_AutoColor = 1
  CT_Interval = 600000
  CT_Shuffle = 1
End Sub

Sub CreateTheme(ByVal FileName)
  Dim ThemeFile
  Set ThemeFile = Fso.OpenTextFile(FileName, 2, True)
  WriteHeader ThemeFile
  WriteThemeMain ThemeFile
  WriteIcons ThemeFile
  WriteCursors ThemeFile
  WriteThemeDesktop ThemeFile
  WriteVisualStyles ThemeFile
  WriteSounds ThemeFile
  WriteExtras ThemeFile
  WriteSlideshow ThemeFile
  ThemeFile.Close
End Sub

Sub WriteHeader(ByRef TextFile)
  TextFile.WriteLine "; Generated by " & ProgramName & " (v" & Version & ")"
  TextFile.WriteLine "; Created at: " & C_Time
  TextFile.WriteLine ""
  TextFile.WriteLine "[MasterThemeSelector]"
  'TextFile.WriteLine "MTSM=RJSPBS"
  TextFile.WriteLine "MTSM=DABJDKT"
  TextFile.WriteLine ""
End Sub

Sub WriteThemeMain(ByRef TextFile)
  TextFile.WriteLine "[Theme]"
  TextFile.WriteLine "DisplayName=" & CT_DisplayName
  TextFile.WriteLine "ThemeId={" & CT_GUID & "}"
  TextFile.WriteLine ""
End Sub

Sub WriteThemeDesktop(ByRef TextFile)
  TextFile.WriteLine "[Control Panel\Desktop]"
  TextFile.WriteLine "Wallpaper=DesktopBackground\" & CT_DefaultBackground
  TextFile.WriteLine "Pattern="
  TextFile.WriteLine "MultimonBackgrounds=0"
  TextFile.WriteLine "PicturePosition=4" ' Resize and crop
  TextFile.WriteLine ""
End Sub

Sub WriteVisualStyles(ByRef TextFile)
  TextFile.WriteLine "[VisualStyles]"
  TextFile.WriteLine "Path=%SystemRoot%\Resources\Themes\Aero\Aero.msstyles"
  TextFile.WriteLine "ColorStyle=NormalColor"
  TextFile.WriteLine "Size=NormalSize"
  TextFile.WriteLine "AutoColorization=" & CT_AutoColor
  TextFile.WriteLine "ColorizationColor=0XC40078D7"
  TextFile.WriteLine "VisualStyleVersion=10"
  TextFile.WriteLine "Transparency=1"
  TextFile.WriteLine ""
End Sub

Sub WriteCursors(ByRef TextFile)
  TextFile.WriteLine "[Control Panel\Cursors]"
  TextFile.WriteLine "AppStarting=%SystemRoot%\cursors\aero_working.ani"
  TextFile.WriteLine "Arrow=%SystemRoot%\cursors\aero_arrow.cur"
  TextFile.WriteLine "Hand=%SystemRoot%\cursors\aero_link.cur"
  TextFile.WriteLine "Help=%SystemRoot%\cursors\aero_helpsel.cur"
  TextFile.WriteLine "No=%SystemRoot%\cursors\aero_unavail.cur"
  TextFile.WriteLine "NWPen=%SystemRoot%\cursors\aero_pen.cur"
  TextFile.WriteLine "SizeAll=%SystemRoot%\cursors\aero_move.cur"
  TextFile.WriteLine "SizeNESW=%SystemRoot%\cursors\aero_nesw.cur"
  TextFile.WriteLine "SizeNS=%SystemRoot%\cursors\aero_ns.cur"
  TextFile.WriteLine "SizeNWSE=%SystemRoot%\cursors\aero_nwse.cur"
  TextFile.WriteLine "SizeWE=%SystemRoot%\cursors\aero_ew.cur"
  TextFile.WriteLine "UpArrow=%SystemRoot%\cursors\aero_up.cur"
  TextFile.WriteLine "Wait=%SystemRoot%\cursors\aero_busy.ani"
  TextFile.WriteLine "DefaultValue=Windows Default"
  TextFile.WriteLine ""
End Sub

Sub WriteIcons(ByRef TextFile)
  TextFile.WriteLine "[CLSID\{20D04FE0-3AEA-1069-A2D8-08002B30309D}\DefaultIcon]"
  TextFile.WriteLine "DefaultValue=%SystemRoot%\System32\imageres.dll,-109"
  TextFile.WriteLine ""
  TextFile.WriteLine "[CLSID\{59031A47-3F72-44A7-89C5-5595FE6B30EE}\DefaultIcon]"
  TextFile.WriteLine "DefaultValue=%SystemRoot%\System32\imageres.dll,-123"
  TextFile.WriteLine ""
  TextFile.WriteLine "[CLSID\{F02C1A0D-BE21-4350-88B0-7367FC96EF3C}\DefaultIcon]"
  TextFile.WriteLine "DefaultValue=%SystemRoot%\System32\imageres.dll,-25"
  TextFile.WriteLine ""
  TextFile.WriteLine "[CLSID\{645FF040-5081-101B-9F08-00AA002F954E}\DefaultIcon]"
  TextFile.WriteLine "Full=%SystemRoot%\System32\imageres.dll,-54"
  TextFile.WriteLine "Empty=%SystemRoot%\System32\imageres.dll,-55"
  TextFile.WriteLine ""
End Sub

Sub WriteSounds(ByRef TextFile)
  TextFile.WriteLine "[Sounds]"
  TextFile.WriteLine "SchemeName=Windows Default"
  TextFile.WriteLine ""
End Sub

Sub WriteSlideshow(ByRef TextFile)
  TextFile.WriteLine "[Slideshow]"
  TextFile.WriteLine "Interval=" & CT_Interval
  TextFile.WriteLine "Shuffle=" & CT_Shuffle
  TextFile.WriteLine "ImagesRootPath=DesktopBackground"
  TextFile.WriteLine ""
End Sub

Sub WriteExtras(ByRef TextFile)
  TextFile.WriteLine "[Control Panel\Cursors.A]"
  TextFile.WriteLine "[Control Panel\Cursors.W]"
  TextFile.WriteLine "DefaultValue=Windows Defailt"
  TextFile.WriteLine ""
  TextFile.WriteLine "[Theme.A]"
  TextFile.WriteLine "[Theme.W]"
  TextFile.WriteLine "DisplayName=" & CT_DisplayName
  TextFile.WriteLine ""
  TextFile.WriteLine "[Control Panel\Desktop.A]"
  TextFile.WriteLine "[Control Panel\Desktop.W]"
  TextFile.WriteLine "Wallpaper=DesktopBackground\" & CT_DefaultBackground
  TextFile.WriteLine ""
End Sub

' *******************
' Packaging Functions
' *******************

Sub CreatePackage(ByVal DirectiveFileName)
  Shell.Run "makecab.exe /F """ & DirectiveFileName & """",, True
  Fso.DeleteFile "setup.inf"
  Fso.DeleteFile "setup.rpt" ' These two are generated by makecab.exe
End Sub

Sub WriteDirectiveFile(ByVal FileName)
  Dim OutputFileName, OutputFolderName
  Dim DirectiveFile
  OutputFolderName = Fso.GetParentFolderName(C_Output)
  If Len(OutputFolderName) = 0 Then
    OutputFolderName = "."
  End If
  OutputFileName = Fso.GetFileName(C_Output)
  Set DirectiveFile = Fso.OpenTextFile(FileName, 2, True)
  DirectiveFile.WriteLine ".OPTION EXPLICIT"
  DirectiveFile.WriteLine ".Set CabinetNameTemplate=" & OutputFileName
  DirectiveFile.WriteLine ".Set DiskDirectory1=" & OutputFolderName
  DirectiveFile.WriteLine ".Set CompressionType=MSZIP"
  DirectiveFile.WriteLine ".Set Cabinet=on"
  DirectiveFile.WriteLine ".Set Compress=on"
  DirectiveFile.WriteLine ".Set CabinetFileCountThreshold=0"
  DirectiveFile.WriteLine ".Set FolderFileCountThreshold=0"
  DirectiveFile.WriteLine ".Set FolderSizeThreshold=0"
  DirectiveFile.WriteLine ".Set MaxCabinetSize=0"
  DirectiveFile.WriteLine ".Set MaxDiskFileCount=0"
  DirectiveFile.WriteLine ".Set MaxDiskSize=0"
  DirectiveFile.WriteLine """" & I_ThemeFileName & """ """ & Fso.GetFileName(I_ThemeFileName) & """"
  WriteDirTree DirectiveFile, C_Source, 0
  DirectiveFile.Close
End Sub

Sub WriteDirTree(ByRef OutFile, ByVal DirPath, ByRef Count)
  Dim Dir, Item
  Set Dir = Fso.GetFolder(DirPath)
  For Each Item In Dir.Files
    Dim Extension, FileExtension, Valid
    FileExtension = Fso.GetExtensionName(Item.Path)
    Valid = False
    
    For Each Extension In I_ValidExtensions
      If LCase(FileExtension) = Extension Then
        Valid = True
        Exit For
      End If
    Next
    
    If Valid Then
      Count = Count + 1
      OutFile.WriteLine """" & Item.Path & """ ""DesktopBackground\" & Count & "." & LCase(Fso.GetExtensionName(Item.Path)) &  """"
      If Count = 1 Then
        CT_DefaultBackground = "1." & LCase(Fso.GetExtensionName(Item.Path))
      End If
    End If
  Next
  For Each Item In Dir.SubFolders
    WriteDirTree OutFile, Fso.GetFolder(Item.Path), Count
  Next
End Sub
