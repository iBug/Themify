' iBug Themify VBScript
Option Explicit

Const ProgramName = "iBug Themify"
Const Version = "1.1"
Const VersionNumber = 8
Dim Title : Title = ProgramName & " v" & Version & "." & VersionNumber

Dim Shell, Fso, LogFile
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")
'Set LogFile = Fso.OpenTextFile("Themify.log", 2, True)

Dim C_Time, C_Source, C_Name, C_Output, C_ConfigName
Dim CT_DisplayName, CT_GUID, CT_DefaultBackground, CT_AutoColor, CT_Color, CT_Interval, CT_Shuffle
Dim I_ConfigFileName, I_ThemeFileName, I_DirectiveFileName, I_DirectiveFileNames, I_ValidExtensions

C_Time = Now()
I_ConfigFileName = Fso.GetSpecialFolder(2) & "\themify.ini"
I_ThemeFileName = Fso.GetSpecialFolder(2) & "\Themify.theme"
I_DirectiveFileName = Fso.GetSpecialFolder(2) & "\themify.ddf"
I_DirectiveFileNames = Array(Fso.GetSpecialFolder(2) & "\themify*.ddf")
I_ValidExtensions = Array("bmp", "jpg", "jpeg", "png")

' Parse command line arguments
If WScript.Arguments.Count > 0 Then
  C_ConfigName = WScript.Arguments(0)
  If Not Fso.FileExists(C_ConfigName) Then
    MsgBox "The config file """ & C_ConfigName & """ does not exist!", 16, Title
    WScript.Quit 1
  End If
  I_ConfigFileName = C_ConfigName
  ' Change working directory to the same as the config file
  Shell.CurrentDirectory = Fso.GetParentFolderName(Fso.GetFile(C_ConfigName).Path)
  GetExistingConfig I_ConfigFileName, False
  MainAutomatic
Else
  GetConfig I_ConfigFileName
  MainAutomatic
End If

' ******************************

Sub MainAutomatic
  GenerateThemeConfig
  WriteDirectiveFile I_DirectiveFileName, I_DirectiveFileNames
  CreateTheme I_ThemeFileName
  CreatePackage I_DirectiveFileName, I_DirectiveFileNames
  'Fso.DeleteFile I_DirectiveFileName
  'Fso.DeleteFile I_ThemeFileName
End Sub

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
    "OutputFile=$default" & vbNewline & vbNewline & _
    "; Interval of slideshow (in milliseconds), default 10 minutes" & vbNewline & _
    "SlideshowInterval=600000" & vbNewline & vbNewline
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
          C_Output = Trim(Config(1))
        Case "slideshowinterval"
          CT_Interval = Trim(Config(1))
        Case Else
          MsgBox "Unknown option """ & Config(0) & """", vbExclamation, Title
          ' And... Do nothing
      End Select
    End If
  Loop
  ConfigFile.Close
End Sub

Sub GetExistingConfig(ByVal ConfigFileName, ByVal DeleteAfter)
  ' Read and apply the config
  ReadConfig ConfigFileName
  If DeleteAfter Then
    Fso.DeleteFile ConfigFileName
  End If
  
  ' Validate config
  If Not ValidateConfig Then
    WScript.Quit 1
  End If
End Sub

Function ValidateConfig
  ValidateConfig = True
  If Not Fso.FolderExists(C_Source) Then
    MsgBox "The source folder """ & C_Source & """ does not exist!", 16, Title
    ValidateConfig = False
    Exit Function
  End If
  If Not IsNumeric(CT_Interval) Then CT_Interval = 600000
  If LCase(C_Output) = "$default" Then C_Output = C_Name
  If Fso.GetExtensionName(C_Output) <> "themepack" _
    And Fso.GetExtensionName(C_Output) <> "deskthemepack" Then
    C_Output = C_Output & ".deskthemepack"
  End If
End Function

Sub GetConfig(ByVal ConfigFileName)
  ' Notify user
  MsgBox "Please set the parameters in the configuration file.", 0, Title
  CreateDefaultConfig ConfigFileName
  
  ' Prompt the user to change the config
  Shell.Run ConfigFileName,, True
  
  GetExistingConfig ConfigFileName, True
End Sub

' *****************
' Library Functions
' *****************

Function CreateGUID
  CreateGUID = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Function GetWorkingDirectory
  GetWorkingDirectory = Fso.GetFolder(".").Path
End Function

Sub WriteLog(ByVal LogMessage)
  'LogFile.WriteLine LogMessage
End Sub

' ********************
' Theme File Functions
' ********************

Sub GenerateThemeConfig
  CT_DisplayName = C_Name
  CT_GUID = CreateGUID()
  CT_DefaultBackground = "1.jpg"
  CT_AutoColor = 1
  CT_Color = "0078D7"
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
  TextFile.WriteLine "ColorizationColor=0XC4" & CT_Color
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
  TextFile.WriteLine "DefaultValue=Windows Default"
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

Sub CreatePackage(ByVal DirectiveFileName, ByVal DirectiveFileNames)
  ' Construct commmand line from functions
  Dim Command, i
  Command = "makecab.exe /F """ & DirectiveFileName & """"
  For i = 1 To UBound(DirectiveFileNames)
    Command = Command & " /F """ & DirectiveFileNames(i) & """"
  Next
  Shell.Run Command,, True
  Fso.DeleteFile "setup.inf"
  Fso.DeleteFile "setup.rpt" ' These two are generated by makecab.exe
End Sub

Sub WriteDirectiveFile(ByVal FileName, ByRef FileNames)
  Dim OutputFileName, OutputFolderName
  Dim DirectiveFile
  OutputFolderName = Fso.GetParentFolderName(C_Output)
  If Len(OutputFolderName) = 0 Then
    OutputFolderName = GetWorkingDirectory()
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
  DirectiveFile.WriteLine ".Set MaxCabinetSize=2147483647"
  DirectiveFile.WriteLine ".Set MaxDiskFileCount=0"
  DirectiveFile.WriteLine ".Set MaxDiskSize=2147483647"
  DirectiveFile.WriteLine """" & I_ThemeFileName & """ """ & Fso.GetFileName(I_ThemeFileName) & """"
  DirectiveFile.Close
  WriteDirTree C_Source, 0, 0, FileNames, Nothing, NewExtendableWriterControl()
End Sub

Function NewExtendableWriterControl()
  ' Array(CurrentNumber, CurrentCount)
  NewExtendableWriterControl = Array(0, -1)
End Function

Sub ExtendableWriter(ByVal Text, ByRef FileNames, ByRef Handle, ByRef Control)
  If Control(1) < 0 Or Control(1) >= 1000 Then
    Dim NewName
    If Not Handle Is Nothing Then Handle.Close
    Control(1) = 0
    Control(0) = Control(0) + 1
    NewName = Replace(FileNames(0), "*", Control(0))
    ReDim Preserve FileNames(UBound(FileNames) + 1)
    FileNames(UBound(FileNames)) = NewName
    Set Handle = Fso.OpenTextFile(NewName, 2, True)
  End If
  Handle.WriteLine Text
  Control(1) = Control(1) + 1
End Sub

Sub WriteDirTree(ByVal DirPath, ByRef Count, ByRef SizeTotal, ByRef FileNames, ByRef Handle, ByRef Control)
  Dim Dir, Item
  Set Dir = Fso.GetFolder(DirPath)
  For Each Item In Dir.Files
    Dim Extension, FileExtension, Valid
    FileExtension = Fso.GetExtensionName(Item.Path)
    Valid = False
    
    If SizeTotal + Item.Size > 2000000000 Then ' CAB can't exceed 2GB
      Valid = False
    Else
      For Each Extension In I_ValidExtensions
        If LCase(FileExtension) = Extension Then
          Valid = True
          Exit For
        End If
      Next
    End If
    
    If Valid Then
      Count = Count + 1
      SizeTotal = SizeTotal + Item.Size
      ExtendableWriter """" & Item.Path & """ ""DesktopBackground\" & Count & "." & LCase(Fso.GetExtensionName(Item.Path)) & """", FileNames, Handle, Control
      If Count = 1 Then
        CT_DefaultBackground = "1." & LCase(Fso.GetExtensionName(Item.Path))
      End If
    End If
  Next
  For Each Item In Dir.SubFolders
    WriteDirTree Fso.GetFolder(Item.Path), Count, SizeTotal, FileNames, Handle, Control
  Next
End Sub
