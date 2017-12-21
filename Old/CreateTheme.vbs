' File: CreateTheme.vbs
Option Explicit

Dim Shell, Fso
Set Shell = CreateObject("WScript.Shell")
Set Fso = CreateObject("Scripting.FileSystemObject")

Dim ThemeFileName, ConfigFileName, ConfigData
Dim ConfigFile
Dim C_Time, C_DisplayName, C_GUID, C_DefaultBackground, C_AutoColor, C_Interval, C_Shuffle

C_Time = Now()
C_DisplayName = "Generated"
C_GUID = CreateGUID()
C_DefaultBackground = "default.jpg"
C_AutoColor = 1
C_Interval = 600000
C_Shuffle = 1
ThemeFileName = C_DisplayName & ".theme"
ConfigFileName = ""

If WScript.Arguments.Count >= 1 Then
  ThemeFileName = WScript.Arguments(0)
End If
If WScript.Arguments.Count >= 2 Then
  ConfigFileName = WScript.Arguments(1)
End If

' Look for config file and read it if exist
If Len(ConfigFileName) > 0 Then
  Set ConfigFile = Fso.OpenTextFile(ConfigFileName, 1, False)
  ConfigData = ConfigFile.ReadAll()
  ConfigFile.Close
  Set ConfigFile = Nothing
End If
' Dangerously apply
ExecuteGlobal ConfigData

WriteTheme ThemeFileName

Sub WriteTheme(ByVal ThemeFileName)
  Dim ThemeFile
  Set ThemeFile = Fso.OpenTextFile(ThemeFileName, 2, True)
  WriteHeader ThemeFile
  WriteThemeMain ThemeFile
  WriteThemeDesktop ThemeFile
  WriteVisualStyles ThemeFile
  WriteSlideshow ThemeFile
  ThemeFile.Close
End Sub

Sub WriteHeader(ByRef TextFile)
  TextFile.WriteLine "; Generated by iBug/Themify"
  TextFile.WriteLine "; Generated: " & C_Time
  TextFile.WriteLine ""
  TextFile.WriteLine "[MasterThemeSelector]"
  TextFile.WriteLine "MTSM=DABJDKT"
  TextFile.WriteLine ""
End Sub

Sub WriteThemeMain(ByRef TextFile)
  TextFile.WriteLine "[Theme]"
  TextFile.WriteLine "DisplayName=" & C_DisplayName
  TextFile.WriteLine "ThemeId={" & C_GUID & "}"
  TextFile.WriteLine ""
End Sub

Sub WriteThemeDesktop(ByRef TextFile)
  TextFile.WriteLine "[Control Panel\Desktop]"
  TextFile.WriteLine "Wallpaper=DesktopBackground\" & C_DefaultBackground
  TextFile.WriteLine "TileWallpaper=0"
  TextFile.WriteLine "WallpaperStyle=10" ' Resize and crop
  TextFile.WriteLine ""
End Sub

Sub WriteVisualStyles(ByRef TextFile)
  TextFile.WriteLine "[VisualStyles]"
  TextFile.WriteLine "Path=%SystemRoot%\Resources\Themes\Aero\Aero.msstyles"
  TextFile.WriteLine "ColorStyle=NormalColor"
  TextFile.WriteLine "Size=NormalSize"
  TextFile.WriteLine "ColorizationColor=0XC40078D7"
  TextFile.WriteLine "AutoColorization=" & C_AutoColor
  TextFile.WriteLine "VisualStyleVersion=10"
  TextFile.WriteLine "Transparency=1"
  TextFile.WriteLine ""
End Sub

Sub WriteSlideshow(ByRef TextFile)
  TextFile.WriteLine "[Slideshow]"
  TextFile.WriteLine "Interval=" & C_Interval
  TextFile.WriteLine "Shuffle=" & C_Shuffle
  TextFile.WriteLine "ImagesRootPath=DesktopBackground"
  TextFile.WriteLine ""
End Sub

'###################
'#  End of Script  #
'###################

' Library Functions

Function CreateGUID
  CreateGUID = Mid(CreateObject("Scriptlet.TypeLib").Guid, 2, 36)
End Function

Sub WriteLog(ByVal LogMessage)
  WScript.Echo LogMessage
End Sub