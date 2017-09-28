' =============================================================================
' MSI Installation Script
' -----------------------
' Author:   Enterprise Application Packaging
' Version:  2.8.0000
'
' Custom Return Codes
' -----------------------
' -101   = Invalid switch
' -2   = Unable to create directory
' 1612 = Unable to find file (e.g. MSI, MST and MSP)
' =============================================================================

'Set up to NOT halt on error
Option Explicit
On Error Resume Next

' -----------------------------------------------------------------------------
' Variable Declaration
' -----------------------------------------------------------------------------

Dim MSIVer, AppCode, DistCode, ProductCode, OldProductCode, Platform
Dim Log, LogPath
Dim WSHShell, oFSO, oExec
Dim Profile, COEPlatform, CodeSrv, CodeDrv, AppChk, CDBuild, AssetTag
Dim MSIError, MSILog, OldMSILog, MSIPath, MSTPath, MsiexecCmd, MSPPath
Dim Switch, Mode, oEnv, x64Key, OSType
Dim CodePath, InstallPath
Dim InstallType, ProductName, RelNum
Dim PREVERSIONFOUND, LogMode, Switch2, UninstLog
Dim OS
'Added in Win7
Dim SystemRoot, UserName

' -----------------------------------------------------------------------------
' Custom Variable Declaration
' -----------------------------------------------------------------------------

' Dim your custom global variables here
Const ForReading = 1, ForWriting = 2, ForAppending = 8
Dim InstOption, RemoveDependencies

' -----------------------------------------------------------------------------
' Package Set Variables
' -----------------------------------------------------------------------------

'> MSIVer should be updated with MSI ProductVersion (e.g. Major.Minor.Build)
MSIVer 		= "N.NN.NNNN"
'> Set these below to the related appcode, dist code, ProductName & RelNum
AppCode 	= "XX"
DistCode 	= "XXXX"
ProductName = "IIS 8 Silent Install for Windows 2012 Server"
RelNum		= "SFR_901" 'RelNum value can only be SFR_XXX, EA_XXX, MSI_ONLY (ie; ACT) or USWM_XXX or DSU-WIN7

PREVERSIONFOUND = False

'> The application's GUID should be updated with each major MSI change to this app
ProductCode = "{NNNNNNNN-NNNN-NNNN-NNNN-NNNNNNNNNNNN}"
Platform = "Server"


' -----------------------------------------------------------------------------
' Declare Objects
' -----------------------------------------------------------------------------

'Create shell reference for some handy install, reg query, etc...usage
Set WSHShell = CreateObject("WScript.Shell")

'Create environment object to retrieve system environment variables
Set oEnv = WSHShell.Environment("SYSTEM")

'Create file system object
Set oFSO = CreateObject("Scripting.FileSystemObject")

SystemRoot =  WSHShell.ExpandEnvironmentStrings("%SystemRoot%")
UserName = WSHShell.ExpandEnvironmentStrings("%USERNAME%")
RemoveDependencies = False
' -----------------------------------------------------------------------------
' Start Logging
' -----------------------------------------------------------------------------

'Ensure log directory exists
CreateFolder("C:\RBFG\LOG")

'Build log path from variables above
LogPath = "C:\RBFG\LOG\" & "IIS8_InstallScript.log"

'Append to log file in C:\RBFG\LOG if it exists, otherwise create it
'ForAppending = 8, Create = True
Set Log = oFSO.OpenTextFile(LogPath, 8, True)

'Create marker between appended logs
Log.WriteLine(String("80","*") & vbCrLf & String("80","*") & vbCrLf & String("80","*"))

'Start logging
Log.WriteLine(vbCrLf & "Start processing: " & GetTimeStamp())


' -----------------------------------------------------------------------------
' Runtime Variables
' -----------------------------------------------------------------------------

If Err Then Err.Clear

'Determine OS type, default is x32
If Is64 Then
	OSType = "x64"
	x64Key = "Wow6432Node\"
Else
	OSType = "x32"
	x64Key = ""
End If

'Read registry keys for COE Server, COE Drive, Profile and Platform info
CodeSrv     = GetFromCOEHive("codesrv")
CodeDrv     = GetFromCOEHive("codedrv")
Profile     = GetFromCOEHive("Profile")
OS			= GetFromCOEHive("OS")
COEPlatform = GetFromCOEHive("Platform")
AssetTag    = GetFromCOEHive("AssetTag")
If Trim(AssetTag) = "" Then AssetTag = GetFromCOEHive("ComputerName")
CDBuild     = GetFromCOEHive("CDBuild")

'Read RBFG registry key, used later in script to see if app is installed
AppChk      = GetAppChk(DistCode)


' -----------------------------------------------------------------------------
' Determine CodePath
' -----------------------------------------------------------------------------

' run from relative path
CodePath = Replace(WScript.ScriptFullName, WScript.ScriptName, "") 
Log.WriteLine("CODEPATH: " & CodePath)

' -----------------------------------------------------------------------------
' MSI Variables
' -----------------------------------------------------------------------------

'build string references with initial variables (at the top of the file)
'so that a simple appcode, or path at the beginning of the vbs will propogate throughout this script
MSILog      = Chr(34) & "c:\rbfg\log\msi" & DistCode & "r" & RelNum & "_" & ProductName & "_" & MSIVer & ".log" & Chr(34)
UninstLog 	= Chr(34) & "c:\rbfg\log\msi" & DistCode & "r" & RelNum & "_" & ProductName & "_" & MSIVer & "_Uninstall.log" & Chr(34)

Call SetInstallPath()

' used by performlocalcopy sub too
Sub SetInstallPath()
	MSIPath     = CodePath & DistCode & ".msi"
	InstallPath = Chr(34) & MSIPath & Chr(34)
	
	'> uncomment the following lines if a transform is required
	'MSTPath     = CodePath & DistCode & ".mst"
	'InstallPath = InstallPath & " TRANSFORMS=" & Chr(34) & MSTPath & Chr(34)
	
	'> uncomment the following line if a patch is required
	'MSPPath = CodePath & DistCode & ".msp"
	
End Sub


' -----------------------------------------------------------------------------
' Determine Call to Execute
' -----------------------------------------------------------------------------

'Check the command line parameters, run the subroutine associated with it
Switch = UCase(WScript.Arguments(0))
Switch2 = UCase(WScript.Arguments(1))

Log.WriteLine("Launch Switch: " & Switch & ", Option = " & Switch2)

'Log system information
Call LogInfo()

'if new source wasn't located...abort installation"
If MSIError = -2 Then
	Log.WriteLine("------------------------------------------------------------------------------------------------")
	Log.WriteLine("MSIError = -2" & vbCrLf & CodePath & DistCode & "_" & MSIVer & ".txt Flag File Not Found in C:\Apps or Q:\" & vbCrLf & "Installation Aborted.") 
	Log.WriteLine("------------------------------------------------------------------------------------------------")
	WScript.Quit(-2)
End If  

Select Case Switch
	Case "PRIST"
		Mode = "/Qn"
		LogMode = " /L*V "
		Call CustomInstall
	Case "UNINST"
		Mode = "/Qn"
		LogMode = " /L*V "
		Call CustomUninstall
	Case "GUIINST"
		Mode = "/Qb-!"
		LogMode = " /L*V "
		MSIError = 0
		Log.WriteLine(vbCrLf & "Switch not applicable: " & Switch)
	Case "GUIUNINST"
		Mode = "/Qb-!"
		LogMode = " /L*V "
		MSIError = 0
		Log.WriteLine(vbCrLf & "Switch not applicable: " & Switch)
	Case "UPDATE"
		Mode = "/Qn"
		LogMode = " /L*V "
		MSIError = 0
		Log.WriteLine(vbCrLf & "Switch not applicable: " & Switch)
	Case "PATCH"
		Mode = "/Qn"
		LogMode = " /L*V "
		'Call Patch
		MSIError = 0
		Log.WriteLine(vbCrLf & "Switch not applicable: " & Switch)
	Case "TEST"
		'call any subroutine to test
		'GetFeatureStatus("IIS-WebServer")
		'Call Install_All
		'Call Install_FTP
		'Call Install_SMTP
	Case Else
		'invalid parameter
		MSIError = -101
End Select

'Check for a script launch with no parameter
If WScript.Arguments(0) = "" Then MSIError = -101
If WScript.Arguments(1) = "" Then MSIError = -102
  
'If the below is true, the vbs was not launched with a 'valid' parameter
If MSIError = -101 Then
	Log.WriteLine(vbCrLf)
	Log.WriteLine("--------------------------------")
	Log.WriteLine("Invalid Command Option provided!")
	Log.WriteLine("--------------------------------")
	Log.WriteLine("Please use the below commandline & Option(s):")
	Log.WriteLine("   ")
	Log.WriteLine("   " & WScript.ScriptName & " Prist ALL silent installation for All option")
	Log.WriteLine("   " & WScript.ScriptName & " Prist FTP  silent installation for FTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Prist SMTP silent installation for SMTP option")	
	Log.WriteLine("   " & WScript.ScriptName & " Uninst ALL silent uninstallation for All option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst ALL_AND_DEP silent uninstallation for All option and It's Dependency Features")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst FTP silent uninstallation for FTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst FTP_AND_DEP silent uninstallation for FTP option and It's Dependency Features")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst SMTP silent uninstallation for SMTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst SMTP_AND_DEP silent uninstallation for SMTP option and It's Dependency Features")
End If

'If the below is true, the vbs was not launched with a 'valid' parameter
If MSIError = -102 Then
	Log.WriteLine(vbCrLf)
	Log.WriteLine("--------------------------------")
	Log.WriteLine("Invalid Command Option provided!")
	Log.WriteLine("--------------------------------")
	Log.WriteLine("Please use the below commandline & Option(s):")
	Log.WriteLine("   ")
	Log.WriteLine("   " & WScript.ScriptName & " Prist ALL silent installation for All option")
	Log.WriteLine("   " & WScript.ScriptName & " Prist FTP  silent installation for FTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Prist SMTP silent installation for SMTP option")	
	Log.WriteLine("   " & WScript.ScriptName & " Uninst ALL silent uninstallation for All option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst ALL_AND_DEP silent uninstallation for All option and It's Dependency Features")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst FTP silent uninstallation for FTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst FTP_AND_DEP silent uninstallation for FTP option and It's Dependency Features")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst SMTP silent uninstallation for SMTP option")
	Log.WriteLine("   " & WScript.ScriptName & " Uninst SMTP_AND_DEP silent uninstallation for SMTP option and It's Dependency Features")
End If

Call Quit()

Sub Quit()

	Log.WriteLine(vbCrLf)
	Log.WriteLine("--------------------------------")
	Log.writeline("Install Error: " & MSIError)
	Log.WriteLine("--------------------------------")
	
	Log.WriteLine(vbCrLf & "End processing: " & GetTimeStamp() & vbCrLf)
	Log.Close
	
	'Script is done, exit and pass the error level
	'* If MSIError = 0, it means success 
	'* If MSIError = 3010, it means success but reboot is required.
	'* Else Install failed.
	WScript.Quit(MSIError)
End Sub


' -----------------------------------------------------------------------------
' Install
' -----------------------------------------------------------------------------

' PRIST   - perform a silent installation (no user inferface)
' GUIINST - perform installation with basic user interface (progress bar displayed)

Sub Install
	
	If Switch <> "UPDATE" Then Log.WriteLine(vbCrLf & GetTimeStamp() & "  * Start " & Switch)
	
	'check for existance of MSI and\or MST
	If CheckFile(MSIPath) And CheckFile(MSTPath) Then	
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing MSI Installation...")
		
		'set the installation command
		MsiexecCmd = "msiexec /I " & InstallPath & " " & Mode & " REBOOT=ReallySuppress" & LogMode & MSILog
		
		'execute the msi command and get its error code
		MSIError = ExecuteMSI(MsiexecCmd)
		
		'log status of the installation
		If MSIError = 0 Or MSIError = 3010 Then
			Log.WriteLine(GetTimeStamp() & "  Installation completed successfully.")
			Log.WriteLine(GetTimeStamp() & "  Return Code: " & MSIError)
			'add AppInfo hive
			'Call AppInfoHive() 		 '<-- Uncomment it for non-msi install
			'add DistCode
			'Call AddDistcode(DistCode)  '<-- Uncomment it for non-msi install
		Else		
			Log.WriteLine(GetTimeStamp() & "  Installation returned error code: " & MSIError & " - See " & MSILog & " for additional details.")
			'add AppInfo hive
			Call AppInfoHive()
		End If
	Else
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Unable to execute installation.")	  	
	End If
	
	If Switch <> "UPDATE" Then Log.WriteLine(vbCrLf & GetTimeStamp() & "  * End " & Switch)
	
End Sub


' -----------------------------------------------------------------------------
' Uninstall
' -----------------------------------------------------------------------------

' UNINST    - perform a silent uninstallation (no user interface)
' GUIUNINST - perform uninstallation with basic user interface (progress bar displayed)

Sub Uninstall
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * Start " & Switch)
	
	'main uninstallation
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing MSI Uninstallation...")
	
	If IsProductInstalled(ProductCode) Then
		MsiexecCmd = "msiexec /X" & ProductCode & " " & Mode & " REBOOT=ReallySuppress" & LogMode & UninstLog
		
		'execute the msi command and get its error code
		MSIError = ExecuteMSI(MsiexecCmd)
		
		'log status of the uninstallation
		If MSIError = 0 Or MSIError = 3010 Then
			Log.WriteLine(GetTimeStamp() & "  Uninstallation completed successfully.")
			Log.WriteLine(GetTimeStamp() & "  Return Code: " & MSIError)	
		
			'delete AppInfo hive
			Call DelAppInfo()
			'delete distribution code from the registry if it still exists after uninstallation
			Call DeleteDistcode(DistCode)
		Else
			Log.WriteLine(GetTimeStamp() & "  Uninstallation returned error code: " & MSIError & " - See " & UninstLog & " for additional details.")
		End If
	Else
		MSIError = 0
		Log.WriteLine(GetTimeStamp() & "  Product not found: " & ProductCode)
	End If
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * End " & Switch)
	
End Sub


' -----------------------------------------------------------------------------
' Major Update
' -----------------------------------------------------------------------------

' silent uninstallation of previous version and installation of the new version (no user interface)

Sub Major
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * Start MAJOR")
	
	'if old version is installed then uninstall old version and install new version
	MSIError = RemoveOldProducts
	
	If PREVERSIONFOUND = True Then
		If MSIError = 0 Or MSIError = 3010 Then	
			'install new version
			Call Install
		Else
			Log.WriteLine(GetTimeStamp() & "  Uninstallation returned error code: " & MSIError & " - See " & OldMSILog &" for additional details.")
			'add AppInfo hive
			Call AppInfoHive()
		End If
	Else
		'if old version is not installed then do nothing
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Previous installation does not exist - Update not executed.")
		MSIError = -1605
		'add AppInfo hive
		Call AppInfoHive()
	End If
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * End MAJOR")
	
End Sub


' -----------------------------------------------------------------------------
' Update
' -----------------------------------------------------------------------------

' silently reinstalls all files, user-specific registry entries and shortcuts (no user interface)

Sub Update
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * Start UPDATE")
		
	'check if old version is installed
	If IsProductInstalled(ProductCode) Then	
	
		'check for existance of MSI
		If CheckFile(MSIPath) Then
			
			Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing MSI Update...")
				
				InstallType = "UPDATE"
					
				'set installation command
				MsiexecCmd = "msiexec /Fvomus " & InstallPath & " " & Mode & LogMode & MSILog
				
				'execute the msi command and get its error code
				MSIError = ExecuteMSI(MsiexecCmd)
				
				'log status of the reinstallation
				If MSIError = 0 Or MSIError = 3010 Then
					Log.WriteLine(GetTimeStamp() & "  Update completed successfully.")
					
					'add AppInfo hive
					Call AppInfoHive()
					
					'add DistCode
					Call AddDistcode(DistCode)

				Else
					Log.WriteLine(GetTimeStamp() & "  Update returned error code: " & MSIError & " - See " & MSILog & " for additional details.")
					'add AppInfo hive
					Call AppInfoHive()
				End If
		Else
			Log.WriteLine(vbCrLf & GetTimeStamp() & "  Unable to execute update.")	  		
		End If
	
	Else
		'if old version is not installed then do nothing
		Log.WriteLine (vbCrLf & GetTimeStamp() & "  Previous installation does not exist - Update not executed.")
		MSIError = -1605
		'add AppInfo hive
		Call AppInfoHive()
	End If
  	
  	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * End UPDATE")

End Sub


' -----------------------------------------------------------------------------
' Patch
' -----------------------------------------------------------------------------

' applies a patch 

Sub Patch
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * Start PATCH")
	
	'check for existance of MSI and MSP
	If CheckFile(MSIPath) And CheckFile(MSPPath) Then
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Patch MSI...")
		
		MsiexecCmd = "msiexec /P " & Chr(34) & MSPPath & Chr(34) & " REINSTALLMODE=omus REBOOT=ReallySuppress " & Mode & LogMode & MSILog
		
		'execute the msi command and get its error code
		MSIError = ExecuteMSI(MsiexecCmd)
		
		'log status of the reinstallation
		If MSIError = 0 Or MSIError = 3010 Then
			Log.WriteLine(GetTimeStamp() & "  Patch completed successfully.")
			'Uncomment below to add AppInfo hive when needed
			'Call AppInfoHive()
		Else
			Log.WriteLine(GetTimeStamp() & "  Patch returned error code: " & MSIError & " - See " & MSILog & " for additional details.")
			'Uncomment below when needed to add AppInfo hive
			'Call AppInfoHive()
		End If
	Else
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Unable to execute patch.")
		'Uncomment below to add AppInfo hive when needed
		'Call AppInfoHive()
	End If
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  * End PATCH")	
End Sub


' -----------------------------------------------------------------------------
' Perform Local Copy
' -----------------------------------------------------------------------------

' copies package to c:\apps

Sub PerformLocalCopy()
	On Error Resume Next
	If Err Then Err.Clear
	
	Dim srcFolder, destFolder
	Dim fSrc, fDest
	Dim srcFolderSize, destFolderSize
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Performing Copy to Local Machine...")
	
	srcFolder  = oFSO.GetParentFolderName(WScript.ScriptFullName)
	destFolder = "C:\apps\" & AppCode & "\" & DistCode & "\" & Platform
	
	' create destination folder if it doesn't exist
	If Not oFSO.FolderExists(destFolder) Then Call CreateFolder(destFolder)
	
	Set fSrc  = oFSO.GetFolder(srcFolder)
	Set fDest = oFSO.GetFolder(destFolder)
	
	' get folder sizes for source comparison
	srcFolderSize  = fSrc.Size
	destFolderSize = fDest.Size
	
	If srcFolderSize = destFolderSize Then
		Log.WriteLine(GetTimeStamp() & "  Folder: " & destFolder)
		Log.WriteLine(GetTimeStamp() & "  Size:   " & FormatNumber(srcFolderSize, 0) & " bytes")
		Log.WriteLine(GetTimeStamp() & "  Local source is up-to-date; copy not required.")
	Else
		' copy is required
		oFSO.DeleteFolder destFolder, True
		If Err Then Err.Clear
		oFSO.CopyFolder srcFolder, destFolder, True
		If Err.Number = 0 Then
			Log.WriteLine(GetTimeStamp() & "  Folder: " & destFolder)
			Log.WriteLine(GetTimeStamp() & "  Size:   " & FormatNumber(srcFolderSize, 0) & " bytes")
			Log.WriteLine(GetTimeStamp() & "  Completed successfully.")
			CodePath = destFolder & "\"
			Call SetInstallPath()
		Else
			Log.WriteLine(GetTimeStamp() & "  Failed.")
			Log.WriteLine(GetTimeStamp() & "  Error Number: " & Err.Number)
			Log.WriteLine(GetTimeStamp() & "  Description:  " & Err.Description)
			Err.Clear
		End If
	End If
End Sub


' -----------------------------------------------------------------------------
' Static Subroutines
' -----------------------------------------------------------------------------

'*** Execute a command ***
Function ExecuteMSI(MsiexecCmd)
	' This receives and runs the msiexec command.
	'Call msi command
	Log.WriteLine(GetTimeStamp() & "  " & MsiexecCmd)
	Set oExec = WSHShell.Exec(MsiexecCmd)
	
	'loop until command has completed
	Do While oExec.Status = 0
		WScript.Sleep 100
	Loop
	
	'return the exit code 
	ExecuteMSI = oExec.ExitCode
End Function

'*** Delete the DistCode from the registry
Sub DeleteDistcode(DistCode)
	'This is a tidy up step.
	On Error Resume Next
	
	Dim regval, dcode
	
	dcode  = UCase(Trim(DistCode))
	regval = "HKLM\Software\" & x64Key & "RBFG\6F20\Desktop_Information\apps\" & dcode
	
	'determine if distcode still exists in the registry
	WSHShell.RegRead regval
	
	If Err = 0 Then	
		'if it still exists attempt to delete it
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Deleting Distribution Code from Registry: " & dcode)
		WSHShell.RegDelete regval
		
		'log status of the distcode deletion
		If Err = 0 Then
			Log.WriteLine(GetTimeStamp() & "  Distribution code has been deleted.")
		Else
			Log.WriteLine(GetTimeStamp() & "  Unable to delete distribution code.")
		End If
	End If
End Sub


'*** Create folder if it does not exist
Sub CreateFolder(strCreateFolder)
	On Error Resume Next
	
	'create recursive folders, if they do not exist
    If oFSO.FolderExists(strCreateFolder) Then
        Exit Sub
    Else
        CreateFolder(oFSO.GetParentFolderName(strCreateFolder))
    End If
    
    oFSO.CreateFolder(strCreateFolder)
    If Err <> 0 Then
		'unable to create log directory, write error to Event Viewer
		WSHShell.LogEvent 1, DistCode & ": MSIError = -2" & vbCrLf & "Unable to create " & strCreateFolder & " directory." & vbCrLf & "Installation aborted."
		MSIError = -2
		Call Quit()
	End If
End Sub


'*** Error checking routine ***
Sub ErrorCheck

	If Err = 0 Then
		Log.WriteLine(GetTimeStamp() & "  Completed successfully.")
	Else
		Log.WriteLine(GetTimeStamp() & "  Failed with error: " & Err & " - " & Err.Description & ".")
		Err.Clear
	End If

End Sub


'*** Determine if this is an x64 machine ***
Function Is64
On Error Resume Next

' This works on Windows 2000, XP, Win 7, as well as Server 2000, 2003, and 2008 (x64)
Dim strComputer, sOSBit 

	Is64 = True 
	sOSBit = WSHShell.RegRead("HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\PROCESSOR_ARCHITECTURE") 
	If ucase(sOSBit) = "X86" Then 
		Is64 = False 
	End If 

End Function 


'*** Log system information ***
Sub LogInfo()
	'Call LogReleaseInfo()
	'Call LogMSIDBInfo(MSIPath, MSTPath)
	Call LogCOEInfo()
	Call LogWinInstInfo()
	Call LogOSInfo()
	Call LogSoftwareInfo()
End Sub

Sub LogReleaseInfo()
'*** Log Release information
	Log.WriteLine(vbCrLf & "RELEASE Info")
	Log.WriteLine("============")
	Log.WriteLine("Release#: " & RelNum)
End Sub

'*** Log Information from COE registry hive
Sub LogCOEInfo()
	Log.WriteLine(vbCrLf & "COE Info")
	Log.WriteLine("========")
	Log.WriteLine("AppChk:   " & AppChk)
	Log.WriteLine("Platform: " & COEPlatform)
	Log.WriteLine("Profile:  " & Profile)
	Log.WriteLine("OS:       " & OS)
	Log.WriteLine("AssetTag: " & AssetTag)
End Sub


'*** Log Windows Installer information ***
Sub LogWinInstInfo()
	Dim strVersion, strPath
	Const SystemFolder = 1
	
	strPath = oFSO.GetSpecialFolder(SystemFolder) & "\msi.dll"
	
	If oFSO.FileExists(strPath) Then
		strVersion = oFSO.GetFileVersion(strPath)		
	Else
		strVersion = "Could not find " & strPath
	End If
	
	Log.WriteLine(vbCrLf & "Windows Installer")
	Log.WriteLine("=================")
	Log.WriteLine("msi.dll Version:   " & strVersion)
End Sub


'*** Log operating system information ***
Sub LogOSInfo()
	On Error Resume Next
	
	Dim strComputer
	Dim objWMIService, colOperatingSystems
	Dim objOperatingSystem
	
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colOperatingSystems = objWMIService.ExecQuery ("Select * from Win32_OperatingSystem")
	
	For Each objOperatingSystem in colOperatingSystems
		Log.WriteLine(vbCrLf & "Operating System")
		Log.WriteLine("================")
		Log.WriteLine("OS:                " & OSType)
		Log.WriteLine("Caption:           " & objOperatingSystem.Caption)
		Log.WriteLine("Version:           " & objOperatingSystem.Version)
		Log.WriteLine("Build Number:      " & objOperatingSystem.BuildNumber)
		Log.WriteLine("OS Type:           " & objOperatingSystem.OSType)
		Log.WriteLine("Service Pack:      " & objOperatingSystem.ServicePackMajorVersion & "." & objOperatingSystem.ServicePackMinorVersion)
		Log.WriteLine("Windows Directory: " & objOperatingSystem.WindowsDirectory)
	Next
	
	Set objWMIService = Nothing
	Set colOperatingSystems = Nothing
End Sub


'*** Log software versions and information ***
Sub LogSoftwareInfo()
	On Error Resume Next
	
	Dim objWMIService, colWMISettings, objWMISetting
	Dim strComputer, strADSIVersion
	
	Log.WriteLine(vbCrLf & "Software Versions & Information")
	Log.WriteLine("===============================")
	Log.WriteLine("Script Host:       " & WScript.FullName)
	Log.WriteLine("WSH Version:       " & WScript.Version)
	Log.WriteLine("VBScript Version:  " & ScriptEngineMajorVersion & "." & ScriptEngineMinorVersion & "." & ScriptEngineBuildVersion)
	
	'wmi version
	strComputer = "."
	Set objWMIService  = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
	Set colWMISettings = objWMIService.ExecQuery("Select * from Win32_WMISetting")
	For Each objWMISetting in colWMISettings
		Log.WriteLine("WMI Version:       " & objWMISetting.BuildVersion)
	Next
	
	'adsi version
	strADSIVersion = WSHShell.RegRead("HKLM\SOFTWARE\Microsoft\Active Setup\Installed Components\{E92B03AB-B707-11d2-9CBD-0000F87A369E}\Version")
	If strADSIVersion = vbEmpty Then
		strADSIVersion = WSHShell.RegRead("HKLM\SOFTWARE\Microsoft\ADs\Providers\LDAP\")
		If strADSIVersion = vbEmpty Then
			strADSIVersion = "ADSI is not installed."
		Else
			strADSIVersion = "2.0"
		End If
	End If
	Log.WriteLine("ADSI Version:      " & strADSIVersion)
	
	Set objWMIService  = Nothing
	Set colWMISettings = Nothing
	
	If Err Then Err.Clear
End Sub


'*** Log MSI Database Information ***
Sub LogMSIDBInfo(strMSIPath, strMSTPath)
	On Error Resume Next
	Dim msi, mst
	
	msi = Replace(strMSIPath, CodePath, "")
	
	If strMSTPath = "" And Not oFSO.FileExists(strMSTPath) Then
		mst = "N/A"
	ElseIf strMSTPath = "" Then
		mst = "None"
	ElseIf Not oFSO.FileExists(strMSTPath) Then
		mst = "Not found"
	Else
		mst = Replace(strMSTPath, CodePath, "")
	End If
	
	Log.WriteLine(vbCrLf & "MSI Database Info")
	Log.WriteLine("=================")
	
	If Not oFSO.FileExists(strMSIPath) Then
		Log.WriteLine("Package not found: " & msi)
	ElseIf InStr(1, msi, ".msi", 1) = 0 Then
		Log.WriteLine("Package: " & msi)
		Log.WriteLine("Version: " & oFSO.GetFileVersion(strMSIPath))
	Else
		Log.WriteLine("MSI Database:    " & msi)
		Log.WriteLine("Transform:       " & mst)
		Log.WriteLine("Product Name:    " & GetMSIProperty(strMSIPath, strMSTPath, "ProductName"))
		Log.WriteLine("Product Version: " & GetMSIProperty(strMSIPath, strMSTPath, "ProductVersion"))
		Log.WriteLine("Product Code:    " & GetMSIProperty(strMSIPath, strMSTPath, "ProductCode"))
	End If
	
	If Err Then Err.Clear
End Sub


' return True  if file is defined and exists
'        True  if file is undefined and doesn't exists (i.e. transforms)
'        False otherwise
Function CheckFile(strFile)
	Dim file, bRet
	
	file = Trim(strfile)
	
	If Not file = "" And oFSO.FileExists(file) Then
		bRet = True
	ElseIf file = "" And Not oFSO.FileExists(file) Then
		bRet = True
	Else
		If Not file = "" Then Log.WriteLine(vbCrLf & GetTimeStamp() & "  File not found: " & file)
		MSIError = 1612
		'add AppInfo hive
		Call AppInfoHive()
		bRet = False
	End If
	
	CheckFile = bRet
End Function


' -----------------------------------------------------------------------------
' Helper Functions
' -----------------------------------------------------------------------------

' return True  if the product is installed
'        False otherwise
Function IsProductInstalled(ByVal strProductCode)
	On Error Resume Next

	Dim regKey, GUID, rc

	If Err Then Err.Clear
	GUID = Trim(strProductCode)
	regKey = "HKLM\SOFTWARE\" & x64Key & "Microsoft\Windows\CurrentVersion\Uninstall\" & GUID
	rc = RegKeyExists(regKey)

	'This logic is added to cover for some 32bits/64bits MSI installed on 64bit OS but
	If (Not rc) And x64Key <> "" Then
		regKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & GUID
		rc = RegKeyExists(regKey)
	End If
	
	IsProductInstalled = rc
	If Err Then Err.Clear

End Function


' returns True  if registry key exists
'         False otherwise
Function RegKeyExists(ByVal strRegistryKey)

	On Error Resume Next
	
	Dim sDescription, oShell, sRegKey

	If Err Then Err.Clear

	Set oShell = CreateObject("WScript.Shell")
	
	sRegKey = Trim(strRegistryKey)
	RegKeyExists = True
	
	sRegKey = Trim(sRegKey)
	If Not Right(sRegKey, 1) = "\" Then sRegKey = sRegKey & "\"
	
	oShell.RegRead "HKEYNotAKey\"
	sDescription = Replace(Err.Description, "HKEYNotAKey\", "")
	
	If Err Then Err.Clear
	oShell.RegRead sRegKey
	RegKeyExists = sDescription <> Replace(Err.Description, sRegKey, "")
	
	On Error Goto 0
	
End Function


' return AppChk value from registry (if exist)
'        empty string otherwise
Function GetAppChk(ByVal strDistCode)

	On Error Resume Next
	If Err Then Err.Clear
	
	Dim dcode, strRet
	
	dcode  = UCase(Trim(strDistCode))
	strRet = ""
	strRet = WSHShell.RegRead("HKLM\Software\RBFG\6F20\Desktop_Information\Apps\" & dcode)
	
	If Err Then
		If Is64 Then
			strRet = WSHShell.RegRead("HKLM\Software\" & x64Key & "RBFG\6F20\Desktop_Information\Apps\" & dcode)
			If Err Then Err.Clear
		Else
			Err.Clear
		End If
	End If
	
	GetAppChk = strRet
End Function


' return string data of specified value from COE hive
'        N/A otherwise
Function GetFromCOEHive(Name)
	If Err Then Err.Clear
	On Error Resume Next
	
	Dim strName, key32, key64
	
	strName = vbEmpty
	key32 = "HKEY_LOCAL_MACHINE\SOFTWARE\COE\" & Name
	key64 = "HKEY_LOCAL_MACHINE\SOFTWARE\" & x64Key & "COE\" & Name
	
	If Is64() Then strName = WSHShell.RegRead(key64)
	If strName = vbEmpty Then
		strName = WSHShell.RegRead(key32)
		If strName = vbEmpty Then strName = "N/A"
	End If
	
	GetFromCOEHive = strName
	If Err Then Err.Clear
End Function


' return string value of property from MSI database
Function GetMSIProperty(strMSIDatabase, strTransform, strProperty)
	If Err Then Err.Clear
	
	Dim objWI, DB, View, record
	Dim strRet
	
	Const ReadOnly = 0
	
	If oFSO.FileExists(strMSIDatabase) Then
		Set objWI = CreateObject("WindowsInstaller.Installer")
		Set DB    = objWI.OpenDatabase(strMSIDatabase, ReadOnly)
		Set View  = DB.OpenView("Select `Value` From Property WHERE `Property` = '" & Trim(strProperty) & "'")
		
		If Not strTransform = "" And oFSO.FileExists(strTransform) Then DB.ApplyTransform strTransform, 0
		
		View.Execute
		Set record = View.Fetch	
		If record Is Nothing Then
			strRet = "Property not found"
		Else
			strRet = record.StringData(1)
		End If
	Else
		strRet = "MSI database not found"
	End If
	
	GetMSIProperty = strRet
	
	Set objWI = Nothing
	Set DB    = Nothing
	Set View  = Nothing
End Function

Function RemoveOldProducts

	On Error Resume Next
	
	Dim OldProdCodes(0), VersionKey, DisplayVersion, NameKey, DisplayName, i, cmd, rc
	Dim UninstRegPath
	
	rc = 0
	
	
	
	OldProdCodes(0) = "{NNNNNNNN-NNNN-NNNN-NNNN-NNNNNNNNNNNN}"  'Current
	'OldProdCodes(1) = "{NNNNNNNN-NNNN-NNNN-NNNN-NNNNNNNNNNNN}"  'Old1
	'OldProdCodes(2) = "{NNNNNNNN-NNNN-NNNN-NNNN-NNNNNNNNNNNN}"  'Old2

	
	For i = 0 To UBound(OldProdCodes)
		If IsProductInstalled(OldProdCodes(i)) Then
			
			
			Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing MSI Uninstallation of previous version...")
		
			InstallType = "UPDATE"
			PREVERSIONFOUND = True
			
			UninstRegPath  = GetUninstProductCodeRegPath(OldProdCodes(i))
			NameKey        = UninstRegPath & "\DisplayName"
			VersionKey     = UninstRegPath & "\DisplayVersion"
			DisplayName    = WSHShell.RegRead(NameKey)
			DisplayVersion = WSHShell.RegRead(VersionKey)
			
			OldMSILog = " c:\RBFG\LOG\msi" & DistCode &  "rOld(v" & DisplayVersion & ").log"
			
			Log.WriteLine(vbCrLf & GetTimeStamp() & "  Removing " & DisplayName & " v" & DisplayVersion & "...")
			cmd = "msiexec /X" & OldProdCodes(i) & " REBOOT=ReallySuppress " & Mode & " " & LogMode & OldMSILog
			rc = ExecuteMSI(cmd)
			If rc = 0 Or rc = 3010 Then
				Log.WriteLine(GetTimeStamp() & "  Completed successfully.")
				'delete AppInfo hive
				Call DelAppInfo()
				'delete distribution code from the registry if it still exists after uninstallation
				Call DeleteDistcode(DistCode)		
			Else
				Log.WriteLine(GetTimeStamp() & "  Failed with " & rc)
			End If
			'Log.WriteLine("Return code: " & rc)
		End If
	Next
	
	RemoveOldProducts = rc
	
End Function

'---------------------
'-- delete app info --
'---------------------
Sub DelAppInfo()
	If Err Then Err.Clear
	On Error Resume Next
	
	Dim AppInstallKey
	Dim arrName(6)
	Dim i, max
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Removing Entry from APPLICATION_INSTALLS Hive...")
	
	AppInstallKey = "HKLM\SOFTWARE\" & x64Key & "APPLICATION_INSTALLS\" & DistCode
	
	arrName(0) = "APPLICATION_NAME"
	arrName(1) = "VERSION_NUMBER"
	arrName(2) = "INSTALL_DATE"
	arrName(3) = "INSTALLED_BY"
	arrName(4) = "INSTALL_STATUS"
	arrName(5) = "INSTALLED_FROM"
	arrName(6) = "RELEASE_EVENT"
	
	max = UBound(arrName)
	For i = 0 To max
		WSHShell.RegDelete(AppInstallKey & "\" & arrName(i))
		Log.WriteLine(GetTimeStamp() & "  Value: " & AppInstallKey & "\" & arrName(i))
		If Err = 0 Then
			Log.WriteLine(GetTimeStamp() & "  Completed successfully.")
		ElseIf Err = -2147024894 Then
			Log.WriteLine(GetTimeStamp() & "  Already removed by MSI.")
		Else
			Log.WriteLine(GetTimeStamp() & "  Failed with error: " & Err & " - " & Err.Description & ".")
			Err.Clear
		End If
	Next
	
	If Err Then Err.Clear
End Sub

'------------------
'-- add app info --
'-------------------
Sub AppInfoHive()
	If Err Then Err.Clear
	
	Dim AppInstallKey, AppHistoryKey
	Dim arrRegValue(6), strData(6), strType
	Dim arrHistory(6)
	Dim arrAppInfo(1)
	Dim i, j, max1, max2
	
	AppInstallKey = "HKLM\SOFTWARE\" & x64Key & "APPLICATION_INSTALLS\" & DistCode
	AppHistoryKey = AppInstallKey & "\History\" & MSIVer
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Adding Entry to APPLICATION_INSTALLS Hive...")
	
	'application_installs
	arrRegValue(0) = AppInstallKey & "\APPLICATION_NAME"
	arrRegValue(1) = AppInstallKey & "\VERSION_NUMBER"
	arrRegValue(2) = AppInstallKey & "\INSTALL_DATE"
	arrRegValue(3) = AppInstallKey & "\INSTALLED_BY"
	arrRegValue(4) = AppInstallKey & "\INSTALL_STATUS"
	arrRegValue(5) = AppInstallKey & "\INSTALLED_FROM"
	arrRegValue(6) = AppInstallKey & "\RELEASE_EVENT"
	
	'history
	arrHistory(0)  = AppHistoryKey & "\APPLICATION_NAME"
	arrHistory(1)  = AppHistoryKey & "\VERSION_NUMBER"
	arrHistory(2)  = AppHistoryKey & "\INSTALL_DATE"
	arrHistory(3)  = AppHistoryKey & "\INSTALLED_BY"
	arrHistory(4)  = AppHistoryKey & "\INSTALL_STATUS"
	arrHistory(5)  = AppHistoryKey & "\INSTALLED_FROM"
	arrHistory(6)  = AppHistoryKey & "\RELEASE_EVENT"
	
	'data
	strData(0) = ProductName
	'strData(1) = GetProductVersion(ProductCode)
	strData(1) = MSIVer
	strData(2) = GetTimeStamp()
	strData(3) = UserName
	If MSIError = 0 Or MSIError = 3010 Then
		strData(4) = "SUCCESS:" & MSIError
	Else
		strData(4) = "FAIL:" & MSIError
	End If
	
	strData(5) = CodePath
	strData(6) = RelNum
	
	strType = "REG_SZ"
	
	arrAppInfo(0) = arrRegValue
	arrAppInfo(1) = arrHistory
	
	max1 = UBound(arrAppInfo)
	For i = 0 To max1
		max2 = UBound(arrAppInfo(i))
		For j = 0 To max2
			WSHShell.RegWrite arrAppInfo(i)(j), strData(j), strType
			Log.WriteLine(GetTimeStamp() & "  Value: " & arrAppInfo(i)(j))
			Log.WriteLine(GetTimeStamp() & "  Data:  " & strData(j))
			ErrorCheck
		Next
		Log.WriteLine()
	Next
	
	If Err Then Err.Clear
End Sub

Sub AddDistcode(DistributionCode)
	On Error Resume Next
	If Err Then Err.Clear
	
	Dim value, rc
	Dim dcode
	
	dcode = UCase(Trim(DistributionCode))
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Adding Distribution Code to Registry: " & dcode)
	
	value = "HKLM\Software\" & x64Key & "RBFG\6F20\Desktop_Information\apps\" & dcode
	
	rc = WSHShell.RegRead(value)
	If Err = 0 Then
		Log.WriteLine(GetTimeStamp() & "  Already in registry: " & value)
	Else
		Err.Clear
	  	WSHShell.RegWrite value, "1", "REG_SZ"
		ErrorCheck
	End If
End Sub

'format number of digits
Function pd(n, totalDigits)
	If totalDigits > Len(n) Then
		pd = String(totalDigits-Len(n),"0") & n
	Else
		pd = n
	End If
End Function

'return time stamp in format YYYY-MM-DD HH:MM:SS
Function GetTimeStamp()
	Dim thisDate, thisTime
	
	thisDate = Year(Date()) & "-" & pd(Month(Date()),2) & "-" & pd(Day(Date()),2)
	thisTime = FormatDateTime(Now(), 4) & ":" & pd(Second(Now()),2)
	
	GetTimeStamp = thisDate & ", " & thisTime
End Function

'to make logging easier
Sub WriteLog(strText)
	Log.WriteLine(GetTimeStamp() & vbTab & "| " & strText)
End Sub

'This Function is added due to some application when it installed on 64bits machine,
'the uninstall product code key path still exist in 32bit location
'Return correct Uninstall Product Code Registry key path if found
'Otherwise return empty string ("")
Function GetUninstProductCodeRegPath(strProductCode)
	On Error Resume Next

	Dim regKey, GUID
	Dim rc
	
	GUID = Trim(strProductCode)
	
	regKey = "HKLM\SOFTWARE\" & x64Key & "Microsoft\Windows\CurrentVersion\Uninstall\" & GUID
	rc = RegKeyExists(regKey)
	
	'If not found and it is 64bit machinne, check x64 key path once more
	If Not rc And x64Key <> "" Then
		regKey = "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\" & GUID
		If Not RegKeyExists(regKey) Then
			regKey = ""
		End If
	End If

	GetUninstProductCodeRegPath = regKey
	
End Function

' -----------------------------------------------------------------------------
' Custom Subroutines
' -----------------------------------------------------------------------------
Function Install_All

	On Error Resume Next
	
	Dim IIS_Feature(54), IIS_Feature_Status(54)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, strStatus
	Dim RebootPending, status
		
	rc = 0
	MSIError = 0
	RebootPending = False	
	IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-HttpErrors"
	IIS_Feature(6)  = "IIS-HttpRedirect"
	IIS_Feature(7)  = "IIS-ASPNET"
	IIS_Feature(8)  = "IIS-NetFxExtensibility"
	IIS_Feature(9)  = "IIS-HealthAndDiagnostics"
	IIS_Feature(10) = "IIS-HttpLogging"	
	IIS_Feature(11) = "IIS-RequestMonitor"
	IIS_Feature(12) = "IIS-Security"		
	IIS_Feature(13) = "IIS-BasicAuthentication"	
	IIS_Feature(14) = "IIS-ClientCertificateMappingAuthentication"	
	IIS_Feature(15) = "IIS-DigestAuthentication"
	IIS_Feature(16) = "IIS-IISCertificateMappingAuthentication" 	
	IIS_Feature(17) = "IIS-IPSecurity"
	IIS_Feature(18) = "IIS-RequestFiltering"
	IIS_Feature(19) = "IIS-URLAuthorization"	
	IIS_Feature(20) = "IIS-WindowsAuthentication"	
	IIS_Feature(21) = "IIS-Performance"	
	IIS_Feature(22) = "IIS-HttpCompressionStatic"	
	IIS_Feature(23) = "IIS-ManagementConsole"	
	IIS_Feature(24) = "IIS-ManagementService"	
	IIS_Feature(25) = "IIS-ManagementScriptingTools"	
	IIS_Feature(26) = "IIS-LoggingLibraries"
	IIS_Feature(27) = "IIS-WebServerManagementTools"
	IIS_Feature(28) = "IIS-CertProvider"
	IIS_Feature(29) = "WAS-WindowsActivationService"
	IIS_Feature(30) = "WAS-ProcessModel"
	IIS_Feature(31) = "WAS-NetFxEnvironment"
	IIS_Feature(32) = "WAS-ConfigurationAPI"
	IIS_Feature(33) = "IIS-ApplicationDevelopment"
	IIS_Feature(34) = "IIS-ISAPIExtensions"
	IIS_Feature(35) = "IIS-ISAPIFilter"
	IIS_Feature(36) = "WCF-HTTP-Activation"		
	IIS_Feature(37) = "NetFx3"
	IIS_Feature(38) = "IIS-NetFxExtensibility45"
	IIS_Feature(39) = "IIS-ApplicationInit"
	IIS_Feature(40) = "NetFx4Extended-ASPNET45"
	IIS_Feature(41) = "IIS-ASPNET45"
	IIS_Feature(42) = "IIS-WebSockets"
	IIS_Feature(43) = "IIS-ServerSideIncludes"
	IIS_Feature(44) = "IIS-Metabase"	
	IIS_Feature(45) = "IIS-WMICompatibility"	
	IIS_Feature(46) = "IIS-LegacyScripts"
	IIS_Feature(47) = "IIS-LegacySnapIn"	
	IIS_Feature(48) = "IIS-ASP"
	IIS_Feature(49) = "IIS-CGI"
	IIS_Feature(50) = "IIS-WebDAV"
	IIS_Feature(51) = "IIS-IIS6ManagementCompatibility"
	IIS_Feature(52) = "IIS-HttpTracing"
	IIS_Feature(53) = "IIS-CustomLogging"
	IIS_Feature(54) = "IIS-HttpCompressionDynamic"	

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Install All option: " & strCmd & vbCrLf )		
	
	'Perform feature enable for all required feature and its dependencies
	strCmd = "cmd /c DISM.exe /Online /Enable-Feature"
	For i = 0 To UBound(IIS_Feature)
		strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 5000

	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			MSIError = 3010
		End If
	Else
		MSIError = rc
	End If

	If rc = 0 Or rc = 3010 Then
		'Using Powershell to enable All feature
		strCmd = "cmd /c PowerShell.exe -File " & Chr(34) & CodePath & "InstallScripts\InstallAllFeatures.ps1" & Chr(34)
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the following command to enable SMTP feature:")
		Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)
		If rc = 0 Or rc = 3010 Then
			If rc = 3010 Then
				MSIError = 3010
			End If
		Else
			MSIError = rc
		End If
		WScript.Sleep 2000
	End If
	
	If rc = 0 Or rc = 3010 Then
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the hardening for ALL option:")

		strCmd = "cmd /c " & Chr(34) & CodePath & "InstallScripts\MoveIISRoot.bat" & Chr(34) & " >> C:\RBFG\Log\MoveIISRoot_bat.log"
		Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)
		
		strCmd = "cmd /c " & Chr(34) & CodePath & "InstallScripts\ASP.bat" & Chr(34) &  " >> C:\RBFG\Log\ASP_bat.log"
		Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)		
	End If			

	'Get All Feature Status	
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\AllFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get All feature status:")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then

		'Display All Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display All feature status: " & vbCrLf )
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If
	
End Function

Function Install_FTP

	On Error Resume Next
	
	Dim IIS_Feature(43), IIS_Feature_Status(43)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, strStatus
	Dim RebootPending, status
		
	rc = 0
	MSIError = 0
	RebootPending = False	
	
	IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-DirectoryBrowsing"
	IIS_Feature(6)  = "IIS-HttpErrors"
	IIS_Feature(7)  = "IIS-HttpRedirect"
	IIS_Feature(8)  = "IIS-ASPNET"
	IIS_Feature(9)  = "IIS-NetFxExtensibility"
	IIS_Feature(10) = "NetFx4Extended-ASPNET45"
	IIS_Feature(11) = "IIS-NetFxExtensibility45"	
	IIS_Feature(12) = "IIS-HealthAndDiagnostics"
	IIS_Feature(13) = "IIS-HttpLogging"	
	IIS_Feature(14) = "IIS-RequestMonitor"
	IIS_Feature(15) = "IIS-Security"		
	IIS_Feature(16) = "IIS-BasicAuthentication"	
	IIS_Feature(17) = "IIS-ClientCertificateMappingAuthentication"
	IIS_Feature(18) = "IIS-DigestAuthentication"
	IIS_Feature(19) = "IIS-IISCertificateMappingAuthentication"		
	IIS_Feature(20) = "IIS-IPSecurity"
	IIS_Feature(21) = "IIS-RequestFiltering"
	IIS_Feature(22) = "IIS-URLAuthorization"	
	IIS_Feature(23) = "IIS-WindowsAuthentication"	
	IIS_Feature(24) = "IIS-Performance"	
	IIS_Feature(25) = "IIS-HttpCompressionStatic"	
	IIS_Feature(26) = "IIS-ManagementConsole"	
	IIS_Feature(27) = "IIS-ManagementService"	
	IIS_Feature(28) = "IIS-ManagementScriptingTools"	
	IIS_Feature(29) = "IIS-LoggingLibraries"
	IIS_Feature(30) = "IIS-WebServerManagementTools"
	IIS_Feature(31) = "IIS-CertProvider"
	IIS_Feature(32) = "WAS-WindowsActivationService"
	IIS_Feature(33) = "WAS-ProcessModel"
	IIS_Feature(34) = "WAS-NetFxEnvironment"
	IIS_Feature(35) = "WAS-ConfigurationAPI"
	IIS_Feature(36) = "IIS-ApplicationDevelopment"
	IIS_Feature(37) = "IIS-ISAPIExtensions"
	IIS_Feature(38) = "IIS-ISAPIFilter"
	IIS_Feature(39) = "WCF-HTTP-Activation"
	IIS_Feature(40) = "NetFx3"
	'Main Feature for FTP Server
	IIS_Feature(41) = "IIS-FTPServer"
	IIS_Feature(42) = "IIS-FTPSvc"
	IIS_Feature(43) = "IIS-FTPExtensibility"
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Install FTP option: " & strCmd & vbCrLf )

	'Perform feature enable for all required feature and its dependencies
	strCmd = "cmd /c DISM.exe /Online /Enable-Feature"
	For i = 0 To UBound(IIS_Feature)
		strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 2000
	
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			MSIError = 3010
		End If
	Else
		MSIError = rc
	End If

	'Using Powershell to enable FTP feature
	strCmd = "cmd /c PowerShell.exe -File " & Chr(34) & CodePath & "InstallScripts\InstallFTPFeatures.ps1" & Chr(34)
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the following command to enable FTP feature:")
	Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			MSIError = 3010
		End If
	Else
		MSIError = rc
	End If
	WScript.Sleep 2000		

	If rc = 0 Or rc = 3010 Then
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the hardening for FTP option:")
		strCmd = "cmd /c " & Chr(34) & CodePath & "InstallScripts\MoveIISRoot.bat" & Chr(34) & " >> C:\RBFG\Log\MoveIISRoot_bat.log"
		Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)		
	End If			

	'Get FTP Feature Status	
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\FTPFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get FTP feature status:")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then

		'Display FTP Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display FTP feature status: " & vbCrLf )
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If

End Function

Function Install_SMTP

	On Error Resume Next
	
	Dim IIS_Feature(47), IIS_Feature_Status(47)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, strStatus
	Dim RebootPending, status
		
	rc = 0
	MSIError = 0
	RebootPending = False
	
	IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-DirectoryBrowsing"
	IIS_Feature(6)  = "IIS-HttpErrors"
	IIS_Feature(7)  = "IIS-HttpRedirect"
	IIS_Feature(8)  = "IIS-ASPNET"
	IIS_Feature(9)  = "IIS-NetFxExtensibility"
	IIS_Feature(10) = "NetFx4Extended-ASPNET45"
	IIS_Feature(11) = "IIS-NetFxExtensibility45"	
	IIS_Feature(12) = "IIS-HealthAndDiagnostics"
	IIS_Feature(13) = "IIS-HttpLogging"	
	IIS_Feature(14) = "IIS-RequestMonitor"
	IIS_Feature(15) = "IIS-Security"		
	IIS_Feature(16) = "IIS-BasicAuthentication"	
	IIS_Feature(17) = "IIS-ClientCertificateMappingAuthentication"
	IIS_Feature(18) = "IIS-DigestAuthentication"
	IIS_Feature(19) = "IIS-IISCertificateMappingAuthentication"		
	IIS_Feature(20) = "IIS-IPSecurity"
	IIS_Feature(21) = "IIS-RequestFiltering"
	IIS_Feature(22) = "IIS-URLAuthorization"	
	IIS_Feature(23) = "IIS-WindowsAuthentication"	
	IIS_Feature(24) = "IIS-Performance"	
	IIS_Feature(25) = "IIS-HttpCompressionStatic"	
	IIS_Feature(26) = "IIS-ManagementConsole"	
	IIS_Feature(27) = "IIS-ManagementService"	
	IIS_Feature(28) = "IIS-ManagementScriptingTools"	
	IIS_Feature(29) = "IIS-LoggingLibraries"
	IIS_Feature(30) = "IIS-WebServerManagementTools"
	IIS_Feature(31) = "IIS-CertProvider"
	IIS_Feature(32) = "WAS-WindowsActivationService"
	IIS_Feature(33) = "WAS-ProcessModel"
	IIS_Feature(34) = "WAS-NetFxEnvironment"
	IIS_Feature(35) = "WAS-ConfigurationAPI"
	IIS_Feature(36) = "IIS-ApplicationDevelopment"
	IIS_Feature(37) = "IIS-ISAPIExtensions"
	IIS_Feature(38) = "IIS-ISAPIFilter"
	IIS_Feature(39) = "WCF-HTTP-Activation"
	IIS_Feature(40) = "NetFx3"
	'Other Dependencies installed by InstallSMTPFeature.ps1
	IIS_Feature(41) = "IIS-LegacySnapIn"	
	IIS_Feature(42) = "IIS-Metabase"
	IIS_Feature(43) = "IIS-IIS6ManagementCompatibility"	
	'SMTP main features
	IIS_Feature(44) = "ServerManager-Core-RSAT"	
	IIS_Feature(45) = "ServerManager-Core-RSAT-Feature-Tools"	
	IIS_Feature(46) = "Smtpsvc-Service-Update-Name"	
	IIS_Feature(47) = "Smtpsvc-Admin-Update-Name"
		
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Install SMTP option: " & vbCrLf )

	'Perform feature enable for all required dependency features
	strCmd = "cmd /c DISM.exe /Online /Enable-Feature"
	For i = 0 To 40
		strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 5000

	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			MSIError = 3010
		End If
	Else
		MSIError = rc
	End If
	
	'Using Powershell to enable SMTP feature
	strCmd = "cmd /c PowerShell.exe -File " & Chr(34) & CodePath & "InstallScripts\InstallSMTPFeatures.ps1" & Chr(34)
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the following command to enable SMTP feature:")
	Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			MSIError = 3010
		End If
	Else
		MSIError = rc
	End If
	WScript.Sleep 2000		

	If rc = 0 Or rc = 3010 Then
		Log.WriteLine(vbCrLf & GetTimeStamp() & "  Execute the hardening for SMTP option:")
		strCmd = "cmd /c " & Chr(34) & CodePath & "InstallScripts\MoveIISRoot.bat" & Chr(34) & " >> C:\RBFG\Log\MoveIISRoot_bat.log"
		Log.WriteLine(GetTimeStamp() & "  Run command = " & strCmd)	
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & "  Return Code = " & rc)		
	End If	
	
	'Get SMTP Feature Status
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\SMTPFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get SMTP feature status:")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then
		
		'Display SMTP Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display SMTP feature status: " & vbCrLf )
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If	

End Function

Function Uninstall_All

	On Error Resume Next
	
	Dim IIS_Feature(53), IIS_Feature_Status(53)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, StrStatus
	Dim RebootPending, status
		
	rc = 0
	MSIError = 0
	RebootPending = False
		IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-HttpErrors"
	IIS_Feature(6)  = "IIS-HttpRedirect"
	IIS_Feature(7)  = "IIS-ASPNET"
	IIS_Feature(8)  = "IIS-NetFxExtensibility"
	IIS_Feature(9)  = "IIS-HealthAndDiagnostics"
	IIS_Feature(10) = "IIS-HttpLogging"	
	IIS_Feature(11) = "IIS-RequestMonitor"
	IIS_Feature(12) = "IIS-Security"		
	IIS_Feature(13) = "IIS-BasicAuthentication"	
	IIS_Feature(14) = "IIS-ClientCertificateMappingAuthentication"	
	IIS_Feature(15) = "IIS-DigestAuthentication"
	IIS_Feature(16) = "IIS-IISCertificateMappingAuthentication" 	
	IIS_Feature(17) = "IIS-IPSecurity"
	IIS_Feature(18) = "IIS-RequestFiltering"
	IIS_Feature(19) = "IIS-URLAuthorization"	
	IIS_Feature(20) = "IIS-WindowsAuthentication"	
	IIS_Feature(21) = "IIS-Performance"	
	IIS_Feature(22) = "IIS-HttpCompressionStatic"	
	IIS_Feature(23) = "IIS-ManagementConsole"	
	IIS_Feature(24) = "IIS-ManagementService"	
	IIS_Feature(25) = "IIS-ManagementScriptingTools"	
	IIS_Feature(26) = "IIS-LoggingLibraries"
	IIS_Feature(27) = "IIS-WebServerManagementTools"
	IIS_Feature(28) = "IIS-CertProvider"
	IIS_Feature(29) = "WAS-WindowsActivationService"
	IIS_Feature(30) = "WAS-ProcessModel"
	IIS_Feature(31) = "WAS-NetFxEnvironment"
	IIS_Feature(32) = "WAS-ConfigurationAPI"
	IIS_Feature(33) = "IIS-ApplicationDevelopment"
	IIS_Feature(34) = "IIS-ISAPIExtensions"
	IIS_Feature(35) = "IIS-ISAPIFilter"
	IIS_Feature(36) = "WCF-HTTP-Activation"		
	IIS_Feature(37) = "NetFx3"
	'Do Not Disable NetFx3 on uninstall
	IIS_Feature(37) = ""		
	IIS_Feature(38) = "IIS-NetFxExtensibility45"
	IIS_Feature(39) = "IIS-ApplicationInit"
	IIS_Feature(40) = "NetFx4Extended-ASPNET45"
	IIS_Feature(41) = "IIS-ASPNET45"
	IIS_Feature(42) = "IIS-WebSockets"
	IIS_Feature(43) = "IIS-ServerSideIncludes"
	IIS_Feature(44) = "IIS-Metabase"	
	IIS_Feature(45) = "IIS-WMICompatibility"	
	IIS_Feature(46) = "IIS-LegacyScripts"
	IIS_Feature(47) = "IIS-LegacySnapIn"	
	'All Main Features	
	IIS_Feature(48) = "IIS-ASP"
	IIS_Feature(49) = "IIS-CGI"
	IIS_Feature(50) = "IIS-WebDAV"
	IIS_Feature(51) = "IIS-IIS6ManagementCompatibility"
	IIS_Feature(52) = "IIS-HttpTracing"
	IIS_Feature(53) = "IIS-CustomLogging"
	IIS_Feature(54) = "IIS-HttpCompressionDynamic"
		
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Uninstall All option: " & strCmd)
	
	'Perform feature disable for all required feature
	strCmd = "cmd /c DISM.exe /Online /Disable-Feature"
	For i = UBound(IIS_Feature) To 48 Step -1
		If IIS_Feature(i) <> "" Then
			strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
		End If
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all main features: ")	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 5000
	
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			RebootPending = True
		End If
	Else
		MSIError = 1603
	End If
	
	If RemoveDependencies Then	
		'Perform feature disable for all dependencies
		strCmd = "cmd /c DISM.exe /Online /Disable-Feature"
		For i = 47 To 0 Step -1
			If IIS_Feature(i) <> "" Then
				strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
			End If
		Next
		strCmd = strCmd & " /NoRestart"
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all dependency features: ")
		Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
		WScript.Sleep 5000		
		If rc = 0 Or rc = 3010 Then
			If rc = 3010 Then
				RebootPending = True
			End If
		Else
			MSIError = 1603
		End If
	End If
	
	If MSIError = 0 Then
		If RebootPending Then
			MSIError = 3010
		End If
	End If				

	'Get All Feature Status	
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\AllFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get All feature status: ")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then

		'Display All Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display All feature status: " & vbCrLf )
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If
		
End Function

Function Uninstall_FTP

	On Error Resume Next
	
	Dim IIS_Feature(43), IIS_Feature_Status(43)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, strStatus
	Dim RebootPending, status
		
	rc = 0
	MSIError = 0
	RebootPending = False

	IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-DirectoryBrowsing"
	IIS_Feature(6)  = "IIS-HttpErrors"
	IIS_Feature(7)  = "IIS-HttpRedirect"
	IIS_Feature(8)  = "IIS-ASPNET"
	IIS_Feature(9)  = "IIS-NetFxExtensibility"
	IIS_Feature(10) = "NetFx4Extended-ASPNET45"
	IIS_Feature(11) = "IIS-NetFxExtensibility45"	
	IIS_Feature(12) = "IIS-HealthAndDiagnostics"
	IIS_Feature(13) = "IIS-HttpLogging"	
	IIS_Feature(14) = "IIS-RequestMonitor"
	IIS_Feature(15) = "IIS-Security"		
	IIS_Feature(16) = "IIS-BasicAuthentication"	
	IIS_Feature(17) = "IIS-ClientCertificateMappingAuthentication"
	IIS_Feature(18) = "IIS-DigestAuthentication"
	IIS_Feature(19) = "IIS-IISCertificateMappingAuthentication"		
	IIS_Feature(20) = "IIS-IPSecurity"
	IIS_Feature(21) = "IIS-RequestFiltering"
	IIS_Feature(22) = "IIS-URLAuthorization"	
	IIS_Feature(23) = "IIS-WindowsAuthentication"	
	IIS_Feature(24) = "IIS-Performance"	
	IIS_Feature(25) = "IIS-HttpCompressionStatic"	
	IIS_Feature(26) = "IIS-ManagementConsole"	
	IIS_Feature(27) = "IIS-ManagementService"	
	IIS_Feature(28) = "IIS-ManagementScriptingTools"	
	IIS_Feature(29) = "IIS-LoggingLibraries"
	IIS_Feature(30) = "IIS-WebServerManagementTools"
	IIS_Feature(31) = "IIS-CertProvider"
	IIS_Feature(32) = "WAS-WindowsActivationService"
	IIS_Feature(33) = "WAS-ProcessModel"
	IIS_Feature(34) = "WAS-NetFxEnvironment"
	IIS_Feature(35) = "WAS-ConfigurationAPI"
	IIS_Feature(36) = "IIS-ApplicationDevelopment"
	IIS_Feature(37) = "IIS-ISAPIExtensions"
	IIS_Feature(38) = "IIS-ISAPIFilter"
	IIS_Feature(39) = "WCF-HTTP-Activation"
	IIS_Feature(40) = "NetFx3"
	'Do not Disable NetFx3 on uninstallation
	IIS_Feature(40) = ""	

	'Main Feature for FTP Server
	IIS_Feature(41) = "IIS-FTPServer"
	IIS_Feature(42) = "IIS-FTPSvc"
	IIS_Feature(43) = "IIS-FTPExtensibility"	

	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Uninstall FTP option: " & strCmd)

	'Perform feature disable for all required feature
	strCmd = "cmd /c DISM.exe /Online /Disable-Feature"
	For i = UBound(IIS_Feature) To 41 Step -1
		If IIS_Feature(i) <> "" Then
			strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
		End If
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all main features: ")	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 2000
	
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			RebootPending = True
		End If
	Else
		MSIError = 1603
	End If
	
	If RemoveDependencies Then	
		'Perform feature disable for all dependencies
		strCmd = "cmd /c DISM.exe /Online /Disable-Feature"
		For i = 40 To 0 Step -1
			If IIS_Feature(i) <> "" Then
				strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
			End If		
		Next
		strCmd = strCmd & " /NoRestart"
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all dependency features: ")
		Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
		WScript.Sleep 2000		
		If rc = 0 Or rc = 3010 Then
			If rc = 3010 Then
				RebootPending = True
			End If
		Else
			MSIError = 1603
		End If
	End If
	
	If MSIError = 0 Then
		If RebootPending Then
			MSIError = 3010
		End If
	End If	
	
	'Get FTP Feature Status	
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\FTPFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get FTP feature status:")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then

		'Display FTP Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display FTP feature status:")
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If
		
End Function

Function Uninstall_SMTP

	On Error Resume Next
	
	Dim IIS_Feature(47), IIS_Feature_Status(47)
	Dim i, rc
	Dim strCmd
	Dim FeatureName
	Dim ReadLine, tmpFile, getFile, targetFile
	Dim CurFeature, statusLine, strStatus
	Dim RebootPending, status		
	
	rc = 0
	MSIError = 0
	RebootPending = False
	
	IIS_Feature(0)  = "IIS-WebServerRole"
	IIS_Feature(1)  = "IIS-WebServer"
	IIS_Feature(2)  = "IIS-CommonHttpFeatures"
	IIS_Feature(3)  = "IIS-StaticContent"
	IIS_Feature(4)  = "IIS-DefaultDocument"
	IIS_Feature(5)  = "IIS-DirectoryBrowsing"
	IIS_Feature(6)  = "IIS-HttpErrors"
	IIS_Feature(7)  = "IIS-HttpRedirect"
	IIS_Feature(8)  = "IIS-ASPNET"
	IIS_Feature(9)  = "IIS-NetFxExtensibility"
	IIS_Feature(10) = "NetFx4Extended-ASPNET45"
	IIS_Feature(11) = "IIS-NetFxExtensibility45"	
	IIS_Feature(12) = "IIS-HealthAndDiagnostics"
	IIS_Feature(13) = "IIS-HttpLogging"	
	IIS_Feature(14) = "IIS-RequestMonitor"
	IIS_Feature(15) = "IIS-Security"		
	IIS_Feature(16) = "IIS-BasicAuthentication"	
	IIS_Feature(17) = "IIS-ClientCertificateMappingAuthentication"
	IIS_Feature(18) = "IIS-DigestAuthentication"
	IIS_Feature(19) = "IIS-IISCertificateMappingAuthentication"		
	IIS_Feature(20) = "IIS-IPSecurity"
	IIS_Feature(21) = "IIS-RequestFiltering"
	IIS_Feature(22) = "IIS-URLAuthorization"	
	IIS_Feature(23) = "IIS-WindowsAuthentication"	
	IIS_Feature(24) = "IIS-Performance"	
	IIS_Feature(25) = "IIS-HttpCompressionStatic"	
	IIS_Feature(26) = "IIS-ManagementConsole"	
	IIS_Feature(27) = "IIS-ManagementService"	
	IIS_Feature(28) = "IIS-ManagementScriptingTools"	
	IIS_Feature(29) = "IIS-LoggingLibraries"
	IIS_Feature(30) = "IIS-WebServerManagementTools"
	IIS_Feature(31) = "IIS-CertProvider"
	IIS_Feature(32) = "WAS-WindowsActivationService"
	IIS_Feature(33) = "WAS-ProcessModel"
	IIS_Feature(34) = "WAS-NetFxEnvironment"
	IIS_Feature(35) = "WAS-ConfigurationAPI"
	IIS_Feature(36) = "IIS-ApplicationDevelopment"
	IIS_Feature(37) = "IIS-ISAPIExtensions"
	IIS_Feature(38) = "IIS-ISAPIFilter"
	IIS_Feature(39) = "WCF-HTTP-Activation"
	IIS_Feature(40) = "NetFx3"
	'Do not uninstall NetFx3 on unistallation
	IIS_Feature(40) = ""
	IIS_Feature(41) = "IIS-LegacySnapIn"	
	IIS_Feature(42) = "IIS-Metabase"
	IIS_Feature(43) = "IIS-IIS6ManagementCompatibility"
	
	'SMTP main features
	IIS_Feature(44) = "ServerManager-Core-RSAT"	
	IIS_Feature(45) = "ServerManager-Core-RSAT-Feature-Tools"	
	IIS_Feature(46) = "Smtpsvc-Service-Update-Name"	
	IIS_Feature(47) = "Smtpsvc-Admin-Update-Name"
	
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform Uninstall SMTP option: ")

	'Perform feature disable for all required feature
	strCmd = "cmd /c DISM.exe /Online /Disable-feature"
	For i = UBound(IIS_Feature) To 44 Step -1
		If IIS_Feature(i) <> "" Then
			strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
		End If		
	Next
	strCmd = strCmd & " /NoRestart"	
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all main features: ")	
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	WScript.Sleep 5000
	
	If rc = 0 Or rc = 3010 Then
		If rc = 3010 Then
			RebootPending = True
		End If
	Else
		MSIError = 1603
	End If
	
	If RemoveDependencies Then	
		'Perform feature disable for all dependencies
		strCmd = "cmd /c DISM.exe /Online /Disable-Feature"
		For i = 43 To 0 Step -1
			If IIS_Feature(i) <> "" Then
				strCmd = strCmd & " /Featurename:" & IIS_Feature(i)
			End If		
		Next
		strCmd = strCmd & " /NoRestart"
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Perform disable feature for all dependency features: ")
		Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
		rc = WSHShell.Run(strCmd, 0, True)
		Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
		WScript.Sleep 5000		
		If rc = 0 Or rc = 3010 Then
			If rc = 3010 Then
				RebootPending = True
			End If
		Else
			MSIError = 1603
		End If
	End If
	
	If MSIError = 0 Then
		If RebootPending Then
			MSIError = 3010
		End If
	End If
	
	'Get SMTP Feature Status
	CreateFolder("C:\Temp")
	getFile = "C:\Temp\SMTPFeatureStatus.txt"
	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If
	Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Get SMTP feature status:")
	strCmd = "cmd /c DISM.exe /Online /Get-Features /Format:Table > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Err.Clear
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then
		
		'Display SMTP Feature Status
		Log.WriteLine(vbCrLf & GetTimeStamp() & vbTab & "Display SMTP feature status: " & vbCrLf )
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				
				'Get all main feature status
				For i = 0 To UBound(IIS_Feature)
					FeatureName = IIS_Feature(i)
					If UCase(FeatureName) = UCase(CurFeature) Then
						strStatus = Trim(statusLine(1))
						Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )
						IIS_Feature_Status(i) = strStatus
					End If
				Next
			End If
		Loop
	
	End If							
	
End Function

Function GetFeatureStatus(FeatureName)

	On Error Resume Next

	Dim ReadLine, strFound, tmpFile, getFile, targetFile
	Dim strCmd, rc
	Dim statusLine, strStatus
	Dim CurFeature

	strFound = False
	strStatus = ""
	getFile = "C:\Temp\GetFeatureStatus.txt"

	If oFSO.FileExists(getFile) Then
		oFSO.DeleteFile(getFile)
	End If

	strCmd = "cmd /c DISM.exe /online /Get-Features /Format:Table  > " & getFile
	Log.WriteLine(GetTimeStamp() & vbTab & "Get " & FeatureName & " status: ")
	Log.WriteLine(GetTimeStamp() & vbTab & "Execute command: " & strCmd )
	rc = WSHShell.Run(strCmd, 0, True)
	Log.WriteLine(GetTimeStamp() & vbTab & "Return code = " & rc )
	
	Set tmpFile = oFSO.GetFile(getFile)
	Set targetFile = oFSO.OpenTextFile(tmpFile, ForReading, True)
		
	If Err = 0 Then
		Do Until targetFile.AtEndOfStream
			ReadLine = targetFile.ReadLine
			ReadLine = Trim(ReadLine)
			If ReadLine <> "" Then
				statusLine = Split(ReadLine, "|")
				CurFeature = Trim(statusLine(0))
				If UCase(FeatureName) = UCase(CurFeature) Then
					StrFound = True
					strStatus = Trim(statusLine(1))
					Log.WriteLine(GetTimeStamp() & vbTab & FeatureName & " status = " & strStatus )

					Exit Do
				End If
			End If
		Loop
	
		'Close file 
		targetFile.Close	
	End If
	
	GetFeatureStatus = StrStatus

End Function

Sub CustomInstall

	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Install...")
	
	Select Case Switch2
		Case "ALL"
			Call Install_ALL
		Case "FTP"
			Call Install_FTP
		Case "SMTP"
			Call Install_SMTP
		Case Else
			'invalid parameter
			MSIError = -102
			Log.WriteLine(vbCrLf & "Option is not applicable or not provided: " & InstOption)

	End Select	

	'this is where you would put any custom actions from a previous installation script, common to Prist and GUIInst
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Exiting Custom Install.")
End Sub

Sub CustomUninstall

	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing CustomUninstall...")
	
	Select Case Switch2
		Case "ALL"
			Call Uninstall_All
		Case "ALL_AND_DEP"
			RemoveDependencies = True
			Call Uninstall_All
		Case "FTP"
			Call Uninstall_FTP
		Case "FTP_AND_DEP"
			RemoveDependencies = True
			Call Uninstall_FTP			
		Case "SMTP"
			Call Uninstall_SMTP
		Case "SMTP_AND_DEP"
			RemoveDependencies = True
			Call Uninstall_SMTP			
		Case Else
			'invalid parameter
			MSIError = -101
			Log.WriteLine(vbCrLf & "Option is not applicable or not provided: " & InstOption)

	End Select	

	'this is where you would put any custom actions from a previous installation script, common to Prist and GUIInst
	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Exiting CustomUninstall.")
	
End Sub

'Sub CustomUninstall
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Uninstall...")
'	this is where you would put any custom actions from a previous uninstallation script, common to Uninst and GUIUninst
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom Uninstall.")
'End Sub

'Sub Major_CustomUninstall
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Major_Custom Uninstall...")
'	this is where you would put any custom actions from the previous Major script
'	Log.WriteLine(GetTimeStamp() & "  Exiting Major_Custom Uninstall.")
'End Sub

'Sub CustomUpdate
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Update...")
'	this is where you would put any custom actions from a previous update script
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom Update.")
'End Sub

'Sub CustomPatch
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Patch...")
'	this is where you would put any custom actions from a previous update script
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom Patch.")
'End Sub

'Sub CustomPrist
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Prist...")
'	this is where you would put any custom actions specific to Prist 
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom Prist.")
'End Sub

'Sub CustomUninst
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom Uninst...")
'	this is where you would put any custom actions specific to Uninst
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom Uninst.")
'End Sub

'Sub CustomGUIInst
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom GUIInst...")
'	this is where you would put any custom actions specific to GUIInst
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom GUIInst.")
'End Sub

'Sub CustomGUIUninst
'	Log.WriteLine(vbCrLf & GetTimeStamp() & "  Executing Custom GUIUninst...")
'	this is where you would put any custom actions specific to GUIUninst
'	Log.WriteLine(GetTimeStamp() & "  Exiting Custom GUIUninst.")
'End Sub
