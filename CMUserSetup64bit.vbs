On Error Resume Next
Dim fso 'As New Scripting.FileSystemObject
Dim WshShell
Set fso = CreateObject("Scripting.FileSystemObject")
Set WshShell = CreateObject("WScript.Shell")

'set dynamic sourcepath
Dim sourcepath
Dim objFile 
Set objFile = fso.GetFile(Wscript.ScriptFullName) 


'set environment variable paths for Windows 
Dim alluserspath
alluserspath = WshShell.ExpandEnvironmentStrings("%ALLUSERSPROFILE%")

Dim appdata
appdata = WshShell.ExpandEnvironmentStrings("%APPDATA%")

Dim userprofile
userprofile = WshShell.ExpandEnvironmentStrings("%USERPROFILE%")

Dim programdata
programdata = WshShell.ExpandEnvironmentStrings("%PROGRAMDATA%")

Dim publicprofile
publicprofile = WshShell.ExpandEnvironmentStrings("%PUBLIC%")

Dim programfiles
programfiles = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

Dim installroot
installroot = WshShell.ExpandEnvironmentStrings("%PROGRAMFILES%")

Dim localappdata
localappdata = WshShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

'Check if Content Manager has been installed onto the computer, then run, else quit
If fso.FileExists(installroot & "\Micro Focus\Content Manager\trim.exe") = False Then
Wscript.Quit
End If

Dim objWMIService
Dim colItems
Dim objItem

'Content Manager User Settings Version Numbering
CONST CMSettingsVersion = "1.01"
Dim CMCurrentSettingsVersion

On Error Resume Next
CMCurrentSettingsVersion = WshShell.RegRead ("HKCU\Software\Micro Focus\Content Manager\CMSettingsVersion")
On Error Goto 0
If CMCurrentSettingsVersion = CMSettingsVersion Then
	'Script has already run, just exit
	WScript.quit
End If

'Add Content Manager SendTo Shortcut

If fso.FileExists(appdata & "\Microsoft\Windows\SendTo\HP Records Manager.lnk") = true Then
	fso.DeleteFile appdata & "\Microsoft\Windows\SendTo\HP Records Manager.lnk"
End If

If fso.FileExists(appdata & "\Microsoft\Windows\SendTo\HPRM Desktop.lnk") = true Then
	fso.DeleteFile appdata & "\Microsoft\Windows\SendTo\HPRM Desktop.lnk"
End If

If fso.FileExists(appdata & "\Microsoft\Windows\SendTo\Content Manager Desktop.lnk") = true Then
	fso.DeleteFile appdata & "\Microsoft\Windows\SendTo\Content Manager Desktop.lnk"
End If


If fso.FileExists(appdata & "\Microsoft\Windows\SendTo\Content Manager.lnk") = false Then
	fso.CopyFile programdata & "\Microsoft\Windows\Start Menu\Programs\Content Manager\Content Manager.lnk", appdata & "\Microsoft\Windows\SendTo\Content Manager.lnk"
End If

'Create HKEY_CURRENT_USER Dataset registry settings
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\AutoGetGlobal", 1, "REG_DWORD" 
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\AutomaticCheckIn\BypassCheckInOfLimboDocs", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\DefaultDB", "AM", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\LoadDefaultDB", "1", "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\AU\Name", "Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\AU\PrimaryURL", "https://client.aumprd.cm.kapish.cloud", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\AU\SecondaryURL", "", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\AU\AuthMechanism", "4", "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Datasets\AU\SupportedAuthMechanisms", "0;1;2;4;", "REG_SZ"

'Create HKEY_CURRENT_USER User Configuration registry settings
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\DataPaths\TopDrawerDataPath", localappdata & "\Micro Focus\Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\UserLexPath", appdata & "\Micro Focus\Content Manager\lex", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OpenDocuments\ConfirmWhenProcessOpen", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OpenDocuments\MinimumDelay", 15, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OpenDocuments\MinimumDelayDiscard", 15, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OpenDocuments\CheckinPollingInterval", 15, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OpenDocuments\MinimumCloseInterval", 15, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\OfficeAddins\UseNativeUI", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\IgnoreAllCapsWords", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\IgnoreMixedDigits", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\CaseSensitive", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\AutoCorrect", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\ReportDoubledWords", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\TopDrawer\AutoCleanupContainers", 1, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Micro Focus\Content Manager\TipOfTheDay\English\Start", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\Lex\MainLexPath", installroot & "\Micro Focus\Content Manager\lex", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\WhatsNew\Shown24.4", 1, "REG_DWORD"

'Force .tr5 to use new CM Path
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Classes\TRIM5.Record.Reference\Shell\Open\command\", chr(34) & "C:\Program Files\Micro Focus\Content Manager\trim.exe" & chr(34) & " " & chr(34) & "%1" & chr(34), "REG_SZ"

'Delete Auto Get Global registry key for Dataset to force Get Global on next start-up
WshShell.Run("REG.EXE DELETE ""HKEY_CURRENT_USER\Software\Micro Focus\Content Manager\DBID\AU"" /V AutoGetGlobal /F")

'Enable Office Integration
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\TRIMOfficeIntegrationOutlook\CommandLineSafe", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\TRIMOfficeIntegrationOutlook\Description", "The Add-in allows users of Microsoft Outlook to save important email messages into HPE Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\TRIMOfficeIntegrationOutlook\FriendlyName", "Content Manager Add-in for Microsoft Outlook", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\TRIMOfficeIntegrationOutlook\LoadBehavior", 3, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Outlook\Addins\TRIMOfficeIntegrationOutlook\Manifest", installroot & "\Micro Focus\Content Manager\TRIMOfficeIntegrationOutlook.vsto|vstolocal", "REG_SZ"

WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\AddIns\TRIMOfficeIntegrationWord\CommandLineSafe", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\AddIns\TRIMOfficeIntegrationWordDescription", "The Add-in allows users of Microsoft Word to save and open documents from HPE Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\AddIns\TRIMOfficeIntegrationWordFriendlyName", "Content Manager Add-in for Microsoft Word", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\AddIns\TRIMOfficeIntegrationWordLoadBehavior", 3, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Word\AddIns\TRIMOfficeIntegrationWordManifest", installroot & "\Micro Focus\Content Manager\TRIMOfficeIntegrationWord.vsto|vstolocal", "REG_SZ"

WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\TRIMOfficeIntegrationPowerPoint\CommandLineSafe", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\TRIMOfficeIntegrationPowerPoint\Description", "The Add-in allows users of Microsoft PowerPoint to save and open documents from HPE Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\TRIMOfficeIntegrationPowerPoint\FriendlyName", "Content Manager Add-in for Microsoft PowerPoint", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\TRIMOfficeIntegrationPowerPoint\LoadBehavior", 3, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\PowerPoint\Addins\TRIMOfficeIntegrationPowerPoint\Manifest", installroot & "\Micro Focus\Content Manager\TRIMOfficeIntegrationPowerPoint.vsto|vstolocal", "REG_SZ"

WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\TRIMOfficeIntegrationExcel\CommandLineSafe", 0, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\TRIMOfficeIntegrationExcel\Description", "The Add-in allows users of Microsoft Excel to save and open documents from HPE Content Manager", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\TRIMOfficeIntegrationExcel\FriendlyName", "Content Manager Add-in for Microsoft Excel", "REG_SZ"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\TRIMOfficeIntegrationExcel\LoadBehavior", 3, "REG_DWORD"
WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\TRIMOfficeIntegrationExcel\Manifest", installroot & "\Micro Focus\Content Manager\TRIMOfficeIntegrationExcel.vsto|vstolocal", "REG_SZ"


On Error Resume Next

officever = WshShell.RegRead("HKCU\Software\Microsoft\Office\16.0\")
	if err.number = 0 then
		WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList\TRIMOfficeIntegrationOutlook", 1, "REG_DWORD"
		WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\PowerPoint\Resiliency\DoNotDisableAddinList\TRIMOfficeIntegrationPowerPoint", 1, "REG_DWORD"
		WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList\TRIMOfficeIntegrationExcel", 1, "REG_DWORD"
		WshShell.RegWrite "HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Word\Resiliency\DoNotDisableAddinList\TRIMOfficeIntegrationWord", 1, "REG_DWORD"
		WshShell.RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\Security\DisableHyperlinkWarning", 1, "REG_DWORD"

	end if
err.Clear



'Create Custom Dictionary Set-up
If fso.FolderExists(appdata & "\Micro Focus") = False Then
	fso.CreateFolder appdata & "\Micro Focus"
End if

If fso.FolderExists(appdata & "\Micro Focus\Content Manager") = False Then
	fso.CreateFolder appdata & "\Micro Focus\Content Manager"
End if

If fso.FolderExists(appdata & "\Micro Focus\Content Manager\OfficeIntegration") = False Then
	fso.CreateFolder appdata & "\Micro Focus\Content Manager\OfficeIntegration"
End if

If fso.FolderExists(appdata & "\Micro Focus\Content Manager\Lex") = False Then
	fso.CreateFolder appdata & "\Micro Focus\Content Manager\Lex"
End if

If fso.FileExists(appdata & "\Micro Focus\Content Manager\Lex\UChange.tlx") = False Then
	fso.CopyFile installroot & "\Micro Focus\Content Manager\Lex\UChange.tlx", appdata & "\Micro Focus\Content Manager\Lex\UChange.tlx"
End If

If fso.FileExists(appdata & "\Micro Focus\Content Manager\Lex\UExclude.tlx") = False Then
	fso.CopyFile installroot & "\Micro Focus\Content Manager\Lex\UExclude.tlx", appdata & "\Micro Focus\Content Manager\Lex\UExclude.tlx"
End If

If fso.FileExists(appdata & "\Micro Focus\Content Manager\Lex\Userdic.tlx") = False Then
	fso.CopyFile installroot & "\Micro Focus\Content Manager\Lex\Userdic.tlx", appdata & "\Micro Focus\Content Manager\Lex\Userdic.tlx"
End If

If fso.FileExists(appdata & "\Micro Focus\Content Manager\Lex\USuggest.tlx") = False Then
	fso.CopyFile installroot & "\Micro Focus\Content Manager\Lex\USuggest.tlx", appdata & "\Micro Focus\Content Manager\Lex\USuggest.tlx"
End If

If fso.FileExists(appdata & "\Micro Focus\Content Manager\Lex\UIgnore.tlx") = False Then
	fso.CopyFile installroot & "\Micro Focus\Content Manager\Lex\UIgnore.tlx", appdata & "\Micro Focus\Content Manager\Lex\UIgnore.tlx"
End If


	fso.CopyFile installroot & "\Micro Focus\Content Manager\preferences", appdata & "\Micro Focus\Content Manager\OfficeIntegration\preferences", true


'Fix Preferences File
Dim objFilepreferences
Const ForReading = 1
Const ForWriting = 2

Set objFilepreferences = fso.OpenTextFile(appdata & "\Micro Focus\Content Manager\OfficeIntegration\preferences", ForReading)

strText = objFilepreferences.ReadAll
objFilepreferences.Close
strNewText = Replace(strText, "<MyDocumentsFolder />", "<MyDocumentsFolder>" & localappdata & "\Micro Focus\Content Manager</MyDocumentsFolder>")
Set objFilepreferences =  fso.OpenTextFile(appdata & "\Micro Focus\Content Manager\OfficeIntegration\preferences", ForWriting)
objFilepreferences.WriteLine strNewText
objFilepreferences.Close



'Write to the registry to say that this version of the script has run (run once per user per machine)

WshShell.RegWrite "HKCU\Software\Micro Focus\Content Manager\CMSettingsVersion", CMSettingsVersion,"REG_SZ"
