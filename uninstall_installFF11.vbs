'* ======================================================================================'*
'* Script Name		: uninstall_install_ff.vbs
'* Purpose  		: Uninstall Firefox 11 and installs latest version
'* Notes			: Compliance 
'* Usage			: N/A
'* Modification Log	: N/A
'* 	Date			Author		  		 Version		Change
'* 28/07/17  		Mihhail Shapovalov   1.0			Initial Release
'* 08/09/17			Mihhail Shapovalov 	 1.1			Adding more uninstall options, when helper.exe doesnt exist
'* ======================================================================================'*

On Error Resume Next

' Global Variables

Dim oShell, objFSO, oExec
Dim fldr32, fldr64, strProfile, strProfilelocal, strNewFF, Uninstallfolder, InstallString, Args, strVersion
strComputer = "." 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

fldr64 = "C:\Program Files (x86)\Mozilla Firefox"
fldr32 = "C:\Program Files\Mozilla Firefox"

appDataLocation = oShell.ExpandEnvironmentStrings("%APPDATA%")
localappDataLocation = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

strProfile = appDataLocation & "\Mozilla"
strProfilelocal = localappDataLocation & "\Mozilla"

strNewFF = "\Firefox Setup 52.2.1esr.exe"
Uninstallfolder = "\uninstall\helper.exe"
Args = " -ms"
strCurDir = oShell.CurrentDirectory

' Function to delete folders
Function DeleteFolder(strFolderPath)
	Dim objFSO
	Set objFSO = CreateObject ("Scripting.FileSystemObject")
	If objFSO.FolderExists(strFolderPath) Then
	  	objFSO.DeleteFolder(strFolderPath), True 
	End If
End Function 

' Function to force close FF
Function CloseFF
Set oExec = oShell.Exec("taskkill /f /fi ""imagename eq firefox.exe""")
Do While oExec.Status = 0
     WScript.Sleep 100
Loop
End Function 

' Function to uninstall old FF
Function UninstallFF (fldr)
	result1 = Msgbox ("We have found a vulnerable Firefox v11, which need to be uninstalled. Would you like to uninstall Firefox now?", vbYesNo, "Symantec IT")
	If result1 = vbNo Then
		Msgbox "Thank you. We will try again later.",  vbYes, "Symantec IT"
		Wscript.Quit(1)
	Else
		Msgbox "Thank you. We will proceed with Firefox uninstall now.",  vbYes, "Symantec IT"
		CloseFF
		WScript.Sleep 100
		itemGUID = "{10135883-A696-4A7F-BFDB-04AC20013E88}"
		Return = oShell.Run("msiexec.exe /x" & itemGUID & " /qb-!", 1, true)
		If return <> 0 or return <> 3010 Then
			UninstallString = Chr(34) & fldr &  Uninstallfolder & Chr(34) & Args
			oShell.Run UninstallString, 1, True
		End If
	End If
End Function 
	
' Function to Install latest FF version
Function InstallFF		
result = Msgbox ("We have uninstalled outdated version of Firefox 11. Would you like to install latest version of Firefox now?", vbYesNo, "Symantec IT")
	If result = vbNo Then 
		MsgBox "Thank you. If you still need Firefox, you can find it on Software Download page", vbYes, "Symantec IT"
		objLog.WriteLine "User cancelled installation of new FF"
	Else
		MsgBox "Firefox version 52.2.1 will be installed now", vbYes, "Symantec IT"
		InstallString = Chr(34) & strCurDir & strNewFF & Chr(34) & Args
		intReturn = oShell.Run (InstallString, 1, True)
		If intReturn <> 0 Then 
   			MsgBox "Error running program"
   			objLog.WriteLine "Error running program"
		Else
			MsgBox "Firefox version 52.2.1 installation completed", vbYes, "Symantec IT"
			objLog.WriteLine "Firefox version 52.2.1 installation completed"
		End If
	End If
End Function 

' Function for 64-bit

Function DetectAndRun
strVersion64 = objFSO.GetFileVersion("C:\Program Files (x86)\Mozilla Firefox\firefox.exe")
strVersion = CInt(Left(strVersion64, 2))
If (strVersion <> 0) and (strVersion < 12) Then
	objLog.WriteLine "FF v11 has been found"
	UninstallFF (fldr64)
	objLog.WriteLine "Uninstalled FF"
	WScript.Sleep 1000
	DeleteFolder fldr64
	objLog.WriteLine "REMOVED FF Folder"
	WScript.Sleep 1000
	DeleteFolder (strProfile)
	objLog.WriteLine "REMOVED APPDATA/ROAMING Folder"
	WScript.Sleep 1000
	DeleteFolder (strProfilelocal)
	objLog.WriteLine "REMOVED APPDATA/LOCAL Folder"
	'WScript.Sleep 100
	InstallFF
	Set strVersion = Nothing
End If
End Function 

' Function for 32-bit
Function DetectAndRun32
strVersion32 = objFSO.GetFileVersion("C:\Program Files\Mozilla Firefox\firefox.exe")
strVersion = CInt(Left(strVersion64, 2))
If (strVersion <> 0) and (strVersion < 12) Then
	objLog.WriteLine "FF v11 has been found"
	UninstallFF (fldr32)
	objLog.WriteLine "Uninstalled FF"
	WScript.Sleep 1000
	DeleteFolder (fldr32)
	objLog.WriteLine "REMOVED FF Folder"
	DeleteFolder (strProfile)
	objLog.WriteLine "REMOVED APPDATA/ROAMING Folder"
	DeleteFolder (strProfilelocal)
	objLog.WriteLine "REMOVED APPDATA/LOCAL Folder"
	'WScript.Sleep 100
	InstallFF
	Set strVersion = Nothing
End If
End Function 


' Main method
' check if FF 32 or 64
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile("c:\temp\ff11removelog.txt")

If (objFSO.FolderExists(fldr64)) Then
	objLog.WriteLine "64-bit OS"
	If (objFSO.FileExists("C:\Program Files (x86)\Mozilla Firefox\uninstall\helper.exe")) Then
		DetectAndRun
  	Else  
		objFSO.CreateFolder "C:\Program Files (x86)\Mozilla Firefox\uninstall"
		objFSO.CopyFile "helper.exe", "C:\Program Files (x86)\Mozilla Firefox\uninstall\"
		DetectAndRun
	End If
ElseIf (objFSO.FolderExists(fldr32)) Then
	objLog.WriteLine "32-bit OS"
	If (objFSO.FileExists("C:\Program Files\Mozilla Firefox\uninstall\helper.exe")) Then
		DetectAndRun32
	Else
		objFSO.CreateFolder "C:\Program Files\Mozilla Firefox\uninstall"
		objFSO.CopyFile "helper.exe", "C:\Program Files\Mozilla Firefox\uninstall"
		DetectAndRun32
	End If
Else 
objLog.WriteLine "FF 11 was not found"
End If