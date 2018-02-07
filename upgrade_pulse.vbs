'* ======================================================================================'*
'* Script Name		: upgrade_pulse.vbs
'* Purpose  		: Upgrade Pulse Secure
'* Notes			: Software Upgrade
'* Usage			: N/A
'* Modification Log	: N/A
'* Date				Author		  		 Version		Change
'* 10/01/18  		Mihhail Shapovalov   1.0			Initial Release
'* ======================================================================================'*

On Error Resume Next

' Global Variables

Dim oShell, objFSO, oExec
Dim fldr32, fldr64, strProfile, strProfilelocal, strNewFF, Uninstallfolder, InstallString, Args, strVersion
strComputer = "." 

Set objFSO = CreateObject("Scripting.FileSystemObject")
Set oShell = CreateObject("WScript.Shell")

fldr64 = "C:\Program Files (x86)\Common Files\Pulse Secure\JamUI"
fldr32 = "C:\Program Files\Common Files\Pulse Secure\JamUI"

appDataLocation = oShell.ExpandEnvironmentStrings("%APPDATA%")
localappDataLocation = oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%")

strProfile = appDataLocation & "\Mozilla"
strProfilelocal = localappDataLocation & "\Mozilla"

strNewFF = "\Firefox Setup 52.2.1esr.exe"
Uninstallfolder = "\uninstall\helper.exe"
Args = " -ms"
strCurDir = oShell.CurrentDirectory

' Function to force close Pulse
Function ClosePulse
Set oExec = oShell.Exec("taskkill /f /fi ""imagename eq pulse.exe""")
Do While oExec.Status = 0
     WScript.Sleep 100
Loop
End Function 

' Function to Install latest FF version
Function InstallPulse		
result = Msgbox ("We to upgrade your Pulse Secure VPN Client to latest version 5.3.1.1183. Your current connection via VPN will be closed.", vbYes, "Symantec IT")
	InstallString = Chr(34) & strCurDir & strNewFF & Chr(34) & Args
	intReturn = oShell.Run (InstallString, 1, True)
	If intReturn <> 0 Then 
   		MsgBox "Error running program"
   		objLog.WriteLine "Error running program"
	Else
		MsgBox "Firefox version 52.2.1 installation completed", vbYes, "Symantec IT"
		objLog.WriteLine "Firefox version 52.2.1 installation completed"
	End If
End Function 

' Function for 64-bit

Function DetectAndRun
strVersion64 = objFSO.GetFileVersion("C:\Program Files (x86)\Common Files\Pulse Secure\JamUI\Pulse.exe")
strVersion = CInt(Left(strVersion64, 2))
If (strVersion <> 0) and (strVersion < 5.3.1.1183) Then
	objLog.WriteLine "Pulse Secure version lower than 5.3.1.1183 has been found"
	InstallPulse
	Set strVersion = Nothing
End If
End Function 

' Function for 32-bit
Function DetectAndRun32
strVersion32 = objFSO.GetFileVersion("C:\Program Files\Common Files\Pulse Secure\JamUI\Pulse.exe")
strVersion = CInt(Left(strVersion64, 2))
If (strVersion <> 0) and (strVersion < 5.3.1.1183) Then
	objLog.WriteLine "Pulse Secure version lower than 5.3.1.1183 has been found"
	InstallPulse
	Set strVersion = Nothing
End If
End Function 


' Main method
' check if FF 32 or 64
Set ObjFSO = CreateObject("Scripting.FileSystemObject")
Set objLog = objFSO.CreateTextFile("c:\temp\pulsesecureupgradelog.txt")

If (objFSO.FolderExists(fldr64)) Then
	objLog.WriteLine "64-bit OS"
	If (objFSO.FileExists("C:\Program Files (x86)\Common Files\Pulse Secure\JamUI\Pulse.exe")) Then
		DetectAndRun
  	Else  
		objLog.WriteLine "No installation found"
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