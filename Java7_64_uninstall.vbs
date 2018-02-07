'  FILENAME: Java6U7uninstall.vbs
'  AUTHOR: Mihhail Shapovalov
'  SYNOPSIS: This script looks for Java6U7 and removes it
'  DESCRIPTION: Searches add remove programs for J6U7 removes it
'  NOTES: - Must edit strCurrentVersion to match the version you want to uninstall
'   - if called with a computer name will, run against remote machine
'   - logs to local path defined in strLogPath
'   - requires admin priv
'  EXAMPLE: Java6u7uninstall.vbs
'  EXAMPLE: Java6u7uninstall.vbs \\workststion
'  INPUTS: \\workststion (optional)
'  RETURNVALUE: logs to value in strLogPath
'  ChangeLog:
'   2017-11-9: mihhail
 
'On Error Resume Next
Option Explicit
DIM objFSO, strComputer, objWMIService, colInstalledVersions, oShell, strVersion
DIM objVersion, strLogPath, strLogName, strPath, strExecQuery
 
IF WScript.Arguments.Count > 0 then
    strComputer = replace(WScript.Arguments(0),"\\","")
ELSE
    strComputer = "."
END If
 
Dim stCurrentVersion : stCurrentVersion = Array("{26A24AE4-039D-4CA4-87B4-2F06417080FF}","{26A24AE4-039D-4CA4-87B4-2F03217080FF}","{26A24AE4-039D-4CA4-87B4-2F06417079FF}","{26A24AE4-039D-4CA4-87B4-2F03217079FF}","{26A24AE4-039D-4CA4-87B4-2F06417076FF}","{26A24AE4-039D-4CA4-87B4-2F03217076FF}","{26A24AE4-039D-4CA4-87B4-2F06417075FF}","{26A24AE4-039D-4CA4-87B4-2F03217075FF}","{26A24AE4-039D-4CA4-87B4-2F06417072FF}","{26A24AE4-039D-4CA4-87B4-2F03217072FF}","{26A24AE4-039D-4CA4-87B4-2F06417071FF}","{26A24AE4-039D-4CA4-87B4-2F03217071FF}","{26A24AE4-039D-4CA4-87B4-2F06417067FF}","{26A24AE4-039D-4CA4-87B4-2F03217067FF}","{26A24AE4-039D-4CA4-87B4-2F06417060FF}","{26A24AE4-039D-4CA4-87B4-2F03217060FF}","{26A24AE4-039D-4CA4-87B4-2F86417055FF}","{26A24AE4-039D-4CA4-87B4-2F83217055FF}","{26A24AE4-039D-4CA4-87B4-2F83217051FF}","{26A24AE4-039D-4CA4-87B4-2F83217051FF}","{26A24AE4-039D-4CA4-87B4-2F86417045FF}","{26A24AE4-039D-4CA4-87B4-2F83217045FF}","{26A24AE4-039D-4CA4-87B4-2F86417040FF}","{26A24AE4-039D-4CA4-87B4-2F83217040FF}","{26A24AE4-039D-4CA4-87B4-2F86417025FF}","{26A24AE4-039D-4CA4-87B4-2F83217025FF}","{26A24AE4-039D-4CA4-87B4-2F86417021FF}","{26A24AE4-039D-4CA4-87B4-2F83217021FF}","{26A24AE4-039D-4CA4-87B4-2F86417017FF}","{26A24AE4-039D-4CA4-87B4-2F83217017FF}","{26A24AE4-039D-4CA4-87B4-2F86417016FF}","{26A24AE4-039D-4CA4-87B4-2F83217016FF}","{26A24AE4-039D-4CA4-87B4-2F86417015FF}","{26A24AE4-039D-4CA4-87B4-2F83217015FF}","{26A24AE4-039D-4CA4-87B4-2F86417014FF}","{26A24AE4-039D-4CA4-87B4-2F83217014FF}","{26A24AE4-039D-4CA4-87B4-2F86417013FF}","{26A24AE4-039D-4CA4-87B4-2F83217013FF}","{26A24AE4-039D-4CA4-87B4-2F86417012FF}","{26A24AE4-039D-4CA4-87B4-2F83217012FF}","{26A24AE4-039D-4CA4-87B4-2F86417011FF}","{26A24AE4-039D-4CA4-87B4-2F83217011FF}","{26A24AE4-039D-4CA4-87B4-2F86417010FF}","{26A24AE4-039D-4CA4-87B4-2F83217010FF}","{26A24AE4-039D-4CA4-87B4-2F86417009FF}","{26A24AE4-039D-4CA4-87B4-2F83217009FF}","{26A24AE4-039D-4CA4-87B4-2F86417008FF}","{26A24AE4-039D-4CA4-87B4-2F83217008FF}","{26A24AE4-039D-4CA4-87B4-2F86417007FF}","{26A24AE4-039D-4CA4-87B4-2F83217007FF}","{26A24AE4-039D-4CA4-87B4-2F86417006FF}","{26A24AE4-039D-4CA4-87B4-2F83217006FF}","{26A24AE4-039D-4CA4-87B4-2F86417005FF}","{26A24AE4-039D-4CA4-87B4-2F83217005FF}","{26A24AE4-039D-4CA4-87B4-2F86417004FF}","{26A24AE4-039D-4CA4-87B4-2F83217004FF}","{26A24AE4-039D-4CA4-87B4-2F86417003FF}","{26A24AE4-039D-4CA4-87B4-2F83217003FF}","{26A24AE4-039D-4CA4-87B4-2F86417002FF}","{26A24AE4-039D-4CA4-87B4-2F83217002FF}","{26A24AE4-039D-4CA4-87B4-2F86417001FF}","{26A24AE4-039D-4CA4-87B4-2F83217001FF}")

strExecQuery = "Select * from Win32_Product Where Name LIKE '%Java 2 Runtime Environment%' OR Name LIKE '%J2SE Runtime Environment%' OR Name LIKE '%Java(TM)%' OR Name Like '%Java 7%'"
KillProc
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colInstalledVersions = objWMIService.ExecQuery (strExecQuery)
Set oShell = CreateObject("WScript.Shell")
strPath = oShell.ExpandEnvironmentStrings("%USERPROFILE%")
strLogPath = strPath & "\LOGS\"

strLogName = "Java_Uninstall.log"
 
LogIt String(120, "_")
LogIt String(120, "¯")

For Each objVersion in colInstalledVersions
	For Each strVersion In stCurrentVersion
		If objVersion.IdentifyingNumber = strVersion Then
			msgbox "We found vulnerable Java Version 7 installed. We will remove it now." ,0, "Symantec IT"
			LogIt Now() & ": " &replace(strComputer,".","localhost") & ": Uninstalling: Java version" & objVersion.IdentifyingNumber
	    	objVersion.Uninstall()
	    	MsgBox "Java Version 7 uninstall is completed" ,0, "Symantec IT"
		end If
	Next
Next
LogIt String(120, "_")
LogIt String(120, "¯")
LogIt String(120, " ")
 
Sub LogIt (strLineToWrite)
    DIM ts
    If Not objFSO.FolderExists(strLogPath) Then MakeDir(strLogPath)
    Set ts = objFSO.OpenTextFile(strLogPath & strLogName, 8, True)
    ts.WriteLine strLineToWrite
    ts.close
End Sub
 
Function MakeDir (strPath)
    Dim strParentPath
    On Error Resume Next
    strParentPath = objFSO.GetParentFolderName(strPath)
 
  If Not objFSO.FolderExists(strParentPath) Then MakeDir strParentPath
    If Not objFSO.FolderExists(strPath) Then objFSO.CreateFolder strPath
    On Error Goto 0
  MakeDir = objFSO.FolderExists(strPath)
End Function
 
Sub KillProc()
   '# kills jusched.exe and jqs.exe if they are running.  These processes will cause the installer to fail.
   Dim wshShell
   Set wshShell = CreateObject("WScript.Shell")
   wshShell.Run "Taskkill /F /IM jusched.exe /T", 0, True
   wshShell.Run "Taskkill /F /IM jqs.exe /T", 0, True
End Sub