Option Explicit

Dim oReg, objWMIService, colOperatingSystems, objOperatingSystem, strAltirisInstallKey, strAltirisServerKey
Const HKEY_LOCAL_MACHINE = &H80000002
Const strKeyPath = "SOFTWARE\Symantec\Symantec Endpoint Protection\SMC\SSHelper"
Const strComputer = "."
Const strResultName = "AltirisUpdateResult"
Const X64BIT = "x64"
Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & strComputer & "\root\default:StdRegProv")

'Get OS type for x86 or x64 information to select the right registry keys
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

Set colOperatingSystems = objWMIService.ExecQuery _
    ("Select * from Win32_ComputerSystem")
    
For Each objOperatingSystem in colOperatingSystems
  If (InStr(objOperatingSystem.SystemType, X64BIT) <> 0) Then      'x64
          strAltirisInstallKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Altiris\Altiris Agent\InstallDir"
          strAltirisServerKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Altiris\Altiris Agent\Servers\"            
  Else    'x86
          strAltirisInstallKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Altiris\Altiris Agent\InstallDir"
          strAltirisServerKey = "HKEY_LOCAL_MACHINE\SOFTWARE\Altiris\Altiris Agent\Servers\"    
  End If
Next

If isAltirisUpdated() = True Then
          oReg.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strResultName, "1"
		  wscript.echo "Test OK"
Else
          oReg.SetStringValue HKEY_LOCAL_MACHINE, strKeyPath, strResultName, "0"

End If

Function isAltirisUpdated()
	Dim objWSHShell
	Dim strKey
	Dim strTimeReg
	Dim strTimeFile
	
	On Error GoTo 0
    On Error Resume Next
    
	Set objWSHShell = WScript.CreateObject("WScript.Shell")

	
	'Get Current NS server
	Dim strVal_NSS
	strKey = strAltirisServerKey
	strVal_NSS = objWSHShell.RegRead(strKey)
	If Err.Number <> 0 Or strVal_NSS = "" Then 
		isAltirisUpdated = True
		On Error GoTo 0
		Exit Function
	End If 

    Dim strVAL_LP 'Basic Inventory Last Post
    Dim strVal_LR 'Policy Last Requested
          strVAL_LP = objWSHShell.RegRead(strKey & strVal_NSS & "\Basic Inventory Last Post")
          strVal_LR = objWSHShell.RegRead(strKey & strVal_NSS & "\Policy Last Requested")
          If Err.Number <> 0 Then 
                   isAlitirsUpdated = True
                   On Error GoTo 0
		           Exit Function
          End If 
          
          if strVAL_LP > strVal_LR Then
              strTimeReg = strVAL_LP
    Else
              strTimeReg = strVal_LR
          End If
          
    'Read file time
          'Key of Altiris Agent Installed
          Dim strVal_AAI
          strKey = strAltirisInstallKey
          strVal_AAI = objWSHShell.RegRead(strKey)
          If Err.Number <> 0 Or strVal_AAI = "" Then 
                   isAlitirsUpdated = True
                   On Error GoTo 0
		           Exit Function
          End If 

    Dim objFSO
    Dim objFDir
    Dim objFc
    Dim objF
    
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFDir = objFSO.GetFolder( strVal_AAI & "\Queue\" & strVal_NSS )
    Set objFC = objFDir.Files
          If Err.Number <> 0 Then 
                   isAltirisUpdated = True
                   On Error GoTo 0
		           Exit Function
          End If 

    Dim strFile
    Dim stDateTime
    Dim strCurFileTime
    Dim sendTime

    For Each objF in objFC
        strFile = objF.Name

        '2007-09-25 13:34:01 -420"
        stDateTime = objF.DateLastModified
              If Err.Number <> 0 Then 
                       isAlitirsUpdated = True
                       On Error GoTo 0
		               Exit Function
             End If 
 
        strCurFileTime = year(stDateTime) & "-" & right("00" & month(stDateTime),2) & "-" & right( "00" & day(stDateTime),2) & " " & right("00" & hour(stDateTime),2) &":"& right("00" & minute(stDateTime),2) &":"& right("00" & second(stDateTime),2)   
   
        If strTimeFile = "" Then
            strTimeFile = strCurFileTime
        Elseif strCurFileTime < strTimeFile Then
            strTimeFile = strCurFileTime
        End if
        
    Next


    If (strTimeFile<>"") And (strTimeFile < strTimeReg) Then 
          isAltirisUpdated = False
	Else
          isAltirisUpdated = True
	End IF

End Function

