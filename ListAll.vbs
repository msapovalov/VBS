'
' List all BIOS settings
'
On Error Resume Next
Dim colItems

strComputer = "LOCALHOST"     ' Change as needed.
strOptions
Set objWMIService = GetObject("WinMgmts:" _
    &"{ImpersonationLevel=Impersonate}!\\" & strComputer & "\root\wmi")
Set colItems = objWMIService.ExecQuery("Select * from Lenovo_BiosSetting")

For Each objItem in colItems
    If Len(objItem.CurrentSetting) > 0 Then
        Setting = ObjItem.CurrentSetting
        StrItem = Left(ObjItem.CurrentSetting, InStr(ObjItem.CurrentSetting, ",") - 1)
        StrValue = Mid(ObjItem.CurrentSetting, InStr(ObjItem.CurrentSetting, ",") + 1, 256)
		
	Set selItems = objWMIService.ExecQuery("Select * from Lenovo_GetBiosSelections")
	For Each objItem2 in selItems
		objItem2.GetBiosSelections StrItem + ";", strOptions
	Next
		
        WScript.Echo StrItem
	WScript.Echo "  current setting  = " + StrValue
	WScript.Echo "  possible settings = " + strOptions
	WScript.Echo
    End If
Next
