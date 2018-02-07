on error resume next

Dim Installer
Dim StrProd_name
Dim oshell
Dim error_code
Dim intprod_len
Dim StrFile_name
Dim Strreboot_Prop
Dim Strlog_file_path
Dim StrRun

StrProd_name = "EV Client"


Set oshell = createobject("wscript.shell")
Set Productcodes = createObject  ("System.collections.ArrayList")

Productcodes.Add "{7A8DE510-7894-4CB5-BC8D-B5143F59A3F5}"
Productcodes.Add "{ADB69895-CF01-484B-BD16-2CE382DD1355}"
Productcodes.Add "{3C6EDD62-4EFF-4B37-B19F-099903227DB1}"
Productcodes.Add "{8FD240B2-3940-4744-B775-3C0544A47FD8}"
Productcodes.Add "{D2B3C9C6-BA0C-439D-AA18-44F707AAAEA2}"
Productcodes.Add "{CC0380C2-A27A-4FEA-A087-74DA45AB40D4}"
Productcodes.Add "{AAB9FBED-B455-4C90-B6F5-7EF93DC91931}"
Productcodes.Add "{899757CA-419A-4B53-8C33-7E909E8D4415}"
Productcodes.Add "{980746E3-7AA4-4982-8336-BEE639E5F3F7}"
Productcodes.Add "{7322AE82-54A0-4181-AF70-0085B8EE3DCB}"
Productcodes.Add "{56AC1D6D-DF2E-4F46-B43D-A8B2A01935BB}"
Productcodes.Add "{73705DEE-C579-4C92-9CBC-A633B3F6A9D2}"
Productcodes.Add "{EA263B74-0878-4DE2-90E8-C2E09ECE59C7}"
Productcodes.Add "{A1BF7E8A-D480-40A4-AA64-67447A9755B4}"
Productcodes.Add "{A21A0045-45FA-4078-8429-892BB3115361}"


logdir = "C:\Windows\Temp\"
'msgbox logdir

Logfile = logdir & "EvClient.log"

'msgbox Logfile

error_code = 0


For each productcode in productcodes

Set Installer = CreateObject ("WindowsInstaller.Installer")

If Installer.productState(productcode) = 5 then

	'Msgbox "We are closing your Outlook Application for uninstalling EV Client,please save your outlook data if any before pressing OK"

		Dim oOL 'Outlook.Application
		Set oOL = GetObject(, "Outlook.Application")
		If oOL Is Nothing Then
    			'no need to do anything, Outlook is not running
		Else
    			Outlook running
    			For i = oOL.Inspectors.count To 1 Step -1
        			oOL.Inspectors.Item(i).CurrentItem.Save
        			oOL.Inspectors.Item(i).Close True
    			Next
    			oOL.Session.Logoff
    			oOL.Quit
		End If


		'msgbox ProductCode
		StrRun = "msiexec /x " & ProductCode & " /norestart /l*v " & logfile & " /qn"
		'msgbox productcode

		'msgbox "Found"
		error_code = oshell.Run(StrRun,1,true)
		if error_code = 3010 then error_code =0
		msgbox "EV Client uninstall successfully."
		WScript.Quit(error_code)

	End if
		Set oOL = Nothing
		

next
		if error_code <> 0 then 
			Set ObjFSO = CreateObject("Scripting.FileSystemObject")
               		 Set objLog = objFSO.CreateTextFile("c:\windows\temp\EV_client_Error.log")
               		 objLog.WriteLine "Error While Unistalling EV Client"

           		Msgbox"Error While Unistalling EV Client.Please direct to PitStop Or MS Fix tool"         	
			WScript.Quit(error_code)
 		end if

WScript.Quit(error_code)







