Const HKEY_CURRENT_USER = &H80000001 
Const HKEY_LOCAL_MACHINE = &H80000002 

strComputer = "." 
  
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & _ 
    strComputer & "\root\default:StdRegProv") 
	
Set WshShell = Wscript.CreateObject("WScript.Shell")

paths = Array ("SOFTWARE\Wow6432Node\PGP Corporation\PGP","SOFTWARE\PGP Corporation\PGP","SOFTWARE\Wow6432Node\PGP Corporation\Common\AllUsers","SOFTWARE\PGP Corporation\Common\AllUsers")
strValueName = "MACHINEGUID"

Set fileSystemObject = CreateObject("Scripting.FileSystemObject")
outfile = "C:\Windows\Guidresetlog.txt"
Set objFile = fileSystemObject.CreateTextFile(outfile, true)
objFile.WriteLine "PGP Guid reset script starting "


For Each item In paths
	strValueName = "MACHINEGUID"
	strVersionKey = "SOFTWARE\Wow6432Node\PGP Corporation\Common"
	strVersionValue = "PRODUCTVERSION"
	oReg.GetStringValue HKEY_LOCAL_MACHINE,item,strValueName,strValue
	If Not IsNull(strValue) Then
		'checking if GUID is in the list
		objFile.WriteLine item & " MACHINEGUID found, validating"
		choice = strValue
		options = Array("{F7E30C81-F0A9-468B-82FE-A646D028E31A}", _
						"{E86B4900-CB51-4B64-BD1D-B11184367928}", _
						"{897ABD28-C059-4F9E-8570-03425D9030D9}", _
						"{9F9E9BA5-C2EF-4B19-B3D3-BE1F8BAAB259}", _
						"{00B75736-EBE5-4925-A10F-3900EF010173}", _
						"{104185D8-5CAA-4946-9C08-3A76C7DB837A}", _
						"{5212CEC1-A0AF-475E-B0E0-EAA679097114}", _
						"{E05ACFD0-A45D-4E6B-8963-625A3ECFA68A}", _
						"{40237F29-72B8-4D7A-A1C2-1821D8ED7171}", _
						"{0C17D92A-D184-4498-BB73-0F5850690601}", _
						"{CAD1233D-F7D3-4E62-8588-2136A77FE9F4}", _
						"{EB2DD362-BB2A-4BC4-AEBB-4E15EE9CD864}", _
						"{F35FE368-53D4-46A8-ACC1-8F0ED273643E}", _
						"{2A0B959C-A8B0-4A89-B090-BFE5D6C39B30}", _
						"{73088087-5DC0-409C-8032-B1D15367242F}", _
						"{BFF60092-59EA-4F1F-AA0E-A7E39A55B515}", _
						"{3C82D26F-77DA-40C6-A943-1E999F0BD0E9}", _
						"{2E0BD825-2380-4F45-8953-8C3B9F47E013}", _
						"{3F1854B6-4FD2-4DC2-A88C-846C7E9F6DBA}", _
						"{BFB7F7B3-0CC4-4A7F-A4FA-5DF85FD0AAA5}", _
						"{E05ACFD0-A45D-4E6B-8963-625A3ECFA68A}", _
						"{A2F9527D-05E6-4C69-A8F9-398980EC6364}", _
						"{EBBF2857-5098-4BE7-918A-8E6FF412E0FF}", _
						"{C0EF6205-EC5F-4B34-AE48-6A33E0435542}", _
						"{2D8A1D67-61CF-41F1-BEAE-40510AF87478}", _
						"{3DA3F7F6-C074-4AFC-A10D-CA5C1FECC4C3}", _
						"{7B8DD3C6-986E-45C6-AC50-31C3A38FB074}", _
						"{05106c6c-46d3-4175-876c-61bae4d356c5}", _
						"{0562a412-a995-4035-8860-f82e19af3224}", _
						"{121599de-397e-444a-88ff-b523f8ad4860}", _
						"{128893cc-499f-4d19-9846-bbc96a6bb804}", _
						"{1ecbd892-5aa3-484e-b117-009f8569e3bd}", _
						"{21efe12e-3148-4278-b77b-3e6904348883}", _
						"{22aacd9c-c9b8-450e-b926-91b8fb6e985a}", _
						"{264c9bee-0654-43c5-90eb-970e08bf0e45}", _
						"{28b2babf-dae4-4a49-8856-06d848e69a25}", _
						"{29a8aa60-8ad8-400c-82eb-bf489c719a18}", _
						"{2d8a1d67-61cf-41f1-beae-40510af87478}", _
						"{2e5be61f-b472-4603-84de-567004ef0c39}", _
						"{2fb68a99-748c-43ce-a7e7-90667ce46efa}", _
						"{36cef7d4-a869-40d9-8bb0-dc4cc5e15f49}", _
						"{3905e216-ac14-4b15-a820-679474cf4c24}", _
						"{3912d1ef-f762-48dc-943c-ecd245279318}", _
						"{39e5db3d-3cd3-4980-8e0b-9f11421e0189}", _
						"{3e34c8cb-8e71-4339-950a-6bd6e4fc6479}", _
						"{4832d9ee-66d2-43fb-affc-90eb00543393}", _
						"{4cc2cd5d-f607-47da-94c1-2734a558a68f}", _
						"{4e0f44b8-76d8-4347-942a-60bcf7dc5445}", _
						"{526c3751-3f4a-45c3-88a6-505acaca29c7}", _
						"{54be68a3-0d3f-4731-b0f4-ffd520475265}", _
						"{56a07db4-4664-4a1f-a109-1f76de75d453}", _
						"{56d6cef0-aecb-43e4-ae96-63aa78b1becf}", _
						"{5774b00f-cfc3-4a97-903d-f2beaa27e90a}", _
						"{59e8777b-1df6-45ac-9aaf-a69e14781b48}", _
						"{5b5d0a51-04c9-48f2-9973-2e3af2658a86}", _
						"{5fd901c9-52ff-4dc3-9214-ce14e9b60ff0}", _
						"{625ab93f-2a9e-40f1-bf81-d5c02e768907}", _
						"{661a4ba5-a4db-4a59-bd81-a4c27b7d1eb3}", _
						"{67afaf4e-6461-4cbe-be35-13141898e6c0}", _
						"{6d38e383-6a32-45ea-bd1a-e242e34888fc}", _
						"{6e68c920-03d5-441c-b224-88c2910344ee}", _
						"{710d517a-eb1d-40fb-8fe4-d0001d901df2}", _
						"{771aa4fd-85bd-457f-b41c-7558c2aad906}", _
						"{77327c18-5ab3-4a09-9733-1c265f0929c2}", _
						"{786877a9-4812-4093-b9da-bb31fa576061}", _
						"{7f6973d4-a5be-4ab3-8424-7965b0b26472}", _
						"{860aa1ce-e91c-46a5-8cf0-a7dbb822da81}", _
						"{89a45077-7fd8-4d73-a5da-25d2f698cfbd}", _
						"{8ccffc0a-d7e1-42ba-a8c6-c2b72c4666cc}", _
						"{93a63c73-02b9-430c-b71d-935faaf2cec8}", _
						"{952a09cf-efd5-4f2b-8887-44c69a3303af}", _
						"{95e3f41d-0bae-4865-9256-b92576c109a6}", _
						"{9d7c7360-63ed-4060-b7f8-f1931bbbcf86}", _
						"{a2f9527d-05e6-4c69-a8f9-398980ec6364}", _
						"{a356bb32-2738-4db0-bb0a-df4593dcd802}", _
						"{a7b4eac1-ae9d-4818-8783-0535ba27f0a3}", _
						"{a940324b-d986-416e-b58b-3eb6dded7e8d}", _
						"{aa68fe2d-a88b-4c6b-9047-11c16d8614d7}", _
						"{af96687f-89bd-4468-b5a9-63adb91e33e6}", _
						"{af9b8c7a-d1d2-4ce9-a823-ee739006652b}", _
						"{b0d0215d-ed24-4082-9079-c80edbf5da5b}", _
						"{b0d99286-174b-400b-ad8d-72f92b41a3f9}", _
						"{b2e099d6-4bad-4340-9c7e-0b029835725e}", _
						"{b5d062cf-3a52-466a-a8df-62818a5f012f}", _
						"{b6e35be2-958d-44c2-9def-5d59ec3c7cd3}", _
						"{b8c87772-8e5b-455a-8a93-5bed4a06795a}", _
						"{b8e04b7d-f13a-4dcd-8c00-0532b28b4774}", _
						"{b98ae8d3-3632-4e42-ab6f-1d72ea15abd2}", _
						"{ba666900-3b88-457b-81a0-6505cc97f6e9}", _
						"{bbf75fe8-3a2a-47c5-837b-b9e25bce3a21}", _
						"{c0795906-44ad-49a2-873e-c96a34d74c3e}", _
						"{c0dc6157-1f18-422b-a8f7-5dd8bc114f6f}", _
						"{c12a8cd3-e8ff-4fa6-9ff7-13e2ed9aeb6c}", _
						"{cacd5e35-0241-49aa-bdd1-bfcc8296a74a}", _
						"{cd1044d4-3f91-4a3b-9c51-ae1c6819948c}", _
						"{d2aef0cb-8627-438f-b579-ee4af52cb9cb}", _
						"{d35f18dd-4a6f-429a-8b57-ec38e271a4d7}", _
						"{d53ff548-ae66-4384-8b04-4e969bbe596f}", _
						"{d69663f8-a8cb-40bf-9f72-d6db270981f6}", _
						"{d9311ed9-e2a4-4ecc-bae2-d27d571a6fc5}", _
						"{d94c903a-2aba-413b-91f4-901c09237edd}", _
						"{de4660b4-d02d-49fb-a4ca-99a2f766755f}", _
						"{df18ad8d-adbf-494e-ae0e-61787d4b2574}", _
						"{e25d825c-795e-4a8f-ac3b-05acc13b3f80}", _
						"{e941b5ba-7d60-4ede-94ae-d69d827c449a}", _
						"{eccdca61-0558-4ab6-865d-a1545bfa05eb}", _
						"{eefca10c-f008-4dd4-9b16-0930f3d952c2}", _
						"{f177c326-2dd5-4cb2-88a5-40a32a8617a5}", _
						"{f3a32ade-77b9-4b38-bae2-00ea8a7ff562}", _
						"{f5c8edde-d442-488c-9474-0cd2d31bb40d}", _
						"{f9c856d6-ddd0-4626-b0f9-4a3fd016d3ef}")
			For j=Lbound(options) to Ubound(options)
			If InStr(options(j), choice) <> 0 Then
				objFile.WriteLine item & " Guid checked and found a duplicate " & choice
				'Checking version of PGP
				oReg.GetStringValue HKEY_LOCAL_MACHINE,strVersionKey,strVersionValue,strVersion	
				objFile.WriteLine strVersion
				Select Case strVersion 
				Case "10.4.1.490"
					objFile.WriteLine "Run PGP GUID RESET"
					WshShell.Run "PGPEnroll.bat"
				Case "10.3.2.16127"
					objFile.WriteLine "Run PGP GUID RESET"
					WshShell.Run "PGPEnroll2.bat"
				Case "10.3.2.21495"
					objFile.WriteLine "Run PGP GUID RESET"
					WshShell.Run "PGPEnroll3.bat"
				Case "10.4.0.1211"
					objFile.WriteLine "Run PGP GUID RESET"
					WshShell.Run "PGPEnroll.bat"
				End Select
			End If
			Next
			bjFile.WriteLine "MACHINEGUID NOT A DUPLICATE."
	Else 
		objFile.WriteLine item & " MACHINEGUID NOT FOUND. PGP NOT INSTALLED"
	End If
Next
