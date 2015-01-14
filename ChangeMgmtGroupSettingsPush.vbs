'  --- [ChangeMgmtGroupSettingsPush.vbs] Visual Basic Script  ---
'
' Author(s):        Ryan Irujo
' Inception:        01.14.2014
' Last Modified:    01.14.2014
'
' Description:      Visual Basic Script that can add/remove Connection Settings in the SCOM 2012 Microsoft Monitoring Agent.
'                   This Script is responsible for copying the 'ChangeMgmtGroupSettingsLocalhost.vbs' File to a remote host
'                   and then executing the Script to change the Connections Settings on the SCOM 2012 Microsoft Monitoring
'                   Agent on that host.
'
'
' Syntax:           cscript ChangeMgmtGroupSettingsPush.vbs 
'
' Example:          cscript ChangeMgmtGroupSettingsPush.vbs 


' Declaring Constant Variables
Const INPUT_FILE_NAME   = ".\ChangeMgmtHostnames.txt"
Const FOR_READING       = 1
Const OverwriteExisting = TRUE

' Reading the list of Hosts to Modify from a Text File.
Set objFSOTextFile = CreateObject("Scripting.FileSystemObject")
Set objFSOHostNames = objFSOTextFile.OpenTextFile(INPUT_FILE_NAME,FOR_READING)
strComputers = objFSOHostNames.ReadAll
objFSOHostNames.Close

' The List of all the Hostnames is added into an Array.
arrComputers = Split(strComputers,vbCrLf)


' Iterating through the Array.
For Each strComputer in arrComputers

	' Checking if the Folder exists on the remote Host.
	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
	Set colFolders    = objWMIService.ExecQuery("Select * From Win32_Directory Where Name = 'C:\\Temp'")

	If colFolders.Count = 0 Then
		Wscript.Echo "'C:\Temp' does not exist on [" & strComputer & "]."
		
		' Creating the C:\Temp Directory on the Host if it doesn't exist.
		Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
		createFolder = objWMIService.Create("cmd.exe /c md C:\Temp", Null, Null, intProcessID)	
		
		If createFolder = 0 Then
			Wscript.Echo "C:\Temp was Successfully Created on [" & strComputer & "]."
			
			' Need to Sleep for 1 Second after the 'C:\Temp' Directory is created in order for it to be accessible by the rest of the Script.
			WScript.Sleep(1000)	
		End If
		
		If createFolder <> 0 Then
			Wscript.Echo "C:\Temp could not be created on [" & strComputer & "]. -- " & Err.Description
		End If
		
	Else
		Wscript.Echo "'C:\Temp' already exists on [" & strComputer & "]."
	End If



	' Copying over the Files and Running the VBScripts.
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.CopyFile "ChangeMgmtGroupSettingsLocalhost.vbs", "\\" & strComputer & "\C$\Temp\", OverWriteExisting

	Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2:Win32_Process")
	changeMgmtSettings = objWMIService.Create("cscript C:\Temp\ChangeMgmtGroupSettingsLocalhost.vbs", null, null, intProcessID)

	If changeMgmtSettings = 0 Then
		Wscript.Echo "Successfully updated the SCOM Agent Connection Settings on [" & strComputer & "]."
	End If

	If changeMgmtSettings <> 0 Then
		Wscript.Echo "Failed to update the SCOM Agent Connection Settings on [" & strComputer & "]. -- " & Err.Description
	End If
Next

