'  --- [ChangeMgmtGroupSettingsLocalhost.vbs] Visual Basic Script  ---
'
' Author(s):        Ryan Irujo
' Inception:        01.13.2015
' Last Modified:    01.13.2015
'
' Description:      Visual Basic Script that can add/remove Connection Settings in the SCOM 2012 Microsoft Monitoring Agent.
'                   Results of the run are returned to a File called 'ChamgeMgmtGroupSettingsLocalhostResults.txt' in 
'                   'C:\Temp'
'
'
' Syntax:           cscript C:\Temp\ChangeMgmtGroupSettingsLocalhost.vbs 
'
' Example:          cscript C:\Temp\ChangeMgmtGroupSettingsLocalhost.vbs

' Creating Results File.
Set objFSO         = CreateObject("Scripting.FileSystemObject")
Set objResultsFile = objFSO.CreateTextFile("C:\Temp\ChangeMgmtGroupSettingsLocalhostResults.txt", True)


' Declaring SCOM Agent Variable(s)
Dim objMSConfig
Set objMSConfig = CreateObject("AgentConfigManager.MgmtSvcCfg")

' Retrieving Hostname of Computer.
Set wshShell = WScript.CreateObject("Wscript.Shell")
strComputerName = wshShell.ExpandEnvironmentStrings( "%COMPUTERNAME%" )


' Adding Connection Settings for a Management Group to the SCOM Agent Configuration
Call objMSConfig.AddManagementGroup("TESTMGMTGROUP01", "SCOMSERVER101.scom.local",5723)
If Err.Number <> 0 Then
	ObjResultsFile.WriteLine "Unable to update the SCOM Agent Connection Settings on [" & strComputerName & "] --> " & Err.Description
End If

If Err.Number = 0 Then
	ObjResultsFile.WriteLine "Successfully updated the SCOM Agent Connection Settings on [" & strComputerName & "]"
	Err.Clear	
End If

'Remove a management group
'Call objMSConfig.RemoveManagementGroup ("MyManagementGroupToRemove”)
