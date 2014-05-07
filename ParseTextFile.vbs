'  --- [ParseTextFile.vbs] Visual Basic Script  ---
'
' Author(s):        Ryan Irujo
' Inception:        05.06.2014
' Last Modified:    05.06.2014
'
' Description:      Sample Visual Basic Script that Parses out Values from a Text File and returns the results to SCOM.
'
'
' Syntax:          cscript ParseTextFile.vbs 
'
' Example:         cscript ParseTextFile.vbs 

' Declaring SCOM Variables
Set oAPI = CreateObject("MOM.ScriptAPI")
Set oBag = oAPI.CreatePropertyBag()

' Declaring Standard Variables
Dim objFSO
Dim objFile
Dim strWorker
Dim strWorkerValues
Const ForReading = 1

' Settings String Value to look for.
strWorker     = "Dallas-Prod"

' Setting Text File to Read.
Set objFSO    = CreateObject("Scripting.FileSystemObject")
Set objFile   = objFSO.OpenTextFile("D:\Sandbox\Strings.txt", ForReading)

' Retrieving All Lines from the Text File.
Dim arrFileLines()
i = 0
Do Until objFile.AtEndOfStream
Redim Preserve arrFileLines(i)
arrFileLines(i) = objFile.ReadLine
i = i + 1
Loop
ObjFile.Close


' Pulling out the Worker Line Value from the Text File.
For Each strLine in arrFileLines
	arrLines = Split(strLine)
	If strWorker = arrLines(0) Then
	strWorkerValues = strLine
	End If
Next


' Splitting out the Numeric Value From the Matching Entry and Returning the Results to SCOM.
arrValues = Split(strWorkerValues)
If arrValues(1) < 4000 Then
	'WScript.Echo strWorkerValues & " is OK"
	'Call oBag.AddValue("ComputerName",sComputerName)
	Call oBag.AddValue("InstanceName",arrValues(0))
	Call oBag.AddValue("CounterName","CitrixWorkerAvg")
	Call oBag.AddValue("PerfValue",arrValues(1))
	Call oAPI.Return(oBag)
End If

If arrValues(1) >= 4000 Then
	'WScript.Echo strWorkerValues & " is CRITICAL"
	'Call oBag.AddValue("ComputerName",sComputerName)
	Call oBag.AddValue("InstanceName",arrValues(0))
	Call oBag.AddValue("CounterName","CitrixWorkerAvg")
	Call oBag.AddValue("PerfValue",arrValues(1))
	Call oAPI.Return(oBag)
End If
