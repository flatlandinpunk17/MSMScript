loopCount = 0
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
Set outfso = CreateObject("Scripting.FileSystemObject")

'Variables that might need to be worked with to get output correct to create the log file'
batDateLog = Year(Date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2)
Set objArgs = WScript.Arguments
CHKDSKlogFilePath = objArgs(0) & "\MSM-" & batDateLog & ".log"
CHKDSKExitFilePath = objArgs(0) & "\CHKDSKExitCodes." & batDateLog & ".log"
Set ResultFSO = CreateObject("Scripting.FileSystemObject")
Set ExitFSO = CreateObject("Scripting.FileSystemObject")
Set ExitFileFSO = ExitFSO.CreateTextFile(CHKDSKExitFilePath, True)
'Set ResultFileFSO = ResultFSO.CreateTextFile(CHKDSKResultFilePath, True)


'outFile = WScript.Arguments.Item(0) & "\chkdskDrives.log"
'Set outFileFso = outfso.CreateTextFile(outFile,True)


'get the number of drives that are internal hard drives'
For Each objDrive in colDrives
	If (objDrive.DriveType = 2) Then
    		loopCount = loopCount + 1
	End If
Next


'Find out how many drives there are and then dynamically create the array'
numDrives = loopCount - 1

Dim strDrives()
Dim strChkDskResults()

ReDim strDrives(numDrives)
ReDim strChkDskResults(numDrives)

loopCount = 0
For Each objDrive in colDrives
	If (objDrive.DriveType = 2) Then
    		strDrives(loopCount) = objDrive.DriveLetter
    		loopCount = loopCount + 1
	End If
Next
Set WshShell = Wscript.CreateObject("Wscript.Shell")
i=0
For Each item In strDrives
	CHKDSKResultFilePath = objArgs(0) & "\CHKDSKResult." & strDrives(i) & "." & batDateLog & ".log"
	Set ResultFileFSO = ResultFSO.CreateTextFile(CHKDSKResultFilePath, True)


	strCheckDisk = "chkdsk " & strDrives(i) & ":"
	
	Set oExec = WshShell.Exec(strCheckDisk)

	Do While oExec.Status = 0
		strOutPut = oExec.StdOut.ReadAll
		ResultFileFSO.Write strOutPut & vbCrLf
		WScript.Sleep 100
	Loop
	Select Case oExec.ExitCode
		Case "0"
			strOutPutExit = " No errors found."
		Case "1"
			strOutPutExit = " Errors were found but fixed."
		Case "2"
			strOutPutExit = " Please run CHKDSK " & strDrives(i) & ": /F /R"
		Case "3"
			strOutPutExit = " Please run CHKDSK " & strDrives(i) & ": /F /R"
	End Select

 	ExitFileFSO.Write "Exit Code: " & oExec.ExitCode & " for drive " & strDrives(i) & strOutPutExit & vbCrLf
 	i = i + 1
Next