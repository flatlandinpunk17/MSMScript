loopCount = 0
batDateLog = Year(Date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2)
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set colDrives = objFSO.Drives
Set outfso = CreateObject("Scripting.FileSystemObject")
Set objArgs = WScript.Arguments
outFile = objArgs(0) & "\DriveSpace." & batDateLog & ".log"
Set outFileFso = outfso.CreateTextFile(outFile,True)

'Wscript.Echo WScript.Arguements(0)

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
'ReDim strChkDskResults(numDrives)

loopCount = 0
For Each objDrive in colDrives
	If (objDrive.DriveType = 2) Then
    		strDrives(loopCount) = objDrive.DriveLetter
    		loopCount = loopCount + 1
	End If
Next

i = 0
For Each item In strDrives
	outFileFso.Write"Drive " & strDrives(i) & " has " & round(objfso.GetDrive(strDrives(i)).FreeSpace / objfso.GetDrive(strDrives(i)).TotalSize * 100,2) & "% of free space" & vbCrLf
	outFileFso.Write"Drive " & strDrives(i) & " has " & round(objfso.GetDrive(strDrives(i)).FreeSpace/1024/1024/1024,2) & "GB free of " & round(objfso.GetDrive(strDrives(i)).TotalSize/1024/1024/1024,2) & "GB" & vbCrLf
	i = i + 1
Next
