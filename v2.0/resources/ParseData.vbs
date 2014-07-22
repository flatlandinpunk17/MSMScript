'Set date correct to be able to open log files correctly'
batDateLog = Year(Date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2)
Set objArgs = WScript.Arguments
MSMlogFilePath = objArgs(0) & "\MSM-" & batDateLog & ".log"
MSMResultFilePath = objArgs(0) & "\MSMResult." & batDateLog & ".log"

'Set and open Result files Remember to make one for each file we will open There has to be a better way to do this'
Set ResultFSO = CreateObject("Scripting.FileSystemObject")
Set ResultFileFSO = ResultFSO.CreateTextFile(MSMResultFilePath, True)

Set OpenedFSO = CreateObject("Scripting.FileSystemObject")


'Create Result Header

ResultFileFSO.Write "***********************************************************" & vbCrLf
ResultFileFSO.Write "*     Results of MSM Script for " & Date & "                 *" & vbCrLf
ResultFileFSO.Write "***********************************************************" & vbCrLf

'Remember to have & vbCrLf after each line to start a new line
ResultFileFSO.Write "Disk Drive Space:" & vbCrLf
'Get Drive Space'
outFile = objArgs(0) & "\DriveSpace." & batDateLog & ".log"
Set OpenedFileFSO = OpenedFSO.OpenTextFile(outFile)
Do Until OpenedFileFSO.AtEndOfStream
	strLine = OpenedFileFSO.ReadLine
	ResultFileFSO.Write strLine & vbCrLf
Loop
OpenedFileFSO.Close
ResultFileFSO.Write "***********************************************************" & vbCrLf
'Get CHKDSK Error Codes'
ResultFileFSO.Write "CHKDSK Exit Codes:" & vbCrLf
outFile = objArgs(0) & "\CHKDSKExitCodes." & batDateLog & ".log"
Set OpenedFileFSO = OpenedFSO.OpenTextFile(outFile)
Do Until OpenedFileFSO.AtEndOfStream
	strLine = OpenedFileFSO.ReadLine
	ResultFileFSO.Write strLine & vbCrLf
Loop
OpenedFileFSO.Close
ResultFileFSO.Write "***********************************************************" & vbCrLf
'Get NetStat Entries'
ResultFileFSO.Write "Netstat:" & vbCrLf
outFile = objArgs(0) & "\Netstat." & batDateLog & ".log"
Set OpenedFileFSO = OpenedFSO.OpenTextFile(outFile)
intLine = 0
Do Until OpenedFileFSO.AtEndOfStream
	strLine = OpenedFileFSO.ReadLine
	intLine = intLine + 1
Loop
ResultFileFSO.Write "NetStat has " & intLine & " Entries" & vbCrLf
OpenedFileFSO.Close
ResultFileFSO.Write "***********************************************************" & vbCrLf

'BIOS information'
ResultFileFSO.Write "BIOS info :" & vbCrLf
outFile = objArgs(0) & "\SysInfo." & batDateLog & ".log"
Set OpenedFileFSO = OpenedFSO.OpenTextFile(outFile)
Do Until OpenedFileFSO.AtEndOfStream
	strLine = OpenedFileFSO.ReadLine
	ResultFileFSO.Write strLine & vbCrLf
Loop
OpenedFileFSO.Close
ResultFileFSO.Write "***********************************************************" & vbCrLf
ResultFileFSO.Write "*    Please remember to actually look through the logs    *" & vbCrLf
ResultFileFSO.Write "***********************************************************" & vbCrLf
ResultFileFSO.Close