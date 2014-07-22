batDateLog = Year(Date) & "-" & Right("0" & Month(date), 2) & "-" & Right("0" & Day(date), 2)
Set objArgs = WScript.Arguments
SysInfoFilePath = objArgs(0) & "\SysInfo." & batDateLog & ".log"
Set ResultFSO = CreateObject("Scripting.FileSystemObject")

Set WshShell = Wscript.CreateObject("Wscript.Shell")

return = WshShell.Run("cmd /c systeminfo > " & SysInfoFilePath,0,True)
Set ResultFileFSO = ResultFSO.OpenTextFile(SysInfoFilePath)
Do Until ResultFileFSO.AtEndOfStream
	strLine = ResultFileFSO.ReadLine
	If InStr(strLine,"System Model:              Virtual Machine") Then
		strBios = "System is a VM. Please check host."
	End If
	If InStr(strLine,"BIOS") > 0 Then
		If strBios <> "System is a VM. Please check host." Then
			strBios = strLine
		End If
	End If
Loop

ResultFileFSO.Close

return = WshShell.Run("cmd /c del " & SysInfoFilePath,0,True)
Set ResultFileFSO = ResultFSO.CreateTextFile(SysInfoFilePath, True)
ResultFileFSO.Write strBios
ResultFileFSO.Close
'This section takes the strBios and checks to see if it threw Dell - 1 as the bios version
'If it has it then runs omreport chassis bios to get the bios from Open Manage
If InStr(strBios, "BIOS Version:              DELL   - 1")  > 0 Then
	intBios = 1
	'return = WshShell.Run("cmd /c del C:\Temp\MSMLog\bios.txt",0,True)
	intOMReturn = WshShell.Run("cmd /c omreport chassis bios > " & SysInfoFilePath,0,True)
End If