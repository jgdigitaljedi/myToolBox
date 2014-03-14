Const ForAppending = 8
DIM fso, objCompNames, strComputer
Set fso = CreateObject("Scripting.FileSystemObject")
Set objCompNames = fso.OpenTextFile("c:\scripts\Check4File\computers.txt")
Set objResultsFile = fso.OpenTextFile("c:\scripts\Check4File\results.txt", ForAppending, True)

Do Until objCompNames.AtEndOfStream
	strComputer = objCompNames.ReadLine

	If (fso.FileExists("\\" & strComputer & "\c$\Program Files\Java\jre1.6.0_35\README.txt")) Then
		objResultsFile.Writeline(strComputer & " has the requested file!")
	Else
		objResultsFile.Writeline(strComputer & " DOES NOT HAVE THE FILE!!!*************************")
	End If
Loop

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run "notepad.exe c:\scripts\Check4File\results.txt"
Set objShell = Nothing 

objCompNames.Close
MsgBox "Done!"

WScript.Quit()
