'<*************** dimensions and constants ***************>
Dim objExcel, objFso, objFile, objSheet, dtmYesterday
Const ForReading = 1
Const xlAscending = 1
Const ForWriting = 2
Dim a,z
Set objShell = CreateObject("WScript.Shell")

'<*************** user interaction for date entry ***************>
dtmLastDay = Day(Date()) - 1
dtmToday = WeekDayName(WeekDay(Now))

If dtmToday = "Monday" Then
	dtmYesterday = dtmLastDay - 2
	Call YTDinfo
Else
	dtmYesterday = dtmLastDay
End If

strTeenth = Left(dtmYesterday, 1)
If strTeenth = "1" Then
	strLast = "4"
Else
	strLast = Right(dtmYesterday, 1)
End If
	

Select Case(strLast)
	Case "1"
		dtmYesterday = dtmYesterday & "st"
	Case "2"
		dtmYesterday = dtmYesterday & "nd"
	Case "3"
		dtmYesterday = dtmYesterday & "rd"
	Case "4", "5", "6", "7", "8", "9", "0"
		dtmYesterday = dtmYesterday & "th"
End Select

strRow = InputBox("Enter the yesterday's date like it appears in the spreadsheet. ex = '26th'", "Running Yesterday's numbers", dtmYesterday)
	If strRow = "" Then
		WScript.Quit
	End If

'<*************** open report and create Excel doc and set variables ***************>
'On Error Resume Next
Set objFso = CreateObject("Scripting.FileSystemObject")
Set objFile = objFso.OpenTextFile("L:\Joey_Thumb\JoeyScripts\shortkeys\reporting.txt", ForReading)
	If Err <> 0 Then
		WScript.Echo "Couldn't reach the reporting file. The connection is probably screwy. Try again or something. Error#= " & Err.Number
	End If

strXlPath = "C:\Users\pg8077\Desktop\Sorted_Report.xlsx"
Set objExcel = CreateObject("Excel.Application")
Set objWkBk = objExcel.Workbooks.Open(strXlPath)
objExcel.Visible = True
Set objSheet = objExcel.ActiveWorkbook.Worksheets(1)
objSheet.Cells.Clear
irow = 1
icol = 1

'<*************** check to see if anyone even launched it yesterday and if not to enter a zero for launches ***************>
On Error Resume Next
strCheck = objFile.ReadLine
If strCheck = ""  or Err.Number <> 0 Then
	WScript.Echo "It doesn't like like anyone launched it yesterday. :-( "
	objShell.Run "notepad.exe " & objFile
	objWkBk.Close
	strMonth = InputBox("Type the name of the current month as it appears on the spreadsheet.", "Where to add zero entry for launches.")
	Set objMaster = objExcel.Workbooks.Open("c:\Users\pg8077\Desktop\ToolBox_Stats.xlsx")
	Set objMSheet = objMaster.Worksheets(strMonth)
	objMaster.Worksheets(strMonth).Activate
	Set objBegin = objExcel.Range("A1").Find(strRow)
	objBegin.Select
	a = objExcel.ActiveCell.Row
	z = 2
	objMSheet.Cells(a, z).Value = "0"
	objMaster.Save
	objExcel.UserControl = True
	strMasterPath = "c:\Users\pg8077\Desktop\ToolBox_Stats.xlsx"
	strAlma = "L:\Joey_Thumb\JoeyScripts\shortkeys\ToolBox_Stats.xlsx"
	objFso.CopyFile strMasterPath, strAlma, TRUE
	objFile.Close
	WScript.Quit	
End If

'<*************** write lines to temp excel doc is above test was passed***************>
Set objFile = objFso.OpenTextFile("L:\Joey_Thumb\JoeyScripts\shortkeys\reporting.txt", ForReading)
Do While Not objFile.AtEndOfStream
	strLine = objFile.Readline
	strDelimit = Split(strLine, "::")
		For Each strSection in strDelimit
			objExcel.Cells(irow, icol).Value = strSection
			icol = icol + 1
		Next
	irow = irow + 1
	icol = 1
Loop

'<*************** format and sort imported results ***************>
objSheet.Columns("A:A").NumberFormat = "dd-mmm"
objSheet.Columns("B:B").ColumnWidth = 85
Set objRange = objSheet.UsedRange
Set objRange2 = objExcel.Range("B1")
objRange.Sort objRange2,xlAscending,,,,,,xlYes

'<*************** split date in column A to later reference correct sheet in master doc ***************>
strValue = objSheet.Range("A1").Text
arrText = Split(strValue, "-")
strWhich = TRIM(arrText(1))

'<*************** convert month into format that corresponds with the month names on master spreadsheet ***************>
Select Case strWhich
Case "Jan"
	strMonth = "Jan"
Case "Feb"
	strMonth = "Feb"
Case "Mar"
	strMonth = "Mar"
Case "Apr"
	strMonth = "Apr"
Case "May"
	strMonth = "May"
Case "Jun"
	strMonth = "Jun"
Case "Jul"
	strMonth = "Jul"
Case "Aug"
	strMonth = "Aug"
Case "Sep"
	strMonth = "Sep"
Case "Oct"
	strMonth = "Oct"
Case "Nov"
	strMonth = "Nov"
Case "Dec"
	strMonth = "Dec"
End Select

'<*************** save sorted excel document to make easy to reference later ***************>
objExcel.DisplayAlerts = False
objExcel.Save

'<*************** open master excel report and search for correct worksheet ***************>
Set objMaster = objExcel.Workbooks.Open("c:\Users\pg8077\Desktop\ToolBox_Stats.xlsx")
Set objMSheet = objMaster.Worksheets(strMonth)
objMaster.Worksheets(strMonth).Activate

'<*************** search for row in worksheet that corresponds with date ***************>
Set objBegin = objExcel.Range("A1").Find(strRow)
objBegin.Select
a = objExcel.ActiveCell.Row
z = ""

'<*************** reference cases to add values to correct columns in row ***************>
LastRow = objSheet.UsedRange.Rows(objSheet.UsedRange.Rows.Count).Row
	For Each cell in objSheet.Range("B1:B" & LastRow)
		strData = LTRIM(cell.Value)
		Select Case strData
			Case "Alma share reached.", "Alma_Info was not reached. Phrasekeys library based off last successful update."
				z = 2
				AddCell(z)				
			Case "Local machine rebooted", "Remote machine rebooted"
				z = 16
				AddCell(z)
			Case "Local machine scheduled reboot", "Remote machine scheduled reboot"
				z = 17
				AddCell(z)
			Case "Local machine shutdown","Remote machine shutdown"
				z = 18
				AddCell(z)
			Case "Local machine scheduled shutdown", "Remote machine scheduled shutdown"
				z = 19
				AddCell(z)
			Case "User specified task scheduled on local machine", "User specified task scheduled on remote machine", "User specified task scheduled on a list of remote machines"
				z = 47
				AddCell(z)
			Case "Remote machine pinged"
				z = 6
				AddCell(z)				
			Case "Local IE files cleaned"
				z = 24
				AddCell(z)
			Case "Mapped drives and printers listed"
				z = 12
				AddCell(z)
			Case "Machine info displayed"
				z = 13
				AddCell(z)
			Case "Checked to see who is logged in"
				z = 14
				AddCell(z)								
			Case "SCCM uninstalled"
				z = 22
				AddCell(z)				
			Case "WMI checked and repaired"
				z = 21
				AddCell(z)				
			Case "File search completed"
				z = 8
				AddCell(z)
			Case "BIOS type info displayed"
				z = 20
				AddCell(z)
			Case "MsiExec repair done"
				z = 23
				AddCell(z)
			Case "Temp directories emptied"
				z = 25
				AddCell(z)				
			Case "Defrag run"
				z = 49
				AddCell(z)
			Case "Autorun disabled"
				z = 15
				AddCell(z)
			Case "Reverse IP lookup done"
				z = 27
				AddCell(z)
			Case "Specified File copyed to remote mahcine"
				z = 26
				AddCell(z)				
			Case "Local ipconfig run"
				z = 7
				AddCell(z)	
			Case "Remote serial number search done"
				z = 46
				AddCell(z)	
			Case "User/group added to local admin group"
				z = 50
				AddCell(z)
			Case "'Open Cmd Prompt Here' Added to right-click menu"
					'do nothing for now
			Case "PST file search completed"
				z = 48
				AddCell(z)
			Case "Phrasekeys Notes generator used"
				z = 4
				AddCell(z)	
			Case "Phrasekeys email generator used"
				z = 5
				AddCell(z)	
			Case "Paths changed for IBM launchers"
				z = 37
				AddCell(z)
			Case "Log file viewed"
				z = 30
				AddCell(z)
			Case "Log file cleared"
				z = 31
				AddCell(z)
			Case "Bug reported"
				z = 32
				AddCell(z)
			Case "RUtil launched"
				z = 33
				AddCell(z)
			Case "Tivoli launched"
				z = 34
				AddCell(z)
			Case "MSTSC launched"
				z = 29
				AddCell(z)				
			Case "Admin cmd prompt launched"
				z = 9
				AddCell(z)
			Case "Paste Tool launched"
				z = 38
				AddCell(z)
			Case "RTA Launched"
				z = 36
				AddCell(z)
			Case "WMA Launched"
				z = 35
				AddCell(z)
			Case "Instructional ppsx viewed", "Text Area ppsx viewed", "Power Options ppsx viewed", "Phrasekeys ppsx viewed", "Scripted Tasks ppsx viewed"
				z = 11
				AddCell(z)	
			Case "Changelog viewed", "Changelog not viewed because of bad connection"
				z = 10
				AddCell(z)
			Case "Toolbox update downloaded"
				z = 40
				AddCell(z)
			Case "Update check ran but latest version already installed", "Update check failed because Alma wasn't reached"
				z = 39
				AddCell(z)
			Case "PHRASEKEYS STANDALONE launched"
				z = 43
				AddCell(z)
			Case "PHRASEKEYS STANDALONE: Notes generator used"
				z = 44
				AddCell(z)
			Case "PHRASEKEYS STANDALONE: Email generator used"
				z = 45
				AddCell(z)
			Case "run.vbs reg key removed"
				z = 51
				AddCell(z)
			Case "User/group added to local admin group"
				z = 50
				AddCell(z)
			Case "User added to remote users group"
				z = 52
				AddCell(z)
			Case "Installed IE8 and Customization package"
				z = 53
				AddCell(z)
			Case "NIC speed checked"
				z = 54
				AddCell(z)
			Case "Installed Software Store Service"
				z = 55
				AddCell(z)
			Case "Time submitted in CLAIM tracker"
				z = 56
				AddCell(z)
			Case "Time viewed in CLAIM tracker"
				z = 57
				AddCell(z)
			Case "CCMA Active X Controls installed."
				z = 58
				AddCell(z)
			Case "Qfiniti registry fix applied."
				z = 28
				AddCell(z)
			Case "Submit area saved", "Submit area opened", "Submit area cleared", "CompName area saved", "CompName area opened", "CompName area cleared"
				'do nothing
			Case Else
				strAnswer = MsgBox("Task '" & strData & "' is out of scope. Do you want to add it to the 'errors' column?", vbYesNo+vbInformation, "Task not listed")
				If strAnswer = vbYes Then
					z = 3
					AddCell(z)
				End If
				'errors from log file	
		End Select
	
	Next
	
	Function AddCell(z)
		irow = a
		icol = z
		cellData = objMSheet.Cells(irow, icol).Value
			If cellData = "" Then
				cellData = "0"
			End If		
		cellData = cellData + 1
		objMSheet.Cells(irow, icol).Value = cellData
	End Function

'<*************** writing info to excel sheet ***************>	
objFile.Close
Set objFile = objFso.OpenTextFile("L:\Joey_Thumb\JoeyScripts\shortkeys\reporting.txt", ForWriting)
objFile.Write ""
ObjFile.Close

'If dtmToday = "Monday" Then
'	YTDinfo
'End If	

strSave = MsgBox("Would you like to save the newly added data to the master spreadsheet?", vbYesNo+vbInformation, "Happy with the results?")
If strSave = vbYes Then
	objMaster.refreshall
	objMaster.Save
	objWkBk.Close
	objExcel.UserControl = True
Else
	MsgBox "You will not be prompted to save at close if you decide you want to save now. You will have to do it manually.", vbExclamation, "WARNING!!"
	objExcel.UserControl = True
End If

strMasterPath = "c:\Users\pg8077\Desktop\ToolBox_Stats.xlsx"
strAlma = "L:\Joey_Thumb\JoeyScripts\shortkeys\ToolBox_Stats.xlsx"
objFso.CopyFile strMasterPath, strAlma, TRUE

'<*************** saving weekly ticket data to YTD excel sheet ***************>


Sub YTDinfo
'<*************** setting it all up ***************>
	Dim strCol 
	Set objFso = CreateObject("Scripting.FileSystemObject")
	Set objExcel = CreateObject("Excel.Application")
	Set objCTBk = objExcel.Workbooks.Open("C:\Phrasekeys\CLAIM Tracker.xlsx")
	Set objSheet = objExcel.ActiveWorkbook.Worksheets(3)
	objSheet.Cells.Copy
	Set objYTD = objExcel.Workbooks.Open("C:\Phrasekeys\YTD Ticket info.xlsx")
'<*************** adding new week's data as new sheet to YTD WkBk ***************>
	Set objNewSheet = objYTD.Worksheets.Add
	With objNewSheet
		'.Range("A1").PasteSpecial -4122
		.Range("A1").PasteSpecial -4163
		.Name = cstr(month(date)) & "-" & dtmYesterday
	End With
'<*************** figuring out how many workdays to add to totals sheet and adding it ***************>
	strSheetMonth = Split(objNewSheet.Name, "-")
	strMonthRow = strSheetMonth(0) + 1
	strWeekL = MsgBox("Was last week a regular 5 day work week?", vbYesNo+vbExclamation, "Regular 5 day week?")
	If strWeekL = vbYes Then
		strDaysWork = objYTD.Sheets("Totals").Cells(strMonthRow, 5).Value
		strDaysWork = strDaysWork + 5
		objYTD.Sheets("Totals").Cells(strMonthRow, 5).Value = strDaysWork
	Else
		strDaysWork = objYTD.Sheets("Totals").Cells(strMonthRow, 5).Value
		strShortWeek = InputBox("Enter the number of days worked last week.", "Days worked")
		strDaysWork = strDaysWork + strShortWeek
		objYTD.Sheets("Totals").Cells(strMonthRow, 5).Value = strDaysWork
	End If
'<*************** counting number of tickets closed last week to add it to the Totals sheet ***************>
	strSheMon = cstr(month(date)) & "-" & dtmYesterday
	x = 1
	'Set objCell = objYTD.Sheets(strSheMon).Cells(x, 1)
	strCounted = objYTD.Sheets(strSheMon).Cells(x, 1).Value
	Do While Not IsEmpty(strCounted)
		x = x + 1
		strCounted = objYTD.Sheets(strSheMon).Cells(x, 1).Value
		If strCounted = "" or strCounted = " " Then
			Exit Do
		End If
		strSplit = Split(strCounted, "-")
		strLTask = LCase(strSplit(1))
		strTask = TRIM(strLTask)
		Select Case strTask
			Case "ccma"
				strCol = 8
			Case "qfiniti"
				strCol = 9
			Case "reimage"
				strCol = 6
			Case "lr", "lease roll"
				strCol = 10
			Case "ru", "ar", "admin"
				strCol = 7
			Case Else
				strCol = 11
		End Select
		'MsgBox strTask & " " & strCol
		strOldVal = objYTD.Sheets("Totals").Cells(strMonthRow, strCol).Value
		'MsgBox strOldVal
		If strOldVal = "" Then
			strOldVal = 0
		End If
		strOldVal = strOldVal + 1
		objYTD.Sheets("Totals").Cells(strMonthRow, strCol).Value = strOldVal
	Loop
	strCounted = x - 1
	strOldTC = objYTD.Sheets("Totals").Cells(strMonthRow, 2).Value
	objYTD.Sheets("Totals").Cells(strMonthRow, 2).Value =  strOldTC + strCounted

'<*************** wrapping it up nice and neatly ***************>
	objYTD.Sheets("Totals").Move objYTD.Sheets(1)
	objYTD.Save
	objCTBk.Close
	objYTD.Close
	
	strYTDPath = "c:\Phrasekeys\YTD Ticket Info.xlsx"
	strAlmaYTD = "L:\Joey_Thumb\JoeyScripts\shortkeys\YTd Ticket Info.xlsx"
	objFso.CopyFile strYTDPath, strAlmaYTD, TRUE
End Sub

WScript.Quit
