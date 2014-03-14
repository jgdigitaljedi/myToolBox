'Microsoft (R) Windows Script Host Version 5.8
'Copyright (C) Microsoft Corporation. All rights reserved.

'******* error Control *******
On error Resume next
WScript.Timeout = 180
msg=""
intFails = 0
intCCM = 0


'CN = trim(inputbox("Enter a Workstation Name or IP Address:" & vbCrLf & "Leave blank for local inventory!", "COPS WMI Evaluation Utility"))
CN = WScript.Arguments.Item(0)
IF CN = "" or ISNULL(CN) THEN
	CN = WScript.CreateObject("Wscript.Network").computername
ELSE
	Set objPing = GetObject("winmgmts:").Get("Win32_PingStatus.Address='" & CN & "'")
	With objPing
		IF isnull(.StatusCode) or .StatusCode > 0 THEN
			wscript.echo "*****  WARNGING ***** " & vbCR& "Could not ping [ " & ucase(CN) & " ]" & vbCR &" Please check your INPUT!"
			wscript.quit
		END IF
	End With
		Set FSO = CreateObject("Scripting.FileSystemObject") 
		If FSO.FolderExists("\\" & CN & "\admin$\system32") Then  
			strCheckAccess = 0  ' user has access 
		else 
			strCheckAccess = 1  ' user does not have access
			wscript.echo "*****  WARNGING ***** " & vbCR& "Could not access [ " & ucase(CN) & " ]" & vbCR &" You must have administrative rights to the machine!" 
			wscript.quit
		End If
END IF

msg = "*****COPS WMI Evaluation Utility*****" & vbCr & vbCr & "**** PC Settings: **** " & vbCrLf
msg = msg & "PC Name - " & CN & " " & vbCrLf


Set ObjSrv = GetObject("winmgmts:\\" & CN & "\root\cimv2")
for each Object in ObjSrv.ExecQuery("select SerialNumber from Win32_SystemEnclosure",,48)
	SN1 = ucase(trim(Object.SerialNumber))
next

for each Object in ObjSrv.ExecQuery("select SerialNumber from Win32_BIOS",,48)
	SN2 = ucase(trim(Object.SerialNumber))
Next

IF SN1 = SN2 THEN
	SN = SN1
ELSE
	IF ISNULL(SN1) or SN1 = "" THEN
		SN = SN2
	ELSEIF ISNULL(SN2) or SN2 = "" THEN
		SN = SN1
	ELSE
		SN = SN1 & " - or - " & SN2
	END IF
END IF
msg = msg & "PC Serial Number - " & SN & " " & vbCrLf


msg = msg  & vbCrLf & "**** WMI Checks Start: **** " & vbCrLf

msg2 = "Win32_Process"
msg3 = "Caption"
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = object.Caption
Next
set nothing = ReportBack(msg1,msg2)



msg2 = "Win32_ComputerSystem"
msg3 = "Caption"
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = object.Caption
Next
set nothing = ReportBack(msg1,msg2)


msg2 = "Win32_LogicalDisk"
msg3 = "Caption"
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = object.Caption
Next
set nothing = ReportBack(msg1,msg2)


msg2 = "Win32_NetworkAdapterConfiguration"
msg3 = "Caption"
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = object.Caption
Next
set nothing = ReportBack(msg1,msg2)


msg2 = "Win32_BIOS"
msg3 = "Caption"
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = object.Caption
Next
set nothing = ReportBack(msg1,msg2)

strValue = ""
HKEY_LOCAL_MACHINE = &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & CN & "\root\default:StdRegProv")
strKeyPath = "SOFTWARE\Microsoft\Windows NT\CurrentVersion"
strValueName  = "ProductId"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
set nothing = ReportBack(strValue,"StdRegProv")
IF strValue <> "" Then strStdRegProv = "OK"



msg2 = "CCM_Client"
msg3 = "ClientId"
Set ObjSrv = GetObject("winmgmts:\\" & CN & "\root\ccm")
for each Object in ObjSrv.ExecQuery("Select " & msg3 & " from " & msg2 & " ",,48)
	msg1 = Object.ClientId
next
set nothing = ReportBack(msg1,msg2)


strValue = ""
HKEY_LOCAL_MACHINE = &H80000002
Set oReg=GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & CN & "\root\default:StdRegProv")
strKeyPath = "Software\AT&T\WMIcare"
strValueName  = "WMIStatus"
oReg.GetStringValue HKEY_LOCAL_MACHINE,strKeyPath,strValueName,strValue
IF strValue = "" or ISNULL(strValue) = TRUE THEN strValue = "Not Installed"
IF strValue = "Not Installed" and strStdRegProv <> "OK" THEN strValue = "Not Installed or WMI Registry is corrupt!"
msg = msg & " " & strValue & " - WMICare Status" & vbCrLf


title1 = "COPS WMI Evaluation - " & DATE
ReadRep = MsgBox(msg , vbOkOnly+vbINFO, title1)
	If ReadRep = vbOk Then
		If intFails <> 0 Then
			FixIT = MsgBox("There were " & intFails & " WMI error(s) found. Would you like to try and fix it now?", vbYesNo+vbExclamation, "Fix WMI errors?")
				If FixIT = vbYes Then
					var1 = "WMIFix.bat"
					var2 = CN
					Set objShell = CreateObject("WScript.Shell")
					strCommando = "cmd.exe /c C:\Phrasekeys\Subs\CopyWMIFix.bat"
					objShell.Run strCommando & " " & var1 & " " & var2
				Else
					'WScript.Quit
				End If
		Else
			'WScript.Quit
		End If
			If intCCM = 1 Then
				InstCCM = MsgBox("Target machine does not have the CCM Client installed. Would you like to install it now?", vbYesNo+vbExclamation, "Install SCCM?")
					If InstCCM = vbYes Then
						var1 = "InstallCCM.bat"
						var2 = CN
						Set objShell = CreateObject("WScript.Shell")
						strCommanda = "cmd.exe /c C:\Phrasekeys\Subs\CopyWMIFix.bat"
						objShell.Run strCommanda & " " & var1 & " " & var2
					Else
						'WScript.Quit
					End If
			Else
				'WScript.Quit
			End If
		If strValue = "Not Installed or WMI Registry is corrupt!" Then
		
			Set objWMIService = GetObject("winmgmts:\\" & CN & "\root\cimv2")
			Set colOperatingSystem = objWMIService.ExecQuery("Select * from Win32_OperatingSystem")
				For Each objOperatingSystem in colOperatingSystem
					Version = objOperatingSystem.Version
				Next
				
			If InStr(1, Version, "6.1") > 0 Then
				WScript.Quit
			Else		
				AttCare = MsgBox("AT&T's WMI Care package is not installed on the target machine. Would you like to install that now? (This package is only for Windows XP machines)",_
				vbYesNo+vbExclamation, "Install WMI Care?")
					If AttCare = vbYes Then
						var1 = "WMICare.bat"
						var2 = CN
						Set objShell = CreateObject("WScript.Shell")
						strCommande = "cmd.exe /c C:\Phrasekeys\Subs\CopyWMIFix.bat"
						objShell.Run strCommande & " " & var1 & " " & var2
					Else
						WScript.Quit
					End If
			End If
		Else
			WScript.Quit
		End If
	End If

'******* Functions Area *******
Function ReportBack(msg1,msg2)
	IF msg1 <> "" Then
	  msg = msg & " PASSED  - "& msg2 & vbCrLf
	Else
	  msg = msg & "*FAILED* - "& msg2 & vbCrLf
	  intFails = intFails + 1
		If msg2 ="CCM_Client" Then
			intFails = intFails - 1
			intCCM = intCCM + 1
		End If
	END IF
	msg1 = ""
	msg2 = ""
	msg3 = ""
End Function
