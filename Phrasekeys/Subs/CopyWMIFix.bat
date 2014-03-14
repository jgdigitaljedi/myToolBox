@echo off
:: Variable
:: -------- Location of Log file
	set log=C:\Phrasekeys\replog.txt
	
:: Variable
:: -------- Location to copy file on each workstation...
	SET dest=C$\temp\

:: Variable
:: -------- Location of installation file on server...
	SET source=C:\Phrasekeys\Subs
	
:: Execution

:run

echo Source = %source%, dest = %dest%, CPUName = %2

Echo Checking to see if %2 is alive...
ping -n 2 %2 | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %2 did not respond to ping >>%log%&goto :EOF

xcopy %source%\%1 \\%2\%dest% /g /h /r /y || echo %2 msi Copy Failed >>%log%
pause

If "%1"=="WMIFix.bat" (
Echo %2 now has WMIFix.bat in the Temp directory. You will need to login or remote into machine to run it correctly.
pause
goto :eof)

If "%1"=="InstallCCM.bat" (
echo %2 now has the CCM installer copied to the Temp directory. Will now attempt to run it as soon as possible.
goto :schedule)

If "%1"=="WMICare.bat" (
xcopy %source%\wmc0100w.exe \\%2\%dest% /g /h /r /y || echo %2 msi Copy Failed >>%log%
echo %2 now has ATT WMI Care package copied down. Will now attempt to run a semi-silent installation as soon as possible.
goto :schedule)


:schedule
pause
SET CURRENTTIME=%TIME%
for /F "tokens=1 delims=:" %%h in ('echo %CURRENTTIME%') do (set /a HR=%%h)
for /F "tokens=2 delims=:" %%m in ('echo %CURRENTTIME%') do (set /a MIN=%%m + 2)

IF %MIN% GEQ 60 (
	SET /a MIN=%MIN%-60 
	SET /a HR=%HR%+1)

IF %HR% GTR 24 SET HR=00

IF %MIN% LEQ 9 (
	SET MIN=0%MIN%)

IF %HR% LEQ 9 (
	SET HR=0%HR%)
:: Execution

@echo off

at \\%2 %HR%:%MIN% C:\temp\%1
echo %2 now has %1 in the Temp directory and will run it in less than 2 minutes. Run the WMI Check in about 5 minutes to see if it was successful.
pause
:eof

