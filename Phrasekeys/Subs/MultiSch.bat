@Echo off

:: Modify variables as needed.

:: Variable
:: --------Location and name of Computer name list...	
	SET cpus=%1

:: Variable
:: -------- Location of Log file
	set log=c:\Phrasekeys\replog.txt
	
:: Variable
:: -------- pull variables from app before computer name list is opened
	Set task=%2
	Set clock=%3
	
:: Execution

@echo off

FOR /F "tokens=*" %%a in (%cpus%) DO set cpuname=%%a&call :run

goto :eof

:run

Echo Checking to see if %cpuname% is alive...

ping -n 2 %cpuname% | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %cpuname% did not respond to ping >>%log%&goto :EOF

at \\%cpuname% %task% %clock%

pause

:eof

