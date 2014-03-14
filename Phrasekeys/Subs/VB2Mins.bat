@Echo off

:: Modify variables as needed.

:: Variable
:: --------Location and name of Computer name list...	
::	SET cpus=C:\Phrasekeys\Subs\Comps.txt
@echo off
:: Variable
:: -------- Location of Log file
set log=C:\Phrasekeys\replog.txt

:: Setting execution time for <2 minutes from now	
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

::FOR /F %%i in (%cpus%) DO set cpuname=%%i&call :run

::goto :eof

:run

:: Variable
:: -------- Modify the time for excution on all AT commands.  Use military time.


Echo Checking to see if %1 is alive...

ping -n 2 %1 | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %1 did not respond to ping while trying to execute %2>>%log%&goto :EOF

at \\%1 %HR%:%MIN% wscript.exe //Nologo //B %2
echo %1 will execute %2 in less than 2 minutes from now
pause
:eof

