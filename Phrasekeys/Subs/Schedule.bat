@Echo off

:: Modify variables as needed.

:: Variable
:: --------Location and name of Computer name list...	
::	SET cpus=C:\Phrasekeys\Subs\Comps.txt
@echo off
:: Variable
:: -------- Location of Log file
set log=C:\Phrasekeys\replog.txt

:: Execution

@echo off

::FOR /F %%i in (%cpus%) DO set cpuname=%%i&call :run

::goto :eof

:run

:: Variable
:: -------- Modify the time for excution on all AT commands.  Use military time.


Echo Checking to see if %1 is alive...

ping -n 2 %1 | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %1 did not respond to ping while trying to schedule task %3>>%log%&goto :EOF

at \\%1 %2 %3
echo %3 will execute on %1 at %2
pause
:eof

