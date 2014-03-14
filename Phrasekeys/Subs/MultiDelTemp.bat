@echo off

:: Modify variables as needed.

:: Variable
:: --------Location and name of Computer name list...	
	SET cpus=%1

:: Variable
:: -------- Location of Log file
	set log=c:\Phrasekeys\replog.txt
	
:: Variable
:: -------- Location of temp directory...
	SET dest=c$\windows\temp

:: Variable
::--------- Location of 2nd directory...
	SET other=c$\temp

:: Execution

FOR /F "tokens=*" %%a in (%cpus%) DO set cpuname=%%a&call :run

goto :eof

:run

echo dest = %dest% and %other%, CPUName = %cpuname%

Echo Checking to see if %cpuname% is alive...
ping -n 2 %cpuname% | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %cpuname% did not respond to ping >>%log%&goto :EOF

del \\%cpuname%\%dest%\*.* /F /S /Q || echo %cpuname% host del Failed >>%log%
del \\%cpuname%\%other%\*.* /F /S /Q || echo %cpuname% host del Failed >>%log%
C:\Phrasekeys\Subs\DelProf2.exe /I /c:\\%cpuname% /d:1

:eof

