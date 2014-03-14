@echo off
:: Delete Windows temp, profiles, and temp
:: Joey Gauthier


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

::goto :eof

:run

echo dest = %dest% & %other%, %1

Echo Checking to see if %1 is alive...
ping -n 2 %1 | find /I "Reply" >NUL

If "%ErrorLevel%"=="1" echo %1 did not respond to ping while attempting Clean Temp Directories >>%log%&goto :EOF

del \\%1\%dest%\*.* /F /S /Q || echo %1 host del Failed while attempting Clean Temp Directories>>%log%
del \\%1\%other%\*.* /F /S /Q || echo %1 host del Failed while attempting Clean Temp Directories>>%log%
C:\Phrasekeys\Subs\DelProf2.exe /I /c:\\%1 /d:1


pause

>>%log%
:eof

