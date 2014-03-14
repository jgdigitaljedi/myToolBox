@Echo off

:: Clear local Temporary Internet Files and History
Echo.
Echo The following errors are normal so don't panic:
Echo Press Ctrl+C to abort or any other key to continue . . .
PAUSE>NUL

del /S /Q "C:\Temp\*.*"
del /S /Q "C:\Windows\Temp\*.*"
pause
Echo Deleting Java caches.

::This part was kind of specific to my client. Change as needed.
"%programfiles%\Java\jre1.6.0_41\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.6.0_35\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.6.0_26\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.5.0_22\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.6.0_20\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\j2re1.4.2_19\javaws\javaws" -Xclearcache -uninstall

pause
