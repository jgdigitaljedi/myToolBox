@ECHO OFF 
Echo.
ECHO IT IS NOT RECCOMENDED TO RUN THIS MORE THAN ONCE A DAY!
Echo.
Echo ClearCache for XP
Echo Removes temporary files FROM ALL PROFILES.
ECHO Removes Java cache for currently logged in profile.
ECHO Uninstalls Java Web applications
ECHO Deletes contents of c:\Temp directory.
Echo.
Echo Temporary files will be removed from the following locations:
Echo.
Echo   1. [Profile]\Local Settings\Temporary Internet Files
Echo   2. [Profile]\Local Settings\Temp
Echo   3. [Profile]\Cookies
Echo   4. [Profile]\Recent
Echo   5. [SystemRoot]\Temp
ECHO   6. [Profile]\Application Data\Sun\Java\Deployment\cache
Echo.
Echo This batch file will close automatically when finished.
Echo.
Echo Press Ctrl+C to abort or any other key to continue . . .
Echo.
PAUSE > NUL
:: Clear local Temporary Internet Files and History
Echo.
Echo The following errors are normal so don't panic:
RD /S /Q "%Userprofile%\Local Settings\Temporary Internet Files"

::Delete Java cache
Echo Deleting Java caches and uninstalling applets
"%programfiles%\Java\jre1.6.0_35\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.6.0_26\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.5.0_22\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\jre1.6.0_20\bin\javaws" -Xclearcache -uninstall
"%programfiles%\Java\j2re1.4.2_19\javaws\javaws" -Xclearcache -uninstall
:: Clear all other temporary cache files in all profiles
SET SRC1=C:\Documents and Settings
SET SRC2=Local Settings\Temporary Internet Files
::SET SRC3=Local Settings\History
SET SRC4=Local Settings\Temp
SET SRC5=Cookies
SET SRC6=Recent
Echo.
Echo About to delete files from Internet Explorer "Temporary Internet files"
Echo This may take a few minutes.  Please wait...
Echo The following error is normal:
FOR /D %%X IN ("%SRC1%\*") DO FOR /D %%Y IN ("%%X\%SRC2%\*.*") DO RMDIR /S /Q "%%Y"
Echo.
Echo About to delete files from Internet Explorer "History"
Echo This may take a few minutes.  Please wait...
Echo The following error is normal:
Echo About to delete files from "Local settings\temp"
FOR /D %%X IN ("%SRC1%\*") DO FOR /D %%Y IN ("%%X\%SRC4%\*.*") DO RMDIR  /S /Q "%%Y"
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC4%\*.*") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "Cookies"
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC5%\*.*") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "Recent" i.e. what appears in Start/Documents/My Documents
FOR /D %%X IN ("%SRC1%\*") DO FOR  %%Y IN ("%%X\%SRC6%\*.lnk") DO DEL /F /S /Q "%%Y"
Echo About to delete files from "[SystemRoot]\Temp"
CD /D %SystemRoot%\Temp
DEL /F /Q *.*
RD /S /Q "%SystemRoot%\Temp"
CD /D %SystemRoot%
MD Temp
del C:\Windows\Temp\*.* /F /S /Q 
:: Done!
Echo Done!
pause
EXIT
