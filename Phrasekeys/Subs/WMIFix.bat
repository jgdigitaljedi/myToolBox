@echo on

:: BatchGotAdmin
:-------------------------------------
::REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------

::taskkill /f /im explorer.exe
cd /d %windir%\System32\Wbem
net stop winmgmt

sc sdset winmgmt D:(A;;CCDCLCSWRPWPDTLOCRRC;;;BA)(A;;CCLCSWLOCRRC;;;AU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;DA)(A;;CCDCLCSWRPWPDTLOCRRC;;;PU)(A;;CCDCLCSWRPWPDTLOCRSDRCWDWO;;;SY)

REG IMPORT %windir%\WBEM.reg

c:
cd %windir%\system32\wbem
rd /S /Q repository
regsvr32 /s %systemroot%\system32\scecli.dll
regsvr32 /s %systemroot%\system32\userenv.dll
mofcomp cimwin32.mof
mofcomp cimwin32.mfl
mofcomp rsop.mof
mofcomp rsop.mfl

winmgmt /clearadap
winmgmt /kill
winmgmt /unregserver
winmgmt /regserver
winmgmt /resyncperf

del %windir%\System32\Wbem\Repository /Q
del %windir%\System32\Wbem\AutoRecover /Q

for /f %%s in (’dir /b /s *.dll’) do regsvr32 /s %%s
for /f %%s in (’dir /b *.mof’) do mofcomp %%s
for /f %%s in (’dir /b *.mfl’) do mofcomp %%s
mofcomp -n:root\cimv2\applications\exchange wbemcons.mof
mofcomp -n:root\cimv2\applications\exchange smtpcons.mof
mofcomp exmgmt.mof
mofcomp exwmi.mof

wmiadap.exe /Regsvr32
wmiapsrv.exe /Regsvr32
wmiprvse.exe /Regsvr32

net start winmgmt

c:\windows\system32\shutdown -r -f -t 00
