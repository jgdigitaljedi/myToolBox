@echo on
sc config msiserver start= demand
echo Stopping Windows Installer Service
net stop msiserver
echo Unregistering Windows Installer
msiexec /unregister
echo Re-Registering Windows Installer
msiexec /regserver
echo Restarting Windows Installer Service
net start msiserver
