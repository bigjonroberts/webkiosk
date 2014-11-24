echo "Windows Kiosk Web Refresher requires Windows 7 or later with Chrome web browser pre-installed"

rem install python

set pyver=3.3.4
set pydir=%cd%\python3
set msi=python-%pyver%

rem replace these hash values if you use a different version!!!
if defined ProgramFiles(x86) (
    set msi=%msi%.amd64
)
set msi=%msi%.msi
set downloadurl="https://www.python.org/ftp/python/%pyver%/%msi%"
rem powershell -Command "(New-Object Net.WebClient).DownloadFile('%downloadurl%', '%msi%')"
rem mkdir python3
rem start /wait msiexec /i %msi% /qb TARGETDIR=%CD%\python3



rem install python windows extensions

if defined ProgramFiles(x86) (
	set pywin32=pywin32-219.win-amd64-py3.3.exe
) else (
	set pywin32=pywin32-219.win-py3.3.exe
)

set downloadurl="http://downloads.sourceforge.net/project/pywin32/pywin32/Build%20219/%pywin32%?ts=1413328340&use_mirror=iweb"
powershell -Command "Invoke-WebRequest -Uri %downloadurl% -OutFile %pywin32% -UserAgent [Microsoft.PowerShell.Commands.PSUserAgent]::FireFox"
rem rempowershell -Command "(New-Object Net.WebClient).DownloadFile('%downloadurl%', '%pywin32%')"
%pywin32%

rem get pip

rem %pydir%\python get-pip.py



rem install selenium

rem %pydir%\python -m pip install -U selenium



rem complete the rest of the installation in python ;-)

%pydir%\python install.py

goto :done

echo "Windows Kiosk Web Refresher requires Windows 7 or later with Chrome web browser pre-installed"

:bail
echo "There was an error. Installation incomplete.

:done

