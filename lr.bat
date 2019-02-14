:: TR3 - re-write to be effective on newer windows versions.
:: Incident Response Collection
:: converting what I can from invoke-ir powershell to batch...
:: http://www.invoke-ir.com/
::
:: Jim Price
:: 2/9/19
::
::	Script start...
@echo off
::	Change Directory to the script's one, since Stupid windows will execute at "\windows\system32" when chosen to run as administrator...
cd /d "%~d0%~p0"

::Get Date and Time
For /F "tokens=1,2,3,4 delims=/ " %%A in ('Date /t') do @(
Set FullDate="%%D-%%C-%%B"
)

For /F "tokens=1,2,3 delims=: " %%A in ('time /t') do @(
Set FullTime="%%A-%%B"
)

set dirname="%computername%_%FullDate%_%FullTime%"
mkdir "%dirname%"
cd "%dirname%"
set var=%computername%.txt


type ..\commands.txt >> %var%

for /f "tokens=*" %%a in (..\commands.txt) do @(
echo. >> %var%
echo. >> %var%
echo ======================================== >> %var%
echo ======================================== >> %var%
echo %%a >> %var%
echo ======================================== >> %var%
echo ======================================== >> %var%
cmd /c "%%a" >> %var%
)

::extract tools - tools used in tools.txt
"c:\program files (x86)"\tanium\tanium clienttools\StdUtils\7za.exe e "c:\program files (x86)\tanium\tanium clienttools\tr3\deps.zip"  -o"\%dirname%\" > %dirname%\deps.log
::
cd ..
cd Tools
autorunsc.exe -a -f -c >> ..\%dirname%\autoruns.csv
handle.exe -a -u >> ..\%dirname%\handles.txt
HiJackThis.EXE /silentautolog
move hijackthis.log ..\%dirname%\
listdlls.exe -accepteula >> ..\%dirname%\dlls.txt
psinfo -accepteula -s -h -d >> ..\%dirname%\psinfo.txt
pslist -accepteula -t >> ..\%dirname%\pslist.txt
psgetsid accepteula >> ..\%dirname%\psgetsid.txt
psloggedon.exe -accepteula >> ..\%dirname%\psloggedon.txt
Tcpvcon.exe -accepteula -a -c -n >> ..\%dirname%\tcpvcon.txt
psfile.exe -accepteula >> ..\%dirname%\openfiles.txt
psloglist.exe -accepteula -g %computername% >> ..\%dirname\
procdump.exe -accepteula -ma >> ..\%dirname%\processes
streams.exe -accepteula -s c:\ >> ..\%dirname%\streams.txt


::Get system info
@echo off
echo +++++Incident Response Collection+++++ >> %computername%.txt
echo +++++Date and Time+++++ >> %computername%.txt
date /t >> %computername%.txt
time /t >> %computername%.txt
echo +++++System Information: Time Zone, Installed Software, OS Version, Uptime, File System+++++ >> %computername%.txt
systeminfo >> %computername%.txt
echo +++++User Accounts+++++ >> %computername%.txt
net user >> %computername%.txt
echo +++++Groups+++++ >> %computername%.txt
net localgroup >> %computername%.txt
echo +++++Network Interfaces+++++ >> %computername%.txt
ipconfig /all >> %computername%.txt
echo +++++Routing Table+++++ >> %computername%.txt
route print >> %computername%.txt
echo +++++ARP Table+++++ >> %computername%.txt
arp -a >> %computername%.txt
echo +++++DNS Cache+++++ >> %computername%.txt
ipconfig /displaydns >> %computername%.txt
echo +++++Network Connections+++++ >> %computername%.txt
netstat -abn >> %computername%.txt
netstat -ano >> %computername%.txt
