@echo off
setlocal
cd %~dp0

IF NOT EXIST bld\certpass.txt GOTO BADPASS
set /p certpass=<bld\certpass.txt

IF NOT EXIST bld\ubCert.cer util\makecert.exe -r -pe -sky signature -n CN=UsherBot -sv bld\ubCert.pvk bld\ubCert.cer
IF NOT EXIST bld\ubCert.pfx util\pvk2pfx.exe -pvk bld\ubCert.pvk -pi %certpass% -spc bld\ubCert.cer -pfx bld\ubCert.pfx -po %certpass% -f

IF "%1" == "" GOTO ERROR
util\signtool.exe sign /f bld\ubCert.pfx /p %certpass% %1
exit /b 0

:BADPASS
echo ERROR: Please save your certificate password to certpass.txt
exit /b 1

:ERROR
echo ERROR: Please provide the path to the exe file to sign
exit /b 1