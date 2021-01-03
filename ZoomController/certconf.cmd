@echo off
setlocal
cd %~dp0

IF NOT EXIST certpass.txt GOTO BADPASS
set /p certpass=<certpass.txt

IF NOT EXIST ubCert.cer util\makecert.exe -r -pe -sky signature -n CN=UsherBot -sv exclude\ubCert.pvk exclude\ubCert.cer
IF NOT EXIST ubCert.pfx util\pvk2pfx.exe -pvk exclude\ubCert.pvk -pi %certpass% -spc exclude\ubCert.cer -pfx exclude\ubCert.pfx -po %certpass% -f

IF "%1" == "" GOTO ERROR
util\signtool.exe sign /f exclude\ubCert.pfx /p %certpass% %1
exit /b 0

:BADPASS
echo ERROR: Please save your certificate password to certpass.txt
exit /b 1

:ERROR
echo ERROR: Please provide the path to the exe file to sign
exit /b 1