@echo off
setlocal
set PYTHONPATH=C:\Python\Anaconda\3.7-64
call "%PYTHONPATH%\Scripts\activate.bat" "%PYTHONPATH%"
"%PYTHONPATH%\python.exe" "%~dp0ChatterBot.py" %*