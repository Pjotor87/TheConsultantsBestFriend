@echo off

echo.
echo Deleting any existing old builds
echo.

rmdir "%~dp0\build" /S /Q
rmdir "%~dp0\dist" /S /Q
del "%~dp0\dist\convertdagboktotxt.exe"
del "%~dp0\convertdagboktotxt.exe"
del "%~dp0\convertdagboktotxt.spec"

echo.
echo Closing in
echo.
echo 	3
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	2
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	1
ping 1.1.1.1 -n 1 -w 2000 > nul
echo.

pyinstaller.exe --onefile --windowed convertdagboktotxt.py

echo.
echo Closing in
echo.
echo 	3
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	2
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	1
ping 1.1.1.1 -n 1 -w 2000 > nul
echo.

copy /y "%~dp0\dist\convertdagboktotxt.exe" "%~dp0\convertdagboktotxt.exe"

echo.
echo Closing in
echo.
echo 	3
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	2
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	1
ping 1.1.1.1 -n 1 -w 2000 > nul

echo.
echo Clearing build files
echo.

rmdir "%~dp0\build" /S /Q
rmdir "%~dp0\dist" /S /Q
del "%~dp0\dist\convertdagboktotxt.exe"
del "%~dp0\convertdagboktotxt.spec"

echo.
echo Closing in
echo.
echo 	3
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	2
ping 1.1.1.1 -n 1 -w 1000 > nul
echo 	1
ping 1.1.1.1 -n 1 -w 2000 > nul