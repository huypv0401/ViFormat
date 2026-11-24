@echo off
REM Go ViFormat khoi Word tren Windows.
setlocal
tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo Hay dong Word truoc khi go, roi chay lai file nay.
  exit /b 1
)
del /F /Q "%AppData%\Microsoft\Word\STARTUP\ViFormat.dotm" 2>nul
del /F /Q "%AppData%\Microsoft\Word\STARTUP\~$Format.dotm" 2>nul
reg delete "HKCU\Software\VB and VBA Program Settings\ViFormat" /f >nul 2>&1
echo Da go ViFormat. Mo lai Word, tab ViFormat se khong con.
endlocal
