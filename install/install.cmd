@echo off
REM Cai ViFormat cho Word tren Windows: copy ViFormat.dotm vao thu muc STARTUP.
setlocal
set "SRC=%~dp0ViFormat.dotm"
if not exist "%SRC%" set "SRC=%~dp0..\ViFormat.dotm"
if not exist "%SRC%" (
  echo Khong thay ViFormat.dotm
  exit /b 1
)
set "DEST=%AppData%\Microsoft\Word\STARTUP"
if not exist "%DEST%" mkdir "%DEST%"

tasklist /FI "IMAGENAME eq WINWORD.EXE" 2>nul | find /I "WINWORD.EXE" >nul
if not errorlevel 1 (
  echo Hay dong Word truoc khi cai, roi chay lai file nay.
  exit /b 1
)

copy /Y "%SRC%" "%DEST%\ViFormat.dotm" >nul
powershell -NoProfile -Command "Unblock-File -Path '%DEST%\ViFormat.dotm'" >nul 2>&1
echo Da cai ViFormat vao %DEST%
echo Mo Word, tab ViFormat se xuat hien. Neu Word bao macro bi chan, vao
echo File ^> Options ^> Trust Center ^> Trusted Locations va them thu muc STARTUP neu can.
endlocal
