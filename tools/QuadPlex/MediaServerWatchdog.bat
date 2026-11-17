@echo off
rem ============================================================================
rem  MediaServerWatchdog.bat
rem  Makes sure Plex, Emby and Jellyfin are running on QuadPlex. If one is not,
rem  it starts it. Meant to run from Task Scheduler every few minutes.
rem
rem  Each server is checked like this:
rem    1. If a Windows service with that name exists, start it if it isn't running.
rem    2. If there is no service, look for the program's process and launch the
rem       program if it isn't running.
rem
rem  Find the real service names on QuadPlex with:
rem    sc query type= service state= all | findstr /i "plex emby jellyfin"
rem  (the line that says SERVICE_NAME is the one to use below)
rem
rem  Only actions and problems are logged, so the log stays small.
rem ============================================================================
setlocal EnableExtensions

set "LOG=%~dp0MediaServerWatchdog.log"

rem ---- Settings: edit these to match QuadPlex --------------------------------
rem  Name           Service name       Process name            Program path
call :Check "Plex"     "PlexService"   "Plex Media Server.exe" "%ProgramFiles%\Plex\Plex Media Server\Plex Media Server.exe"
call :Check "Emby"     "EmbyServer"    "EmbyServer.exe"        "%AppData%\Emby-Server\system\EmbyServer.exe"
call :Check "Jellyfin" "JellyfinServer" "jellyfin.exe"         "%ProgramFiles%\Jellyfin\Server\jellyfin.exe"
rem ----------------------------------------------------------------------------

endlocal
exit /b 0


rem ============================================================================
rem  :Check  Name  ServiceName  ProcessName  ProgramPath
rem ============================================================================
:Check
set "NAME=%~1"
set "SVC=%~2"
set "PROC=%~3"
set "EXE=%~4"

rem --- Installed as a Windows service? ---
if not "%SVC%"=="" (
    sc query "%SVC%" >nul 2>&1
    if not errorlevel 1 goto :CheckService
)

rem --- Not a service: check for the running program instead ---
tasklist /FI "IMAGENAME eq %PROC%" /NH 2>nul | find /I "%PROC%" >nul
if not errorlevel 1 exit /b 0

if not exist "%EXE%" (
    call :Log "%NAME%: not running, and no service '%SVC%' or program '%EXE%' was found. Check the settings at the top of the script."
    exit /b 1
)

call :Log "%NAME%: not running. Launching %EXE%"
start "" "%EXE%"
exit /b 0

:CheckService
sc query "%SVC%" | find "RUNNING" >nul
if not errorlevel 1 exit /b 0

call :Log "%NAME%: service '%SVC%' is not running. Starting it."
net start "%SVC%" >nul 2>&1
if errorlevel 1 (
    call :Log "%NAME%: FAILED to start service '%SVC%'. Is the task running with admin rights?"
    exit /b 1
)
call :Log "%NAME%: service '%SVC%' started."
exit /b 0


:Log
echo [%date% %time%] %~1
>>"%LOG%" echo [%date% %time%] %~1
exit /b 0
