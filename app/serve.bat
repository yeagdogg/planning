@echo off
rem Planning Workbench launcher — pattern lifted from wfmodelingworkbench:
rem ONE port argument drives the port AND both websocket origins, so the
rem "403 Refusing websocket connection" trap is structurally impossible.
rem
rem     app\serve.bat          (port 8602)
rem     app\serve.bat 8611     (any port)
setlocal
set PORT=%1
if "%PORT%"=="" set PORT=8602
"%~dp0.venv\Scripts\python.exe" -m panel serve "%~dp0main.py" ^
    --address 127.0.0.1 --port %PORT% --show ^
    --allow-websocket-origin localhost:%PORT% ^
    --allow-websocket-origin 127.0.0.1:%PORT% ^
    --num-threads 4
endlocal
