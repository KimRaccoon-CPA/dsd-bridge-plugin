@echo off
rem dsd-bridge 실행기 - setup.ps1 이 만든 venv(%LOCALAPPDATA%\dsd-bridge\.venv)의 파이썬으로 이 폴더의 dsd_bridge.py 를 실행한다
set PYTHONUTF8=1
"%LOCALAPPDATA%\dsd-bridge\.venv\Scripts\python.exe" "%~dp0dsd_bridge.py" %*
