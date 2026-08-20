@echo off
rem Prezlab PPT QC: LAN server launcher (Windows box).
rem Runs from any location; logs to data\lan-server.log. A second copy
rem exits quietly when the port is already served.
cd /d "%~dp0.."
.venv\Scripts\python.exe -m qc.web --lan --port 8000 >> data\lan-server.log 2>&1
