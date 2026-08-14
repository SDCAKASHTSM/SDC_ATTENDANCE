@echo off
title SDC Attendance Server

cd /d C:\Users\DELL\SDC_ATTENDANCE

call venv\Scripts\activate.bat

start "" http://127.0.0.1:5000

python app.py

pause