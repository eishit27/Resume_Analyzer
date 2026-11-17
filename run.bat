@echo off
title Tkinter Application Launcher
echo Launching GUI Front-End Application...
python main.py
if errorlevel 1 (
    echo.
    echo Application exited with an error. Check Python setup.
    pause
)
