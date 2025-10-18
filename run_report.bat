
@echo off
setlocal enabledelayedexpansion

REM === Adjust PYTHON_HOME if you use a virtualenv ===
REM Example using system Python on PATH:
set PYTHON_EXE=python

REM Change to script directory
cd /d %~dp0

REM Run
%PYTHON_EXE% auto_report.py 1>> run.log 2>&1
