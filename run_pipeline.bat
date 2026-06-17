@echo off
cd /d "%~dp0"
echo =======================================================================
echo          University Admission Status Automation Pipeline (2027)
echo =======================================================================
echo.
python run_pipeline.py
echo.
echo =======================================================================
echo Pipeline finished.
pause
