@echo off
cd /d "%~dp0"
title EPC Automation - Split
echo.
echo  ============================================
echo   EPC AUTOMATION - GENERATE PM FILES
echo   Reliance Retail | EPC Department
echo  ============================================
echo.
python split.py
