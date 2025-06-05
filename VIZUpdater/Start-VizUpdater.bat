@echo off
title Visualizer Update Script
mode con cols=80 lines=40
:loop
cls
pwsh -nop -nol -ep Bypass -f "%~dp0\VIZUpdater.ps1"
goto :loop
