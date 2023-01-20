@echo off
title Visualizer Update Script
mode con cols=80 lines=40
:loop
cls
pwsh -NoProfileLoadTime -nol -ep Bypass -f VIZUpdater.ps1
goto :loop
