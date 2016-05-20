@echo off
mode con:cols=90 lines=50

powershell Set-ExecutionPolicy Unrestricted
powershell .\powershell.ps1
powershell Set-ExecutionPolicy Default