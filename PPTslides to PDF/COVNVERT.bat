@echo off
SET Directory=%~dp0
SET ps1path=%Directory%converter.ps1
PowerShell -NoProfile -ExecutionPolicy Bypass -Command "& '%ps1path%'";

