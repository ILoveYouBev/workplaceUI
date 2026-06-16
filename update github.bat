@echo off
title CSR Tool - Auto Git Push

cd /d "%~dp0"

echo ============================
echo CSR TOOL AUTO GIT PUSH
echo ============================

if not exist .git (
    echo ERROR: Not a git repository
    pause
    exit /b
)

echo.
echo Staging changes...
git add .

echo.
git status

echo.
set /p msg=Commit message: 
if "%msg%"=="" set msg=Auto update

echo.
echo Committing...
git commit -m "%msg%"

echo.
echo Pushing...
git push origin main

echo.
echo DONE!
pause