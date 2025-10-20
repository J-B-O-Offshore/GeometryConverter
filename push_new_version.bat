@echo off
REM =========================================
REM Git Commit + Tag + Push Script (Retag)
REM =========================================

REM Check if version number is provided as argument
IF "%~1"=="" (
    echo Usage: %~nx0 VERSION_NUMBER
    echo Example: %~nx0 v1.0.0
    exit /b 1
)

SET VERSION=%~1

echo.
echo =================================
echo   Committing and tagging version %VERSION%
echo =================================
echo.

REM Add all changes
git add .

REM Commit with message
git commit -m "Release %VERSION%"

REM Check if tag already exists
git rev-parse %VERSION% >nul 2>&1
IF %ERRORLEVEL%==0 (
    echo Tag %VERSION% already exists. Deleting old tag...
    git tag -d %VERSION%
    git push origin :refs/tags/%VERSION%
)

REM Create new tag
git tag %VERSION%

REM Push commits
git push origin main

REM Push new tag
git push origin %VERSION%

echo.
echo Done! Committed and retagged version %VERSION%.
pause