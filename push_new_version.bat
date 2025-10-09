@echo on
REM ================================
REM Git Commit + Tag + Push Script
REM ================================

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

REM Create tag
git tag %VERSION%

REM Push commits
git push origin main

REM Push tags
git push origin %VERSION%

echo.
echo Done! Pushed commit and tag %VERSION%.
pause