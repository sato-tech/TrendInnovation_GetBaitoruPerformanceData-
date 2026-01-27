@echo off
setlocal enabledelayedexpansion

REM Display initial message
echo.
echo ========================================
echo   Loop Test Execution
echo ========================================
echo.
echo [INFO] Starting batch file...
echo.

REM Set code page to UTF-8
chcp 65001 >nul 2>&1
title Test Loop
color 0A

REM Change to batch file directory
echo [INFO] Changing directory...
cd /d "%~dp0"
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] Failed to change directory.
    echo Error code: !ERRORLEVEL!
    echo Batch file path: %~dp0
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)

echo [INFO] Working directory: %CD%
echo.

REM Check package.json
echo [INFO] Checking package.json...
if not exist "package.json" (
    echo [ERROR] package.json not found.
    echo Current directory: %CD%
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)
echo [INFO] package.json found.
echo.

REM Check Node.js
echo [INFO] Checking Node.js...
where node >nul 2>&1
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] Node.js not found.
    echo Please install Node.js.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)
echo [INFO] Node.js found.
echo.

REM Check npm
echo [INFO] Checking npm...
where npm >nul 2>&1
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] npm not found.
    echo Please install npm.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)
echo [INFO] npm found.
echo.

REM Check node_modules
echo [INFO] Checking node_modules...
if exist "node_modules" (
    echo [INFO] node_modules found.
    echo.
    goto :continue_test
)

echo [WARNING] node_modules not found.
echo Install dependencies? (Y/N)
set /p INSTALL_CHOICE=
if /i "!INSTALL_CHOICE!"=="Y" (
    echo.
    echo [INFO] Running npm install...
    call npm install
    if !ERRORLEVEL! NEQ 0 (
        echo [ERROR] npm install failed.
        echo Error code: !ERRORLEVEL!
        echo.
        echo Press any key to exit...
        pause
        exit /b 1
    )
    echo [SUCCESS] Dependencies installed successfully.
    echo.
) else (
    echo [INFO] Test skipped.
    echo.
    echo Press any key to exit...
    pause
    exit /b 0
)

:continue_test

REM Ask for number of companies to process
echo ========================================
echo.
echo [INFO] How many companies do you want to process?
echo.
set /p LOOP_COUNT="Enter number of companies (1 or more): "

REM Validate input
if "!LOOP_COUNT!"=="" (
    echo [ERROR] No number entered.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)

REM Remove any spaces from input
set "LOOP_COUNT=!LOOP_COUNT: =!"

REM Check if input is a valid number using set /a
set /a TEST_NUM=!LOOP_COUNT! 2>nul
if !ERRORLEVEL! NEQ 0 (
    echo [ERROR] Invalid number entered: !LOOP_COUNT!
    echo Please enter a positive integer.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)

REM Verify the conversion worked correctly
set /a TEST_NUM=!LOOP_COUNT!
if !TEST_NUM! NEQ !LOOP_COUNT! (
    echo [ERROR] Invalid number entered: !LOOP_COUNT!
    echo Please enter a positive integer.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)

REM Check if number is greater than 0
if !LOOP_COUNT! LEQ 0 (
    echo [ERROR] Number must be greater than 0.
    echo.
    echo Press any key to exit...
    pause
    exit /b 1
)

echo.
echo [INFO] Processing !LOOP_COUNT! companies...
echo.
echo ========================================
echo.

REM Execute npm run test:loop with the count as argument
call npm run test:loop !LOOP_COUNT!
set EXIT_CODE=!ERRORLEVEL!

echo.
echo ========================================
if !EXIT_CODE! EQU 0 (
    echo [SUCCESS] Test completed successfully.
    color 0A
) else (
    echo [ERROR] Test failed. Exit code: !EXIT_CODE!
    color 0C
    echo.
    echo Troubleshooting:
    echo 1. Check if .env file is configured correctly
    echo 2. Check if credentials.json exists in project root
    echo 3. Check if input Excel file exists
    echo 4. Check error messages above
)
echo ========================================
echo.
echo Press any key to exit...
pause
