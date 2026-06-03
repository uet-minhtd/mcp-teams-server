@echo off
title MCP Teams Server (Dev)
echo ========================================
echo   MCP Teams Server - Dev Mode
echo ========================================
echo.

:: Check if node_modules exists
if not exist "node_modules" (
    echo Installing dependencies...
    call npm install
    if errorlevel 1 (
        echo ERROR: npm install failed!
        pause
        exit /b 1
    )
)

:: Start in dev mode with tsx (auto-reload)
echo Starting in dev mode...
echo.
call npm run dev
pause
