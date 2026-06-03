@echo off
title MCP Teams Server
echo ========================================
echo   MCP Teams Server - Starting...
echo ========================================
echo.

:: Check if node_modules exists
if not exist "node_modules" (
    echo [1/3] Installing dependencies...
    call npm install
    if errorlevel 1 (
        echo ERROR: npm install failed!
        pause
        exit /b 1
    )
) else (
    echo [1/3] Dependencies already installed. Skipping.
)

:: Build TypeScript
echo [2/3] Building TypeScript...
call npm run build
if errorlevel 1 (
    echo ERROR: Build failed!
    pause
    exit /b 1
)

:: Start the server in HTTP mode
echo [3/3] Starting server (HTTP mode)...
echo.
node dist/index.js --http
pause
