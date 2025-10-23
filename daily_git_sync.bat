@echo off
setlocal enabledelayedexpansion

set "VAULT_PATH=C:\Users\yuten\OneDrive\ƒhƒLƒ…ƒƒ“ƒg\Obsidian Vault"

echo ====================================
echo Obsidian Vault Git Backup
echo ====================================
echo.

cd /d "%VAULT_PATH%"
if errorlevel 1 (
    echo Error: Cannot access vault directory
    pause
    exit /b 1
)

echo Vault: %VAULT_PATH%
echo.

if not exist ".git" (
    echo Error: Not a Git repository
    pause
    exit /b 1
)

git status --porcelain > temp_status.txt 2>&1
if errorlevel 1 (
    echo Error: git status failed
    del temp_status.txt
    pause
    exit /b 1
)

for /f %%a in (temp_status.txt) do set "HAS_CHANGES=1"
del temp_status.txt

if not defined HAS_CHANGES (
    echo No changes detected. Skipping backup.
    echo.
    pause
    exit /b 0
)

echo Changes detected:
echo.
git status --short
echo.

echo Staging changes...
git add -A
if errorlevel 1 (
    echo Error: git add failed
    pause
    exit /b 1
)

echo Committing...
git commit -m "Auto backup: %date% %time:~0,5%"
if errorlevel 1 (
    echo Error: git commit failed
    pause
    exit /b 1
)

echo Pushing to remote...
git push
if errorlevel 1 (
    echo Error: git push failed
    pause
    exit /b 1
)

echo.
echo ====================================
echo Backup completed!
echo ====================================
echo.
pause