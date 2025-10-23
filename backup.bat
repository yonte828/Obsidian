@echo off

cd /d "C:\Users\yuten\OneDrive\ƒhƒLƒ…ƒƒ“ƒg\Obsidian Vault"
if not %errorlevel%==0 goto error

if not exist .git goto norepository

git status --porcelain > status.tmp
for /f %%i in (status.tmp) do set HASCHANGE=1
del status.tmp

if not defined HASCHANGE goto nochange

echo Changes detected
git status --short
echo.

git add -A
if not %errorlevel%==0 goto error

git commit -m "Auto backup %date% %time:~0,5%"
if not %errorlevel%==0 goto error

git push
if not %errorlevel%==0 goto error

echo Backup completed
exit /b 0

:nochange
echo No changes detected
exit /b 0

:norepository
echo Error: Not a Git repository
exit /b 1

:error
echo Error occurred
exit /b 1