@echo off
rem ============================================================
rem  AI Closer - share the manifest folder for Excel sideloading
rem
rem  Excel's trusted add-in catalog must be a UNC network share.
rem  A local path (C:\...\manifests) is accepted by the Trust
rem  Center dialog and then silently never lists the add-in.
rem
rem  Run this from an ADMINISTRATOR prompt: net share needs it.
rem ============================================================
setlocal

rem Resolve manifests\ relative to this script, so the repo can
rem live anywhere and be cloned under any name.
for %%I in ("%~dp0..\manifests") do set "FOLDER=%%~fI"

if not exist "%FOLDER%" mkdir "%FOLDER%"
if not exist "%FOLDER%\manifest.xml" (
  copy /Y "%~dp0..\manifest.xml" "%FOLDER%\manifest.xml" >nul
  if errorlevel 1 (
    echo Could not copy manifest.xml into "%FOLDER%".
    pause
    exit /b 1
  )
)

rem Re-sharing the same name fails; drop any previous one first.
net share OfficeAddins /delete >nul 2>&1

net share OfficeAddins="%FOLDER%" /grant:Everyone,READ
if errorlevel 1 (
  echo.
  echo Failed to create the share.
  echo Run this from an Administrator command prompt.
  pause
  exit /b 1
)

echo.
echo Shared "%FOLDER%"
echo.
echo Catalog path to paste into Excel:
echo     \\localhost\OfficeAddins
echo.
echo Next: File ^> Options ^> Trust Center ^> Trust Center Settings ^>
echo       Trusted Add-in Catalogs. Paste the path above, click Add
echo       catalog, tick "Show in Menu", OK, OK. Then close Excel
echo       completely and reopen it.
echo.
echo Finally: Insert ^> My Add-ins ^> Shared Folder tab ^> AI Closer.
pause
