@echo off
title Activate Microsoft Office VL (2010, 2016, 2019, 2021)
cls
echo ============================================================================
echo # Project: Activate Office Volume-License editions via KMS
echo ============================================================================

:: ──────────────────────────────────────────────────────────────────────────
:: 1) Find ospp.vbs and detect Office version folder
:: ──────────────────────────────────────────────────────────────────────────
set "OSPPDir="
for %%P in (
  "%ProgramFiles%\Microsoft Office\Office16"
  "%ProgramFiles(x86)%\Microsoft Office\Office16"
  "%ProgramFiles%\Microsoft Office\Office15"
  "%ProgramFiles(x86)%\Microsoft Office\Office15"
  "%ProgramFiles%\Microsoft Office\Office14"
  "%ProgramFiles(x86)%\Microsoft Office\Office14"
) do if exist "%%~P\ospp.vbs" set "OSPPDir=%%~P"

if not defined OSPPDir (
  echo ERROR: No Office 2010/2016/2019/2021 installation found.
  pause
  exit /b
)
cd /d "%OSPPDir%"

:: ──────────────────────────────────────────────────────────────────────────
:: 2) Set license-folder under Office\root\Licenses16
:: ──────────────────────────────────────────────────────────────────────────
set "Lic16=%OSPPDir%\..\root\Licenses16"
set "OfficeVer="

:: ──────────────────────────────────────────────────────────────────────────
:: 3) Detect 2021, then 2019, then 2016 .xrm-ms files
:: ──────────────────────────────────────────────────────────────────────────
if exist "%Lic16%\ProPlus2021VL*.xrm-ms" (
  set "OfficeVer=2021"
  set "GVLK=FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"
) else if exist "%Lic16%\ProPlus2019VL*.xrm-ms" (
  set "OfficeVer=2019"
  set "GVLK=NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"
) else if exist "%Lic16%\ProPlus2016VL*.xrm-ms" (
  set "OfficeVer=2016"
  set "GVLK=XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"
)

if not defined OfficeVer (
  echo ERROR: No Office 2016/2019/2021 VL license files found under:
  echo    %Lic16%
  pause
  exit /b
)

echo Detected Office %OfficeVer% VL.
echo License folder: %Lic16%
echo GVLK to install: %GVLK%
echo.

:: ──────────────────────────────────────────────────────────────────────────
:: 4) Install all matching .xrm-ms license files
:: ──────────────────────────────────────────────────────────────────────────
echo Installing VL license file(s)...
for %%F in ("%Lic16%\*%OfficeVer%VL*.xrm-ms") do (
  cscript //nologo ospp.vbs /inslic:"%%~F" >nul
)

:: ──────────────────────────────────────────────────────────────────────────
:: 5) Inject the GVLK
:: ──────────────────────────────────────────────────────────────────────────
echo Installing GVLK…
cscript //nologo ospp.vbs /unpkey:6MWKP >nul
cscript //nologo ospp.vbs /inpkey:%GVLK% >nul

:: ──────────────────────────────────────────────────────────────────────────
:: 6) Prepare KMS activation
:: ──────────────────────────────────────────────────────────────────────────
echo.
echo Preparing KMS activation…
cscript //nologo slmgr.vbs /ckms      >nul
cscript //nologo ospp.vbs /setprt:1688 >nul

:: ──────────────────────────────────────────────────────────────────────────
:: 7) Loop through public KMS hosts
:: ──────────────────────────────────────────────────────────────────────────
set /a attempt=1
:TryKMS
if %attempt%==1 set "KMSHost=kms.03k.org"
if %attempt%==2 set "KMSHost=kms.loli.beer"
if %attempt%==3 set "KMSHost=kms.digiboy.ir"
if %attempt%==4 goto KMSFailed

echo Attempt #%attempt%: contacting %KMSHost% …
cscript //nologo ospp.vbs /sethst:%KMSHost% >nul
cscript //nologo ospp.vbs /act | find /i "successful" >nul
if %errorlevel%==0 (
  echo.
  echo ============================================================================
  echo Activation successful via %KMSHost%!
  echo ============================================================================
  goto ActivationDone
)

echo Failed on %KMSHost%. Retrying…
set /a attempt+=1
goto :TryKMS

:KMSFailed
echo.
echo ============================================================================
echo ERROR: All KMS hosts failed.
echo Please check network connectivity or supply different KMS servers.
echo ============================================================================
pause
exit /b

:ActivationDone
echo.
echo You can verify status with:
echo   cscript //nologo ospp.vbs /dstatusall
echo.
pause
exit /b
