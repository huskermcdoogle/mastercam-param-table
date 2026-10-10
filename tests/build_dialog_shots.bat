@echo off
setlocal

rem Builds dialog_shots.exe: the user manual's pictures of the add-in's own windows
rem (dump window, load preview, undo, batch dump), drawn without Mastercam - see
rem tests\dialog_shots.cpp. tools\make_help_images.ps1 builds and runs it.
rem MFC, no Mastercam SDK. Into %PT_TEST_OUT% (else %TEMP%\ParamTableTests), as the tests.

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo ERROR: vswhere.exe not found - is Visual Studio installed?
    exit /b 1
)

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -property installationPath`) do set "VSDIR=%%I"
if not defined VSDIR (
    echo ERROR: Visual Studio not found.
    exit /b 1
)

call "%VSDIR%\VC\Auxiliary\Build\vcvars64.bat" >nul
if errorlevel 1 (
    echo ERROR: vcvars64 failed.
    exit /b 1
)

if defined PT_TEST_OUT ( set "OUT=%PT_TEST_OUT%" ) else ( set "OUT=%TEMP%\ParamTableTests" )
rem Its objects in a folder of their own: the tests build the same sources without MFC.
if not exist "%OUT%\dialog_shots" md "%OUT%\dialog_shots"

echo === dialog_shots (build) ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /MD /D_AFXDLL /DUNICODE /D_UNICODE ^
    /Fo"%OUT%\dialog_shots\\" /Fe"%OUT%\dialog_shots.exe" ^
    "%~dp0dialog_shots.cpp" "%~dp0..\src\DumpDialog.cpp" "%~dp0..\src\BatchDialog.cpp" "%~dp0..\src\Shop.cpp" ^
    "%~dp0..\src\Inspect.cpp" "%~dp0..\src\Preview.cpp" "%~dp0..\src\PreviewTicks.cpp" ^
    "%~dp0..\src\Pick.cpp" "%~dp0..\src\FileRules.cpp" "%~dp0..\src\Impact.cpp" "%~dp0..\src\Undo.cpp" ^
    "%~dp0..\src\Plan.cpp" "%~dp0..\src\Csv.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\XlsxRead.cpp" ^
    /link /MANIFEST:EMBED
if errorlevel 1 (
    echo dialog_shots: BUILD FAILED
    exit /b 1
)
echo built %OUT%\dialog_shots.exe
exit /b 0
