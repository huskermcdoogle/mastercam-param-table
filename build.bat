@echo off
setlocal

rem Builds the Parameter Table Tool C-Hook (x64 Release by default).
rem Usage:  build.bat [Release^|Debug]
rem
rem Mastercam keeps a C-Hook RESIDENT once it has loaded it, so the link fails
rem with LNK1104 (file locked) while Mastercam is running. Close it first.

set "CFG=%~1"
if "%CFG%"=="" set "CFG=Release"

set "VSWHERE=%ProgramFiles(x86)%\Microsoft Visual Studio\Installer\vswhere.exe"
if not exist "%VSWHERE%" (
    echo ERROR: vswhere.exe not found - is Visual Studio installed?
    exit /b 1
)

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -requires Microsoft.Component.MSBuild -find MSBuild\**\Bin\MSBuild.exe`) do set "MSBUILD=%%I"
if not defined MSBUILD (
    echo ERROR: MSBuild.exe not found.
    exit /b 1
)

for /f "usebackq delims=" %%I in (`"%VSWHERE%" -latest -property installationPath`) do set "VSPATH=%%I"

rem THE SDK NEEDS THE v143 TOOLSET. Visual Studio 2026's own v145 cannot compile the
rem SDK's headers (BCGCBProInc.h rejects _MSC_VER >= 1950, and UniqueArray_CH.h puts
rem `const` on a constructor - hard error C7731). v143 is registered under MSBuild's
rem v170 tree, so that tree's targets are used for it.
set "VCT=%VSPATH%\MSBuild\Microsoft\VC\v170\"
if not exist "%VCT%Platforms\x64\PlatformToolsets\v143" (
    echo ERROR: the v143 toolset is not installed.
    echo Add "MSVC v143 - VS 2022 C++ x64/x86 build tools" and its MFC component
    echo in the Visual Studio Installer.
    exit /b 1
)

rem AND NOT JUST ANY v143. The SDK's UniqueArray_CH.h puts `const` on a constructor,
rem which MSVC 14.44 and 14.51 both reject (C7731) - it was accepted by the compiler
rem the SDK was built with. So the build is PINNED to an older 14.x build: v143
rem alone selects the newest installed one, which is the one that fails.
set "TOOLVER="
for /d %%D in ("%VSPATH%\VC\Tools\MSVC\14.40*") do set "TOOLVER=%%~nxD"
if not defined TOOLVER (
    echo ERROR: the 14.40 build of the v143 toolset is not installed.
    echo Add "MSVC v143 - VS 2022 C++ x64/x86 build tools ^(v14.40-17.10^)" and its
    echo "C++ MFC for v143 build tools ^(v14.40-17.10^)" in the Visual Studio Installer.
    exit /b 1
)

rem MFC is a separate Visual Studio component and the Mastercam SDK cannot build
rem without it - CString/CArray/CWnd appear in the SDK's own class definitions.
if not exist "%VSPATH%\VC\Tools\MSVC\%TOOLVER%\atlmfc\include\afxwin.h" (
    echo ERROR: no MFC headers under %TOOLVER% ^(atlmfc\include\afxwin.h^).
    echo Install "C++ MFC for v143 build tools ^(v14.40-17.10^)" in the Visual Studio Installer.
    exit /b 1
)

rem The link tracking log is deleted before every build. MSBuild's own record of
rem the last link was measured to report "up to date" over a stale DLL on the
rem sister project; only the LINK logs go, so compilation stays incremental.
set "TLOG=%LOCALAPPDATA%\ParamTable\obj\x64\%CFG%\ParamTable.tlog"
if exist "%TLOG%\link.read.1.tlog" del /q "%TLOG%\link.*.tlog" >nul 2>&1

echo Building ParamTable ^(%CFG%^|x64, v143 %TOOLVER%^)...
"%MSBUILD%" "%~dp0ParamTable.vcxproj" /nologo /v:minimal /p:Configuration=%CFG% /p:Platform=x64 /p:PlatformToolset=v143 /p:VCToolsVersion=%TOOLVER% "/p:VCTargetsPath=%VCT%\"
if errorlevel 1 (
    echo.
    echo Build FAILED.
    exit /b 1
)

if not exist "%~dp0build\x64\%CFG%\ParamTable.dll" (
    echo Build FAILED - no ParamTable.dll was produced.
    exit /b 1
)

echo.
echo Build complete: %~dp0build\x64\%CFG%\ParamTable.dll
endlocal
