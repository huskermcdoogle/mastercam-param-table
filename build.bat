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

rem ALWAYS LINK. MSBuild's own record of the last link was measured to report
rem "up to date" over a stale DLL (and the intermediates now sit in a folder per
rem checkout, so the old trick of deleting its link logs no longer found them):
rem the DLL goes before every build, so the link always runs. It takes seconds;
rem compilation stays incremental.
if exist "%~dp0build\x64\%CFG%\ParamTable.dll" del /q "%~dp0build\x64\%CFG%\ParamTable.dll"

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

rem A backstop: never report "Build complete" over a DLL older than any source file.
powershell -NoProfile -Command "$d=(Get-Item '%~dp0build\x64\%CFG%\ParamTable.dll').LastWriteTimeUtc; $s=Get-ChildItem '%~dp0src','%~dp0res\vbaProject.bin','%~dp0vba\ribbon.xml' -Recurse -File | Where-Object { $_.LastWriteTimeUtc -gt $d.AddSeconds(2) }; if ($s) { 'STALE DLL - newer than it: ' + (($s | Select-Object -First 5 -ExpandProperty Name) -join ', '); exit 1 }"
if errorlevel 1 (
    echo Build FAILED - the DLL is older than the sources. Delete %LOCALAPPDATA%\ParamTable\obj and build again.
    exit /b 1
)
echo.
echo Build complete: %~dp0build\x64\%CFG%\ParamTable.dll
endlocal
