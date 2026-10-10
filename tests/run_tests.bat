@echo off
setlocal

rem Compiles and runs the SDK-free unit tests - no Mastercam, no part.

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

rem Build output OUTSIDE Dropbox: Dropbox locks a just-written .obj while it
rem syncs it, and the next compile that writes the same file then fails.
if defined PT_TEST_OUT ( set "OUT=%PT_TEST_OUT%" ) else ( set "OUT=%TEMP%\ParamTableTests" )
if not exist "%OUT%" md "%OUT%"

set FAILED=0

echo === csv_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\csv_test.exe" ^
    "%~dp0csv_test.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\csv_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === plan_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\plan_test.exe" ^
    "%~dp0plan_test.cpp" "%~dp0..\src\Csv.cpp" "%~dp0..\src\Plan.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\plan_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === load_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\load_test.exe" ^
    "%~dp0load_test.cpp" "%~dp0..\src\Impact.cpp" "%~dp0..\src\Undo.cpp" "%~dp0..\src\PreviewTicks.cpp" ^
    "%~dp0..\src\Plan.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\load_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === xlsx_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\xlsx_test.exe" ^
    "%~dp0xlsx_test.cpp" "%~dp0..\src\Xlsx.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\xlsx_test.exe" "%OUT%\sample.xlsx"
    if errorlevel 1 set FAILED=1
)

echo === xlsx_read_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\xlsx_read_test.exe" ^
    "%~dp0xlsx_read_test.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\XlsxRead.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\xlsx_read_test.exe" "%~dp0fixtures"
    if errorlevel 1 set FAILED=1
)

echo === estimate_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\estimate_test.exe" ^
    "%~dp0estimate_test.cpp" "%~dp0..\src\Estimate.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\estimate_test.exe" "%OUT%"
    if errorlevel 1 set FAILED=1
)

echo === inspect_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\inspect_test.exe" ^
    "%~dp0inspect_test.cpp" "%~dp0..\src\Inspect.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\inspect_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === filerules_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\filerules_test.exe" ^
    "%~dp0filerules_test.cpp" "%~dp0..\src\FileRules.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\filerules_test.exe" "%OUT%"
    if errorlevel 1 set FAILED=1
)

echo === macro_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\macro_test.exe" ^
    "%~dp0macro_test.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\Summary.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\macro_test.exe" "%OUT%" "%~dp0..\res\vbaProject.bin" "%~dp0..\vba\ribbon.xml"
    if errorlevel 1 set FAILED=1
)

echo === partconfig_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\partconfig_test.exe" ^
    "%~dp0partconfig_test.cpp" "%~dp0..\src\PartConfig.cpp" "%~dp0..\src\Summary.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\XlsxRead.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\partconfig_test.exe" "%OUT%"
    if errorlevel 1 set FAILED=1
)

echo === holder_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\holder_test.exe" ^
    "%~dp0holder_test.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\holder_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === summary_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\summary_test.exe" ^
    "%~dp0summary_test.cpp" "%~dp0..\src\Summary.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\XlsxRead.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\summary_test.exe" "%OUT%"
    if errorlevel 1 set FAILED=1
)

echo === history_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\history_test.exe" ^
    "%~dp0history_test.cpp" "%~dp0..\src\History.cpp" "%~dp0..\src\Summary.cpp" "%~dp0..\src\Xlsx.cpp" "%~dp0..\src\XlsxRead.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\history_test.exe" "%OUT%"
    if errorlevel 1 set FAILED=1
)

rem tools\check_summary.ps1 then drives summary_sample.xlsx and history_sample.xlsx in real Excel.

echo === pick_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\pick_test.exe" ^
    "%~dp0pick_test.cpp" "%~dp0..\src\Pick.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\pick_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === ribbon_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\ribbon_test.exe" ^
    "%~dp0ribbon_test.cpp" "%~dp0..\src\Ribbon.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\ribbon_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === stocksim_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\stocksim_test.exe" ^
    "%~dp0stocksim_test.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\stocksim_test.exe"
    if errorlevel 1 set FAILED=1
)

echo === mill_test ===
cl /nologo /EHsc /W4 /O2 /std:c++17 /utf-8 /Fo"%OUT%\\" /Fe"%OUT%\mill_test.exe" ^
    "%~dp0mill_test.cpp" "%~dp0..\src\MillMrr.cpp" "%~dp0..\src\Xform.cpp" "%~dp0..\src\Csv.cpp"
if errorlevel 1 ( set FAILED=1 ) else (
    "%OUT%\mill_test.exe"
    if errorlevel 1 set FAILED=1
)

rem The manual's pictures of the add-in's own windows: built only (it needs a dump workbook -
rem tools\make_help_images.ps1 runs it), so a change that ties the windows (src\Preview.cpp,
rem src\DumpDialog.cpp) to the Mastercam SDK again shows here, not at picture time.
call "%~dp0build_dialog_shots.bat"
if errorlevel 1 set FAILED=1

if %FAILED%==1 (
    echo.
    echo TESTS FAILED
    exit /b 1
)
echo.
echo all tests passed
exit /b 0
