@echo off
rem Get the current directory name (project folder name)
for %%F in ("%cd%") do set project_name=%%~nxF

rem Get the pyinstaller path
set "PROJECT_DIR=%cd%"
set "VENV_PATH=%PROJECT_DIR%\.venv\Scripts\pyinstaller.exe"

rem Run PyInstaller with the specified arguments
%VENV_PATH% --onefile --clean --noconsole --icon=Excel_Analyzer_Icon.ico --add-data="Excel_Analyzer_Icon.png;." --add-data="Excel_Analyzer_Icon.ico;." --name "%project_name%" main.py

rem Check if the PyInstaller build was successful
if exist "dist\%project_name%.exe" (
    echo Build completed successfully.
) else (
    echo Build failed.
    exit /b 1
)

rem Clean up the build and spec files
rmdir /s /q build
del /q "%project_name%.spec"

echo Cleanup complete.

