@echo off
REM ============================================================
REM  Build ScreenCaptureUtility.exe with PyInstaller
REM  Put this file next to main.py (and icon.ico), then
REM  double-click it or run:  .\build.bat
REM ============================================================

cd /d "%~dp0"

REM Use the project's virtual environment if it exists
if exist "venv\Scripts\activate.bat" (
    call "venv\Scripts\activate.bat"
)

echo.
echo === Installing build requirements ===
python -m pip install --upgrade pip
python -m pip install pyinstaller pillow python-docx imageio-ffmpeg PyAudioWPatch
if errorlevel 1 goto :error

echo.
echo === Cleaning previous build ===
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist ScreenCaptureUtility.spec del ScreenCaptureUtility.spec

set ICON_ARG=
if exist icon.ico set ICON_ARG=--icon icon.ico

echo.
echo === Building ScreenCaptureUtility.exe ===
python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name ScreenCaptureUtility ^
    --hidden-import docx ^
    --collect-data docx ^
    --hidden-import imageio_ffmpeg ^
    --collect-all imageio_ffmpeg ^
    --hidden-import pyaudiowpatch ^
    --collect-all pyaudiowpatch ^
    --exclude-module cv2 ^
    --exclude-module numpy ^
    %ICON_ARG% ^
    main.py
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Done!  Your app is here:
echo    %~dp0dist\ScreenCaptureUtility.exe
echo ============================================================
pause
exit /b 0

:error
echo.
echo *** Build failed - see the messages above. ***
pause
exit /b 1