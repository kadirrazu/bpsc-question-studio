@echo off
setlocal
cd /d "%~dp0"
if exist "dist\BPSC-Question-Studio" (
  echo dist\BPSC-Question-Studio already exists. Rename or move it before rebuilding.
  goto failed
)
py -3.12 -c "import platform, struct, sys; sys.exit(not (struct.calcsize('P') == 8 and platform.machine().upper() in ('AMD64', 'X86_64')))"
if errorlevel 1 (
  echo This build needs Windows x64 Python 3.12 and the py launcher.
  goto failed
)
if not exist ".build\v051\tools\Scripts\python.exe" (
  py -3.12 -m venv ".build\v051\tools"
  if errorlevel 1 goto failed
)
set "BUILD_PY=.build\v051\tools\Scripts\python.exe"
set "BUILD_ROOT=.build\v051"
set "EMBEDDED_PYTHON=%BUILD_ROOT%\python-3.12.10-embed-amd64.zip"
echo Installing the build compiler...
"%BUILD_PY%" -m pip install --only-binary=:all: "ziglang==0.17.1"
if errorlevel 1 goto failed
if not exist "%EMBEDDED_PYTHON%" (
  echo Downloading the Windows x64 embedded Python runtime...
  curl.exe --fail --location --retry 2 "https://www.python.org/ftp/python/3.12.10/python-3.12.10-embed-amd64.zip" --output "%EMBEDDED_PYTHON%.part"
  if errorlevel 1 goto failed
  move /Y "%EMBEDDED_PYTHON%.part" "%EMBEDDED_PYTHON%" >nul
  if errorlevel 1 goto failed
)
"%BUILD_PY%" -c "import hashlib,pathlib,sys; sys.exit(0 if hashlib.sha256(pathlib.Path(sys.argv[1]).read_bytes()).hexdigest() == '4acbed6dd1c744b0376e3b1cf57ce906f9dc9e95e68824584c8099a63025a3c3' else 'Embedded Python checksum mismatch. Remove the cached ZIP, then rebuild.')" "%EMBEDDED_PYTHON%"
if errorlevel 1 goto failed
echo Downloading the application wheels...
"%BUILD_PY%" -m pip download --only-binary=:all: --platform win_amd64 --implementation cp --python-version 312 --abi cp312 --dest "%BUILD_ROOT%\wheels" -r requirements.txt -c packaging\portable-constraints.txt
if errorlevel 1 goto failed
echo Compiling the launcher...
"%BUILD_PY%" -m ziglang rc /FO%BUILD_ROOT%\branding.res packaging\branding.rc
if errorlevel 1 goto failed
"%BUILD_PY%" -m ziglang cc -target x86_64-windows-gnu -O2 -municode -Wl,--subsystem,windows packaging\launcher.c "%BUILD_ROOT%\branding.res" -o "%BUILD_ROOT%\Launcher.exe"
if errorlevel 1 goto failed
echo Assembling the portable application...
"%BUILD_PY%" packaging\assemble_portable.py --python-zip "%EMBEDDED_PYTHON%" --wheels "%BUILD_ROOT%\wheels" --launcher "%BUILD_ROOT%\Launcher.exe" --output "dist\BPSC-Question-Studio"
if errorlevel 1 goto failed
echo.
echo Build complete. Run "dist\BPSC-Question-Studio\BPSC Question Studio.exe".
echo Keep the entire dist\BPSC-Question-Studio folder when distributing the app.
pause
exit /b 0

:failed
echo.
echo Build stopped. See the error above.
pause
exit /b 1
