@echo on
setlocal EnableExtensions

REM ---- Force UTF-8 for Python/stdout even under Task Scheduler ----
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
set "LANG=en_US.UTF-8"
set "LC_ALL=en_US.UTF-8"

chcp 65001 >nul

REM ==== move to script folder ====
cd /d "%~dp0" || exit /b 1

REM ==== CONFIG ====
set "INPUT_PATH=C:\Users\Najeeb.Jabareen\Desktop\job\projects\wissotzky\source_file_stage1\QS.xlsx"
set "OUTPUT_DIR=C:\Users\Najeeb.Jabareen\Desktop\job\projects\wissotzky\outputs_files_stage1"
REM ==============

if not exist "%OUTPUT_DIR%" mkdir "%OUTPUT_DIR%"

REM ==== LOG FILE ====
set "LOGDIR=%~dp0_logs"
set "TSLOG=%LOGDIR%\task.log"
echo ==== %DATE% %TIME% ==== >> "%TSLOG%"
echo [TS] Starting via Task Scheduler >> "%TSLOG%"
if not exist "%LOGDIR%" mkdir "%LOGDIR%"
set "LOG=%LOGDIR%\run.log"
echo ==== %DATE% %TIME% ==== >> "%LOG%"

REM ==== choose python ====
set "PYEXE=%~dp0.venv\Scripts\python.exe"
if not exist "%PYEXE%" set "PYEXE=python"

echo Using: %PYEXE% >> "%LOG%" 2>&1
"%PYEXE%" -V >> "%LOG%" 2>&1

REM ==== run pipeline (force UTF-8 at runtime too) ====
"%PYEXE%" -X utf8 -W ignore::FutureWarning -m pipeline.run_stage1 ^
  --input "%INPUT_PATH%" ^
  --output-dir "%OUTPUT_DIR%" ^
  --drop-empty ^
  --split-by-manager ^
  --market-private ^
  --market-tedmiti ^
  --pivot-private ^
  --pivot-tedmiti ^
  --by-agent >> "%LOG%" 2>&1

echo ExitCode=%ERRORLEVEL% >> "%LOG%" 2>&1
echo ExitCode=%ERRORLEVEL% >> "%TSLOG%" 2>&1
endlocal
exit /b %ERRORLEVEL%
