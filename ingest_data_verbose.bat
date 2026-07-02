@echo on
setlocal EnableDelayedExpansion

echo ============================================
echo   VENDOR INGESTION JOB STARTING
echo   %DATE%  %TIME%
echo ============================================

REM ====== PATHS ======
set INPUT_DIR=C:\Users\Work\Documents\_data_pipeline\input_files
set BUNDLE_DIR=C:\Users\Work\Documents\_data_pipeline
set LOG_DIR=C:\Users\Work\Documents\_data_pipeline\logs
set ARCHIVE_DIR=C:\Users\Work\Documents\_data_pipeline\archive
set ERROR_DIR=C:\Users\Work\Documents\_data_pipeline\error_files

REM ====== DATABASE ======
set CONN_STR=DRIVER={ODBC Driver 17 for SQL Server};SERVER=100.68.124.105;DATABASE=pim;UID=Ashley;PWD=SecurePassword1;

REM ====== RUN METADATA ======
set RUN_ID=%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%_%TIME:~0,2%%TIME:~3,2%%TIME:~6,2%
set RUN_ID=%RUN_ID: =0%

echo.
echo INPUT DIRECTORY:   %INPUT_DIR%
echo MAPPINGS DIRECTORY: %BUNDLE_DIR%\vendor_mappings
echo LOG DIRECTORY:     %LOG_DIR%
echo ARCHIVE DIRECTORY: %ARCHIVE_DIR%
echo ERROR DIRECTORY:   %ERROR_DIR%
echo RUN ID:            %RUN_ID%
echo.

echo Checking Python...
python --version
if errorlevel 1 (
    echo ❌ Python not found!
    pause
    exit /b
)

echo.
echo Checking input files...
dir "%INPUT_DIR%" *.xlsx *.xlsm *.xlsb *.xls

echo.
echo ============================================
echo   STARTING INGESTION SCRIPT
echo ============================================
echo.

python "%BUNDLE_DIR%\ingest_vendors.py" ^
  --input "%INPUT_DIR%" ^
  --mappings "%BUNDLE_DIR%\vendor_mappings" ^
  --conn "%CONN_STR%" ^
  --run-id %RUN_ID% ^
  --log-dir "%LOG_DIR%" ^
  --archive-dir "%ARCHIVE_DIR%" ^
  --error-dir "%ERROR_DIR%" ^
  --report-out "ingest_report.json"

set EXITCODE=%ERRORLEVEL%

echo.
echo ============================================
echo   SCRIPT FINISHED  (Exit Code: %EXITCODE%)
echo   %DATE%  %TIME%
echo ============================================

if %EXITCODE% NEQ 0 (
    echo ❌ The script exited with errors.
) else (
    echo ✅ Ingestion completed.
)

echo.
echo Recent log files:
dir "%LOG_DIR%" /od

echo.
pause
endlocal
