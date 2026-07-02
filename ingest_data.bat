@echo off
setlocal

REM ===== EDIT THESE =====
set INPUT_DIR=C:\Users\Work\Documents\_data_pipeline\input_files
set BUNDLE_DIR=C:\Users\Work\Documents\_data_pipeline
set LOG_DIR=C:\Users\Work\Documents\_data_pipeline\logs
set ARCHIVE_DIR=C:\Users\Work\Documents\_data_pipeline\archive
set ERROR_DIR=C:\Users\Work\Documents\_data_pipeline\error_files

set CONN_STR=DRIVER={ODBC Driver 17 for SQL Server};SERVER=100.68.124.105;DATABASE=pim;UID=Ashley;PWD=SecurePassword1;
set RUN_ID=20260202

REM ===== RUN =====
python "%BUNDLE_DIR%\ingest_vendors.py" ^
  --input "%INPUT_DIR%" ^
  --mappings "%BUNDLE_DIR%\vendor_mappings" ^
  --conn "%CONN_STR%" ^
  --run-id %RUN_ID% ^
  --log-dir "%LOG_DIR%" ^
  --archive-dir "%ARCHIVE_DIR%" ^
  --error-dir "%ERROR_DIR%" ^
  --report-out "ingest_report.json"

pause
endlocal
