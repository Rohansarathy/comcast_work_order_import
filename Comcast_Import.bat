@echo OFF
echo Starting batch file...

cd /d "C:\Users\Administrator\OneDrive - ITG Communications, LLC\Work Order Import\work_order_import_comcast_upload"
echo Current directory is: %CD%

if errorlevel 1 (
    echo Failed to change directory
    exit /b 1
)

echo Running Python...
"C:\Program Files\Python312\python.exe" "C:\Users\Administrator\OneDrive - ITG Communications, LLC\Work Order Import\work_order_import_comcast_upload\Work_order_main.py"

echo Python finished with errorlevel %ERRORLEVEL%

if errorlevel 1 (
    echo Robot execution failed
    exit /b 1
)

echo Test execution completed successfully
exit /b 0
