@echo off
echo NIRF Data MongoDB Import Tool
echo ==============================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Error: Python is not installed or not in PATH.
    echo Please install Python and try again.
    goto :end
)

REM Check if All-XML directory exists
if not exist "All-XML" (
    echo Error: All-XML directory not found. Please make sure it exists and contains XML files.
    echo.
    pause
    goto :end
)

echo Found All-XML directory with XML files for processing.
echo.

REM Install required packages if needed
echo Checking required packages...
pip show pymongo >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo Installing required packages...
    pip install pymongo tqdm
    if %ERRORLEVEL% NEQ 0 (
        echo Failed to install required packages.
        goto :end
    )
)

echo.
echo Importing XML data to MongoDB...
echo This will process all XML files and create a structured master database.
echo.

REM Run the Python script with default parameters for batch processing
python xml_to_mongodb.py All-XML localhost 27017 nirf_database individuals master_database

echo.
if %ERRORLEVEL% EQU 0 (
    echo Import completed successfully.
) else (
    echo Import failed with errors.
)

:end
pause