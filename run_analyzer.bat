@echo off
echo SharePoint Permissions Analyzer
echo ================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher and try again
    pause
    exit /b 1
)

REM --- VIRTUAL ENVIRONMENT HANDLING ---
set "VENV_PATH=venv"
set "ACTIVATE_SCRIPT=%VENV_PATH%\Scripts\activate.bat"
set "PYTHON_EXE=python"
set "VENV_ACTIVE=false"

if exist "%ACTIVATE_SCRIPT%" (
    echo Virtual environment found. Activating...
    call "%ACTIVATE_SCRIPT%"
    if errorlevel 1 (
        echo Error: Failed to activate virtual environment. Proceeding without it.
        set "VENV_ACTIVE=false"
    ) else (
        echo Virtual environment activated.
        REM Update PYTHON_EXE to point to the venv's python
        set "PYTHON_EXE=%VENV_PATH%\Scripts\python.exe"
        set "VENV_ACTIVE=true"
    )
) else (
    echo No virtual environment found. Creating one...
    python -m venv %VENV_PATH%
    if errorlevel 1 (
        echo Error: Failed to create virtual environment. Using system Python.
        set "VENV_ACTIVE=false"
    ) else (
        echo Virtual environment created. Activating...
        call "%ACTIVATE_SCRIPT%"
        if errorlevel 1 (
            echo Error: Failed to activate virtual environment. Using system Python.
            set "VENV_ACTIVE=false"
        ) else (
            echo Virtual environment activated.
            set "PYTHON_EXE=%VENV_PATH%\Scripts\python.exe"
            set "VENV_ACTIVE=true"
            echo Upgrading pip...
            "%PYTHON_EXE%" -m pip install --upgrade pip
        )
    )
)

echo.
echo Checking dependencies...
REM Use the determined python executable to check/install dependencies
"%PYTHON_EXE%" -c "import msal" >nul 2>&1
if errorlevel 1 (
    echo Installing required dependencies...
    "%PYTHON_EXE%" -m pip install -r requirements.txt
    if errorlevel 1 (
        echo Error: Failed to install dependencies.
        echo Please ensure you have internet access and try again.
        pause
        exit /b 1
    )
)

echo.
echo Checking Azure AD credentials...

REM Check if credentials exist
"%PYTHON_EXE%" -c "from credential_manager import CredentialManager; cm = CredentialManager(); exit(0 if cm.has_credentials() else 1)" >nul 2>&1
if errorlevel 1 (
    echo No Azure AD credentials found.
    echo.
    echo Setting up Azure AD app registration...
    echo This will guide you through creating an Azure AD app registration.
    echo.
    pause
    "%PYTHON_EXE%" main.py --setup
    if errorlevel 1 (
        echo Setup failed. Please try again.
        pause
        exit /b 1
    )
    echo.
    echo Setup complete! Now running the analyzer...
    echo.
) else (
    echo Azure AD credentials found.
)

echo Starting SharePoint Permissions Analyzer...
echo.

REM Check if diagnostic mode is requested
if "%1"=="--diagnose" (
    echo Running diagnostic mode...
    "%PYTHON_EXE%" diagnose_permissions.py
    echo.
    echo Diagnostic complete. Press any key to exit...
    pause >nul
    goto :cleanup
)

REM Run the main application using the determined python executable
"%PYTHON_EXE%" main.py %*

echo.
echo Analysis complete. Press any key to exit...
pause >nul

:cleanup
REM Deactivate virtual environment if it was activated by this script
if "%VENV_ACTIVE%"=="true" (
    echo Deactivating virtual environment...
    call deactivate
)
