# SharePoint Permissions Analyzer - PowerShell Launcher
# This script provides additional options and better error handling

param(
    [string]$LogLevel = "INFO",
    [string]$OutputFile = "sharepoint_analysis_report.html",
    [string]$TenantId = "",
    [string]$LogFile = "",
    [switch]$Help
)

if ($Help) {
    Write-Host "SharePoint Permissions Analyzer - PowerShell Launcher" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Usage: .\run_analyzer.ps1 [options]" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "Options:" -ForegroundColor Yellow
    Write-Host "  -LogLevel <LEVEL>     Set logging level (DEBUG, INFO, WARNING, ERROR)" -ForegroundColor White
    Write-Host "  -OutputFile <FILE>    Set output HTML file name" -ForegroundColor White
    Write-Host "  -TenantId <ID>        Set Microsoft 365 tenant ID" -ForegroundColor White
    Write-Host "  -LogFile <FILE>       Log to file instead of console" -ForegroundColor White
    Write-Host "  -Help                 Show this help message" -ForegroundColor White
    Write-Host ""
    Write-Host "Examples:" -ForegroundColor Yellow
    Write-Host "  .\run_analyzer.ps1" -ForegroundColor White
    Write-Host "  .\run_analyzer.ps1 -LogLevel DEBUG -OutputFile my_report.html" -ForegroundColor White
    Write-Host "  .\run_analyzer.ps1 -TenantId your-tenant-id" -ForegroundColor White
    exit 0
}

Write-Host "SharePoint Permissions Analyzer" -ForegroundColor Cyan
Write-Host "================================" -ForegroundColor Cyan
Write-Host ""

# Check if Python is installed
try {
    $pythonVersion = python --version 2>&1
    Write-Host "Python version: $pythonVersion" -ForegroundColor Green
} catch {
    Write-Host "Error: Python is not installed or not in PATH" -ForegroundColor Red
    Write-Host "Please install Python 3.8 or higher and try again" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check for virtual environment
$venvPath = "venv"
$activateScript = Join-Path $venvPath "Scripts\activate.bat"
$pythonExe = "python"
$venvActive = $false

if (Test-Path $activateScript) {
    Write-Host "Virtual environment found. Activating..." -ForegroundColor Yellow
    try {
        & $activateScript
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Virtual environment activated." -ForegroundColor Green
            $pythonExe = Join-Path $venvPath "Scripts\python.exe"
            $venvActive = $true
        } else {
            Write-Host "Error: Failed to activate virtual environment. Proceeding without it." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error: Failed to activate virtual environment. Proceeding without it." -ForegroundColor Yellow
    }
} else {
    Write-Host "No virtual environment found. Creating one..." -ForegroundColor Yellow
    try {
        python -m venv $venvPath
        if ($LASTEXITCODE -eq 0) {
            Write-Host "Virtual environment created. Activating..." -ForegroundColor Green
            & $activateScript
            if ($LASTEXITCODE -eq 0) {
                Write-Host "Virtual environment activated." -ForegroundColor Green
                $pythonExe = Join-Path $venvPath "Scripts\python.exe"
                $venvActive = $true
                Write-Host "Upgrading pip..." -ForegroundColor Yellow
                & $pythonExe -m pip install --upgrade pip
            } else {
                Write-Host "Error: Failed to activate virtual environment. Using system Python." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Error: Failed to create virtual environment. Using system Python." -ForegroundColor Yellow
        }
    } catch {
        Write-Host "Error: Failed to create virtual environment. Using system Python." -ForegroundColor Yellow
    }
}

# Check if requirements are installed
Write-Host "Checking dependencies..." -ForegroundColor Yellow
try {
    $msalCheck = & $pythonExe -c "import msal" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Installing required dependencies..." -ForegroundColor Yellow
        & $pythonExe -m pip install -r requirements.txt
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Error: Failed to install dependencies" -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
    } else {
        Write-Host "Dependencies are up to date" -ForegroundColor Green
    }
} catch {
    Write-Host "Error checking dependencies: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

# Check for Azure AD credentials
Write-Host "Checking Azure AD credentials..." -ForegroundColor Yellow
try {
    $credCheck = & $pythonExe -c "from credential_manager import CredentialManager; cm = CredentialManager(); exit(0 if cm.has_credentials() else 1)" 2>&1
    if ($LASTEXITCODE -ne 0) {
        Write-Host "No Azure AD credentials found." -ForegroundColor Yellow
        Write-Host ""
        Write-Host "Setting up Azure AD app registration..." -ForegroundColor Yellow
        Write-Host "This will guide you through creating an Azure AD app registration." -ForegroundColor Yellow
        Write-Host ""
        Read-Host "Press Enter to continue"
        
        & $pythonExe main.py --setup
        if ($LASTEXITCODE -ne 0) {
            Write-Host "Setup failed. Please try again." -ForegroundColor Red
            Read-Host "Press Enter to exit"
            exit 1
        }
        Write-Host ""
        Write-Host "Setup complete! Now running the analyzer..." -ForegroundColor Green
        Write-Host ""
    } else {
        Write-Host "Azure AD credentials found." -ForegroundColor Green
    }
} catch {
    Write-Host "Error checking credentials: $_" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Starting SharePoint Permissions Analyzer..." -ForegroundColor Green
Write-Host ""

# Build command arguments
$args = @("main.py")
if ($LogLevel) { $args += "--log-level", $LogLevel }
if ($OutputFile) { $args += "--output", $OutputFile }
if ($TenantId) { $args += "--tenant-id", $TenantId }
if ($LogFile) { $args += "--log-file", $LogFile }

# Run the main application
try {
    & $pythonExe @args
    $exitCode = $LASTEXITCODE
    
    Write-Host ""
    if ($exitCode -eq 0) {
        Write-Host "Analysis completed successfully!" -ForegroundColor Green
        if (Test-Path $OutputFile) {
            $fullPath = Resolve-Path $OutputFile
            Write-Host "Report saved to: $fullPath" -ForegroundColor Green
        }
    } else {
        Write-Host "Analysis completed with errors (Exit code: $exitCode)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "Error running the analyzer: $_" -ForegroundColor Red
    $exitCode = 1
}

# Deactivate virtual environment if it was activated by this script
if ($venvActive) {
    Write-Host "Deactivating virtual environment..." -ForegroundColor Yellow
    deactivate
}

Write-Host ""
Read-Host "Press Enter to exit"
exit $exitCode
