<#
.SYNOPSIS
    Setup script to install all required PowerShell modules and dependencies for Power BI Report Export and Conversion.

.DESCRIPTION
    This script installs:
    - NuGet package provider
    - Sets PSGallery as trusted repository
    - MicrosoftPowerBIMgmt modules (with AcceptLicense)
    - Any other required dependencies

.NOTES
    Run this script once before using the export and conversion tools.
    Requires administrator privileges for some installations.
#>

[CmdletBinding()]
param()

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Power BI Report Tools - Dependency Setup" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# Set execution policy for current process
Write-Host "Setting execution policy for this session..." -ForegroundColor Yellow
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force

# Function to install NuGet provider
function Install-NuGetProvider {
    Write-Host "Checking NuGet package provider..." -ForegroundColor Yellow
    try {
        $nuget = Get-PackageProvider -Name NuGet -ListAvailable -ErrorAction SilentlyContinue
        if (-not $nuget) {
            Write-Host "  Installing NuGet provider (minimum version 2.8.5.201)..." -ForegroundColor Cyan
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser -ErrorAction Stop | Out-Null
            Write-Host "  NuGet provider installed successfully" -ForegroundColor Green
        } else {
            Write-Host "  NuGet provider already installed (Version: $($nuget.Version))" -ForegroundColor Green
        }
    }
    catch {
        Write-Host "  Failed to install NuGet provider: $_" -ForegroundColor Red
        throw
    }
}

# Function to set PSGallery as trusted
function Set-PSGalleryTrusted {
    Write-Host "Configuring PSGallery repository..." -ForegroundColor Yellow
    try {
        $repo = Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue
        if ($repo) {
            if ($repo.InstallationPolicy -ne 'Trusted') {
                Write-Host "  Setting PSGallery as trusted repository..." -ForegroundColor Cyan
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
                Write-Host "  PSGallery set as trusted" -ForegroundColor Green
            } else {
                Write-Host "  PSGallery already trusted" -ForegroundColor Green
            }
        } else {
            Write-Host "  PSGallery repository not found" -ForegroundColor Red
        }
    }
    catch {
        Write-Host "  Failed to configure PSGallery: $_" -ForegroundColor Red
        throw
    }
}

# Function to install a PowerShell module
function Install-RequiredModule {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ModuleName,
        [Parameter(Mandatory = $false)]
        [string]$MinimumVersion
    )
    
    Write-Host "Checking module: $ModuleName..." -ForegroundColor Yellow
    try {
        $module = Get-Module -ListAvailable -Name $ModuleName | Sort-Object Version -Descending | Select-Object -First 1
        
        if ($module) {
            if ($MinimumVersion -and $module.Version -lt [version]$MinimumVersion) {
                Write-Host "  Module installed but outdated (Current: $($module.Version), Required: $MinimumVersion)" -ForegroundColor Yellow
                Write-Host "  Updating $ModuleName..." -ForegroundColor Cyan
                Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -AcceptLicense -ErrorAction Stop
                Write-Host "  $ModuleName updated successfully" -ForegroundColor Green
            } else {
                Write-Host "  $ModuleName already installed (Version: $($module.Version))" -ForegroundColor Green
            }
        } else {
            Write-Host "  Installing $ModuleName..." -ForegroundColor Cyan
            if ($MinimumVersion) {
                Install-Module -Name $ModuleName -MinimumVersion $MinimumVersion -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -AcceptLicense -ErrorAction Stop
            } else {
                Install-Module -Name $ModuleName -Scope CurrentUser -Force -AllowClobber -Repository PSGallery -AcceptLicense -ErrorAction Stop
            }
            Write-Host "  $ModuleName installed successfully" -ForegroundColor Green
        }
        
        # Verify installation
        Import-Module $ModuleName -ErrorAction Stop
        Write-Host "  $ModuleName imported and verified" -ForegroundColor Green
    }
    catch {
        Write-Host "  Failed to install/import $ModuleName : $_" -ForegroundColor Red
        throw
    }
}

# Main installation process
try {
    Write-Host "Step 1: Installing NuGet Provider" -ForegroundColor Cyan
    Write-Host "-----------------------------------" -ForegroundColor Cyan
    Install-NuGetProvider
    Write-Host ""

    Write-Host "Step 2: Configuring PSGallery Repository" -ForegroundColor Cyan
    Write-Host "-----------------------------------" -ForegroundColor Cyan
    Set-PSGalleryTrusted
    Write-Host ""

    Write-Host "Step 3: Installing Power BI Management Modules" -ForegroundColor Cyan
    Write-Host "-----------------------------------" -ForegroundColor Cyan
    
    # Install main Power BI management module (includes all sub-modules)
    Install-RequiredModule -ModuleName "MicrosoftPowerBIMgmt"
    
    # Explicitly install sub-modules to ensure they're available
    Install-RequiredModule -ModuleName "MicrosoftPowerBIMgmt.Reports"
    Install-RequiredModule -ModuleName "MicrosoftPowerBIMgmt.Workspaces"
    Install-RequiredModule -ModuleName "MicrosoftPowerBIMgmt.Profile"
    
    Write-Host ""

    Write-Host "Step 4: Installing Optional Modules" -ForegroundColor Cyan
    Write-Host "-----------------------------------" -ForegroundColor Cyan
    
    # PS2EXE for building console to executable (optional)
    Write-Host "Installing PS2EXE (optional - for building .exe)..." -ForegroundColor Yellow
    try {
        Install-RequiredModule -ModuleName "PS2EXE"
    }
    catch {
        Write-Host "  ! PS2EXE installation skipped (non-critical)" -ForegroundColor Yellow
    }
    
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "All dependencies installed successfully!" -ForegroundColor Green
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "Next steps:" -ForegroundColor White
    Write-Host "  1. Run run_exporter_console.ps1 to export and convert reports" -ForegroundColor White
    Write-Host "  2. Or run export_reports.ps1 directly with parameters" -ForegroundColor White
    Write-Host ""
}
catch {
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Setup failed!" -ForegroundColor Red
    Write-Host "========================================" -ForegroundColor Red
    Write-Host "Error: $_" -ForegroundColor Red
    Write-Host ""
    Write-Host "Please address the error and run this script again." -ForegroundColor Yellow
    exit 1
}

Write-Host "Press Enter to close..." -ForegroundColor Cyan
[void] (Read-Host)
