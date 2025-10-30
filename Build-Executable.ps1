# Build script for Account Status Checker
# This script converts the PowerShell script to a standalone .exe

# Ensure ps2exe is installed
if (-not (Get-Module -ListAvailable -Name ps2exe)) {
    Write-Host "Installing ps2exe module..." -ForegroundColor Yellow
    Install-Module -Name ps2exe -Scope CurrentUser -Force
}

Import-Module ps2exe

# Define paths
$scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "Windows-Update-Repair.ps1"
$exePath = Join-Path -Path $PSScriptRoot -ChildPath "Windows Update Repair.exe"
$iconPath = Join-Path -Path $PSScriptRoot -ChildPath "icon.ico"  # Optional: Add your own icon

# Prompt for version or use date-based default
$defaultVersion = Get-Date -Format "yy.MM.dd.1"
$userVersion = Read-Host "Enter version number (press Enter for default: $defaultVersion)"
$version = if ([string]::IsNullOrWhiteSpace($userVersion)) { $defaultVersion } else { $userVersion }

Write-Host "Building executable..." -ForegroundColor Cyan
Write-Host "Source: $scriptPath" -ForegroundColor Gray
Write-Host "Output: $exePath" -ForegroundColor Gray
Write-Host "Version: $version" -ForegroundColor Gray

# Convert to EXE with options
$params = @{
    InputFile = $scriptPath
    OutputFile = $exePath
    NoConsole = $false  # Hide console for clean GUI-only mode
    NoOutput = $false   # Suppress output messages
    NoError = $true    # Suppress error messages (they'll show in GUI MessageBoxes)
    RequireAdmin = $true  # No admin rights needed for AD queries
    x64 = $true  # Build for 64-bit
    Verbose = $true
    Version = $version
}

# Add icon if it exists
if (Test-Path $iconPath) {
    $params.IconFile = $iconPath
    Write-Host "Using icon: $iconPath" -ForegroundColor Green
}

try {
    Invoke-ps2exe @params
    
    if (Test-Path $exePath) {
        Write-Host "`nBuild successful!" -ForegroundColor Green
        Write-Host "Executable created: $exePath" -ForegroundColor Green
        Write-Host "File size: $([math]::Round((Get-Item $exePath).Length / 1MB, 2)) MB" -ForegroundColor Gray
        
        # Test if file is accessible
        $exeInfo = Get-Item $exePath
        Write-Host "`nExecutable Details:" -ForegroundColor Cyan
        Write-Host "  Name: $($exeInfo.Name)" -ForegroundColor Gray
        Write-Host "  Created: $($exeInfo.CreationTime)" -ForegroundColor Gray
        Write-Host "  Path: $($exeInfo.FullName)" -ForegroundColor Gray
    } else {
        Write-Host "`nBuild failed - executable not found" -ForegroundColor Red
    }
} catch {
    Write-Host "`nError during build:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
}

