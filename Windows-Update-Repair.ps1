# =================================================================================
#
# NAME: Windows-Update-Repair.ps1
#
# AUTHOR: Enhanced Version
#
# COMMENT: This script repairs Windows Update issues by stopping services,
#          clearing cache, re-registering DLLs, and resetting network settings.
#          It includes error checking and user feedback for each step.
#          It must be run as an Administrator.
#
# =================================================================================

# Step 1: Check for Administrator Privileges
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Warning "This script requires Administrator privileges to run."
    Write-Host "Please right-click the script file and select 'Run as Administrator'."
    Read-Host "Press any key to exit..."
    exit
}

# Clear the screen for a clean output
Clear-Host

# --- Script Header ---
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "     Windows Update Repair Utility" -ForegroundColor Cyan
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host
Write-Host "This script will repair Windows Update by resetting services and cache." -ForegroundColor White
Write-Host "The process can take 5-10 minutes. Please do not close this window." -ForegroundColor Yellow
Write-Host

# Function to pause and exit on error
function Pause-And-Exit {
    Write-Host
    Read-Host "Press ENTER to exit the script."
    exit
}

# --- Step 1: Stop Windows Update Services ---
Write-Host "--- Step 1 of 5: Stopping Windows Update Services ---" -ForegroundColor Green

$services = @("BITS", "wuauserv", "appidsvc", "cryptsvc")
foreach ($service in $services) {
    try {
        $serviceObj = Get-Service -Name $service -ErrorAction Stop
        if ($serviceObj.Status -eq 'Running') {
            Write-Host "Stopping service: $service" -ForegroundColor Gray
            Stop-Service -Name $service -Force -ErrorAction Stop
        } else {
            Write-Host "Service already stopped: $service" -ForegroundColor Gray
        }
    } catch {
        Write-Warning "Failed to stop service $service`: $($_.Exception.Message)"
    }
}

Write-Host "Services stopped successfully." -ForegroundColor Green
Write-Host

# --- Step 2: Clear Windows Update Cache ---
Write-Host "--- Step 2 of 5: Clearing Windows Update Cache ---" -ForegroundColor Green

$cachePaths = @(
    "$env:SystemRoot\SoftwareDistribution",
    "$env:SystemRoot\System32\catroot2"
)

foreach ($path in $cachePaths) {
    try {
        if (Test-Path "$path\*") {
            Write-Host "Clearing cache: $path" -ForegroundColor Gray
            Remove-Item "$path\*" -Recurse -Force -ErrorAction Stop
            Write-Host "Cache cleared: $path" -ForegroundColor Gray
        } else {
            Write-Host "Cache already empty: $path" -ForegroundColor Gray
        }
    } catch {
        Write-Warning "Failed to clear cache $path`: $($_.Exception.Message)"
    }
}

Write-Host "Cache clearing completed." -ForegroundColor Green
Write-Host

# --- Step 3: Re-register Windows Update DLLs ---
Write-Host "--- Step 3 of 5: Re-registering Windows Update DLLs ---" -ForegroundColor Green

$dlls = @(
    "atl.dll", "urlmon.dll", "mshtml.dll", "shdocvw.dll", "browseui.dll",
    "jscript.dll", "vbscript.dll", "scrrun.dll", "msxml.dll", "msxml3.dll",
    "msxml6.dll", "actxprxy.dll", "softpub.dll", "wintrust.dll", "dssenh.dll",
    "rsaenh.dll", "gpkcsp.dll", "sccbase.dll", "slbcsp.dll", "cryptdlg.dll",
    "oleaut32.dll", "ole32.dll", "shell32.dll", "initpki.dll", "wuapi.dll",
    "wuaueng.dll", "wuaueng1.dll", "wucltui.dll", "wups.dll", "wups2.dll",
    "wuweb.dll", "qmgr.dll", "qmgrprxy.dll", "wucltux.dll", "muweb.dll", "wuwebv.dll"
)

$successCount = 0
$totalDlls = $dlls.Count

Write-Host "Registering $totalDlls DLL files..." -ForegroundColor Gray

foreach ($dll in $dlls) {
    try {
        $result = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", $dll -Wait -PassThru -WindowStyle Hidden
        if ($result.ExitCode -eq 0) {
            $successCount++
        } else {
            Write-Warning "Failed to register: $dll (Exit Code: $($result.ExitCode))"
        }
    } catch {
        Write-Warning "Error registering $dll`: $($_.Exception.Message)"
    }
}

Write-Host "DLL registration completed: $successCount/$totalDlls successful" -ForegroundColor Green
Write-Host

# --- Step 4: Reset Network Settings ---
Write-Host "--- Step 4 of 5: Resetting Network Settings ---" -ForegroundColor Green

try {
    Write-Host "Resetting Winsock catalog..." -ForegroundColor Gray
    $winsockResult = netsh winsock reset
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Winsock reset successful" -ForegroundColor Gray
    } else {
        Write-Warning "Winsock reset failed with exit code $LASTEXITCODE"
    }
} catch {
    Write-Warning "Error resetting Winsock: $($_.Exception.Message)"
}

try {
    Write-Host "Resetting WinHTTP proxy settings..." -ForegroundColor Gray
    $proxyResult = netsh winhttp reset proxy
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Proxy settings reset successful" -ForegroundColor Gray
    } else {
        Write-Warning "Proxy reset failed with exit code $LASTEXITCODE"
    }
} catch {
    Write-Warning "Error resetting proxy: $($_.Exception.Message)"
}

Write-Host "Network settings reset completed." -ForegroundColor Green
Write-Host

# --- Step 5: Start Windows Update Services ---
Write-Host "--- Step 5 of 5: Starting Windows Update Services ---" -ForegroundColor Green

foreach ($service in $services) {
    try {
        Write-Host "Starting service: $service" -ForegroundColor Gray
        Start-Service -Name $service -ErrorAction Stop
        
        # Wait a moment and verify service started
        Start-Sleep -Seconds 2
        $serviceObj = Get-Service -Name $service
        if ($serviceObj.Status -eq 'Running') {
            Write-Host "Service started successfully: $service" -ForegroundColor Gray
        } else {
            Write-Warning "Service may not have started properly: $service (Status: $($serviceObj.Status))"
        }
    } catch {
        Write-Warning "Failed to start service $service`: $($_.Exception.Message)"
    }
}

Write-Host "Service startup completed." -ForegroundColor Green
Write-Host

# --- Final Message ---
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "Windows Update repair process completed!" -ForegroundColor Cyan
Write-Host "It is recommended to restart your computer now." -ForegroundColor Yellow
Write-Host "After restart, try running Windows Update again." -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host
Read-Host "Press ENTER to exit the script."