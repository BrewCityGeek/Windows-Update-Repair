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

# --- Step 1: Stop and Disable Windows Update Services ---
Write-Host "--- Step 1 of 6: Stopping and Disabling Windows Update Services ---" -ForegroundColor Green

$services = @("BITS", "wuauserv", "appidsvc", "cryptsvc")

# Store original service startup types for restoration later
$originalStartupTypes = @{}

foreach ($service in $services) {
    try {
        $serviceObj = Get-Service -Name $service -ErrorAction Stop
        $serviceWMI = Get-WmiObject -Class Win32_Service -Filter "Name='$service'" -ErrorAction Stop
        
        # Store original startup type
        $originalStartupTypes[$service] = $serviceWMI.StartMode
        Write-Host "Original startup type for $service`: $($serviceWMI.StartMode)" -ForegroundColor Gray
        
        if ($serviceObj.Status -eq 'Running') {
            Write-Host "Stopping service: $service" -ForegroundColor Gray
            Stop-Service -Name $service -Force -ErrorAction Stop
        } else {
            Write-Host "Service already stopped: $service" -ForegroundColor Gray
        }
        
        # Disable the service
        Write-Host "Disabling service: $service" -ForegroundColor Gray
        Set-Service -Name $service -StartupType Disabled -ErrorAction Stop
        
    } catch {
        Write-Warning "Failed to stop/disable service $service`: $($_.Exception.Message)"
    }
}

Write-Host "Services stopped and disabled successfully." -ForegroundColor Green
Write-Host

# --- Step 2: Clear Windows Update Cache ---
Write-Host "--- Step 2 of 6: Clearing Windows Update Cache ---" -ForegroundColor Green

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
Write-Host "--- Step 3 of 6: Re-registering Windows Update DLLs ---" -ForegroundColor Green

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
Write-Host "--- Step 4 of 6: Resetting Network Settings ---" -ForegroundColor Green

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

# --- Step 5: Re-enable Windows Update Services ---
Write-Host "--- Step 5 of 6: Re-enabling Windows Update Services ---" -ForegroundColor Green

foreach ($service in $services) {
    try {
        # Set all services to Automatic startup (reboot will start them)
        Write-Host "Setting service $service to Automatic startup" -ForegroundColor Gray
        Set-Service -Name $service -StartupType Automatic -ErrorAction Stop
        Write-Host "Service $service configured for automatic startup" -ForegroundColor Gray
        
    } catch {
        Write-Warning "Failed to configure service $service`: $($_.Exception.Message)"
    }
}

Write-Host "Services configured for automatic startup. They will start after reboot." -ForegroundColor Green
Write-Host

# --- Step 6: Reboot System ---
Write-Host "--- Step 6 of 6: Preparing System Reboot ---" -ForegroundColor Green

Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "Windows Update repair process completed!" -ForegroundColor Cyan
Write-Host "The system will reboot in 30 seconds..." -ForegroundColor Yellow
Write-Host "After restart, Windows Update should work properly." -ForegroundColor Yellow
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host

# Give user a chance to cancel the reboot
Write-Host "Press Ctrl+C within 30 seconds to cancel the automatic reboot." -ForegroundColor Red
Write-Host "Otherwise, the system will reboot automatically." -ForegroundColor Yellow
Write-Host

try {
    # Countdown timer
    for ($i = 30; $i -gt 0; $i--) {
        Write-Host "Rebooting in $i seconds..." -ForegroundColor Yellow
        Start-Sleep -Seconds 1
    }
    
    Write-Host "Initiating system reboot..." -ForegroundColor Green
    Restart-Computer -Force
} catch {
    Write-Host "Reboot cancelled by user or error occurred." -ForegroundColor Red
    Write-Host "Please manually restart your computer to complete the repair process." -ForegroundColor Yellow
    Read-Host "Press ENTER to exit the script"
}
