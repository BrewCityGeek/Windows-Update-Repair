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

# Function to stop services with retry logic
function Stop-ServiceWithRetry {
    param(
        [string]$ServiceName,
        [int]$MaxRetries = 3,
        [int]$WaitSeconds = 2
    )
    
    for ($i = 1; $i -le $MaxRetries; $i++) {
        try {
            $service = Get-Service -Name $ServiceName -ErrorAction Stop
            if ($service.Status -eq 'Running') {
                Write-Host "Attempt $i`: Stopping service: $ServiceName" -ForegroundColor Gray
                Stop-Service -Name $ServiceName -Force -ErrorAction Stop
                
                # Wait and verify the service stopped
                Start-Sleep -Seconds $WaitSeconds
                $service = Get-Service -Name $ServiceName -ErrorAction Stop
                if ($service.Status -eq 'Stopped') {
                    Write-Host "  ✓ Service stopped: $ServiceName" -ForegroundColor DarkGreen
                    return $true
                }
            } else {
                Write-Host "Service already stopped: $ServiceName" -ForegroundColor Gray
                return $true
            }
        } catch {
            Write-Warning "Attempt $i failed to stop service $ServiceName`: $($_.Exception.Message)"
            if ($i -lt $MaxRetries) {
                Start-Sleep -Seconds $WaitSeconds
            }
        }
    }
    
    Write-Warning "Failed to stop service $ServiceName after $MaxRetries attempts"
    return $false
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
        
        # Use retry logic to stop the service
        Stop-ServiceWithRetry -ServiceName $service
        
        # Disable the service
        Write-Host "Disabling service: $service" -ForegroundColor Gray
        Set-Service -Name $service -StartupType Disabled -ErrorAction Stop
        
    } catch {
        Write-Warning "Failed to stop/disable service $service`: $($_.Exception.Message)"
    }
}

Write-Host "Services stopped and disabled successfully." -ForegroundColor Green
Write-Host "Waiting 3 seconds for services to fully stop..." -ForegroundColor Gray
Start-Sleep -Seconds 3
Write-Host

# --- Step 2: Clear Windows Update Cache ---
Write-Host "--- Step 2 of 6: Clearing Windows Update Cache ---" -ForegroundColor Green

$cachePaths = @(
    "$env:SystemRoot\SoftwareDistribution",
    "$env:SystemRoot\System32\catroot2"
)

foreach ($path in $cachePaths) {
    try {
        if (Test-Path $path) {
            Write-Host "Clearing cache: $path" -ForegroundColor Gray
            
            # First, try to remove all contents
            if (Test-Path "$path\*") {
                Remove-Item "$path\*" -Recurse -Force -ErrorAction SilentlyContinue
            }
            
            # Then try to remove the directory itself and recreate it
            # This ensures any locked files are handled properly
            $tempPath = "$path" + "_temp"
            try {
                # Rename the directory (this often works even when deletion fails)
                if (Test-Path $path) {
                    Rename-Item $path $tempPath -ErrorAction SilentlyContinue
                    # Create new empty directory
                    New-Item -Path $path -ItemType Directory -Force | Out-Null
                    # Try to remove the renamed directory
                    Remove-Item $tempPath -Recurse -Force -ErrorAction SilentlyContinue
                }
            } catch {
                # If rename fails, just ensure the directory exists and is empty
                if (-not (Test-Path $path)) {
                    New-Item -Path $path -ItemType Directory -Force | Out-Null
                }
            }
            
            Write-Host "Cache cleared: $path" -ForegroundColor Gray
        } else {
            Write-Host "Cache directory not found, creating: $path" -ForegroundColor Gray
            New-Item -Path $path -ItemType Directory -Force | Out-Null
        }
    } catch {
        Write-Warning "Failed to clear cache $path`: $($_.Exception.Message)"
        # Ensure directory exists even if clearing failed
        try {
            if (-not (Test-Path $path)) {
                New-Item -Path $path -ItemType Directory -Force | Out-Null
            }
        } catch {
            Write-Warning "Could not recreate directory: $path"
        }
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
$skippedCount = 0

Write-Host "Registering $totalDlls DLL files..." -ForegroundColor Gray

foreach ($dll in $dlls) {
    try {
        # First check if the DLL exists in System32
        $system32Path = "$env:SystemRoot\System32\$dll"
        $syswow64Path = "$env:SystemRoot\SysWOW64\$dll"
        
        $dllPath = $null
        if (Test-Path $system32Path) {
            $dllPath = $system32Path
        } elseif (Test-Path $syswow64Path) {
            $dllPath = $syswow64Path
        }
        
        if ($dllPath) {
            Write-Host "Registering: $dll" -ForegroundColor Gray
            $result = Start-Process -FilePath "regsvr32.exe" -ArgumentList "/s", "`"$dllPath`"" -Wait -PassThru -WindowStyle Hidden
            if ($result.ExitCode -eq 0) {
                $successCount++
                Write-Host "  ✓ Successfully registered: $dll" -ForegroundColor DarkGreen
            } else {
                Write-Warning "  ✗ Failed to register: $dll (Exit Code: $($result.ExitCode))"
            }
        } else {
            Write-Host "  - Skipping $dll (not found on this system)" -ForegroundColor DarkYellow
            $skippedCount++
        }
    } catch {
        Write-Warning "  ✗ Error registering $dll`: $($_.Exception.Message)"
    }
}

Write-Host "DLL registration completed:" -ForegroundColor Green
Write-Host "  ✓ Successful: $successCount" -ForegroundColor Green
Write-Host "  ✗ Failed: $($totalDlls - $successCount - $skippedCount)" -ForegroundColor Red
Write-Host "  - Skipped (not found): $skippedCount" -ForegroundColor Yellow
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
