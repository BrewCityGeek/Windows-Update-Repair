# Usage Guide - Windows Update Repair Tool

## Prerequisites

Before running the script, ensure you meet these requirements:

### System Requirements
- Windows 10, Windows 11, or Windows Server
- Administrator privileges
- Active internet connection (recommended)
- At least 1GB free disk space

### Important Notes
⚠️ **Close Windows Update** settings page before running  
⚠️ **Save your work** and close programs before running  
⚠️ **Disable antivirus** temporarily if it interferes  

## Running the Script

### Method 1: PowerShell (Recommended)
1. Press `Windows + X` and select "Windows PowerShell (Admin)" or "Terminal (Admin)"
2. Navigate to the script directory:
   ```powershell
   cd "c:\scripts\System Repair Scripts"
   ```
3. Run the script:
   ```powershell
   .\Windows-Update-Repair.ps1
   ```

### Method 2: File Explorer
1. Navigate to the script location in File Explorer
2. Right-click on `Windows-Update-Repair.ps1`
3. Select **"Run with PowerShell"**
4. If prompted, choose **"Run as Administrator"**

### Method 3: Run Dialog
1. Press `Windows + R` to open Run dialog
2. Type: `powershell -ExecutionPolicy Bypass -File "c:\scripts\System Repair Scripts\Windows-Update-Repair.ps1"`
3. Press Enter

## What to Expect

### Execution Timeline
| Step | Process | Typical Duration |
|------|---------|------------------|
| 1 | Stop Services | 30-60 seconds |
| 2 | Clear Cache | 1-3 minutes |
| 3 | Register DLLs | 2-4 minutes |
| 4 | Reset Network | 30-60 seconds |
| 5 | Start Services | 30-60 seconds |

### Screen Output
The script provides color-coded feedback:
- 🔵 **Cyan**: Headers and completion messages
- 🟢 **Green**: Step progress and success
- 🟡 **Yellow**: Warnings and recommendations
- ⚪ **Gray**: Individual action details
- 🔴 **Red**: Errors (script continues despite errors)

### Sample Output
```
=================================================
     Windows Update Repair Utility
=================================================

This script will repair Windows Update by resetting services and cache.
The process can take 5-10 minutes. Please do not close this window.

--- Step 1 of 5: Stopping Windows Update Services ---
Stopping service: BITS
Stopping service: wuauserv
Stopping service: appidsvc
Stopping service: cryptsvc
Services stopped successfully.

--- Step 2 of 5: Clearing Windows Update Cache ---
Clearing cache: C:\Windows\SoftwareDistribution
Cache cleared: C:\Windows\SoftwareDistribution
Clearing cache: C:\Windows\System32\catroot2
Cache cleared: C:\Windows\System32\catroot2
Cache clearing completed.

--- Step 3 of 5: Re-registering Windows Update DLLs ---
Registering 35 DLL files...
DLL registration completed: 35/35 successful

--- Step 4 of 5: Resetting Network Settings ---
Resetting Winsock catalog...
Winsock reset successful
Resetting WinHTTP proxy settings...
Proxy settings reset successful
Network settings reset completed.

--- Step 5 of 5: Starting Windows Update Services ---
Starting service: BITS
Service started successfully: BITS
Starting service: wuauserv
Service started successfully: wuauserv
Starting service: appidsvc
Service started successfully: appidsvc
Starting service: cryptsvc
Service started successfully: cryptsvc
Service startup completed.

=================================================
Windows Update repair process completed!
It is recommended to restart your computer now.
After restart, try running Windows Update again.
=================================================
```

## Troubleshooting

### Common Issues

#### "Execution Policy" Error
If you see an execution policy error:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### "Access Denied" Error
- Ensure you're running as Administrator
- Close Windows Update settings page
- Try running from an elevated PowerShell prompt

#### Services Won't Stop
- Close Windows Update settings in Control Panel
- End Windows Update related processes in Task Manager
- Restart computer and try again

#### Cache Clearing Fails
- Some files may be in use - this is normal
- Script will continue and skip locked files
- Restart computer if many files couldn't be deleted

#### DLL Registration Failures
- Some DLLs may not exist on your system - this is normal
- Script tracks success/failure count
- Failures don't prevent script completion

#### Services Won't Restart
- Check Windows Event Viewer for service errors
- Restart computer to reset service states
- Run `services.msc` to manually start services

### Post-Repair Steps

#### After Successful Completion
1. **Restart your computer** (highly recommended)
2. **Open Windows Update** (Settings > Update & Security > Windows Update)
3. **Click "Check for updates"** 
4. **Install any pending updates**
5. **Run the script again** if issues persist

#### Verification Steps
To verify the repair was successful:

```powershell
# Check service status
Get-Service BITS, wuauserv, appidsvc, cryptsvc

# Check Windows Update
Get-WindowsUpdate -Online

# Test network connectivity
Test-NetConnection -ComputerName download.windowsupdate.com -Port 80
```

### Advanced Troubleshooting

#### Manual Service Reset
If services still won't start:
```powershell
# Reset Windows Update services manually
sc config wuauserv start= auto
sc config BITS start= auto
sc config cryptsvc start= auto
sc config appidsvc start= auto

# Start services
net start BITS
net start wuauserv
net start cryptsvc
net start appidsvc
```

#### Alternative Cache Locations
If standard cache clearing doesn't work:
```powershell
# Additional cache locations to clear manually
Remove-Item "$env:LocalAppData\Microsoft\Windows\WindowsUpdate\*" -Recurse -Force
Remove-Item "$env:ProgramData\USOPrivate\*" -Recurse -Force
```

#### Registry Reset (Advanced Users Only)
⚠️ **Warning**: Only for experienced users
```powershell
# Reset Windows Update registry keys
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate" /f
reg delete "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update" /f
```

## Specific Error Codes

### Common Windows Update Error Codes This Script Fixes

| Error Code | Description | How Script Helps |
|------------|-------------|------------------|
| 0x80070005 | Access Denied | Service and permission reset |
| 0x8024402F | Connection failed | Network settings reset |
| 0x80240034 | Update not applicable | Cache and component reset |
| 0x8024A105 | Update service not available | Service restart |
| 0x80244007 | Server not available | Network and proxy reset |
| 0x80240016 | Update failed | Complete component reset |

## Best Practices

### Before Running
1. **Create system restore point**
2. **Close all programs**
3. **Ensure stable power supply**
4. **Check available disk space**

### During Execution
1. **Don't close the PowerShell window**
2. **Don't run other update tools simultaneously**
3. **Don't interrupt the process**
4. **Monitor for error messages**

### After Running
1. **Restart the computer**
2. **Test Windows Update immediately**
3. **Check for driver updates**
4. **Verify system stability**

## When to Seek Additional Help

Contact IT support if:
- Script fails repeatedly with same errors
- Windows Update still doesn't work after restart
- System becomes unstable after running script
- Critical services won't start after repair
- Multiple error codes persist

## Alternative Solutions

If this script doesn't resolve your issues:

### Built-in Windows Tools
- **Windows Update Troubleshooter**: Settings > Update & Security > Troubleshoot
- **DISM**: Use the Windows system file repair script
- **SFC**: System File Checker for corrupted files

### Microsoft Tools
- **Windows Update MiniTool**: Third-party update manager
- **Windows Repair Tool**: Microsoft's official repair utility
- **System File Checker**: Built-in Windows utility

### Manual Reset Process
1. Download fresh Windows Update Agent
2. Manually reset all Windows Update components
3. Re-register all Windows Update related services
4. Reset Windows Update policies