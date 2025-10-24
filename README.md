# Windows Update Repair Tool

A PowerShell script that fixes common Windows Update issues by resetting services, clearing cache, re-registering DLLs, and restoring network settings.

## Overview

This script addresses Windows Update problems by performing a comprehensive reset of the Windows Update system components. It's designed to resolve issues like stuck updates, error codes, and corrupted update components.

## Features

- ✅ **Administrator Privilege Check** - Ensures script runs with proper permissions
- ✅ **Service Management** - Safely stops and restarts Windows Update services
- ✅ **Cache Clearing** - Removes corrupted update cache files
- ✅ **DLL Re-registration** - Fixes corrupted Windows Update components
- ✅ **Network Reset** - Restores Winsock and proxy settings
- ✅ **Error Handling** - Continues operation even if individual steps fail
- ✅ **Progress Feedback** - Clear status updates and colored output
- ✅ **Service Verification** - Confirms services restart successfully

## System Requirements

- **Operating System**: Windows 10, Windows 11, or Windows Server
- **Permissions**: Administrator privileges required
- **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+
- **Time**: Allow 5-10 minutes for completion

## Repair Process

The script performs these steps in sequence:

1. **Stop Services** - Safely stops Windows Update related services
   - BITS (Background Intelligent Transfer Service)
   - Windows Update Service (wuauserv)
   - Application Identity (appidsvc)
   - Cryptographic Services (cryptsvc)

2. **Clear Cache** - Removes corrupted cache files
   - SoftwareDistribution folder contents
   - Catroot2 folder contents

3. **Re-register DLLs** - Fixes Windows Update component registration
   - 35+ critical Windows Update and system DLLs
   - Progress tracking with success/failure counts

4. **Reset Network** - Restores network connectivity settings
   - Winsock catalog reset
   - WinHTTP proxy settings reset

5. **Restart Services** - Brings Windows Update system back online
   - Verifies each service starts successfully

## Files

- `Windows-Update-Repair.ps1` - Main repair script

## Quick Start

1. Download the script to your computer
2. Right-click `Windows-Update-Repair.ps1`
3. Select "Run with PowerShell" or "Run as Administrator"
4. Follow the on-screen prompts
5. Restart your computer when prompted
6. Try running Windows Update again

## When to Use

Run this script when experiencing:

### Common Windows Update Issues
- Windows Update stuck at certain percentages
- Error codes: 0x80070005, 0x8024402F, 0x80240034, 0x8024A105
- "Windows Update can't currently check for updates"
- Updates downloading but failing to install
- Windows Update service not starting

### Symptoms This Script Addresses
- Update downloads fail repeatedly
- Windows Update shows no available updates when there should be
- Update service crashes or becomes unresponsive
- Error messages about corrupted update components
- Network connectivity issues affecting updates

## Safety

This script is safe because:
- Uses only standard Windows utilities and commands
- Doesn't modify system files directly
- Only resets Windows Update components to default state
- Can be safely run multiple times
- Includes comprehensive error handling

## Compatibility

| Windows Version | Compatibility | Notes |
|----------------|---------------|-------|
| Windows 11 | ✅ Full | All features supported |
| Windows 10 | ✅ Full | All features supported |
| Windows Server 2022 | ✅ Full | All features supported |
| Windows Server 2019 | ✅ Full | All features supported |
| Windows Server 2016 | ✅ Full | All features supported |

## Author

Enhanced version with comprehensive error handling and user feedback.

## License

This script is provided as-is for educational and repair purposes. Use at your own risk.