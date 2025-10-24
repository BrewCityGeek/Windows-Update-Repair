# Build Instructions for Windows Update Repair Tool

This document provides instructions for building the Windows Update Repair PowerShell script into a standalone executable.

## Prerequisites

- **PowerShell 5.1 or later** (Windows PowerShell or PowerShell Core)
- **Internet connection** (for downloading the ps2exe module)
- **Windows operating system** (64-bit recommended)

## Build Dependencies

The build process requires the `ps2exe` PowerShell module, which converts PowerShell scripts to standalone executables.

### Automatic Installation

The build script will automatically install the `ps2exe` module if it's not already present:

```powershell
Install-Module -Name ps2exe -Scope CurrentUser -Force
```

### Manual Installation (Optional)

If you prefer to install the module manually before building:

```powershell
Install-Module -Name ps2exe -Scope CurrentUser
```

## Build Process

### Using the Build Script

1. **Navigate to the project directory**:
   ```powershell
   cd "c:\scripts\Windows-Update-Repair"
   ```

2. **Run the build script**:
   ```powershell
   .\Build-Executable.ps1
   ```

### Build Configuration

The build script (`Build-Executable.ps1`) is configured with the following parameters:

- **Input File**: `Windows-Update-Repair.ps1`
- **Output File**: `Windows Update Repair.exe`
- **Console Mode**: Enabled (shows console output)
- **Admin Rights**: Required (`RequireAdmin = $true`)
- **Architecture**: 64-bit (`x64 = $true`)
- **Error Handling**: Suppressed in console (errors shown in GUI MessageBoxes)

### Optional: Custom Icon

To use a custom icon for the executable:

1. Place an `.ico` file named `icon.ico` in the project directory
2. The build script will automatically detect and use it

## Build Output

### Successful Build

When the build completes successfully, you'll see:

```
Build successful!
Executable created: c:\scripts\Windows-Update-Repair\Windows Update Repair.exe
File size: [X.XX] MB

Executable Details:
  Name: Windows Update Repair.exe
  Created: [timestamp]
  Path: [full path to executable]
```

### Build Artifacts

The build process creates:

- **`Windows Update Repair.exe`** - The standalone executable
- The original PowerShell script remains unchanged

## Troubleshooting

### Common Issues

1. **Module Installation Fails**
   ```
   Error: Unable to install ps2exe module
   ```
   **Solution**: Run PowerShell as Administrator and try again

2. **Build Permission Denied**
   ```
   Error: Access denied during build
   ```
   **Solution**: Ensure the output directory is writable and no antivirus is blocking the process

3. **PowerShell Execution Policy**
   ```
   Error: Execution of scripts is disabled
   ```
   **Solution**: 
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```

4. **Missing Source Script**
   ```
   Error: Cannot find Windows-Update-Repair.ps1
   ```
   **Solution**: Ensure you're running the build script from the correct directory

### Build Verification

To verify the executable was built correctly:

1. **Check file existence**:
   ```powershell
   Test-Path ".\Windows Update Repair.exe"
   ```

2. **Check file properties**:
   ```powershell
   Get-Item ".\Windows Update Repair.exe" | Select-Object Name, Length, CreationTime
   ```

3. **Test execution** (optional):
   ```powershell
   & ".\Windows Update Repair.exe"
   ```

## Advanced Build Options

### Manual ps2exe Usage

For advanced users who want to customize the build process:

```powershell
Invoke-ps2exe -InputFile "Windows-Update-Repair.ps1" -OutputFile "Custom-Name.exe" -NoConsole:$false -RequireAdmin:$true -x64:$true
```

### Available ps2exe Parameters

- `-NoConsole` - Hide console window (set to `$true` for GUI-only mode)
- `-NoOutput` - Suppress output messages
- `-NoError` - Suppress error messages
- `-RequireAdmin` - Require administrator privileges
- `-x64` - Build for 64-bit architecture
- `-IconFile` - Path to custom icon file
- `-Verbose` - Show detailed build information

## Distribution

The resulting executable (`Windows Update Repair.exe`) is self-contained and can be:

- Copied to other Windows machines without requiring PowerShell or additional modules
- Distributed via email, USB drives, or network shares
- Run directly by double-clicking (will prompt for administrator privileges if required)

## Notes

- The executable requires administrator privileges to perform Windows Update repairs
- The original PowerShell script functionality is preserved in the executable
- The build process creates a portable executable that doesn't require PowerShell to be installed on target machines
- File size is typically 5-15 MB depending on the complexity of the source script

## Support

If you encounter issues during the build process:

1. Verify all prerequisites are met
2. Check the troubleshooting section above
3. Ensure you have the latest version of the `ps2exe` module
4. Run the build script with verbose output for detailed error information