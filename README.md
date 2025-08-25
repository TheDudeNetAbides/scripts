# Hyper-V VM Export Tool

A comprehensive PowerShell script for exporting Hyper-V virtual machines with progress tracking, flexible destination options, and robust error handling.

## Features

- **Real-time Progress Tracking**: Shows actual Hyper-V export progress using CIM/WMI
- **Multiple Export Modes**: Shutdown, Save State, or Online export
- **Flexible Destinations**: Local folders, USB drives, or network shares (UNC paths)
- **Automatic Prerequisites Check**: Validates admin rights, PowerShell version, and Hyper-V module
- **Comprehensive Logging**: Creates timestamped transcript and error logs
- **Space Validation**: Checks available disk space before export
- **Interactive Mode**: Prompts for missing parameters when run without arguments

## Prerequisites

- Windows Server 2016+ or Windows 10/11 Pro/Enterprise with Hyper-V
- PowerShell 5.1 (Windows PowerShell recommended over PowerShell 7)
- Administrator privileges
- Hyper-V PowerShell module installed
- PowerShell execution policy allowing script execution (RemoteSigned or Unrestricted)

## Required Permissions and Security

### Administrator Rights Required
This script **must** run with administrator privileges because it uses the following privileged operations:
- Starting and stopping VMs
- Accessing VM configuration files
- Creating export directories in system locations
- Querying CIM/WMI for Hyper-V job status
- Managing VM states (shutdown, save, power operations)

### PowerShell Execution Policy
The script needs to execute PowerShell code, which requires an appropriate execution policy:

**Option 1: Set Execution Policy (Recommended)**
```powershell
# Set for current user only
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned

# Or set system-wide (requires admin)
Set-ExecutionPolicy -Scope LocalMachine -ExecutionPolicy RemoteSigned
```

**Option 2: Bypass for Single Execution**
```powershell
PowerShell.exe -ExecutionPolicy Bypass -File .\Export-HyperVVM.ps1
```

**Option 3: Unblock Downloaded Files**
If downloaded from the internet, unblock the script:
```powershell
Unblock-File -Path .\Export-HyperVVM.ps1
```

### Required PowerShell Cmdlets

The script uses the following PowerShell cmdlets that require specific modules and permissions:

#### Hyper-V Module Cmdlets
- `Get-VM` - Query VM information
- `Get-VMHardDiskDrive` - Get VM disk information
- `Stop-VMGuest` - Graceful VM shutdown
- `Stop-VM` - Force VM power off
- `Save-VM` - Save VM state
- `Export-VM` - Core VM export functionality
- `Import-Module Hyper-V` - Load Hyper-V module

#### System and Storage Cmdlets
- `Get-Disk` - Query physical disk information
- `Get-Partition` - Query disk partitions
- `Get-Volume` - Query volume/drive information
- `Get-Item` - File system operations
- `Get-ChildItem` - Directory enumeration
- `New-Item` - Create directories
- `Test-Path` - Path validation

#### CIM/WMI Cmdlets (for Progress Tracking)
- `Get-CimInstance` - Query Hyper-V management objects
- `Get-CimAssociatedInstance` - Query related CIM objects

#### PowerShell Core Cmdlets
- `Start-Job` / `Get-Job` / `Remove-Job` - Background job management
- `Start-Transcript` / `Stop-Transcript` - Session logging
- `Write-Progress` - Progress bar display
- `Read-Host` - User input
- `Write-Host` / `Write-Warning` / `Write-Error` - Output display

## Usage

### Command Line Parameters

```powershell
.\Export-HyperVVM.ps1 [-VMName <string>] [-DestinationPath <string>] [-Mode <string>] [-LogPath <string>]
```

#### Parameters

- **VMName** (optional): Name of the VM to export. If omitted, you'll be prompted.
- **DestinationPath** (optional): Export destination path. If omitted, interactive selection menu appears.
- **Mode** (optional): Export mode. Valid values:
  - `Shutdown` (default): Gracefully shut down VM before export
  - `Save`: Save VM state before export
  - `Online`: Export while VM is running (less consistent)
- **LogPath** (optional): Custom path for log files. Defaults to `Documents\HyperV-Exports\Logs`

### Examples

#### Basic Usage
```powershell
# Interactive mode - prompts for all parameters
.\Export-HyperVVM.ps1

# Export specific VM with default settings
.\Export-HyperVVM.ps1 -VMName "MyVM"

# Full parameter specification
.\Export-HyperVVM.ps1 -VMName "MyVM" -DestinationPath "D:\Exports" -Mode "Shutdown"
```

#### Advanced Examples
```powershell
# Export to network share
.\Export-HyperVVM.ps1 -VMName "WebServer" -DestinationPath "\\backup-server\vm-exports" -Mode "Save"

# Custom log location
.\Export-HyperVVM.ps1 -VMName "Database" -LogPath "C:\Logs\HyperV"

# Online export (VM stays running)
.\Export-HyperVVM.ps1 -VMName "ProductionVM" -Mode "Online"
```

## Export Modes Explained

### Shutdown Mode (Recommended)
- Attempts graceful guest shutdown first
- Falls back to forced power-off if needed
- Ensures data consistency
- Best for production environments

### Save Mode
- Saves VM state to disk
- Preserves running applications and memory state
- VM can be quickly restored to exact state
- Good for development/testing scenarios

### Online Mode
- Exports while VM continues running
- Fastest option but may have consistency issues
- Only use when VM downtime is not acceptable
- Consider application-level consistency mechanisms

## Destination Options

When run interactively, the script offers three destination types:

1. **Local Folder**: Any accessible local directory (C:\Exports, D:\Backups, etc.)
2. **USB Drive**: Automatically detects USB-attached storage devices
3. **Network Share**: UNC paths to shared network storage

## Log Files

The script creates two log files in the specified log directory:

- **Transcript Log**: Complete PowerShell session transcript
- **Error Log**: Dedicated error messages with timestamps

Default location: `%USERPROFILE%\Documents\HyperV-Exports\Logs`

## Export Structure

Exported VMs are organized as follows:
```
<DestinationPath>\
└── <VMName>-Export-<timestamp>\
    ├── Virtual Machines\          # VM configuration files
    ├── Virtual Hard Disks\        # VHD/VHDX files
    └── Snapshots\                 # Checkpoint files (if any)
```

## Troubleshooting

### Common Issues

1. **"Run this in an elevated Windows PowerShell 5.1 session"**
   - Solution: Right-click PowerShell and "Run as Administrator"

2. **"Execution policy is tight"**
   - The script offers to set CurrentUser policy to RemoteSigned
   - Alternative: Run with `-ExecutionPolicy Bypass`

3. **"PowerShell 7 detected"**
   - Use Windows PowerShell 5.1 for better Hyper-V cmdlet compatibility
   - Found in: `%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe`

4. **"Hyper-V PowerShell module not found"**
   - Install Hyper-V management tools via Windows Features
   - Or use: `Enable-WindowsOptionalFeature -Online -FeatureName Microsoft-Hyper-V-Management-PowerShell`

5. **"Execution cannot be loaded because running scripts is disabled"**
   - Check execution policy: `Get-ExecutionPolicy`
   - Set policy: `Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser`
   - Or run with bypass: `PowerShell.exe -ExecutionPolicy Bypass -File .\Export-HyperVVM.ps1`

6. **"Not enough free space"**
   - Check available disk space on destination drive
   - Consider using external storage or network share

### Performance Tips

- Use local SSDs for fastest export performance
- Ensure adequate free space (VM size + 10% overhead)
- Close unnecessary applications during large exports
- Consider using Save mode for faster subsequent exports

## Version History

- **v6**: Fixed ordering, real progress tracking, and execution policy prompts
- Enhanced CIM-based progress monitoring
- Improved USB drive detection
- Better error handling and logging

## License

This script is provided as-is for educational and operational use. Test thoroughly in non-production environments before use.
