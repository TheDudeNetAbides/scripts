# Hyper-V VM Management Suite

A comprehensive PowerShell toolkit for complete Hyper-V virtual machine lifecycle management. This suite includes both export and import scripts with progress tracking, flexible storage options, and robust error handling.

## Suite Overview

| Script | Primary Function | Key Features |
|--------|------------------|--------------|
| **Export-HyperVVM.ps1** | Create VM backups/exports | Real-time progress, multiple VM states, space validation |
| **Import-HyperVVM.ps1** | Restore VMs from exports | Conflict resolution, multiple import modes, auto-discovery |

## Features

### Export Tool Features
- **Real-time Progress Tracking**: Shows actual Hyper-V export progress using CIM/WMI
- **Multiple Export Modes**: Shutdown, Save State, or Online export
- **Flexible Destinations**: Local folders, USB drives, or network shares (UNC paths)
- **Space Validation**: Checks available disk space before export
- **VM State Management**: Graceful shutdown with fallback options

### Import Tool Features
- **Real-time Progress Tracking**: Shows actual Hyper-V import progress using CIM/WMI
- **Multiple Import Types**: Register in-place, Copy, or Move operations
- **Flexible Sources**: Local folders, USB drives, or network shares (UNC paths)
- **Automatic VM Discovery**: Analyzes export structure and suggests VM names
- **Conflict Resolution**: Detects existing VMs with same names and prompts for alternatives
- **Space Validation**: Checks available disk space before import

### Shared Features
- **Automatic Prerequisites Check**: Validates admin rights, PowerShell version, and Hyper-V module
- **Comprehensive Logging**: Creates timestamped transcript and error logs
- **Interactive Mode**: Prompts for missing parameters when run without arguments
- **USB Drive Detection**: Smart detection of USB-attached storage devices
- **Network Share Support**: Full UNC path support with connectivity validation

## Prerequisites

- Windows Server 2016+ or Windows 10/11 Pro/Enterprise with Hyper-V
- PowerShell 5.1 (Windows PowerShell recommended over PowerShell 7)
- Administrator privileges
- Hyper-V PowerShell module installed
- PowerShell execution policy allowing script execution (RemoteSigned or Unrestricted)

## Required Permissions and Security

### Administrator Rights Required
Both scripts **must** run with administrator privileges because they use the following privileged operations:
- Starting, stopping, and managing VMs
- Importing and exporting VM configuration files
- Accessing VM configuration files in system locations
- Creating export/import directories in system locations
- Querying CIM/WMI for Hyper-V job status
- Managing VM states (shutdown, save, power operations)
- Moving and copying virtual disk files

### PowerShell Execution Policy
The scripts need to execute PowerShell code, which requires an appropriate execution policy:

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
PowerShell.exe -ExecutionPolicy Bypass -File .\Import-HyperVVM.ps1
```

**Option 3: Unblock Downloaded Files**
If downloaded from the internet, unblock the scripts:
```powershell
Unblock-File -Path .\Export-HyperVVM.ps1
Unblock-File -Path .\Import-HyperVVM.ps1
```

### Required PowerShell Cmdlets

The scripts use the following PowerShell cmdlets that require specific modules and permissions:

#### Hyper-V Module Cmdlets
- `Export-VM` - Core VM export functionality (Export script)
- `Import-VM` - Core VM import functionality (Import script)
- `Get-VM` - Query VM information
- `Get-VMHardDiskDrive` - Get VM disk information
- `Stop-VMGuest` - Graceful VM shutdown (Export script)
- `Stop-VM` - Force VM power off (Export script)
- `Save-VM` - Save VM state (Export script)
- `Rename-VM` - Rename imported VMs (Import script)
- `Move-VMStorage` - Move VM storage locations (Import script)
- `Get-VMHost` - Query Hyper-V host settings
- `Import-Module Hyper-V` - Load Hyper-V module

#### System and Storage Cmdlets
- `Get-Disk` - Query physical disk information
- `Get-Partition` - Query disk partitions
- `Get-Volume` - Query volume/drive information
- `Get-PSDrive` - Query PowerShell drives (Import script)
- `Get-Item` - File system operations
- `Get-ChildItem` - Directory enumeration
- `New-Item` - Create directories
- `Test-Path` - Path validation
- `Split-Path` - Path manipulation (Import script)

#### CIM/WMI Cmdlets (for Progress Tracking)
- `Get-CimInstance` - Query Hyper-V management objects
- `Get-CimAssociatedInstance` - Query related CIM objects

#### PowerShell Core Cmdlets
- `Start-Job` / `Get-Job` / `Remove-Job` - Background job management
- `Start-Transcript` / `Stop-Transcript` - Session logging
- `Write-Progress` - Progress bar display
- `Read-Host` - User input
- `Write-Host` / `Write-Warning` / `Write-Error` - Output display

### Network Share Access
If using UNC paths, ensure:
- Current user has appropriate permissions (read for import sources, write for export destinations)
- Network connectivity to target server
- Appropriate SMB/CIFS protocols enabled

## Usage

### Export Script Usage

#### Command Line Parameters
```powershell
.\Export-HyperVVM.ps1 [-VMName <string>] [-DestinationPath <string>] [-Mode <string>] [-LogPath <string>]
```

#### Export Parameters
- **VMName** (optional): Name of the VM to export. If omitted, you'll be prompted.
- **DestinationPath** (optional): Export destination path. If omitted, interactive selection menu appears.
- **Mode** (optional): Export mode. Valid values:
  - `Shutdown` (default): Gracefully shut down VM before export
  - `Save`: Save VM state before export
  - `Online`: Export while VM is running (less consistent)
- **LogPath** (optional): Custom path for log files. Defaults to `Documents\HyperV-Exports\Logs`

#### Export Examples
```powershell
# Interactive mode - prompts for all parameters
.\Export-HyperVVM.ps1

# Export specific VM with default settings
.\Export-HyperVVM.ps1 -VMName "MyVM"

# Full parameter specification
.\Export-HyperVVM.ps1 -VMName "WebServer" -DestinationPath "D:\Exports" -Mode "Shutdown"

# Export to network share
.\Export-HyperVVM.ps1 -VMName "Database" -DestinationPath "\\backup-server\vm-exports" -Mode "Save"

# Online export (VM stays running)
.\Export-HyperVVM.ps1 -VMName "ProductionVM" -Mode "Online"
```

### Import Script Usage

#### Command Line Parameters
```powershell
.\Import-HyperVVM.ps1 [-SourcePath <string>] [-VMName <string>] [-DestinationPath <string>] [-ImportType <string>] [-VhdDestinationPath <string>] [-SnapshotFilePath <string>] [-GenerateNewId] [-LogPath <string>]
```

#### Import Parameters
- **SourcePath** (optional): Path to the VM export folder. If omitted, interactive selection menu appears.
- **VMName** (optional): Name for the imported VM. If omitted, estimated from export folder name.
- **DestinationPath** (optional): Destination path for VM configuration files. Defaults to Hyper-V default location.
- **ImportType** (optional): Import operation type. Valid values:
  - `Copy` (default): Copy all files to new location
  - `Register`: Register VM in-place (files stay where they are)
  - `Move`: Move files to new location
- **VhdDestinationPath** (optional): Destination for virtual disk files. Defaults to Hyper-V VHD default location.
- **SnapshotFilePath** (optional): Path to snapshot files if importing specific snapshot.
- **GenerateNewId** (switch): Generate new VM ID to avoid conflicts.
- **LogPath** (optional): Custom path for log files. Defaults to `Documents\HyperV-Imports\Logs`

#### Import Examples
```powershell
# Interactive mode - prompts for all parameters
.\Import-HyperVVM.ps1

# Import from specific export folder
.\Import-HyperVVM.ps1 -SourcePath "D:\VMExports\WebServer-Export-20241225_120000"

# Import with custom name and destination
.\Import-HyperVVM.ps1 -SourcePath "\\nas\vmexports\DB-Server-Export-20241225_120000" -VMName "DatabaseServer" -DestinationPath "E:\VMs"

# Register VM in-place (no file copying)
.\Import-HyperVVM.ps1 -SourcePath "D:\VMBackups\MyVM-Export-20241225_120000" -ImportType "Register"

# Import with new ID to avoid conflicts
.\Import-HyperVVM.ps1 -SourcePath "\\server\share\VM-Export" -GenerateNewId -VMName "ClonedVM"
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

## Import Types Explained

### Copy Mode (Recommended)
- Copies all VM files to new destination locations
- Original export files remain unchanged
- Safe for creating VM copies or moving between systems
- Requires additional disk space
- Best for production environments

### Register Mode
- Registers VM using existing file locations
- No file copying occurs - VM uses files where they are
- Fastest import option
- Useful when files are already in desired location
- Risk: Original export becomes the live VM

### Move Mode
- Moves files from source to destination
- Original export location is emptied
- Good for permanent VM relocation
- Requires write access to source location
- Use with caution - original export is consumed

## Storage Options

When run interactively, both scripts offer three storage types:

1. **Local Folder**: Any accessible local directory
2. **USB Drive**: Automatically detects USB-attached storage devices
3. **Network Share**: UNC paths to shared network storage

## VM Name Conflict Resolution (Import Only)

The import script automatically detects if a VM with the target name already exists and provides options:
- Enter a new unique name for the imported VM
- Cancel the import operation
- Shows details of existing conflicting VM (state, location)

## Log Files

Both scripts create two log files in the specified log directory:

- **Transcript Log**: Complete PowerShell session transcript
- **Error Log**: Dedicated error messages with timestamps

Default locations:
- Export: `%USERPROFILE%\Documents\HyperV-Exports\Logs`
- Import: `%USERPROFILE%\Documents\HyperV-Imports\Logs`

## File Structure

### Export Structure
Exported VMs are organized as follows:
```
<DestinationPath>\
└── <VMName>-Export-<timestamp>\
    ├── Virtual Machines\          # VM configuration files
    ├── Virtual Hard Disks\        # VHD/VHDX files
    └── Snapshots\                 # Checkpoint files (if any)
```

### Import Structure Expected
The import script expects VM exports in standard Hyper-V export format:
```
<SourcePath>\
├── Virtual Machines\          # VM configuration files (.vmcx or .exp)
├── Virtual Hard Disks\        # VHD/VHDX files
└── Snapshots\                 # Checkpoint files (if any)
```

## Troubleshooting

### Common Issues

1. **"Run this in an elevated Windows PowerShell 5.1 session"**
   - Solution: Right-click PowerShell and "Run as Administrator"

2. **"Execution policy is tight"**
   - Both scripts offer to set CurrentUser policy to RemoteSigned
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
   - Or run with bypass: `PowerShell.exe -ExecutionPolicy Bypass -File .\<script>.ps1`

6. **"Not enough free space"**
   - Check available disk space on destination/source drives
   - Consider using external storage or network share
   - Use Register mode for imports when files are already in suitable locations

7. **"Invalid VM export: Missing 'Virtual Machines' folder"** (Import only)
   - Verify the source path points to a complete VM export
   - Check that all export files were copied/transferred correctly

8. **"VM with name 'X' already exists"** (Import only)
   - The script will prompt for a new name or allow cancellation
   - Use `-GenerateNewId` switch to create unique VM instances

### Performance Tips

- Use local SSDs for fastest performance
- Ensure adequate free space (VM size + 20% overhead for imports, + 10% for exports)
- Use Register mode for imports when files are already in optimal locations
- Consider separate fast storage for VHDs vs VM configuration files
- Close unnecessary applications during large operations
- Monitor disk space on destination/source drives

### Best Practices

#### For Exports
- **Always test exports** with non-critical VMs first
- **Use Shutdown mode** for production VMs to ensure consistency
- **Monitor available space** on destination drives
- **Verify export completion** before deleting original VMs
- **Document export locations** and retention policies

#### For Imports
- **Always test imports** with non-critical exports first
- **Use Copy mode** for production VM imports to preserve originals
- **Verify VM functionality** after import before deleting source exports
- **Use meaningful VM names** to avoid conflicts
- **Monitor disk space** on destination drives
- **Document VM locations** and naming conventions

## Script Comparison

| Feature | Export Tool | Import Tool |
|---------|-------------|-------------|
| **Primary Function** | Create VM backups/exports | Restore VMs from exports |
| **VM State Handling** | Shutdown/Save/Online modes | Automatic state restoration |
| **Progress Tracking** | Real-time export progress | Real-time import progress |
| **Conflict Handling** | Overwrites existing exports | Detects and resolves name conflicts |
| **File Operations** | Always creates new files | Copy/Register/Move options |
| **Space Requirements** | Source VM size + 5% overhead | Import size + 10-20% overhead |
| **Interactive Features** | Destination selection | Source browsing + conflict resolution |

## Workflow Examples

### Complete VM Migration Workflow
```powershell
# Step 1: Export VM from source system
.\Export-HyperVVM.ps1 -VMName "WebServer" -DestinationPath "\\nas\migration" -Mode "Shutdown"

# Step 2: Import VM on destination system  
.\Import-HyperVVM.ps1 -SourcePath "\\nas\migration\WebServer-Export-20241225_120000" -ImportType "Copy"
```

### VM Backup and Restore Workflow
```powershell
# Daily backup
.\Export-HyperVVM.ps1 -VMName "ProductionDB" -DestinationPath "\\backup-server\daily" -Mode "Save"

# Restore from backup if needed
.\Import-HyperVVM.ps1 -SourcePath "\\backup-server\daily\ProductionDB-Export-20241225_020000" -VMName "ProductionDB-Restored"
```

### VM Cloning Workflow
```powershell
# Export original VM
.\Export-HyperVVM.ps1 -VMName "Template" -DestinationPath "C:\Templates" -Mode "Shutdown"

# Import as clone with new ID
.\Import-HyperVVM.ps1 -SourcePath "C:\Templates\Template-Export-20241225_120000" -GenerateNewId -VMName "Clone1"
```

## Version History

- **v6.1**: Revalidated both scripts with PSScriptAnalyzer 
- **v6.0**: Fixed ordering, real progress tracking, and execution policy prompts
- Enhanced CIM-based progress monitoring for both scripts
- Improved USB drive detection and export/import folder browsing
- Better error handling, logging, and user experience

## License

These scripts are provided as-is for educational and operational use. Test thoroughly in non-production environments before use.

## Support

For issues, feature requests, or contributions:
1. Verify all prerequisites are met
2. Check troubleshooting section
3. Review log files for detailed error information
4. Ensure scripts are running with Administrator privileges
5. Validate PowerShell execution policy settings
