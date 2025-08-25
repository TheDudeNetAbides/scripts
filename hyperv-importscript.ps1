# =====================================================================
# Import-HyperVVM.ps1 - v6.1 (PSScriptAnalyzer validated)
# =====================================================================

[CmdletBinding()]
param(
    [string]$SourcePath,
    [string]$VMName,
    [string]$DestinationPath,
    [ValidateSet('Register','Copy','Move')]
    [string]$ImportType = 'Copy',
    [string]$VhdDestinationPath,
    [string]$SnapshotFilePath,
    [switch]$GenerateNewId,
    [string]$LogPath
)

# Execution Policy assistance (must be AFTER param in a script)
try {
    if (-not $env:__HV_IMPORT_POLICY_AUDITED) {
        $env:__HV_IMPORT_POLICY_AUDITED = '1'
        $restrictive = @('Restricted','AllSigned','Undefined')
        $cuEP = Get-ExecutionPolicy -Scope CurrentUser -ErrorAction SilentlyContinue
        $lmEP = Get-ExecutionPolicy -Scope LocalMachine -ErrorAction SilentlyContinue
        if ($cuEP -in $restrictive -and $lmEP -in $restrictive) {
            Write-Host "Execution policy is tight (CurrentUser=$cuEP, LocalMachine=$lmEP)." -ForegroundColor Yellow
            $ans = Read-Host "Set CurrentUser to RemoteSigned for future runs? (Y/N)"
            if ($ans -match '^(Y|y)$') {
                try {
                    Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
                    Write-Host "Set CurrentUser policy to RemoteSigned." -ForegroundColor Green
                } catch { Write-Warning "Could not set policy: $($_.Exception.Message)" }
            } else {
                Write-Host "Tip: You can always launch with -ExecutionPolicy Bypass on demand." -ForegroundColor Yellow
            }
        }
    }
} catch { }

#region Utilities
function Throw-If($Condition, $Message) { if ($Condition) { throw $Message } }

function Ensure-Admin {
    $isAdmin = ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).
        IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    Throw-If (-not $isAdmin) "Run this in an elevated Windows PowerShell 5.1 session."
}

function Ensure-PSVersion {
    if ($PSVersionTable.PSEdition -eq 'Core') {
        Write-Warning "PowerShell 7 detected. Use Windows PowerShell 5.1 for Hyper-V cmdlets."
    }
}

function Ensure-ModuleHyperV {
    if (-not (Get-Module -ListAvailable -Name Hyper-V)) {
        throw "Hyper-V PowerShell module not found. Install Hyper-V management tools."
    }
    Import-Module Hyper-V -ErrorAction Stop | Out-Null
}

function New-Loggers {
    param([string]$BaseLogDir)
    if (-not $BaseLogDir) {
        $BaseLogDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "HyperV-Imports\Logs"
    }
    New-Item -Path $BaseLogDir -ItemType Directory -Force | Out-Null
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $Transcript = Join-Path $BaseLogDir "Import-$($VMName)-$ts-transcript.log"
    $ErrorLog   = Join-Path $BaseLogDir "Import-$($VMName)-$ts-error.log"
    Start-Transcript -Path $Transcript -Append | Out-Null
    [PSCustomObject]@{ Transcript=$Transcript; ErrorLog=$ErrorLog; Base=$BaseLogDir }
}

function Log-Error { param([string]$Message,[string]$ErrorLog)
    $line = "[{0}] ERROR: {1}" -f (Get-Date), $Message
    $line | Tee-Object -FilePath $ErrorLog -Append | Out-String | Write-Host -ForegroundColor Red
}
function Log-Info { param([string]$Message) Write-Host "[INFO] $Message" }

function Get-USBDriveVolumes {
    # Catch USB-attached SSDs/HDDs that present as DriveType=Fixed
    $usbDisks = Get-Disk | Where-Object { $_.BusType -eq 'USB' -and $_.OperationalStatus -eq 'Online' }
    $vols = foreach ($d in $usbDisks) {
        Get-Partition -DiskNumber $d.Number -ErrorAction SilentlyContinue |
            Get-Volume -ErrorAction SilentlyContinue |
            Where-Object { $_.DriveLetter }
    }
    if ($vols) { $vols | Sort-Object DriveLetter -Unique }
}

function Choose-SourcePath {
    Write-Host ""
    Write-Host "Choose VM export source location:"
    Write-Host "  1) Local folder (e.g. C:\Exports\MyVM-Export-20240101_120000)"
    Write-Host "  2) USB-attached drive"
    Write-Host "  3) UNC network share (\\server\share\path)"
    $choice = Read-Host "Enter 1, 2, or 3"
    switch ($choice) {
        '1' { Read-Host "Enter full path to VM export folder" }
        '2' {
            $usbVols = Get-USBDriveVolumes
            if (-not $usbVols) {
                $usbVols = Get-Volume | Where-Object { $_.DriveType -eq 'Removable' -and $_.DriveLetter } |
                           Sort-Object DriveLetter
            }
            if (-not $usbVols) { throw "No USB-attached or removable drives detected." }
            Write-Host "Available USB or removable volumes:"
            $i = 1; $map = @{}
            foreach ($v in $usbVols) {
                $bus = try { (Get-Partition -DriveLetter $v.DriveLetter | Get-Disk).BusType } catch { "Unknown" }
                $label = if ($v.FileSystemLabel) { $v.FileSystemLabel } else { "(no label)" }
                Write-Host ("  {0}) {1}:\  {2}  [{3}]" -f $i,$v.DriveLetter,$label,$bus)
                $map[$i] = "$($v.DriveLetter):\"
                $i++
            }
            $pick = [int](Read-Host "Pick a number")
            if (-not $map.ContainsKey($pick)) { throw "Invalid selection." }
            $basePath = $map[$pick]
            
            # List available VM export folders
            $exportDirs = Get-ChildItem -Path $basePath -Directory | Where-Object { $_.Name -match "Export" }
            if (-not $exportDirs) { 
                $sub = Read-Host "Enter subfolder path under $basePath"
                Join-Path $basePath $sub
            } else {
                Write-Host "Available VM exports on $basePath"
                $dirMap = @{}
                $j = 1
                foreach ($dir in $exportDirs) {
                    Write-Host ("  {0}) {1}" -f $j, $dir.Name)
                    $dirMap[$j] = $dir.FullName
                    $j++
                }
                $dirPick = [int](Read-Host "Pick export folder number, or 0 to enter custom path")
                if ($dirPick -eq 0) {
                    $sub = Read-Host "Enter custom subfolder path under $basePath"
                    Join-Path $basePath $sub
                } else {
                    if (-not $dirMap.ContainsKey($dirPick)) { throw "Invalid selection." }
                    $dirMap[$dirPick]
                }
            }
        }
        '3' {
            $unc = Read-Host "Enter UNC path to VM export folder like \\server\share\VMExports\MyVM-Export-..."
            if ($unc -notmatch '^(\\\\[^\\]+)\\') { throw "Invalid UNC path." }
            $unc
        }
        default { throw "Invalid choice. Run again and select 1, 2, or 3." }
    }
}

function Validate-VMExport { param([string]$Path)
    Throw-If (-not (Test-Path $Path)) "Source path '$Path' does not exist or is not accessible."
    
    # Look for VM configuration files
    $vmConfigDirs = Get-ChildItem -LiteralPath $Path -Directory -Recurse -ErrorAction SilentlyContinue | 
                    Where-Object { $_.Name -eq 'Virtual Machines' }
    Throw-If (-not $vmConfigDirs) "Invalid VM export: Missing 'Virtual Machines' folder in '$Path'."
    
    $vmxFiles = Get-ChildItem -LiteralPath $vmConfigDirs[0].FullName -Filter "*.vmcx" -ErrorAction SilentlyContinue
    if (-not $vmxFiles) {
        $vmxFiles = Get-ChildItem -LiteralPath $vmConfigDirs[0].FullName -Filter "*.exp" -ErrorAction SilentlyContinue
    }
    Throw-If (-not $vmxFiles) "No VM configuration files found in export."
    
    return $vmxFiles[0].FullName
}

function Get-ExportedVMInfo { param([string]$ConfigPath)
    try {
        # Try to extract VM info from config file path
        $vmDir = Split-Path $ConfigPath -Parent
        $exportRoot = Split-Path $vmDir -Parent
        
        # Look for clues about the original VM name
        $exportDirName = Split-Path $exportRoot -Leaf
        if ($exportDirName -match '(.+)-Export-\d{8}_\d{6}') {
            $vmNameGuess = $Matches[1]
        } else {
            $vmNameGuess = "ImportedVM"
        }
        
        # Check for VHD files to estimate size
        $vhdDirs = Get-ChildItem -LiteralPath $exportRoot -Directory -Recurse | 
                   Where-Object { $_.Name -match 'Virtual Hard Disks' }
        $totalSize = 0
        if ($vhdDirs) {
            $vhdFiles = Get-ChildItem -LiteralPath $vhdDirs[0].FullName -File -ErrorAction SilentlyContinue
            foreach ($vhd in $vhdFiles) { $totalSize += $vhd.Length }
        }
        
        [PSCustomObject]@{
            ConfigFile = $ConfigPath
            ExportRoot = $exportRoot
            EstimatedVMName = $vmNameGuess
            EstimatedSize = $totalSize
            EstimatedSizeGB = [math]::Round($totalSize/1GB,2)
        }
    } catch {
        throw "Failed to analyze VM export: $($_.Exception.Message)"
    }
}

function Choose-DestinationPath {
    $defaultPath = (Get-VMHost).VirtualMachinePath
    Write-Host ""
    Write-Host "Choose VM destination location:"
    Write-Host "  1) Default Hyper-V location ($defaultPath)"
    Write-Host "  2) Custom local path"
    Write-Host "  3) Browse available drives"
    $choice = Read-Host "Enter 1, 2, or 3"
    switch ($choice) {
        '1' { $defaultPath }
        '2' { Read-Host "Enter custom path for VM files" }
        '3' {
            $drives = Get-PSDrive -PSProvider FileSystem | Where-Object { $_.Free -gt 10GB } | Sort-Object Name
            Write-Host "Available drives with >10GB free:"
            $i = 1; $map = @{}
            foreach ($d in $drives) {
                $freeGB = [math]::Round($d.Free/1GB,1)
                Write-Host ("  {0}) {1}:\  Free: {2} GB" -f $i,$d.Name,$freeGB)
                $map[$i] = "$($d.Name):\"
                $i++
            }
            $pick = [int](Read-Host "Pick a drive number")
            if (-not $map.ContainsKey($pick)) { throw "Invalid selection." }
            $sub = Read-Host "Optional subfolder under $($map[$pick]) (blank = root)"
            if ($sub) { Join-Path $map[$pick] $sub } else { $map[$pick] }
        }
        default { throw "Invalid choice. Run again and select 1, 2, or 3." }
    }
}

function Ensure-Destination { param([string]$Path)
    if (-not (Test-Path $Path)) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction SilentlyContinue | Out-Null
    }
    Throw-If (-not (Test-Path $Path)) "Destination path '$Path' is not accessible."
}

function Ensure-FreeSpace { param([string]$DestPath,[int64]$NeededBytes)
    if ($DestPath.StartsWith("\\")) {
        Log-Info "UNC path detected. Free space check skipped."
        return
    }
    $root = (Get-Item -LiteralPath $DestPath).PSDrive
    $free = $root.Free
    $neededGB = [math]::Round($NeededBytes/1GB,2)
    $freeGB   = [math]::Round($free/1GB,2)
    Throw-If ($free -lt $NeededBytes) "Not enough free space on $($root.Name):\  Need ~${neededGB} GB, have ${freeGB} GB."
    Log-Info "Free space OK on $($root.Name):\  Need ~${neededGB} GB, have ${freeGB} GB."
}

function Check-VMNameConflict { param([string]$Name)
    $existingVM = Get-VM -Name $Name -ErrorAction SilentlyContinue
    if ($existingVM) {
        Write-Host "VM with name '$Name' already exists!" -ForegroundColor Yellow
        Write-Host "  Current VM State: $($existingVM.State)"
        Write-Host "  Current VM Path: $($existingVM.Path)"
        $response = Read-Host "Enter new name for imported VM, or 'cancel' to abort"
        if ($response -eq 'cancel') { throw "Import cancelled due to name conflict." }
        return $response
    }
    return $Name
}
#endregion Utilities

# Real Hyper-V import progress via CIM
$HvNs = 'root\virtualization\v2'
function Get-HvVmCim([string]$Name) {
    Get-CimInstance -Namespace $HvNs -ClassName Msvm_ComputerSystem -Filter ("ElementName='{0}'" -f $Name) -ErrorAction SilentlyContinue
}
function Get-ImportConcreteJob {
    Get-CimInstance -Namespace $HvNs -ClassName Msvm_ConcreteJob -ErrorAction SilentlyContinue |
        Where-Object { $_.Caption -match 'Import' -or $_.Description -match 'Import' -or $_.Name -match 'Import' } |
        Sort-Object StartTime -Descending | Select-Object -First 1
}
function Get-ConcreteJobById([string]$InstanceId) {
    Get-CimInstance -Namespace $HvNs -ClassName Msvm_ConcreteJob -ErrorAction SilentlyContinue |
        Where-Object { $_.InstanceID -eq $InstanceId } |
        Select-Object -First 1
}
function JobStateText([uint16]$State) {
    switch ($State) {
        2 {'New'} 3 {'Running'} 4 {'Suspended'} 7 {'Completed'} 8 {'Terminated'}
        9 {'Killed'} 10 {'Error'} default { "State $State" }
    }
}

# Interactive fallbacks if args omitted
if (-not $SourcePath) { $SourcePath = Choose-SourcePath }
if (-not $VMName) { $VMName = $null }  # will be determined from export
if (-not $DestinationPath) { $DestinationPath = $null }  # choose later if blank

# Steps
$stepsList = @(
    "Validating prerequisites",
    "Analyzing VM export",
    "Resolving destination and checking conflicts",
    "Estimating import size and free space", 
    "Importing VM",
    "Final validation and configuration"
)
$steps = $stepsList.Count
$completed = 0

try {
    Write-Progress -Activity "Importing VM" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    Ensure-Admin
    Ensure-PSVersion
    Ensure-ModuleHyperV
    
    # Initialize logging with temporary name
    if (-not $VMName) { $VMName = "VM-Import" }
    $logs = New-Loggers -BaseLogDir $LogPath
    Log-Info "Transcript: $($logs.Transcript)"
    Log-Info "Error log:  $($logs.ErrorLog)"

    # Step 2
    $completed = 1
    Write-Progress -Activity "Importing VM" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    $configFile = Validate-VMExport -Path $SourcePath
    $vmInfo = Get-ExportedVMInfo -ConfigPath $configFile
    Log-Info "Found VM export: $($vmInfo.ExportRoot)"
    Log-Info "Config file: $($vmInfo.ConfigFile)"
    Log-Info "Estimated original name: $($vmInfo.EstimatedVMName)"
    Log-Info "Estimated size: $($vmInfo.EstimatedSizeGB) GB"

    # Use estimated name if not provided
    if ($VMName -eq "VM-Import") { $VMName = $vmInfo.EstimatedVMName }

    # Step 3
    $completed = 2
    Write-Progress -Activity "Importing VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    $VMName = Check-VMNameConflict -Name $VMName
    if (-not $DestinationPath) { $DestinationPath = Choose-DestinationPath }
    Ensure-Destination -Path $DestinationPath

    # Set default VHD path if not specified
    if (-not $VhdDestinationPath) {
        $VhdDestinationPath = (Get-VMHost).VirtualHardDiskPath
    }
    Ensure-Destination -Path $VhdDestinationPath
    Log-Info "VM destination: $DestinationPath"
    Log-Info "VHD destination: $VhdDestinationPath"

    # Step 4
    $completed = 3
    Write-Progress -Activity "Importing VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    $spaceNeeded = [int64]($vmInfo.EstimatedSize * 1.1)  # Add 10% overhead
    Ensure-FreeSpace -DestPath $DestinationPath -NeededBytes $spaceNeeded
    if ($VhdDestinationPath -ne $DestinationPath) {
        Ensure-FreeSpace -DestPath $VhdDestinationPath -NeededBytes $vmInfo.EstimatedSize
    }

    # Step 5 - start import and show progress
    $completed = 4
    Write-Progress -Activity "Importing VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    Log-Info "Starting import with type: $ImportType"
    
    # Build import parameters
    $importParams = @{
        Path = $vmInfo.ConfigFile
        VhdDestinationPath = $VhdDestinationPath
        ErrorAction = 'Stop'
    }
    
    if ($ImportType -eq 'Copy') { $importParams.Copy = $true }
    if ($GenerateNewId) { $importParams.GenerateNewId = $true }
    if ($SnapshotFilePath -and (Test-Path $SnapshotFilePath)) { 
        $importParams.SnapshotFilePath = $SnapshotFilePath 
        Log-Info "Using snapshot file: $SnapshotFilePath"
    }

    $importJobPs = Start-Job -ScriptBlock {
        param($ImportParams, $DestPath)
        Import-Module Hyper-V | Out-Null
        $vm = Import-VM @ImportParams
        if ($DestPath -ne $vm.Path) {
            Move-VMStorage -VMName $vm.Name -DestinationStoragePath $DestPath -ErrorAction Stop
        }
        return $vm
    } -ArgumentList $importParams, $DestinationPath

    # Monitor import progress
    $cimJobInstanceId = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    $lastPct = -1
    
    while ($true) {
        $psState = (Get-Job -Id $importJobPs.Id).State

        # Try to find CIM import job
        if (-not $cimJobInstanceId) {
            $cimJob = Get-ImportConcreteJob
            if ($cimJob) { $cimJobInstanceId = $cimJob.InstanceID }
        } else {
            $cimJob = Get-ConcreteJobById -InstanceId $cimJobInstanceId
        }

        if ($cimJob) {
            $pct = [int]$cimJob.PercentComplete
            $stateText = JobStateText $cimJob.JobState
            if ($pct -ne $lastPct) {
                Write-Progress -Activity "Importing VM '$VMName'" -Status "$stateText... $pct% complete (Hyper-V)" -PercentComplete $pct
                $lastPct = $pct
            }
            if ($cimJob.JobState -in 7,8,9,10) {
                if ($cimJob.JobState -ne 7) {
                    throw "Hyper-V import job ended in state '$stateText'."
                }
                break
            }
        } else {
            # Fallback: show estimated progress
            $elapsed = $sw.Elapsed.TotalMinutes
            $estimatedTotal = ($vmInfo.EstimatedSizeGB / 2)  # Rough estimate: 2GB per minute
            $pct = if ($estimatedTotal -gt 0) { [int][math]::Min(($elapsed / $estimatedTotal) * 100, 95) } else { 50 }
            if ($pct -ne $lastPct) {
                Write-Progress -Activity "Importing VM '$VMName'" -Status "Importing (est.)... $pct% complete" -PercentComplete $pct
                $lastPct = $pct
            }
        }

        if ($psState -in 'Failed','Stopped') { break }
        if ($psState -eq 'Completed' -and (-not $cimJob -or $cimJob.JobState -eq 7)) { break }
        Start-Sleep -Seconds 2
    }

    $importedVM = Receive-Job -Id $importJobPs.Id -ErrorAction Stop
    Remove-Job -Id $importJobPs.Id -Force | Out-Null

    # Step 6
    $completed = 5
    Write-Progress -Activity "Importing VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    
    # Rename VM if needed
    if ($importedVM.Name -ne $VMName) {
        Log-Info "Renaming VM from '$($importedVM.Name)' to '$VMName'"
        Rename-VM -VM $importedVM -NewName $VMName
        $importedVM = Get-VM -Name $VMName
    }

    # Final validation
    $finalVM = Get-VM -Name $VMName -ErrorAction Stop
    $vmDisks = Get-VMHardDiskDrive -VMName $VMName
    
    Write-Progress -Activity "Importing VM '$VMName'" -Completed -Status "Done"
    Log-Info "Import complete!"
    Log-Info "VM Name: $($finalVM.Name)"
    Log-Info "VM State: $($finalVM.State)"  
    Log-Info "VM Path: $($finalVM.Path)"
    Log-Info "Generation: $($finalVM.Generation)"
    Log-Info "Virtual Hard Disks: $($vmDisks.Count)"
    foreach ($disk in $vmDisks) {
        Log-Info "  - $($disk.Path)"
    }

} catch {
    $msg = $_.Exception.Message
    if ($logs -and $logs.ErrorLog) { Log-Error -Message $msg -ErrorLog $logs.ErrorLog }
    else { Write-Error $msg }
} finally {
    try { Stop-Transcript | Out-Null } catch {}
}
