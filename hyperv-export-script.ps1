# =====================================================================
# Modified for public use
# Export-HyperVVM.ps1 - v6 (fixed ordering + real progress + policy prompt)
# =====================================================================

[CmdletBinding()]
param(
    [string]$VMName,
    [string]$DestinationPath,
    [ValidateSet('Shutdown','Save','Online')]
    [string]$Mode = 'Shutdown',
    [string]$LogPath
)

# Execution Policy assistance (must be AFTER param in a script)
try {
    if (-not $env:__HV_EXPORT_POLICY_AUDITED) {
        $env:__HV_EXPORT_POLICY_AUDITED = '1'
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
        $BaseLogDir = Join-Path ([Environment]::GetFolderPath('MyDocuments')) "HyperV-Exports\Logs"
    }
    New-Item -Path $BaseLogDir -ItemType Directory -Force | Out-Null
    $ts = Get-Date -Format "yyyyMMdd_HHmmss"
    $Transcript = Join-Path $BaseLogDir "Export-$($VMName)-$ts-transcript.log"
    $ErrorLog   = Join-Path $BaseLogDir "Export-$($VMName)-$ts-error.log"
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

function Choose-Destination {
    Write-Host ""
    Write-Host "Choose destination type:"
    Write-Host "  1) Local folder (e.g. C:\Exports or D:\Exports)"
    Write-Host "  2) USB-attached drive (SSD/HDD/flash)"
    Write-Host "  3) UNC network share (\\server\share\path)"
    $choice = Read-Host "Enter 1, 2, or 3"
    switch ($choice) {
        '1' { Read-Host "Enter local folder path" }
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
                Write-Host ("  {0}) {1}:\  {2}  Free: {3:N2} GB  [{4}]" -f $i,$v.DriveLetter,$label,($v.SizeRemaining/1GB),$bus)
                $map[$i] = "$($v.DriveLetter):\"
                $i++
            }
            $pick = [int](Read-Host "Pick a number")
            if (-not $map.ContainsKey($pick)) { throw "Invalid selection." }
            $sub = Read-Host "Optional subfolder under $($map[$pick]) (blank = root)"
            if ($sub) { Join-Path $map[$pick] $sub } else { $map[$pick] }
        }
        '3' {
            $unc = Read-Host "Enter UNC path like \\server\share\exports"
            if ($unc -notmatch '^(\\\\[^\\]+)\\') { throw "Invalid UNC path." }
            $unc
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

function Get-VMSizeEstimateBytes { param([string]$Name)
    $total = 0
    $disks = Get-VMHardDiskDrive -VMName $Name -ErrorAction Stop
    foreach ($d in $disks) {
        if ($d.Path -and (Test-Path $d.Path)) { $total += (Get-Item $d.Path).Length }
    }
    [int64]($total * 1.05)  # add 5 percent for config and metadata
}

function Get-FolderSizeBytes { param([string]$Path)
    if (-not (Test-Path $Path)) { return 0 }
    $s = 0
    Get-ChildItem -LiteralPath $Path -Recurse -Force -ErrorAction SilentlyContinue |
        ForEach-Object { if (-not $_.PSIsContainer) { $s += $_.Length } }
    [int64]$s
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

function Wait-VMOff { param([string]$Name,[int]$TimeoutSec=300)
    $sw = [Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $TimeoutSec) {
        if ((Get-VM -Name $Name).State -eq 'Off') { return $true }
        Start-Sleep -Milliseconds 500
    }
    return $false
}

function Prepare-VM { param([string]$Name,[string]$Mode,[string]$ErrorLog)
    $vm = Get-VM -Name $Name -ErrorAction Stop
    switch ($Mode) {
        'Shutdown' {
            if ($vm.State -ne 'Off') {
                Log-Info "Requesting guest shutdown for '$Name'..."
                try { Stop-VMGuest -Name $Name -ErrorAction Stop | Out-Null } catch { }
                if (-not (Wait-VMOff -Name $Name -TimeoutSec 240)) {
                    Log-Info "Guest shutdown did not complete in time. Forcing power off..."
                    Stop-VM -Name $Name -TurnOff -Force -ErrorAction Stop
                    if (-not (Wait-VMOff -Name $Name -TimeoutSec 60)) {
                        throw "Failed to power off VM '$Name'."
                    }
                }
            }
        }
        'Save' {
            if ($vm.State -eq 'Running') {
                Log-Info "Saving VM '$Name' state..."
                Save-VM -Name $Name -ErrorAction Stop
            }
        }
        'Online' {
            Log-Info "Proceeding with online export for '$Name'. Shutdown is safer for consistency."
        }
    }
}
#endregion Utilities

# Real Hyper-V export progress via CIM
$HvNs = 'root\virtualization\v2'
function Get-HvVmCim([string]$Name) {
    Get-CimInstance -Namespace $HvNs -ClassName Msvm_ComputerSystem -Filter ("ElementName='{0}'" -f $Name)
}
function Get-ExportConcreteJobForVm([string]$Name) {
    $vmCim = Get-HvVmCim -Name $Name
    if (-not $vmCim) { return $null }
    Get-CimAssociatedInstance -Namespace $HvNs -InputObject $vmCim -Association Msvm_AffectedJobElement -ErrorAction SilentlyContinue |
        Where-Object { $_.CimClass.CimClassName -eq 'Msvm_ConcreteJob' } |
        Where-Object { $_.Caption -match 'Export' -or $_.Description -match 'Export' -or $_.Name -match 'Export' } |
        Select-Object -First 1
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
if (-not $VMName) { $VMName = Read-Host "Enter VM name to export" }
if (-not $DestinationPath) { $DestinationPath = $null }  # choose later if blank

# Steps
$stepsList = @(
    "Validating prerequisites",
    "Resolving destination",
    "Estimating export size and free space",
    "Preparing VM per selected mode",
    "Exporting VM",
    "Final validation"
)
$steps = $stepsList.Count
$completed = 0

try {
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    Ensure-Admin
    Ensure-PSVersion
    Ensure-ModuleHyperV
    $logs = New-Loggers -BaseLogDir $LogPath
    Log-Info "Transcript: $($logs.Transcript)"
    Log-Info "Error log:  $($logs.ErrorLog)"

    $vm = Get-VM -Name $VMName -ErrorAction Stop
    Log-Info "Found VM: $($vm.Name)  Generation: $($vm.Generation)  State: $($vm.State)"

    # Step 2
    $completed = 1
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    if (-not $DestinationPath) { $DestinationPath = Choose-Destination }
    Ensure-Destination -Path $DestinationPath
    $exportRoot = Join-Path $DestinationPath "$($VMName)-Export-$(Get-Date -Format yyyyMMdd_HHmmss)"
    New-Item -ItemType Directory -Path $exportRoot -Force | Out-Null
    Log-Info "Export folder: $exportRoot"

    # Step 3
    $completed = 2
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    $estimate = Get-VMSizeEstimateBytes -Name $VMName
    $estimateGB = [math]::Max([math]::Round($estimate/1GB,2), 1)
    Log-Info "Estimated export size: ~${estimateGB} GB"
    Ensure-FreeSpace -DestPath $exportRoot -NeededBytes $estimate

    # Step 4
    $completed = 3
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    Prepare-VM -Name $VMName -Mode $Mode -ErrorLog $logs.ErrorLog

    # Step 5 - start export and show actual Hyper-V progress
    $completed = 4
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    Log-Info "Starting export..."
    $exportJobPs = Start-Job -ScriptBlock {
        param($VMName,$ExportPath)
        Import-Module Hyper-V | Out-Null
        Export-VM -Name $VMName -Path $ExportPath -ErrorAction Stop
    } -ArgumentList $VMName, $exportRoot

    # capture the CIM job that represents the export
    $cimJobInstanceId = $null
    $sw = [Diagnostics.Stopwatch]::StartNew()
    do {
        $cimJob = Get-ExportConcreteJobForVm -Name $VMName
        if ($cimJob) { $cimJobInstanceId = $cimJob.InstanceID; break }
        Start-Sleep -Milliseconds 400
    } while ($sw.Elapsed.TotalSeconds -lt 15)

    $lastPct = -1
    while ($true) {
        $psState = (Get-Job -Id $exportJobPs.Id).State

        if ($cimJobInstanceId) {
            $cimJob = Get-ConcreteJobById -InstanceId $cimJobInstanceId
        } else {
            $cimJob = Get-ExportConcreteJobForVm -Name $VMName
            if ($cimJob) { $cimJobInstanceId = $cimJob.InstanceID }
        }

        if ($cimJob) {
            $pct = [int]$cimJob.PercentComplete
            $stateText = JobStateText $cimJob.JobState
            if ($pct -ne $lastPct) {
                Write-Progress -Activity "Exporting VM '$VMName'" -Status "$stateText... $pct% complete (Hyper-V)" -PercentComplete $pct
                $lastPct = $pct
            }
            if ($cimJob.JobState -in 7,8,9,10) {
                if ($cimJob.JobState -ne 7) {
                    throw "Hyper-V export job ended in state '$stateText'."
                }
                break
            }
        } else {
            # Fallback: show estimated progress by folder growth
            $sizeNow = Get-FolderSizeBytes -Path $exportRoot
            $pct = if ($estimate -gt 0) { [int][math]::Min(([double]$sizeNow / [double]$estimate) * 100, 99) } else { 50 }
            if ($pct -ne $lastPct) {
                Write-Progress -Activity "Exporting VM '$VMName'" -Status "Exporting (est.)... $pct% complete" -PercentComplete $pct
                $lastPct = $pct
            }
        }

        if ($psState -in 'Failed','Stopped') { break }
        if ($psState -eq 'Completed' -and -not $cimJob) { break }
        Start-Sleep -Seconds 1
    }

    Receive-Job -Id $exportJobPs.Id -ErrorAction Stop | Out-Null
    Remove-Job -Id $exportJobPs.Id -Force | Out-Null

    # Step 6
    $completed = 5
    Write-Progress -Activity "Exporting VM '$VMName'" -Status $stepsList[$completed] -PercentComplete (($completed/$steps)*100)
    $configDirs = Get-ChildItem -LiteralPath $exportRoot -Directory -Recurse | Where-Object { $_.Name -match 'Virtual Machines' }
    if (-not $configDirs) { throw "Export appears incomplete. Missing 'Virtual Machines' folder." }

    Write-Progress -Activity "Exporting VM '$VMName'" -Completed -Status "Done"
    Log-Info "Export complete. Folder: $exportRoot"

} catch {
    $msg = $_.Exception.Message
    if ($logs -and $logs.ErrorLog) { Log-Error -Message $msg -ErrorLog $logs.ErrorLog }
    else { Write-Error $msg }
} finally {
    try { Stop-Transcript | Out-Null } catch {}
}
