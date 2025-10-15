<#
.NOTES
CyberPipe.ps1  
https://github.com/dwmetz/CyberPipe  
Formerly known as "CSIRT-Collect"  
Author: @dwmetz  

.FUNCTIONALITY
This script performs the following actions:  
- Captures a memory image using DumpIt (Windows x86/x64/ARM64) or Magnet RAM Capture on legacy systems  
- Captures a triage snapshot using MAGNET Response  
- Checks for encrypted volumes  
- Recovers the active BitLocker recovery key (if available)  
- Saves all artifacts, logs, and outputs to USB or designated network path  

.SYNOPSIS
CyberPipe v5.2

**Prerequisites (must be present in the \Tools directory):**  
- [MAGNET Response](https://magnetforensics.com) — `MagnetRESPONSE.exe`
- [Encrypted Disk Detector](https://www.magnetforensics.com/resources/encrypted-disk-detector/) — `EDDv310.exe`
- `CyberPipe5.ps1` should be located adjacent to the `\Tools` directory (on USB or network share)  

**Usage:**  
- Launch PowerShell as Administrator  
- Run `.\CyberPipe.ps1`

.EXAMPLE
.\CyberPipe.ps1
# Runs the default full triage profile with memory, pagefile, volatile data, and system files

.\CyberPipe.ps1 -CollectionProfile RAMOnly
# Captures only RAM and exits

.\CyberPipe.ps1 -CollectionProfile Volatile
# Captures only volatile data (network, registry hives, etc.)

.\CyberPipe.ps1 -CollectionProfile QuickTriage
# Fast triage collection (volatile + system files, no RAM) - completes in ~2 minutes

.\CyberPipe.ps1 -Compress
# Run full triage and compress output to ZIP file

.\CyberPipe.ps1 -Net "\\server\share"
# Run collection to network share instead of local USB drive

.\CyberPipe.ps1 -Net "\\server\share" -CollectionProfile QuickTriage -Compress
# Network collection with specific profile and compression

.NOTES
Virtual Environment Detection:
The script automatically detects if running in a VM (VMware, Hyper-V, VirtualBox, etc.).
This is important because:
- VM memory dumps may not capture hypervisor-level malware
- Memory overcommitment can affect collection completeness
- Nested virtualization may hide attacker infrastructure
- Analysts need to know environment limitations when interpreting results
#>
param (
    [switch]$Elevated,
    [ValidateSet("Volatile","RAMOnly","RAMPage","RAMSystem","QuickTriage","")]
    [string]$CollectionProfile = $env:CYBERPIPE_PROFILE,
    [switch]$Compress,
    [string]$Net = ""
)
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ((Test-Admin) -eq $false)  {
    if ($elevated) {
    } else {
        Write-host " "
        Write-host  "CyberPipe requires Admin permissions (not detected). Exiting."   
        Write-host  " "    
    }
    exit
}
[console]::ForegroundColor="Cyan"
Clear-Host
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "        ╔══════════════════════════════════════════════════╗"
Write-Host "        ║                                                  ║"
Write-Host "        ║     ██████╗██╗   ██╗██████╗ ███████╗██████╗      ║"
Write-Host "        ║    ██╔════╝╚██╗ ██╔╝██╔══██╗██╔════╝██╔══██╗     ║"
Write-Host "        ║    ██║      ╚████╔╝ ██████╔╝█████╗  ██████╔╝     ║"
Write-Host "        ║    ██║       ╚██╔╝  ██╔══██╗██╔══╝  ██╔══██╗     ║"
Write-Host "        ║    ╚██████╗   ██║   ██████╔╝███████╗██║  ██║     ║"
Write-Host "        ║     ╚═════╝   ╚═╝   ╚═════╝ ╚══════╝╚═╝  ╚═╝     ║"
Write-Host "        ║                                                  ║"
Write-Host "        ║           ██████╗ ██╗██████╗ ███████╗            ║"
Write-Host "        ║           ██╔══██╗██║██╔══██╗██╔════╝            ║"
Write-Host "        ║           ██████╔╝██║██████╔╝█████╗              ║"
Write-Host "        ║           ██╔═══╝ ██║██╔═══╝ ██╔══╝              ║"
Write-Host "        ║           ██║     ██║██║     ███████╗            ║"
Write-Host "        ║           ╚═╝     ╚═╝╚═╝     ╚══════╝            ║"
Write-Host "        ║                                                  ║"
Write-Host "        ║          Incident Response Collection            ║"
Write-Host "        ║                     v5.2                         ║"
Write-Host "        ║                                                  ║"
Write-Host "        ╚══════════════════════════════════════════════════╝"
Write-Host ""
Write-Host "           Memory • Triage • Forensics • Chain of Custody"
Write-Host ""
Sleep 2
Write-Host "CyberPipe IR Collection Script v5.2"
Write-Host "https://github.com/dwmetz/CyberPipe"
Write-Host "$([char]0x00A9)2025 @dwmetz | bakerstreetforensics.com"
Write-Host ""
[console]::ForegroundColor="DarkCyan"
## Network Collection Handling
if ($Net) {
    Write-Host -Fore Cyan "Network mode enabled. Mapping drive to $Net..."
    $Networkpath = "Z:\"

    If (Test-Path -Path $Networkpath) {
        Write-Host "Drive Z: already mapped."
    }
    Else {
        try {
            (New-Object -ComObject WScript.Network).MapNetworkDrive("Z:", $Net)
            Start-Sleep -Seconds 2

            If (Test-Path -Path $Networkpath) {
                Write-Host -Fore Green "Drive mapped successfully to $Net"
            }
            Else {
                Write-Host -Fore Red "Error: Drive mapping appeared to succeed but path not accessible."
                exit 1
            }
        }
        catch {
            Write-Host -Fore Red "Error mapping network drive: $($_.Exception.Message)"
            exit 1
        }
    }

    Set-Location $Networkpath
    Write-Host -Fore Cyan "Working from network location: Z:\"
}
## Below is for USB and Network:
$wd = Get-Location
$tstamp = (Get-Date -Format "yyyyMMddHHmm")
$outputpath = "$wd\Collections\$env:COMPUTERNAME-$tstamp"
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
If (Test-Path -Path $wd\Tools) {
}
Else {
        Write-Host " "
        Write-Host -For DarkCyan "Tools directory not present."
        Write-Host " "
        exit 1

    }
    
If (Test-Path -Path Collections) {
    Write-Host  "Collections directory exists."
}
Else {
    $null = mkdir Collections
    If (Test-Path -Path Collections) {
        Write-Host  "Collection directory created."
    }
    Else {
        Write-Host -For DarkCyan "Error creating directory."
        exit 1
    }
}
Set-Location Collections
If (Test-Path -Path $outputpath) {
    Write-Host  "Host directory already exists."
}
Else {
    $null = mkdir $outputpath
    If (Test-Path -Path $outputpath) {
        Write-Host  "Host directory created."
    }
    Else {
        Write-Host -For DarkCyan "Error creating directory."
        exit 1
    }
}

# Validate required tools exist
$requiredTools = @(
    "$wd\Tools\MagnetRESPONSE.exe",
    "$wd\Tools\EDDv310.exe"
)

foreach ($tool in $requiredTools) {
    if (-not (Test-Path $tool)) {
        Write-Host -Fore Red "Required tool not found: $tool"
        exit 1
    }
}

# Check available disk space on both target drive AND system drive (C:)
$targetDrive = (Get-Item $outputpath).PSDrive
$targetFreeSpaceGB = [math]::Round((Get-PSDrive $targetDrive.Name).Free / 1GB, 2)
$systemFreeSpaceGB = [math]::Round((Get-PSDrive C).Free / 1GB, 2)

# Get system RAM to estimate space needed (RAM capture needs ~RAM size in temp space)
$totalRAM_GB = [math]::Round((Get-CimInstance Win32_ComputerSystem).TotalPhysicalMemory / 1GB, 2)

# Calculate required space based on collection profile
# For RAM collections: need space equal to RAM + 5GB overhead
# For non-RAM collections: 10GB minimum
$targetRequired = 10
$systemRequired = 10

# Check if this profile includes RAM capture
# Only default, RAMOnly, RAMPage, and RAMSystem actually capture memory
if ($CollectionProfile -eq "" -or $CollectionProfile -match "^RAM") {
    # RAM-based or default profile - need more system space
    $systemRequired = [math]::Max(($totalRAM_GB + 5), 10)
    $targetRequired = [math]::Max(($totalRAM_GB + 10), 15)
}

Write-Host "System RAM: $totalRAM_GB GB"
Write-Host "Target drive ($($targetDrive.Name):) free space: $targetFreeSpaceGB GB (need $targetRequired GB)"
Write-Host "System drive (C:) free space: $systemFreeSpaceGB GB (need $systemRequired GB)"

if ($targetFreeSpaceGB -lt $targetRequired) {
    Write-Host -Fore Red "Insufficient space on target drive. Available: $targetFreeSpaceGB GB, Required: $targetRequired GB"
    exit 1
}

if ($systemFreeSpaceGB -lt $systemRequired) {
    Write-Host -Fore Red "Insufficient space on system drive (C:). Available: $systemFreeSpaceGB GB, Required: $systemRequired GB"
    Write-Host -Fore Red "MAGNET Response requires approximately RAM-sized space on C: for temporary files during memory capture."
    Write-Host -Fore Yellow "To proceed anyway, use -CollectionProfile Volatile to skip RAM capture."
    exit 1
}

### Pre-Collection Volatile Snapshot (stored in memory for later inclusion in report)
Write-Host -Fore Cyan "Capturing pre-collection volatile snapshot..."
$collection = "$env:COMPUTERNAME-$tstamp"
$preCollectionTime = Get-Date
$snapshotOutput = @()
$snapshotOutput += "=== PRE-COLLECTION VOLATILE SNAPSHOT ==="
$snapshotOutput += "Captured: $($preCollectionTime.ToString())"
$snapshotOutput += ""

# System Uptime
$osInfo = Get-CimInstance Win32_OperatingSystem
$uptime = (Get-Date) - $osInfo.LastBootUpTime
$snapshotOutput += "--- SYSTEM UPTIME ---"
$snapshotOutput += "Last Boot: $($osInfo.LastBootUpTime)"
$snapshotOutput += "Uptime: $($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes"
$snapshotOutput += ""

# Detect Virtual Environment
# Important for analysis: VMs may have memory overcommitment, nested malware, or hypervisor-level threats
# that won't be captured in guest memory dumps. Knowing the environment helps analysts understand
# collection limitations and potential blind spots.
$snapshotOutput += "--- VIRTUALIZATION DETECTION ---"
$computerSystem = Get-CimInstance Win32_ComputerSystem
$snapshotOutput += "Model: $($computerSystem.Model)"
$snapshotOutput += "Manufacturer: $($computerSystem.Manufacturer)"
if ($computerSystem.Model -match "Virtual|VMware|VirtualBox|Hyper-V|QEMU|Xen") {
    $snapshotOutput += "Virtual Environment: DETECTED ($($computerSystem.Model))"
    $snapshotOutput += "Note: VM memory dumps may not capture hypervisor-level activity or overcommitted memory"
} else {
    $snapshotOutput += "Virtual Environment: Physical or Unknown"
}
$snapshotOutput += ""

# Logged-on Users
$snapshotOutput += "--- LOGGED-ON USERS ---"
try {
    $loggedOnUsers = Get-CimInstance Win32_LoggedOnUser -ErrorAction Stop |
        Select-Object -ExpandProperty Antecedent |
        Select-Object -Unique Domain, Name
    foreach ($user in $loggedOnUsers) {
        $snapshotOutput += "$($user.Domain)\$($user.Name)"
    }
} catch {
    $snapshotOutput += "Unable to enumerate logged-on users"
}
$snapshotOutput += ""

# Active Network Connections
$snapshotOutput += "--- ACTIVE NETWORK CONNECTIONS ---"
try {
    $connections = Get-NetTCPConnection -State Established -ErrorAction Stop |
        Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess
    foreach ($conn in $connections) {
        $processName = (Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue).Name
        $snapshotOutput += "$($conn.LocalAddress):$($conn.LocalPort) -> $($conn.RemoteAddress):$($conn.RemotePort) [$processName PID:$($conn.OwningProcess)]"
    }
} catch {
    $snapshotOutput += "Unable to enumerate network connections"
}
$snapshotOutput += ""

# Running Processes (top 20 by memory)
$snapshotOutput += "--- TOP PROCESSES (by memory) ---"
try {
    $processes = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 20
    foreach ($proc in $processes) {
        $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 2)
        $snapshotOutput += "$($proc.Name) [PID:$($proc.Id)] - $memMB MB"
    }
} catch {
    $snapshotOutput += "Unable to enumerate processes"
}

Write-Host -Fore Cyan "Pre-collection snapshot captured (will be included in final report)"
Write-Host ""

### Collection Profiles
switch ($CollectionProfile) {
    "Volatile" {
        $profileName = "Volatile"
        $arguments = "/capturevolatile"
    }
    "RAMSystem" {
        $profileName = "RAM & Critical System Files"
        $arguments = "/captureram /capturesystemfiles"
    }
    "RAMPage" {
        $profileName = "RAM & Pagefile"
        $arguments = "/captureram /capturepagefile"
    }
    "RAMOnly" {
        $profileName = "RAM Dump"
        $arguments = "/captureram"
    }
    "QuickTriage" {
        $profileName = "Quick Triage"
        $arguments = "/capturevolatile /capturesystemfiles"
    }
    default {
        $profileName = "MAGNET Triage"
        $arguments = "/captureram /capturepagefile /capturevolatile /capturesystemfiles"
    }
}

Write-Host ""
$global:progressPreference = 'silentlyContinue'
Write-Host -Fore Cyan  "
Running MAGNET Response...
"
Write-Host ""
Write-Host  "Magnet RESPONSE v1.7
$([char]0x00A9)2021-2024 Magnet Forensics Inc
"
$OS = $(((Get-CimInstance Win32_OperatingSystem).Caption).split('|')[0])
$arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
$name = (Get-CimInstance Win32_OperatingSystem).CSName
Write-host  "
Hostname:           $name
Operating System:   $OS
Architecture:       $arch
Selected Profile:   $profileName
Output Directory:   $outputpath
"
Write-Host -Fore Cyan "
Collecting Artifacts...
"

# Build argument string properly
$magnetArgs = "/accepteula /unattended /silent /caseref:CyberPipe /output:`"$outputpath`" $arguments"
$magnetProcess = Start-Process -FilePath "$wd\Tools\MagnetRESPONSE.exe" -ArgumentList $magnetArgs -PassThru -NoNewWindow

# Progress indicator while MAGNET Response runs
$elapsed = 0
while (-not $magnetProcess.HasExited) {
    Start-Sleep -Seconds 5
    $elapsed += 5
    $minutes = [math]::Floor($elapsed / 60)
    $seconds = $elapsed % 60

    # Show elapsed time and check output folder size (including subfolders)
    try {
        # Force refresh of directory to avoid cached results
        $collectionSize = 0
        Get-ChildItem -Path $outputpath -Recurse -File -Force -ErrorAction SilentlyContinue | ForEach-Object {
            # Force file info refresh by accessing FileInfo directly
            $fileInfo = [System.IO.FileInfo]::new($_.FullName)
            $collectionSize += $fileInfo.Length
        }

        if ($collectionSize -gt 0) {
            # Use MB for collections under 1GB, GB for larger collections
            if ($collectionSize -lt 1GB) {
                $sizeMB = [math]::Round($collectionSize / 1MB, 1)
                Write-Host "`r  [Running: $minutes min $seconds sec | Collected: $sizeMB MB]" -Fore DarkCyan -NoNewline
            } else {
                $sizeGB = [math]::Round($collectionSize / 1GB, 2)
                Write-Host "`r  [Running: $minutes min $seconds sec | Collected: $sizeGB GB]" -Fore DarkCyan -NoNewline
            }
        } else {
            Write-Host "`r  [Running: $minutes min $seconds sec | Collected: 0.0 MB]" -Fore DarkCyan -NoNewline
        }
    }
    catch {
        Write-Host "`r  [Running: $minutes min $seconds sec]" -Fore DarkCyan -NoNewline
    }
}
Write-Host ""  # New line after progress completes

$magnetProcess.WaitForExit()

if ($magnetProcess.ExitCode -ne 0) {
    Write-Host -Fore Red "MAGNET Response failed with exit code: $($magnetProcess.ExitCode)"
    exit 1
}

$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host -Fore Cyan  "** Magnet RESPONSE Completed in $Minutes minutes and $Seconds seconds. **
"
Write-Host -Fore Cyan "Running Encrypted Disk Detector (EDD)...
"
$collection = "$env:COMPUTERNAME-$tstamp"
$eddTempFile = "$outputpath\$collection-edd-temp.txt"
$eddProcess = Start-Process -FilePath "$wd\Tools\EDDv310.exe" -ArgumentList "/batch" -RedirectStandardOutput $eddTempFile -PassThru -Wait -NoNewWindow

if ($eddProcess.ExitCode -ne 0) {
    Write-Host -Fore Yellow "Warning: EDD exited with code $($eddProcess.ExitCode)"
}

Start-Sleep 1
$eddOutput = Get-Content $eddTempFile
$eddOutput | ForEach-Object { Write-Host $_ }
Write-Host -Fore Cyan  "
Checking for BitLocker Keys...
"
# Get all BitLocker volumes, not just C:
$bitlockerVolumes = Get-BitLockerVolume | Where-Object { $_.ProtectionStatus -eq "On" }

$keyOutput = @()
if ($bitlockerVolumes.Count -eq 0) {
    Write-Host -Fore Yellow "No BitLocker protected volumes found."
    $keyOutput += "No BitLocker protected volumes found on $env:computername"
}
else {
    foreach ($volume in $bitlockerVolumes) {
        $keyOutput += "Volume: $($volume.MountPoint)"
        $keyOutput += "Protection Status: $($volume.ProtectionStatus)"

        if ($volume.KeyProtector.Count -gt 0) {
            $keyOutput += "Key Protectors:"
            foreach ($kp in $volume.KeyProtector) {
                $keyOutput += "  Type: $($kp.KeyProtectorType)"
                if ($kp.RecoveryPassword) {
                    $keyOutput += "  Recovery Password: $($kp.RecoveryPassword)"
                }
                if ($kp.KeyProtectorId) {
                    $keyOutput += "  Key ID: $($kp.KeyProtectorId)"
                }
            }
            Write-Host -Fore Cyan "BitLocker key(s) recovered for volume $($volume.MountPoint)"
        }
        else {
            $keyOutput += "No key protectors found for this volume."
            Write-Host -Fore Yellow "No key protectors for volume $($volume.MountPoint)"
        }
        $keyOutput += ""
    }
}
Set-Location ~
$StopWatch.Stop()
$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host -Fore Cyan  "
*** Collection Completed in $Minutes minutes and $Seconds seconds. ***
"

# Generate Comprehensive CyberPipe Report
Write-Host -Fore Cyan "Generating CyberPipe collection report..."
$reportFile = "$outputpath\CyberPipe-Report.txt"
$reportOutput = @()

# Header
$reportOutput += "=" * 80
$reportOutput += "CYBERPIPE INCIDENT RESPONSE COLLECTION REPORT"
$reportOutput += "=" * 80
$reportOutput += ""
$reportOutput += "Host: $env:COMPUTERNAME"
$reportOutput += "Collection Profile: $profileName"
$reportOutput += "Collection Started: $($preCollectionTime.ToString())"
$reportOutput += "Collection Completed: $((Get-Date).ToString())"
$reportOutput += "Duration: $Minutes minutes, $Seconds seconds"
$reportOutput += "Generated by: CyberPipe v5.2"
$reportOutput += "https://github.com/dwmetz/CyberPipe"
$reportOutput += ""
$reportOutput += "=" * 80

# Pre-Collection Volatile Snapshot
$reportOutput += ""
$reportOutput += $snapshotOutput
$reportOutput += ""
$reportOutput += "=" * 80

# Collection Summary
$reportOutput += ""
$reportOutput += "=== COLLECTION SUMMARY ==="
$reportOutput += ""
$allFiles = Get-ChildItem -Path $outputpath -Recurse -File
$totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeGB = [math]::Round($totalSize / 1GB, 2)
$fileCount = $allFiles.Count

$reportOutput += "Total Files Collected: $fileCount"
$reportOutput += "Total Collection Size: $totalSizeGB GB"
$reportOutput += "System RAM: $totalRAM_GB GB"
$reportOutput += "Uptime at Collection: $($uptime.Days) days, $($uptime.Hours) hours"
$reportOutput += "Virtual Environment: $(if ($computerSystem.Model -match 'Virtual|VMware|VirtualBox|Hyper-V|QEMU|Xen') { $computerSystem.Model } else { 'Physical/Unknown' })"
$reportOutput += ""

# Files by Type
$reportOutput += "--- FILES BY TYPE ---"
$filesByType = $allFiles | Group-Object Extension | Sort-Object Count -Descending
foreach ($group in $filesByType) {
    $ext = if ($group.Name) { $group.Name } else { "(no extension)" }
    $groupSize = ($group.Group | Measure-Object -Property Length -Sum).Sum
    $groupSizeGB = [math]::Round($groupSize / 1GB, 3)
    $reportOutput += "$ext : $($group.Count) files ($groupSizeGB GB)"
}
$reportOutput += ""

# Key Artifacts
$reportOutput += "--- KEY ARTIFACTS ---"
$keyArtifacts = $allFiles | Where-Object { $_.Name -match "(\.raw|\.mem|\.dmp|\.txt|\.json|\.csv)" }
foreach ($artifact in $keyArtifacts) {
    if ($artifact.Length -lt 1MB) {
        $sizeKB = [math]::Round($artifact.Length / 1KB, 2)
        $reportOutput += "$($artifact.Name) - $sizeKB KB"
    } else {
        $sizeMB = [math]::Round($artifact.Length / 1MB, 2)
        $reportOutput += "$($artifact.Name) - $sizeMB MB"
    }
}
$reportOutput += ""
$reportOutput += "=" * 80

# Encrypted Disk Detection
$reportOutput += ""
$reportOutput += "=== ENCRYPTED DISK DETECTION ==="
$reportOutput += ""
$reportOutput += $eddOutput
$reportOutput += ""
$reportOutput += "=" * 80

# BitLocker Recovery Keys
$reportOutput += ""
$reportOutput += "=== BITLOCKER RECOVERY KEYS ==="
$reportOutput += ""
$reportOutput += $keyOutput
$reportOutput += ""
$reportOutput += "=" * 80

# SHA256 Hashes
$reportOutput += ""
$reportOutput += "=== SHA256 INTEGRITY HASHES ==="
$reportOutput += ""

Get-ChildItem -Path $outputpath -Recurse -File | ForEach-Object {
    try {
        $hash = Get-FileHash -Path $_.FullName -Algorithm SHA256 -ErrorAction Stop
        $relativePath = $_.FullName.Replace("$outputpath\", "")
        $reportOutput += "$($hash.Hash)  $relativePath"
    }
    catch {
        Write-Host -Fore Yellow "Warning: Could not hash file $($_.Name)"
        $reportOutput += "ERROR: Could not hash $($_.Name)"
    }
}

$reportOutput += ""
$reportOutput += "=" * 80
$reportOutput += "END OF REPORT"
$reportOutput += "=" * 80

$reportOutput | Out-File -FilePath $reportFile -Encoding UTF8
Write-Host -Fore Cyan "Comprehensive report created: CyberPipe-Report.txt"

# Clean up temporary EDD file
if (Test-Path $eddTempFile) {
    Remove-Item $eddTempFile -Force -ErrorAction SilentlyContinue
}

$summary = @{
    Hostname = $name
    OS = $OS
    Architecture = $arch
    Profile = $profileName
    CollectionStarted = $preCollectionTime.ToString()
    CollectionCompleted = (Get-Date).ToString()
    Duration = "$Minutes min $Seconds sec"
    TotalFiles = $fileCount
    TotalSizeGB = $totalSizeGB
    Uptime = "$($uptime.Days) days, $($uptime.Hours) hours"
    VirtualEnvironment = if ($computerSystem.Model -match "Virtual|VMware|VirtualBox|Hyper-V|QEMU|Xen") { $computerSystem.Model } else { "Physical/Unknown" }
    ReportFile = "CyberPipe-Report.txt"
    Status = "Completed"
}
$summary | ConvertTo-Json | Set-Content "$outputpath\collection-summary.json"

# Add CSV header if file doesn't exist
if (-not (Test-Path "$wd\CyberPipe-runs.csv")) {
    Set-Content -Path "$wd\CyberPipe-runs.csv" -Value "Timestamp,Hostname,Profile,Duration"
}

Add-Content -Path "$wd\CyberPipe-runs.csv" -Value "$((Get-Date).ToString()),$env:COMPUTERNAME,$profileName,$Minutes`:$Seconds"

# Optional Compression
if ($Compress) {
    Write-Host -Fore Cyan "Compressing collection..."
    $zipPath = "$wd\Collections\$collection.zip"

    try {
        Compress-Archive -Path $outputpath -DestinationPath $zipPath -CompressionLevel Optimal -Force
        $zipSize = [math]::Round((Get-Item $zipPath).Length / 1GB, 2)
        Write-Host -Fore Green "Collection compressed: $collection.zip ($zipSize GB)"

        # Optionally remove uncompressed folder
        # Uncomment the following lines to auto-delete after compression:
        # Remove-Item -Path $outputpath -Recurse -Force
        # Write-Host -Fore Cyan "Uncompressed collection removed."
    }
    catch {
        Write-Host -Fore Yellow "Warning: Compression failed - $($_.Exception.Message)"
        Write-Host -Fore Yellow "Uncompressed collection remains at: $outputpath"
    }
}

# Validate collection actually succeeded by checking for artifacts
if (-not (Test-Path $outputpath)) {
    Write-Host -Fore Red "Collection failed: Output directory not found."
    exit 1
}

# Check that we have more than just our own generated files (report, summary, log)
$collectedFiles = Get-ChildItem -Path $outputpath -Recurse -File | Where-Object {
    $_.Name -notmatch "(CyberPipe-Report|summary\.json|log\.txt)"
}

if ($collectedFiles.Count -eq 0) {
    Write-Host -Fore Red "Collection failed: No artifacts collected by MAGNET Response."
    Write-Host -Fore Red "Check $outputpath\log.txt for details."
    exit 1
}

# Check if this was a RAM collection and verify memory dump exists
if ($CollectionProfile -match "RAM|^$") {
    $memoryDump = Get-ChildItem -Path $outputpath -Recurse -File | Where-Object {
        $_.Extension -match "\.(raw|mem|dmp|bin)" -and $_.Length -gt 100MB
    }

    if (-not $memoryDump) {
        Write-Host -Fore Yellow "Warning: RAM collection selected but no memory dump found (or dump < 100MB)."
        Write-Host -Fore Yellow "Collection may have failed. Check $outputpath\log.txt for details."
    }
}

# Calculate total collection size for better reporting
$collectionSizeGB = [math]::Round(($collectedFiles | Measure-Object -Property Length -Sum).Sum / 1GB, 2)

Write-Host -Fore Green "Collection succeeded."
Write-Host -Fore Green "  Profile: $profileName"
Write-Host -Fore Green "  Files: $($collectedFiles.Count) | Size: $collectionSizeGB GB"
Write-Host ""
exit 0
