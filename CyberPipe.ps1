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
if ((Test-Admin) -eq $false) {
  if ($elevated) {
  } else {
    Write-host " "
    Write-host "CyberPipe requires Admin permissions (not detected). Exiting."  
    Write-host " "  
  }
  exit
}
[console]::ForegroundColor="Cyan"
Clear-Host

Sleep 1
if ($PSVersionTable.PSEdition -eq 'Core') {
$banner = @'
        ╔════════════════════════════════════════════════════════════════════════════╗
        ║                                                                            ║
        ║     ██████╗██╗   ██╗██████╗ ███████╗██████╗ ██████╗ ██╗██████╗ ███████╗    ║
        ║    ██╔════╝╚██╗ ██╔╝██╔══██╗██╔════╝██╔══██╗██╔══██╗██║██╔══██╗██╔════╝    ║
        ║    ██║      ╚████╔╝ ██████╔╝█████╗  ██████╔╝██████╔╝██║██████╔╝█████╗      ║
        ║    ██║       ╚██╔╝  ██╔══██╗██╔══╝  ██╔══██╗██╔═══╝ ██║██╔═══╝ ██╔══╝      ║
        ║    ╚██████╗   ██║   ██████╔╝███████╗██║  ██║██║     ██║██║     ███████╗    ║
        ║     ╚═════╝   ╚═╝   ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝     ╚═╝╚═╝     ╚══════╝    ║
        ║                                                                            ║
        ║                         Incident Response Collection                       ║
        ║                                   v5.3                                     ║
        ║                                                                            ║
        ╚════════════════════════════════════════════════════════════════════════════╝
                              Memory • Triage • Chain of Custody                      

                             https://github.com/dwmetz/CyberPipe

'@
} else {
  $banner = @'
        +-------------------------------------------------+ 
        |                   CYBERPIPE                     |
        |                                                 |
        |          Incident Response Collection           |
        |                    v5.3                         |
        |                                                 |        
        +-------------------------------------------------+
                Memory - Triage - Chain of Custody  

                https://github.com/dwmetz/CyberPipe    
'@
}

$banner
$copyrightSymbol = [char]0x00A9
$copyrightLine = $copyrightSymbol + '2025 @dwmetz | bakerstreetforensics.com'
Write-Host ''
Write-Host ''
Write-Host $copyrightLine
Write-Host ''
Start-Sleep 1

[console]::ForegroundColor='DarkCyan'
## Network Collection Handling
if ($Net) {
    $netMsg = 'Network mode enabled. Mapping drive to {0}...' -f $Net
    Write-Host -Fore Cyan $netMsg
    $Networkpath = 'Z:\'

    If (Test-Path -Path $Networkpath) {
        Write-Host 'Drive Z: already mapped.'
    }
    Else {
        try {
            (New-Object -ComObject WScript.Network).MapNetworkDrive('Z:', $Net)
            Start-Sleep -Seconds 2

            If (Test-Path -Path $Networkpath) {
                $successMsg = 'Drive mapped successfully to {0}' -f $Net
                Write-Host -Fore Green $successMsg
            }
            Else {
                Write-Host -Fore Red 'Error: Drive mapping appeared to succeed but path not accessible.'
                exit 1
            }
        }
        catch {
            $errMsg = 'Error mapping network drive: {0}' -f $_.Exception.Message
            Write-Host -Fore Red $errMsg
            exit 1
        }
    }

    Set-Location $Networkpath
    Write-Host -Fore Cyan 'Working from network location: Z:\'
}
## Below is for USB and Network:
# Save the script's directory (not current location, which may change)
if ($PSScriptRoot) {
    $wd = $PSScriptRoot
} else {
    # Fallback for PS 2.0
    $wd = Split-Path -Parent $MyInvocation.MyCommand.Path
}
$tstamp = (Get-Date -Format 'yyyyMMddHHmm')
$collectionName = $env:COMPUTERNAME + '-' + $tstamp
$collectionsDir = Join-Path $wd 'Collections'
$outputpath = Join-Path $collectionsDir $collectionName
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
If (Test-Path -Path $wd\Tools) {
}
Else {
    Write-Host ' '
    Write-Host -For DarkCyan 'Tools directory not present.'
    Write-Host ' '
    exit 1

  }
  
If (Test-Path -Path Collections) {
  Write-Host 'Collections directory exists.'
}
Else {
  $null = mkdir Collections
  If (Test-Path -Path Collections) {
    Write-Host 'Collection directory created.'
  }
  Else {
    Write-Host -For DarkCyan 'Error creating directory.'
    exit 1
  }
}
Set-Location Collections
If (Test-Path -Path $outputpath) {
  Write-Host 'Host directory already exists.'
}
Else {
  $null = mkdir $outputpath
  If (Test-Path -Path $outputpath) {
    Write-Host 'Host directory created.'
  }
  Else {
    Write-Host -For DarkCyan 'Error creating directory.'
    exit 1
  }
}

# Validate required tools exist
$magnetPath = Join-Path $wd 'Tools\MagnetRESPONSE.exe'
$eddPath = Join-Path $wd 'Tools\EDDv310.exe'
$requiredTools = @(
  $magnetPath,
  $eddPath
)

foreach ($tool in $requiredTools) {
  if (-not (Test-Path $tool)) {
    $toolMsg = 'Required tool not found: {0}' -f $tool
    Write-Host -Fore Red $toolMsg
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
if ($CollectionProfile -eq "" -or $CollectionProfile -match '^RAM') {
  # RAM-based or default profile - need more system space
  $systemRequired = [math]::Max(($totalRAM_GB + 5), 10)
  $targetRequired = [math]::Max(($totalRAM_GB + 10), 15)
}

$gbUnit = 'GB'
$ramMsg = 'System RAM: {0} {1}' -f $totalRAM_GB, $gbUnit
Write-Host $ramMsg
$targetMsg = 'Target drive {0} free space: {1} {2} (need {3} {4})' -f $targetDrive.Name, $targetFreeSpaceGB, $gbUnit, $targetRequired, $gbUnit
Write-Host $targetMsg
$systemMsg = 'System drive C: free space: {0} {1} (need {2} {3})' -f $systemFreeSpaceGB, $gbUnit, $systemRequired, $gbUnit
Write-Host $systemMsg

if ($targetFreeSpaceGB -lt $targetRequired) {
  $errMsg = 'Insufficient space on target drive. Available: {0} {1} - Required: {2} {3}' -f $targetFreeSpaceGB, $gbUnit, $targetRequired, $gbUnit
  Write-Host -Fore Red $errMsg
  exit 1
}

if ($systemFreeSpaceGB -lt $systemRequired) {
  $errMsg = 'Insufficient space on system drive (C:). Available: {0} {1} - Required: {2} {3}' -f $systemFreeSpaceGB, $gbUnit, $systemRequired, $gbUnit
  Write-Host -Fore Red $errMsg
  Write-Host -Fore Red 'MAGNET Response requires approximately RAM-sized space on C: for temporary files during memory capture.'
  Write-Host -Fore Yellow 'To proceed anyway, use -CollectionProfile Volatile to skip RAM capture.'
  exit 1
}

### Pre-Collection Volatile Snapshot (stored in memory for later inclusion in report)
Write-Host -Fore Cyan 'Capturing pre-collection volatile snapshot...'
$collection = $env:COMPUTERNAME + '-' + $tstamp
$preCollectionTime = Get-Date
$snapshotOutput = @()
$snapshotOutput += '=== PRE-COLLECTION VOLATILE SNAPSHOT ==='
$capturedLine = 'Captured: {0}' -f $preCollectionTime.ToString()
$snapshotOutput += $capturedLine
$snapshotOutput += ''

# System Uptime
$osInfo = Get-CimInstance Win32_OperatingSystem
$uptime = (Get-Date) - $osInfo.LastBootUpTime
$snapshotOutput += '--- SYSTEM UPTIME ---'
$lastBootLine = 'Last Boot: {0}' -f $osInfo.LastBootUpTime
$snapshotOutput += $lastBootLine
$uptimeDetailLine = 'Uptime: {0} days, {1} hours, {2} minutes' -f $uptime.Days, $uptime.Hours, $uptime.Minutes
$snapshotOutput += $uptimeDetailLine
$snapshotOutput += ''

# Detect Virtual Environment
# Important for analysis: VMs may have memory overcommitment, nested malware, or hypervisor-level threats
# that won't be captured in guest memory dumps. Knowing the environment helps analysts understand
# collection limitations and potential blind spots.
$snapshotOutput += '--- VIRTUALIZATION DETECTION ---'
$computerSystem = Get-CimInstance Win32_ComputerSystem
$modelLine = 'Model: {0}' -f $computerSystem.Model
$snapshotOutput += $modelLine
$mfgLine = 'Manufacturer: {0}' -f $computerSystem.Manufacturer
$snapshotOutput += $mfgLine
if ($computerSystem.Model -match 'Virtual|VMware|VirtualBox|Hyper-V|QEMU|Xen') {
  $vmLine = 'Virtual Environment: DETECTED ({0})' -f $computerSystem.Model
  $snapshotOutput += $vmLine
  $snapshotOutput += 'Note: VM memory dumps may not capture hypervisor-level activity or overcommitted memory'
} else {
  $snapshotOutput += 'Virtual Environment: Physical or Unknown'
}
$snapshotOutput += ""

# Logged-on Users
$snapshotOutput += '--- LOGGED-ON USERS ---'
try {
  $loggedOnUsers = Get-CimInstance Win32_LoggedOnUser -ErrorAction Stop |
    Select-Object -ExpandProperty Antecedent |
    Select-Object -Unique Domain, Name
  foreach ($user in $loggedOnUsers) {
    $userString = '{0}\{1}' -f $user.Domain, $user.Name
    $snapshotOutput += $userString
  }
} catch {
  $snapshotOutput += 'Unable to enumerate logged-on users'
}
$snapshotOutput += ""

# Active Network Connections
$snapshotOutput += '--- ACTIVE NETWORK CONNECTIONS ---'
try {
  $connections = Get-NetTCPConnection -State Established -ErrorAction Stop |
    Select-Object LocalAddress, LocalPort, RemoteAddress, RemotePort, OwningProcess
  foreach ($conn in $connections) {
    $processName = (Get-Process -Id $conn.OwningProcess -ErrorAction SilentlyContinue).Name
    $connLine = '{0}:{1} -> {2}:{3} [{4} PID:{5}]' -f $conn.LocalAddress, $conn.LocalPort, $conn.RemoteAddress, $conn.RemotePort, $processName, $conn.OwningProcess
    $snapshotOutput += $connLine
  }
} catch {
  $snapshotOutput += 'Unable to enumerate network connections'
}
$snapshotOutput += ""

# Running Processes (top 20 by memory)
$snapshotOutput += '--- TOP PROCESSES (by memory) ---'
try {
  $processes = Get-Process | Sort-Object WorkingSet64 -Descending | Select-Object -First 20
  foreach ($proc in $processes) {
    $memMB = [math]::Round($proc.WorkingSet64 / 1MB, 2)
    $snapshotOutput += '{0} [PID:{1}] - {2} MB' -f $proc.Name, $proc.Id, $memMB
  }
} catch {
  $snapshotOutput += 'Unable to enumerate processes'
}

Write-Host -Fore Cyan 'Pre-collection snapshot captured (will be included in final report)'
Write-Host ''

### Collection Profiles
switch ($CollectionProfile) {
  'Volatile' {
    $profileName = 'Volatile'
    $arguments = '/capturevolatile'
  }
  'RAMSystem' {
    $profileName = 'RAM and Critical System Files'
    $arguments = '/captureram /capturesystemfiles'
  }
  'RAMPage' {
    $profileName = 'RAM and Pagefile'
    $arguments = '/captureram /capturepagefile'
  }
  'RAMOnly' {
    $profileName = 'RAM Dump'
    $arguments = '/captureram'
  }
  'QuickTriage' {
    $profileName = 'Quick Triage'
    $arguments = '/capturevolatile /capturesystemfiles'
  }
  default {
    $profileName = 'MAGNET Triage'
    $arguments = '/captureram /capturepagefile /capturevolatile /capturesystemfiles'
  }
}

Write-Host ''
$global:progressPreference = 'silentlyContinue'
Write-Host ''
Write-Host -Fore Cyan 'Running MAGNET Response...'
Write-Host ''
Write-Host ''
Write-Host 'Magnet RESPONSE v1.7'
$magnetCopyright = $copyrightSymbol + '2021-2024 Magnet Forensics Inc'
Write-Host $magnetCopyright
Write-Host ''
$OS = $(((Get-CimInstance Win32_OperatingSystem).Caption).split('|')[0])
$arch = (Get-CimInstance Win32_OperatingSystem).OSArchitecture
$name = (Get-CimInstance Win32_OperatingSystem).CSName
Write-Host ''
Write-Host "Hostname:      $name"
Write-Host "Operating System:  $OS"
Write-Host "Architecture:    $arch"
Write-Host "Selected Profile:  $profileName"
Write-Host "Output Directory:  $outputpath"
Write-Host ''
Write-Host ''
Write-Host -Fore Cyan 'Collecting Artifacts...'
Write-Host ''

# Build argument string properly
$outputQuoted = [char]34 + $outputpath + [char]34
$magnetArgs = '/accepteula /unattended /silent /caseref:CyberPipe /output:' + $outputQuoted + ' ' + $arguments
# Use the original working directory we saved
$magnetExePath = Join-Path $wd 'Tools\MagnetRESPONSE.exe'
$magnetProcess = Start-Process -FilePath $magnetExePath -ArgumentList $magnetArgs -PassThru -NoNewWindow

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
        $progressMsg = ' [Running: {0} min {1} sec | Collected: {2} MB]' -f $minutes, $seconds, $sizeMB
        Write-Host ([char]13 + $progressMsg) -Fore DarkCyan -NoNewline
      } else {
        $sizeGB = [math]::Round($collectionSize / 1GB, 2)
        $progressMsg = ' [Running: {0} min {1} sec | Collected: {2} GB]' -f $minutes, $seconds, $sizeGB
        Write-Host ([char]13 + $progressMsg) -Fore DarkCyan -NoNewline
      }
    } else {
      $progressMsg = ' [Running: {0} min {1} sec | Collected: 0.0 MB]' -f $minutes, $seconds
      Write-Host ([char]13 + $progressMsg) -Fore DarkCyan -NoNewline
    }
  }
  catch {
    $progressMsg = ' [Running: {0} min {1} sec]' -f $minutes, $seconds
    Write-Host ([char]13 + $progressMsg) -Fore DarkCyan -NoNewline
  }
}
Write-Host '' # New line after progress completes

$magnetProcess.WaitForExit()

# PS 5.1 Compatibility: Refresh process and validate exit code
try {
  $magnetProcess.Refresh()
} catch {
  # Refresh may not be available in all PS versions, ignore error
}

$magnetExitCode = $magnetProcess.ExitCode

# Additional validation: Check if files were actually collected (more reliable than exit code alone)
Start-Sleep -Seconds 2
$collectedFiles = Get-ChildItem -Path $outputpath -Recurse -File -ErrorAction SilentlyContinue | Where-Object {
  $_.Name -notmatch '(log\.txt|.*temp.*)'
}

if ($magnetExitCode -ne 0) {
  # If exit code is non-zero but files were collected, it may be a PS 5.1 exit code reporting issue
  if ($collectedFiles.Count -gt 0) {
    $warnMsg = 'MAGNET Response reported exit code {0}, but files were collected successfully. Continuing...' -f $magnetExitCode
    Write-Host -Fore Yellow $warnMsg
  } else {
    # Genuine failure - no files and bad exit code
    $errMsg = 'MAGNET Response failed with exit code: {0}' -f $magnetExitCode
    Write-Host -Fore Red $errMsg
    exit 1
  }
}

$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host ''
$magnetMsg = '** Magnet RESPONSE Completed in {0} minutes and {1} seconds. **' -f $Minutes, $Seconds
Write-Host -Fore Cyan $magnetMsg
Write-Host ''
Write-Host -Fore Cyan 'Running Encrypted Disk Detector (EDD)...'
Write-Host ''
$collection = $env:COMPUTERNAME + '-' + $tstamp
$eddTempFileName = $collection + '-edd-temp.txt'
$eddTempFile = Join-Path $outputpath $eddTempFileName
$eddExePath = Join-Path $wd 'Tools\EDDv310.exe'
$eddProcess = Start-Process -FilePath $eddExePath -ArgumentList '/batch' -RedirectStandardOutput $eddTempFile -PassThru -Wait -NoNewWindow

if ($eddProcess.ExitCode -ne 0) {
  $warnMsg = 'Warning: EDD exited with code {0}' -f $eddProcess.ExitCode
  Write-Host -Fore Yellow $warnMsg
}

Start-Sleep 1
$eddOutput = Get-Content $eddTempFile
$eddOutput | ForEach-Object { Write-Host $_ }
Write-Host ''
Write-Host -Fore Cyan 'Checking for BitLocker Keys...'
Write-Host ''
# Get all BitLocker volumes, not just C:
$bitlockerVolumes = Get-BitLockerVolume | Where-Object { $_.ProtectionStatus -eq 'On' }

$keyOutput = @()
if ($bitlockerVolumes.Count -eq 0) {
  Write-Host -Fore Yellow 'No BitLocker protected volumes found.'
  $noBLLine = 'No BitLocker protected volumes found on {0}' -f $env:computername
  $keyOutput += $noBLLine
}
else {
  foreach ($volume in $bitlockerVolumes) {
    $volumeLine = 'Volume: {0}' -f $volume.MountPoint
    $keyOutput += $volumeLine
    $statusLine = 'Protection Status: {0}' -f $volume.ProtectionStatus
    $keyOutput += $statusLine

    if ($volume.KeyProtector.Count -gt 0) {
      $keyOutput += 'Key Protectors:'
      foreach ($kp in $volume.KeyProtector) {
        $typeLine = ' Type: {0}' -f $kp.KeyProtectorType
        $keyOutput += $typeLine
        if ($kp.RecoveryPassword) {
          $pwdLine = ' Recovery Password: {0}' -f $kp.RecoveryPassword
          $keyOutput += $pwdLine
        }
        if ($kp.KeyProtectorId) {
          $idLine = ' Key ID: {0}' -f $kp.KeyProtectorId
          $keyOutput += $idLine
        }
      }
      $blMsg = 'BitLocker key(s) recovered for volume {0}' -f $volume.MountPoint
      Write-Host -Fore Cyan $blMsg
    }
    else {
      $keyOutput += 'No key protectors found for this volume.'
      $noKeyMsg = 'No key protectors for volume {0}' -f $volume.MountPoint
      Write-Host -Fore Yellow $noKeyMsg
    }
    $keyOutput += ""
  }
}
Set-Location ~
$StopWatch.Stop()
$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host ''
$completionMsg = '*** Collection Completed in {0} minutes and {1} seconds. ***' -f $Minutes, $Seconds
Write-Host -Fore Cyan $completionMsg
Write-Host ''

# Generate Comprehensive CyberPipe Report
Write-Host -Fore Cyan 'Generating CyberPipe collection report...'
$reportFile = Join-Path $outputpath 'CyberPipe-Report.txt'
$reportOutput = @()

# Header
$separator = '=' * 80
$reportOutput += $separator
$reportOutput += 'CYBERPIPE INCIDENT RESPONSE COLLECTION REPORT'
$reportOutput += $separator
$reportOutput += ''
$hostLine = 'Host: {0}' -f $env:COMPUTERNAME
$reportOutput += $hostLine
$profileLine = 'Collection Profile: {0}' -f $profileName
$reportOutput += $profileLine
$collectionStarted = 'Collection Started: {0}' -f $preCollectionTime.ToString()
$reportOutput += $collectionStarted
$collectionCompleted = 'Collection Completed: {0}' -f (Get-Date).ToString()
$reportOutput += $collectionCompleted
$durationLine = 'Duration: {0} minutes, {1} seconds' -f $Minutes, $Seconds
$reportOutput += $durationLine
$reportOutput += 'Generated by: CyberPipe v5.3'
$reportOutput += 'https://github.com/dwmetz/CyberPipe'
$reportOutput += ''
$reportOutput += $separator

# Pre-Collection Volatile Snapshot
$reportOutput += ''
$reportOutput += $snapshotOutput
$reportOutput += ''
$reportOutput += $separator

# Collection Summary
$reportOutput += ''
$reportOutput += '=== COLLECTION SUMMARY ==='
$reportOutput += ''
$allFiles = Get-ChildItem -Path $outputpath -Recurse -File
$totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
$totalSizeGB = [math]::Round($totalSize / 1GB, 2)
$fileCount = $allFiles.Count

$filesLine = 'Total Files Collected: {0}' -f $fileCount
$reportOutput += $filesLine
$reportOutput += 'Total Collection Size: {0} GB' -f $totalSizeGB
$reportOutput += 'System RAM: {0} GB' -f $totalRAM_GB
$uptimeLine = 'Uptime at Collection: {0} days, {1} hours' -f $uptime.Days, $uptime.Hours
$reportOutput += $uptimeLine
$virtualEnvLine = 'Virtual Environment: {0}' -f $virtualEnv
$reportOutput += $virtualEnvLine
$reportOutput += ""

# Files by Type
$reportOutput += '--- FILES BY TYPE ---'
$filesByType = $allFiles | Group-Object Extension | Sort-Object Count -Descending
foreach ($group in $filesByType) {
  $ext = if ($group.Name) { $group.Name } else { '(no extension)' }
  $groupSize = ($group.Group | Measure-Object -Property Length -Sum).Sum
  $groupSizeGB = [math]::Round($groupSize / 1GB, 3)
  $reportOutput += '{0} : {1} files {2} GB' -f $ext, $group.Count, $groupSizeGB
}
$reportOutput += ""

# Key Artifacts
$reportOutput += '--- KEY ARTIFACTS ---'
$keyArtifacts = $allFiles | Where-Object { $_.Name -match '(\.raw|\.mem|\.dmp|\.txt|\.json|\.csv)' }
foreach ($artifact in $keyArtifacts) {
  if ($artifact.Length -lt 1MB) {
    $sizeKB = [math]::Round($artifact.Length / 1KB, 2)
    $reportOutput += '{0} {1} KB' -f $artifact.Name, $sizeKB
  } else {
    $sizeMB = [math]::Round($artifact.Length / 1MB, 2)
    $reportOutput += '{0} {1} MB' -f $artifact.Name, $sizeMB
  }
}
$reportOutput += ""
$reportOutput += $separator

# Encrypted Disk Detection
$reportOutput += ""
$reportOutput += '=== ENCRYPTED DISK DETECTION ==='
$reportOutput += ""
$reportOutput += $eddOutput
$reportOutput += ""
$reportOutput += $separator

# BitLocker Recovery Keys
$reportOutput += ""
$reportOutput += '=== BITLOCKER RECOVERY KEYS ==='
$reportOutput += ""
$reportOutput += $keyOutput
$reportOutput += ""
$reportOutput += $separator

# SHA256 Hashes
$reportOutput += ""
$reportOutput += '=== SHA256 INTEGRITY HASHES ==='
$reportOutput += ""

Get-ChildItem -Path $outputpath -Recurse -File | ForEach-Object {
  try {
    $hash = Get-FileHash -Path $_.FullName -Algorithm SHA256 -ErrorAction Stop
    $pathPrefix = $outputpath + '\'
    $relativePath = $_.FullName.Replace($pathPrefix, '')
    $hashLine = '{0} {1}' -f $hash.Hash, $relativePath
    $reportOutput += $hashLine
  }
  catch {
    $warnMsg = 'Warning: Could not hash file {0}' -f $_.Name
    Write-Host -Fore Yellow $warnMsg
    $errLine = 'ERROR: Could not hash {0}' -f $_.Name
    $reportOutput += $errLine
  }
}

$reportOutput += ""
$reportOutput += $separator
$reportOutput += 'END OF REPORT'
$reportOutput += $separator

$reportOutput | Out-File -FilePath $reportFile -Encoding UTF8
Write-Host -Fore Cyan 'Comprehensive report created: CyberPipe-Report.txt'

# Clean up temporary EDD file
if (Test-Path $eddTempFile) {
  Remove-Item $eddTempFile -Force -ErrorAction SilentlyContinue
}

$durationString = '{0} min {1} sec' -f $Minutes, $Seconds
$uptimeString = '{0} days, {1} hours' -f $uptime.Days, $uptime.Hours
$virtualEnv = if ($computerSystem.Model -match 'Virtual|VMware|VirtualBox|Hyper-V|QEMU|Xen') { $computerSystem.Model } else { 'Physical/Unknown' }

$summary = @{
  Hostname = $name
  OS = $OS
  Architecture = $arch
  Profile = $profileName
  CollectionStarted = $preCollectionTime.ToString()
  CollectionCompleted = (Get-Date).ToString()
  Duration = $durationString
  TotalFiles = $fileCount
  TotalSizeGB = $totalSizeGB
  Uptime = $uptimeString
  VirtualEnvironment = $virtualEnv
  ReportFile = 'CyberPipe-Report.txt'
  Status = 'Completed'
}
$summaryFile = Join-Path $outputpath 'collection-summary.json'
$summary | ConvertTo-Json | Set-Content $summaryFile

# Add CSV header if file does not exist
$csvPath = Join-Path $wd 'CyberPipe-runs.csv'
if (-not (Test-Path $csvPath)) {
  $csvHeader = 'Timestamp,Hostname,Profile,Duration'
  Set-Content -Path $csvPath -Value $csvHeader
}

$comma = [char]44
$colon = [char]58
$csvLine = (Get-Date).ToString() + $comma + $env:COMPUTERNAME + $comma + $profileName + $comma + $Minutes + $colon + $Seconds
$csvPath = Join-Path $wd 'CyberPipe-runs.csv'
Add-Content -Path $csvPath -Value $csvLine

# Optional Compression
if ($Compress) {
  Write-Host -Fore Cyan 'Compressing collection...'
  $collectionsDir = Join-Path $wd 'Collections'
  $zipFileName = $collection + '.zip'
  $zipPath = Join-Path $collectionsDir $zipFileName


    # Check collection size first to determine compression strategy
    $collectionSize = (Get-ChildItem -Path $outputpath -Recurse -File | Measure-Object -Property Length -Sum).Sum
    $collectionSizeGB = [math]::Round($collectionSize / 1GB, 2)

    $sizeMsg = '  Collection size: {0} {1}' -f $collectionSizeGB, $gbUnit
    Write-Host $sizeMsg

    # Try 7-Zip first if available (handles large files, better compression)
    $localToolPath = Join-Path $wd 'Tools\7z.exe'
    $sevenZipPaths = @(
        $localToolPath,
        'C:\Program Files\7-Zip\7z.exe',
        'C:\Program Files (x86)\7-Zip\7z.exe'
    )

    $sevenZipExe = $sevenZipPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if ($sevenZipExe) {
        Write-Host -Fore Cyan '  Using 7-Zip for compression (supports large files)...'
        try {
            # Use 7-Zip with ZIP64 format (no size limit)
            # -tzip = ZIP format, -mx5 = medium compression (balance speed/ratio)
            $sourcePattern = $outputpath + '\*'
            $zipPathQuoted = [char]34 + $zipPath + [char]34
            $sourceQuoted = [char]34 + $sourcePattern + [char]34
            $7zArgs = 'a -tzip -mx5 ' + $zipPathQuoted + ' ' + $sourceQuoted
            $7zProcess = Start-Process -FilePath $sevenZipExe -ArgumentList $7zArgs -Wait -PassThru -NoNewWindow

            if ($7zProcess.ExitCode -eq 0 -and (Test-Path $zipPath)) {
                $zipSize = [math]::Round((Get-Item $zipPath).Length / 1GB, 2)
                $successMsg = 'Collection compressed: {0}.zip - {1} {2}' -f $collection, $zipSize, $gbUnit
                Write-Host -Fore Green $successMsg

                # Optionally remove uncompressed folder
                # Uncomment the following lines to auto-delete after compression:
                # Remove-Item -Path $outputpath -Recurse -Force
                # Write-Host -Fore Cyan 'Uncompressed collection removed.'
            } else {
                $errMsg = '7-Zip exited with code {0}' -f $7zProcess.ExitCode
                throw $errMsg
            }
        }
        catch {
            $errMsg = '7-Zip compression failed: {0}' -f $_.Exception.Message
            Write-Host -Fore Red $errMsg
            $remainsMsg = 'Uncompressed collection remains at: {0}' -f $outputpath
            Write-Host -Fore Yellow $remainsMsg
        }
    }
    # Fall back to Compress-Archive only for small collections (< 1.5GB)
    elseif ($collectionSizeGB -lt 1.5) {
        Write-Host -Fore Cyan '  Using built-in compression (small collection)...'
        try {
            Compress-Archive -Path $outputpath -DestinationPath $zipPath -CompressionLevel Optimal -Force
            $zipSize = [math]::Round((Get-Item $zipPath).Length / 1GB, 2)
            $successMsg = 'Collection compressed: {0}.zip - {1} {2}' -f $collection, $zipSize, $gbUnit
            Write-Host -Fore Green $successMsg

            # Optionally remove uncompressed folder
            # Uncomment the following lines to auto-delete after compression:
            # Remove-Item -Path $outputpath -Recurse -Force
            # Write-Host -Fore Cyan 'Uncompressed collection removed.'
        }
        catch {
            $errMsg = 'Compression failed: {0}' -f $_.Exception.Message
            Write-Host -Fore Red $errMsg
            $remainsMsg = 'Uncompressed collection remains at: {0}' -f $outputpath
            Write-Host -Fore Yellow $remainsMsg
        }
    }
    else {
        $tooLargeMsg = 'Collection is too large - {0} GB - for built-in compression.' -f $collectionSizeGB
        Write-Host -Fore Yellow $tooLargeMsg
        Write-Host -Fore Yellow 'PowerShell Compress-Archive cannot create archives larger than 2GB.'
        Write-Host -Fore Yellow ""
        Write-Host -Fore Yellow 'To compress large collections, install 7-Zip:'
        Write-Host -Fore Yellow '  1. Download from https://www.7-zip.org/'
        Write-Host -Fore Yellow '  2. Install to default location, OR'
        $toolsMsg = '  3. Copy 7z.exe to {0}\Tools\' -f $wd
        Write-Host -Fore Yellow $toolsMsg
        Write-Host -Fore Yellow ""
        $remainsMsg = 'Uncompressed collection remains at: {0}' -f $outputpath
        Write-Host -Fore Yellow $remainsMsg
    }
}

# Validate collection actually succeeded by checking for artifacts
if (-not (Test-Path $outputpath)) {
  Write-Host -Fore Red 'Collection failed: Output directory not found.'
  exit 1
}

# Check that we have more than just our own generated files (report, summary, log)
$collectedFiles = Get-ChildItem -Path $outputpath -Recurse -File | Where-Object {
  $_.Name -notmatch '(CyberPipe-Report|summary\.json|log\.txt)'
}

if ($collectedFiles.Count -eq 0) {
  Write-Host -Fore Red 'Collection failed: No artifacts collected by MAGNET Response.'
  $logFile = Join-Path $outputpath 'log.txt'
  $checkMsg = 'Check {0} for details.' -f $logFile
  Write-Host -Fore Red $checkMsg
  exit 1
}

# Check if this was a RAM collection and verify memory dump exists
if ($CollectionProfile -match 'RAM|^$') {
  $memoryDump = Get-ChildItem -Path $outputpath -Recurse -File | Where-Object {
    $_.Extension -match '\.(raw|mem|dmp|bin)' -and $_.Length -gt 100MB
  }

  if (-not $memoryDump) {
    Write-Host -Fore Yellow 'Warning: RAM collection selected but no memory dump found (or dump under 100MB).'
    $logPath = Join-Path $outputpath 'log.txt'
    $logMsg = 'Collection may have failed. Check {0} for details.' -f $logPath
    Write-Host -Fore Yellow $logMsg
  }
}

# Calculate total collection size for better reporting
$collectionSizeGB = [math]::Round(($collectedFiles | Measure-Object -Property Length -Sum).Sum / 1GB, 2)

Write-Host -Fore Green 'Collection succeeded.'
$profileMsg = ' Profile: {0}' -f $profileName
Write-Host -Fore Green $profileMsg
$successMsg = ' Files: {0} - Size: {1} {2}' -f $collectedFiles.Count, $collectionSizeGB, $gbUnit
Write-Host -Fore Green $successMsg
Write-Host ''
exit 0
