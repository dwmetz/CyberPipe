<#
.NOTES
CyberPipe5.ps1  
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
CyberPipe v5.1

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
#>
param (
    [switch]$Elevated,
    [string]$CollectionProfile = $env:CYBERPIPE_PROFILE
)
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ((Test-Admin) -eq $false)  {
    if ($elevated) {
    } else {
        Write-host " "
        Write-host  "CyberPipe requires Admin permissions (not detected). Exiting."   
        Write-host  " "    
    }
    exit
}
[console]::ForegroundColor="Cyan"
Clear-Host
Write-Host ""
Write-Host ""
Write-Host ""
Write-Host "   .',;::cccccc:;.                         ...'''''''..'."
Write-Host "   .;ccclllloooddxc.                   .';clooddoolcc::;:;."
Write-Host "   .:ccclllloooddxo.               .,coxxxxxdl:,'.."
Write-Host "   'ccccclllooodddd'            .,,'lxkxxxo:'."
Write-Host "   'ccccclllooodddd'        .,:lxOkl,;oxo,."
Write-Host "   ':cccclllooodddo.      .:dkOOOOkkd;''."
Write-Host "   .:cccclllooooddo.  ..;lxkOOOOOkkkd;"
Write-Host "   .;ccccllloooodddc:coxkkkkOOOOOOx:."
Write-Host "    'cccclllooooddddxxxxkkkkOOOOx:."
Write-Host "     ,ccclllooooddddxxxxxkkkxlc,."
Write-Host "      ':llllooooddddxxxxxoc;."
Write-Host "       .';:clooddddolc:,.."
Write-Host "           ''''''''''"
Write-Host ""
Write-Host "CyberPipe IR Collection Script v5.1"
Write-Host "https://github.com/dwmetz/CyberPipe"
Write-Host "$([char]0x00A9)2025 @dwmetz | bakerstreetforensics.com"
Write-Host ""
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
[console]::ForegroundColor="DarkCyan"
## Network Collection - uncomment the section below for Network use
<#
$server = "\\hydepark\triage" # Server configuration
Write-Host  "Mapping network drive..."
$Networkpath = "Z:\" 
If (Test-Path -Path $Networkpath) {
    Write-Host  "Drive is mapped."
}
Else {
    # map network drive
    (New-Object -ComObject WScript.Network).MapNetworkDrive("Z:","$server")
    # check mapping again
    If (Test-Path -Path $Networkpath) {
        Write-Host  "Drive has been mapped."
    }
    Else {
        Write-Host -Fore Red "Error mapping drive."
    }
}
Set-Location $Networkpath
## End of Network section
#>

## Below is for USB and Network:
$wd = Get-Location
$tstamp = (Get-Date -Format "yyyyMMddHHmm")
$outputpath = "$wd\Collections\$env:COMPUTERNAME-$tstamp"
If (Test-Path -Path $wd\Tools) {
}
Else {
        Write-Host " "
        Write-Host -For DarkCyan "Tools directory not present."
        Write-Host " "
        exit
        
    }
    
If (Test-Path -Path Collections) {
    Write-Host  "Collections directory exists."
}
Else {
    $null = mkdir Collections
    If (Test-Path -Path Collections) {
        Write-Host  "Collection directory created."
    }
    Else {
        Write-Host -For DarkCyan "Error creating directory."
    }
}
Set-Location Collections
If (Test-Path -Path $outputpath) {
    Write-Host  "Host directory already exists."
}
Else {
    $null = mkdir $outputpath
    If (Test-Path -Path $outputpath) {
        Write-Host  "Host directory created."
    }
    Else {
        Write-Host -For DarkCyan "Error creating directory."
    }
}

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
    default {
        $profileName = "MAGNET Triage"
        $arguments = "/captureram /capturepagefile /capturevolatile /capturesystemfiles"
    }
}

Write-Host ""
$tstamp = (Get-Date -Format "yyyyMMddHHmm")
$global:progressPreference = 'silentlyContinue'
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
Write-Host -Fore Cyan  "
Running MAGNET Response...
"
Write-Host ""
Write-Host  "Magnet RESPONSE v1.7
$([char]0x00A9)2021-2024 Magnet Forensics Inc
"
$OS = $(((gcim Win32_OperatingSystem -ComputerName $server.Name).Name).split('|')[0])
$arch = (get-wmiobject win32_operatingsystem).osarchitecture
$name = (get-wmiobject win32_operatingsystem).csname
Write-host  "
Hostname:           $name
Operating System:   $OS
Architecture:       $arch
Selected Profile:   $profileName
Output Directory:   $outputpath
"
.$wd\Tools\MagnetRESPONSE.exe /accepteula /unattended /silent /caseref:CyberPipe /output:"$outputpath" $arguments 
Write-Host -Fore Cyan "
Collecting Artifacts...
"
Wait-Process -name "MagnetRESPONSE"
$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host -Fore Cyan  "** Magnet RESPONSE Completed in $Minutes minutes and $Seconds seconds. **
"
Write-Host -Fore Cyan "Running Encrypted Disk Detector (EDD)...
"
$collection = "$env:COMPUTERNAME-$tstamp"
.$wd\Tools\EDDv310.exe /batch >> $outputpath\$collection-edd.txt
Start-Sleep 1
Get-Content $outputpath\$collection-edd.txt
Write-Host -Fore Cyan  "
Checking for BitLocker Key...
"
(Get-BitLockerVolume -MountPoint C).KeyProtector > $outputpath\$collection-key.txt 
If ($Null -eq (Get-Content "$outputpath\$collection-key.txt")) {
Write-Host -Fore yellow "
Bitlocker key not identified.
"
Set-Content -Path $outputpath\$collection-key.txt -Value "
No Bitlocker key identified for $env:computername
"
}
Else {
    Write-Host -Fore Cyan "
Bitlocker key recovered.
"
}
Set-Content -Path $outputpath\$collection-complete.txt -Value "Collection complete: $((Get-Date).ToString())"
Set-Location ~
$StopWatch.Stop()
$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host -Fore Cyan  "
*** Collection Completed in $Minutes minutes and $Seconds seconds. ***
"

$summary = @{
    Hostname = $name
    OS = $OS
    Architecture = $arch
    Profile = $profileName
    Timestamp = $tstamp
    Duration = "$Minutes min $Seconds sec"
}
$summary | ConvertTo-Json | Set-Content "$outputpath\collection-summary.json"

Add-Content -Path "$wd\CyberPipe-runs.csv" -Value "$((Get-Date).ToString()),$env:COMPUTERNAME,$profileName,$Minutes`:$Seconds"

if (-not (Test-Path $outputpath)) {
    Write-Host -Fore Red "Collection failed."
    exit 1
}
Write-Host -Fore Green "Collection succeeded."
exit 0
