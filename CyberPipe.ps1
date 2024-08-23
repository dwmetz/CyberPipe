<#
CyberPipe.ps1
https://github.com/dwmetz/CyberPipe
previously named "CSIRT-Collect"
Author: @dwmetz

This script will:
- capture a memory image with DumpIt for Windows, (x32, x64, ARM64), or Magnet RAM capture on legacy systems
- capture a triage image with MAGNET Response,
- check for encrypted disks,
- recover the active BitLocker Recovery key,
- save all artifacts, output and audit logs to USB or source network drive.

Release Notes: 

v5.0 RESPONSE Edition

Prerequisites: (in \Tools directory)
- [MAGNET Response](https://magnetforensics.com) (MagnetRESPONSE.exe)
- [Encrypted Disk Detector](https://www.magnetforensics.com/resources/encrypted-disk-detector/) (EDDv310.exe)
- CyberPipe5.ps1 next to your TOOLS directory (whether on network or USB) 
Operation:
- Open PowerShell as Adminstrator
- Execute ./CyberPipe.ps1

#>
param ([switch]$Elevated)
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
Write-host  "
    .',;::cccccc:;.                         ...'''''''..'.  
   .;ccclllloooddxc.                   .';clooddoolcc::;:;. 
   .:ccclllloooddxo.               .,coxxxxxdl:,'..         
   'ccccclllooodddd'            .,,'lxkxxxo:'.              
   'ccccclllooodddd'        .,:lxOkl,;oxo,.                 
   ':cccclllooodddo.      .:dkOOOOkkd;''.                   
   .:cccclllooooddo.  ..;lxkOOOOOkkkd;                      
   .;ccccllloooodddc:coxkkkkOOOOOOx:.                       
    'cccclllooooddddxxxxkkkkOOOOx:.                         
     ,ccclllooooddddxxxxxkkkxlc,.                           
      ':llllooooddddxxxxxoc;.                               
       .';:clooddddolc:,..                                  
           ''''''''''                                                                                                                 
"                
Write-Host  "CyberPipe IR Collection Script v5.0" 
Write-Host  "https://github.com/dwmetz/CyberPipe"
Write-Host  "@dwmetz | $([char]0x00A9)2024 bakerstreetforensics.com"
Write-Host ""
Write-Host ""
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
[console]::ForegroundColor="DarkCyan"
## Network Collection - uncomment the section below for Network use
<#
$server = "\\server\share" # Server configuration
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
$tstamp = (Get-Date -Format "yyyyMMddHHmm")
$wd = Get-Location
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

### VARIABLE SETUP

$profileName = "Volatile (testing)"
$arguments = "/capturevolatile" 
#>
<#
$profileName = "MAGNET Triage"
$arguments = "/captureram /capturepagefile /capturevolatile /capturesystemfiles" 
#>
<#
$profileName = "RAM Dump"
$arguments = "/captureram"
#>
<#
$profileName = "RAM & Pagefile"
$arguments = "/captureram /capturepagefile"
#>

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
