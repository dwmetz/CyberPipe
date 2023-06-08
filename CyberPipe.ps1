<#
CyberPipe.ps1
https://github.com/dwmetz/CyberPipe
previously named "CSIRT-Collect"
Author: @dwmetz

Function: This script will:
- capture a memory image with DumpIt for Windows, (x32, x64, ARM64)
- capture a triage image with KAPE,
- check for encrypted disks,
- recover the active BitLocker Recovery key,
- save all artifacts, output and audit logs to USB or source network drive.

Prerequisites: (updated for v.4)
- [MAGNET DumpIt for Windows](https://www.magnetforensics.com/resources/magnet-dumpit-for-windows/)
- [KAPE](https://www.sans.org/tools/kape)
- DumpIt.exe (64-bit) in /modules/bin
- DumpIt_arm.exe (DumpIt.exe ARM release) in /modules/bin
- (optional) DumpIt_x86.exe (DumpIt.exe x86 release) in /modules/bin
- [Encrypted Disk Detector](https://www.magnetforensics.com/resources/encrypted-disk-detector/) (EDDv310.exe) in /modules/bin/EDD
- CyberPipe.ps1 next to your KAPE directory (whether on network or USB) and the script will take care of any folder creation necessary.

Execution:
- Open PowerShell as Adminstrator
- Execute ./CyberPipe.ps1

Release Notes:

v4.01 - Memory modules and EDD separated to enable easy commenting-out of memory capture for triage capture only

v4.0 - "One Script to Rule them All"
- Admin permissions check before execution
- Memory acquisition will use Magnet DumpIt for Windows (previously used Magnet RAM Capture).
- Support for x64, ARM64 and x86 architectures.
- Both memory acquistion and triage collection now facilitated via KAPE batch mode with _kape.cli dynamically built during execution.
- Capture directories now named to $hostname-$timestamp to support multiple collections from the same asset without overwriting.
- Alert if Bitlocker key not detected. Both display and (empty) text file updated if encryption key not detected.
- If key is detected it is written to the output file.
- More efficient use of variables for output files rather than relying on renaming functions during operations.
- Now just one script for Network or USB usage. Uncomment the “Network Collection” section for network use.
- Stopwatch function will calculate the total runtime of the collection.
- ASCII art “Ceci n’est pas une pipe.”

#>
param ([switch]$Elevated)
function Test-Admin {
    $currentUser = New-Object Security.Principal.WindowsPrincipal $([Security.Principal.WindowsIdentity]::GetCurrent())
    $currentUser.IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
}
if ((Test-Admin) -eq $false)  {
    if ($elevated) {
    } else {
        Write-host -fore DarkCyan "CyberPipe requires Admin permissions (not detected). Exiting."        
    }
    exit
}
Clear-Host
Write-Host ""
Write-Host ""
Write-Host ""
Write-host -Fore Cyan "
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
Write-Host -Fore Cyan "            CyberPipe IR Collection Script v4.01" 
Write-Host -Fore Gray "          https://github.com/dwmetz/CyberPipe"
Write-Host -Fore Gray "          @dwmetz | $([char]0x00A9)2023 bakerstreetforensics.com"
Write-Host ""
Write-Host ""
$stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
## Network Collection - uncomment the section below for Network use
$server = "\\hydepark\automate\watchfolders\cyberpipe" # Server configuration
Write-Host -Fore Gray "Mapping network drive..."
$Networkpath = "Z:\" 
If (Test-Path -Path $Networkpath) {
    Write-Host -Fore Gray "Drive Exists already."
}
Else {
    # map network drive
    (New-Object -ComObject WScript.Network).MapNetworkDrive("Z:","$server")
    # check mapping again
    If (Test-Path -Path $Networkpath) {
        Write-Host -Fore Gray "Drive has been mapped."
    }
    Else {
        Write-Host -Fore Red "Error mapping drive."
    }
}
Set-Location Z:
#>
## Below is for USB and Network:
$tstamp = (Get-Date -Format "_yyyyMMddHHmm")
$collection = $env:COMPUTERNAME+$tstamp
$wd = Get-Location
If (Test-Path -Path Collections) {
    Write-Host -Fore Gray "Collections directory exists."
}
Else {
    $null = mkdir Collections
    If (Test-Path -Path Collections) {
        Write-Host -Fore Gray "Collection directory created."
    }
    Else {
        Write-Host -For Cyan "Error creating directory."
    }
}
Set-Location Collections
$CollectionHostpath = "$wd\Collections\$collection"
If (Test-Path -Path $CollectionHostpath) {
    Write-Host -Fore Gray "Host directory already exists."
}
Else {
    $null = mkdir $CollectionHostpath
    If (Test-Path -Path $CollectionHostpath) {
        Write-Host -Fore Gray "Host directory created."
    }
    Else {
        Write-Host -For Cyan "Error creating directory."
    }
}
$MemoryCollectionpath = "$CollectionHostpath\Memory"
If (Test-Path -Path $MemoryCollectionpath) {
}
Else {
    $null = mkdir "$CollectionHostpath\Memory"
    If (Test-Path -Path $MemoryCollectionpath) {
    }
    Else {
        Write-Host -For Red "Error creating Memory directory."
    }
}
Write-Host -Fore Gray "Determining OS build info..."
[System.Environment]::OSVersion.Version > $CollectionHostpath\Memory\$env:COMPUTERNAME-profile.txt
Write-Host -Fore Gray "Preparing _kape.cli..."
$dest = "$CollectionHostpath"
Set-Location $wd\KAPE
# MEMORY COLLECTION 
$arm = (Get-WmiObject -Class Win32_ComputerSystem).SystemType -match '(ARM)'
if ($arm -eq "True") {
    Write-Host "ARM detected"
    Set-Content -Path _kape.cli -Value "--msource C:\ --mdest $dest --module DumpIt_Memory_ARM --ul" }
else {
    Set-Content -Path _kape.cli -Value "--msource C:\ --mdest $dest --module DumpIt_Memory --ul" }
#>
Add-Content -Path _kape.cli -Value "--msource C:\ --mdest $dest --module MagnetForensics_EDD --ul" 
Add-Content -Path _kape.cli -Value "--tsource C:\ --tdest $dest --target KapeTriage --vhdx $env:computername --zv false"
Write-host -Fore Gray "Note: DumpIt, EDD & KAPE triage collection processes will launch in separate windows."
Write-host -Fore Cyan "Triage aquisition will initate after memory collection completes."
$null = .\kape.exe 
Set-Location $MemoryCollectionpath
Get-ChildItem -Filter '*memdump*' -Recurse | Rename-Item -NewName {$_.name -replace 'memdump', $collection }
Write-Host -Fore Gray "Checking for BitLocker Key..."
(Get-BitLockerVolume -MountPoint C).KeyProtector > $CollectionHostpath\LiveResponse\$collection-key.txt 
If ($Null -eq (Get-Content "$CollectionHostpath\LiveResponse\$collection-key.txt")) {
Write-Host -Fore yellow "Bitlocker key not identified."
Set-Content -Path $CollectionHostpath\LiveResponse\$collection-key.txt -Value "No Bitlocker key identified for $env:computername"
}
Else {
    Write-Host -fore green "Bitlocker key recovered."
}
Set-Content -Path $CollectionHostpath\collection-complete.txt -Value "Collection complete: $((Get-Date).ToString())"
Set-Location ~
$StopWatch.Stop()
$null = $stopwatch.Elapsed
$Minutes = $StopWatch.Elapsed.Minutes
$Seconds = $StopWatch.Elapsed.Seconds
Write-Host -Fore Cyan "** Collection Completed in $Minutes minutes and $Seconds seconds.**"