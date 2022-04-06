<#
CSIRT-Collect_USB.ps1 v3.1
https://github.com/dwmetz/CSIRT-Collect
Author: @dwmetz

Function: This script will:
- capture a memory image with Magnet Ram Capture, 
- capture a triage image with KAPE,
- check for encrypted disks,
- recover the active BitLocker Recovery key,
save all artifacts directly to the USB device.

Prerequisites:
On the root of the USB:
-CSIRT-Collect_USB.ps1
-folder (empty to start) titled 'Collections'
-KAPE folder from default install. Ensure you have EDD.exe in \modules\bin\EDD
  and verify that the EDD version matches the MagnetForensics_EDD.mkape
-MEMORY folder with MRC.exe (Magnet Ram Capture) and 7za.exe (7zip)

Execution:
-Open PowerShell as Adminstrator
-Navigate to the USB device
-Execute ./CSIRT-Collect_USB.ps1

v3.1 - "Summit Release"
unified code between network and USB versions
contributors dwmetz, stark4n6
#>
Write-Host -Fore Gray "------------------------------------------------------"
Write-Host -Fore Cyan "       CSIRT IR Collection Script - USB, v3.1" 
Write-Host -Fore DarkCyan "       https://github.com/dwmetz/CSIRT-Collect"
Write-Host -Fore Cyan "       @dwmetz | bakerstreetforensics.com"
Write-Host -Fore Gray "------------------------------------------------------"
Start-Sleep -Seconds 3
## Establish collection directory
Set-Location Collections
mkdir $env:computername -Force
Set-Location ..
## capture memory image
.\Memory\MRC.exe /accepteula /go /silent
Start-Sleep -Seconds 5
Write-Host -Fore Cyan "Initiating Magnet Ram Capture."
Write-Host -Fore Cyan "Capturing memory..."
Write-Host -Fore Cyan "This process may take several minutes..."
Wait-Process -name "MRC"
## document the OS build information
Write-Host -Fore Cyan "Determining OS build info..."
[System.Environment]::OSVersion.Version > windows_build.txt
Write-Host -Fore Cyan "Cleaning up"
Get-ChildItem -Filter '*windows_build*' -Recurse | Rename-Item -NewName {$_.name -replace 'windows', $env:computername }
Move-Item -Path .\*.txt -Destination .\Collections\$env:COMPUTERNAME\
Set-Location Memory
Get-ChildItem -Filter 'MagnetRAMCapture*' -Recurse | Rename-Item -NewName {$_.name -replace 'MagnetRAMCapture', $env:computername }
Get-ChildItem -Filter '*.raw' -Recurse | Rename-Item -NewName {$_.name -replace ' - ', '_' }
Get-ChildItem -Filter '*.raw' -Recurse | Rename-Item -NewName {$_.name -replace ' ', '_' }
Move-Item -Path .\*.raw -Destination ..\Collections\$env:COMPUTERNAME\
## execute the KAPE "OS" collection
Set-Location ..
Write-Host -Fore Cyan "Collecting OS artifacts..."
Start-Sleep -Seconds 3
Kape\kape.exe --tsource C: --tdest Collections\$env:COMPUTERNAME --target KapeTriage --vhdx $env:COMPUTERNAME --zv false --module MagnetForensics_EDD --mdest Collections\$env:computername\Decrypt
## Encryption Detection & Recovery
get-content .\Collections\$env:COMPUTERNAME\Decrypt\LiveResponse\EDD.txt
Write-Host -fore cyan "Retrieving BitLocker Keys"
(Get-BitLockerVolume -MountPoint C).KeyProtector > bitlocker_recovery.txt
Get-ChildItem -Filter 'bitlocker*' -Recurse | Rename-Item -NewName {$_.name -replace 'bitlocker', $env:computername }
Move-Item -Path .\*.txt -Destination .\Collections\$env:COMPUTERNAME\Decrypt\LiveResponse
## indicates completion
Set-Content -Path .\Collections\$env:COMPUTERNAME\collection-complete.txt -Value "Collection complete: $((Get-Date).ToString())"
Write-Host -Fore Cyan "** Process Complete **"