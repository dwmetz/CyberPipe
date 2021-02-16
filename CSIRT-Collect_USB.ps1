<#
CSIRT-Collect_USB.ps1
https://github.com/dwmetz/CSIRT-Collect
Doug Metz dwmetz@gmail.com
Function: This script will capture a memory image and a KAPE collection to the USB device.

On the root of the USB:
-CSIRT-Collect_USB.ps1
-folder (empty to start) titled 'Collections'
-folders for KAPE and Memory - see ReadMe for details.

Execution:
-Open PowerShell as Adminstrator
-Navigate to the USB device
-Execute ./CSIRT-Collect_USB.ps1
#>
Write-Host -Fore White "--------------------------------------------------"
Write-Host -Fore Red "       CSIRT IR Collection Script - USB, v1.0" 
Write-Host -Fore Cyan "       (c) 2021 dwmetz@gmail.com" 
Write-Host -Fore White "--------------------------------------------------"
Start-Sleep -Seconds 3
## Establish collection directory
Set-Location Collections
mkdir $env:computername -Force
## capture memory image
Write-Host -Fore Green "Capturing memory..."
Start-Sleep -Seconds 3
\Memory\winpmem.exe \Collections\$env:computername\memdump.raw
## rename the zip file to the hostname of the computer
Write-Host -Fore Green "Renaming file..."
Get-ChildItem -Filter "*memdump*" -Recurse | Rename-Item -NewName {$_.name -replace 'memdump', $env:computername }
## document the OS build information (memory profile)
Write-Host -Fore Green "Determining OS build info..."
Start-Sleep -Seconds 3
[System.Environment]::OSVersion.Version > $env:computername\windowsbuild.txt
Set-Location ..
## execute the KAPE "OS" collection
Write-Host -Fore Green "Collecting OS artifacts..."
Start-Sleep -Seconds 3
Kape\kape.exe --tsource C: --tdest Collections\$env:COMPUTERNAME --target !SANS_Triage --vhdx $env:COMPUTERNAME --zv false
## indicates completion
Set-Content -Path \Collections\$env:COMPUTERNAME\collection-complete.txt -Value "Collection complete: $((Get-Date).ToString())"
Write-Host -Fore Cyan "** Process Complete **"