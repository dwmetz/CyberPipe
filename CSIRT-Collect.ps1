<#
CSIRT-Collect.ps1 v3.0
https://github.com/dwmetz
Author: @dwmetz
Function: This script will 
- map a drive to the "Collections" share, 
- capture a memory image with Magnet Ram Capture,
- capture a triage collection with KAPE, 
- transfer the output back to the network share.

Prerequisites:
Network share location with "Collections" folder. Within 'Collections', 2 subdirectories:
- Memory, containing Magnet Ram Capture (MRC.exe) and CLI version of 7zip (7za.exe)
- KAPE (default directory as installed)

#>
Write-Host -Fore White "--------------------------------------------------"
Write-Host -Fore Cyan "       CSIRT IR Collection Script, v3.0" 
Write-Host -Fore Cyan "       (c) 2021 @dwmetz" 
Write-Host -Fore White "--------------------------------------------------"
Start-Sleep -Seconds 3
## map the network drive and change to that directory
Write-Host -Fore Cyan "Mapping network drive..."
$Networkpath = "X:\" 
If (Test-Path -Path $Networkpath) {
    Write-Host -Fore Cyan "Drive Exists already"
}
Else {
    #map network drive
    (New-Object -ComObject WScript.Network).MapNetworkDrive("X:","\\Synology\Collections") 
    #check mapping again
    If (Test-Path -Path $Networkpath) {
        Write-Host -Fore Cyan "Drive has been mapped"
    }
    Else {
        Write-Host -For Red "Error mapping drive"
    }
}
# create local memory directory
Write-Host -Fore Cyan "Setting up local directory..."
mkdir C:\temp\IR -Force
Set-Location C:\temp\IR
Write-Host -Fore Cyan "Copying tools..."
robocopy "\\Synology\Collections\Memory" . *.exe
Write-Host -Fore Cyan "Capturing RAM Image..."
.\MRC.exe /accepteula /go /silent
Start-Sleep -Seconds 5
Write-Host -Fore Cyan "Waiting for capture to complete..."
Wait-Process -name "MRC"
## document the OS build information
Write-Host -Fore Cyan "Determining OS build info..."
[System.Environment]::OSVersion.Version > windowsbuild.txt
Get-ChildItem -Filter '*windowsbuild*' -Recurse | Rename-Item -NewName {$_.name -replace 'windowsbuild', $env:computername }
Write-Host -Fore Cyan "Zipping the memory image..."
.\7za a -t7z memdump.7z *.raw *.txt -mx1
## clean up files
Write-Host -Fore Cyan "Cleaning up..."
Remove-Item *.raw
Remove-Item *.txt
## rename the zip file to the hostname of the computer
Write-Host -Fore Cyan "Renaming file..."
Get-ChildItem -Filter '*memdump*' -Recurse | Rename-Item -NewName {$_.name -replace 'memdump', $env:computername }
Write-Host -Fore Cyan "RAM Capture Completed."
## create output directory on "IR" share
mkdir X:\$env:COMPUTERNAME
Write-Host -Fore Cyan "Copying memory image to network..."
## copy memory image to network
robocopy . "\\Synology\Collections\$env:COMPUTERNAME" *.7z *.txt
## delete the directory and contents 
Write-Host -Fore Cyan "Removing temporary files"
Set-Location C:\TEMP
Remove-Item -LiteralPath "C:\temp\IR" -Force -Recurse
## create the KAPE directory on the client
Write-Host -Fore Cyan "Creating KAPE directory on host..."
mkdir C:\Temp\KAPE -Force
## execute the KAPE "OS" collection
Write-Host -Fore Cyan "Collecting OS artifacts..."
Set-Location X:\KAPE
.\kape.exe --tsource C: --tdest C:\Temp\KAPE --target KapeTriage --vhdx $env:COMPUTERNAME
## transfer evidence to share
Set-Location C:\Temp\Kape
robocopy . "\\Synology\Collections\$env:COMPUTERNAME"
## delete the local directory and contents 
Write-Host -Fore Cyan "Removing temporary files"
Set-Location C:\TEMP
Remove-Item -LiteralPath "C:\temp\KAPE" -Force -Recurse
Set-Content -Path X:\$env:COMPUTERNAME\transfer-complete.txt -Value "Transfer complete: $((Get-Date).ToString())"
Remove-PSDrive -Name X
Write-Host -Fore Cyan "** Process Complete **"