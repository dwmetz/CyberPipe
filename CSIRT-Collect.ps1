<#
CSIRT-Collections.ps1
https://github.com/dwmetz/CSIRT-Collect
Doug Metz dwmetz@gmail.com
Function: This script will map a drive to the "Collections" share, capture a memory image and a KAPE collection on the computer, and transfer the output back to the network share.
#>
Write-Host -Fore White "--------------------------------------------------"
Write-Host -Fore Cyan "       CSIRT IR Collection Script, v1.5" 
Write-Host -Fore Cyan "       (c) 2021 dwmetz@gmail.com" 
Write-Host -Fore White "--------------------------------------------------"
Start-Sleep -Seconds 3

Set-ExecutionPolicy -Scope CurrentUser Unrestricted

## map the network drive and change to that directory
Write-Host -Fore Green "Mapping network drive..."

$Networkpath = "X:\" 

If (Test-Path -Path $Networkpath) {
    Write-Host -Fore Green "Drive Exists already"
}
Else {
    #map network drive
    (New-Object -ComObject WScript.Network).MapNetworkDrive("X:","\\Synology\Collections") 

    #check mapping again
    If (Test-Path -Path $Networkpath) {
        Write-Host -Fore Green "Drive has been mapped"
    }
    Else {
        Write-Host -For Red "Error mapping drive"
    }
}

# create local memory directory
Write-Host -Fore Green "Setting up local directory..."
mkdir C:\temp\IR -Force
Set-Location C:\temp\IR
Write-Host -Fore Green "Copying tools..."
robocopy "\\Synology\Collections\Memory" . *.exe
## capture memory image
Write-Host -Fore Green "Capturing memory..."
.\winpmem.exe memdump.raw
## zip the memory image
Write-Host -Fore Green "Zipping the memory image..."
.\7za a -t7z memdump.7z memdump.raw -mx1
## delete the raw file
Remove-Item memdump.raw
Write-Host -Fore Green "Deleting raw image..."
## rename the zip file to the hostname of the computer
Write-Host -Fore Green "Renaming file..."
Get-ChildItem -Filter "*memdump*" -Recurse | Rename-Item -NewName {$_.name -replace 'memdump', $env:computername }

## document the OS build information
Write-Host -Fore Green "Determining OS build info..."
[System.Environment]::OSVersion.Version > C:\Temp\IR\windowsbuild.txt
Write-Host -Fore Green "Renaming file..."
Get-ChildItem -Filter "*windowsbuild*" -Recurse | Rename-Item -NewName {$_.name -replace 'windowsbuild', $env:computername }

## create output directory on "Collections" share
mkdir X:\$env:COMPUTERNAME

Write-Host -Fore Green "Copying memory image to network..."

## copy memory image to network
robocopy . "\\Synology\Collections\$env:COMPUTERNAME" *.7z *.txt

## delete the directory and contents 
Write-Host -Fore Green "Removing temporary files"
Set-Location C:\TEMP
Remove-Item -LiteralPath "C:\temp\IR" -Force -Recurse

## create the KAPE directory on the client
Write-Host -Fore Green "Creating KAPE directory on host..."
mkdir C:\Temp\KAPE -Force

## execute the KAPE "OS" collection
Write-Host -Fore Green "Collecting OS artifacts..."
Set-Location X:\KAPE
.\kape.exe --tsource C: --tdest C:\Temp\KAPE --target !SANS_Triage --vhdx $env:COMPUTERNAME


## transfer evidence to share
Set-Location C:\Temp\Kape
robocopy . "\\Synology\Collections\$env:COMPUTERNAME"

## delete the local directory and contents 
Write-Host -Fore Green "Removing temporary files"
Set-Location C:\TEMP
Remove-Item -LiteralPath "C:\temp\KAPE" -Force -Recurse
Set-Content -Path X:\$env:COMPUTERNAME\transfer-complete.txt -Value "Transfer complete: $((Get-Date).ToString())"
Remove-PSDrive -Name X
Write-Host -Fore Cyan "** Process Complete **"

## End