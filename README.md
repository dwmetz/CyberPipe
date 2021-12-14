# CSIRT-Collect
A PowerShell script to collect memory and (triage) disk forensics for incident response investigations.

The script leverages a network share, from which it will access and copy the required executables and subsequently upload the acquired evidence to the same share post-collection.

Permission requirements for said directory will be dependent on the nuances of the environment and what credentials are used for the script execution (interactive vs. automation)

In the demonstration code, a network location of \\Synology\Collections can be seen. This should be changed to reflect the specifics of your environment.

Collections folder needs to include:
- subdirectory KAPE; copy the directory from existing install
- subdirectory MEMORY; 7za.exe command line version of 7zip and Magnet Ram Capture.

For a walkthough of the script https://bakerstreetforensics.com/2021/12/13/adding-ram-collections-to-kape-triage/

## CSIRT-Collect

- Maps to existing network drive -
- - Subdir 1: “Memory” – Winpmem and 7zip executables
- - Subdir 2: ”KAPE” – directory (copied from local install)
- Creates a local directory on asset
- Copies the Memory exe files to local directory
- Captures memory with Magnet Ram Capture
- When complete, ZIPs the memory image
- Renames the zip file based on hostname
- Documents the OS Build Info (no need to determine profile for Volatility)
- Compressed image is copied to network directory and deleted from host after transfer complete
- New temp Directory on asset for KAPE output
- KAPE !SANS_Triage collection is run using VHDX as output format [$hostname.vhdx]
- VHDX transfers to network
- Removes the local KAPE directory after completion
- Writes a “Process complete” text file to network to signal investigators that collection is ready for analysis.

## CSIRT-Collect_USB

Essentially the same functionality as CSIRT-Collect.ps1 with the exception that it is intented to be run from a USB device. The extra compression operations on the memory image and KAPE .vhdx have been removed.
There is a slight change to the folder structure for the USB version.
On the root of the USB:
- CSIRT-Collect_USB.ps1
- folder (empty to start) titled 'Collections'
- folders for KAPE and Memory - same as above

Execution:
-Open PowerShell as Adminstrator
-Navigate to the USB device
-Execute ./CSIRT-Collect_USB.ps1



