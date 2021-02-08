# CSIRT-Collect
A PowerShell script to collect memory and (triage) disk forensics for incident response investigations.

The script leverages a network share, from which it will access and copy the required executables and subsequently upload the acquired evidence to the same share post-collection.

Permission requirements for said directory will be dependent on the nuances of the environment and what credentials are used for the script execution (interactive vs. automation)

In the demonstration code, a network location of \\Synology\Collections can be seen. This should be changed to reflect the specifics of your environment.

Collections folder needs to include:
- subdirectory KAPE; copy the directory from existing install
- subdirectory MEMORY; 7za.exe command line version of 7zip and winpmem.exe


CSIRT-Collect.ps1 Operations:

Maps to existing network drive -

Subdir 1: “Memory” – Winpmem and 7zip executables

Subdir 2: ”KAPE” – directory (copied from local install)

Creates a local directory on asset

Copies the Memory exe files to local directory

Captures memory with Winpmem

When complete, ZIPs the memory image

Renames the zip file based on hostname

Documents the OS Build Info (no need to determine profile for Volatility)

Compressed image is copied to network directory and deleted from host after transfer complete

New temp Directory on asset for KAPE output

KAPE !SANS_Triage collection is run using VHDX as output format [$hostname.vhdx]

VHDX transfers to network

Removes the local KAPE directory after completion

Writes a “Process complete” text file to network to signal investigators that collection is ready for analysis.
