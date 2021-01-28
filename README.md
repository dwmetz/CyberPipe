# CSIRT-Collect
A PowerShell script to collect memory and (triage) disk forensics for incident response investigations.

The script leverages a network share, from which it will access and copy the required executables and subsequently upload the acquired evidence to the same share.

Permission requirements for said directory will be dependent on the nuances of the environment and what credentials are used for the script execution (interactive vs. automation)

For the demonstration a network location of \\Synology\Collections can be seen for reference, and should be changed to reflect the specifics of your environment.