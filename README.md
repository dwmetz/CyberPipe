<div align="center">
 <img style="padding:0;vertical-align:bottom;" height="158" width="311" src="/images/BSF.png"/>
 <p>
  <h2>
   CyberPipe v5
  </h2>
  <h5>
      An easy to use PowerShell script to collect memory and disk forensics for DFIR investigations.
   </h5>
<p>
<p>
 </div>
<div align="center">
  <img style="padding:0;vertical-align:bottom;" height="340" width="526" src="/images/screenshot.png"/>
  <div align="left">
  <h5>
   Functions:
  </h5>

- :ram: Capture a memory image with MAGNET DumpIt for Windows, (x32, x64, ARM64), or MAGNET RAM Capture on legacy systems;
- :computer: Create a Triage collection* with MAGNET Response;
- :closed_lock_with_key: Check for encrypted disks with Encrypted Disk Detector;
- :key: Recover the active BitLocker Recovery key;
- :floppy_disk: Save all artifacts, output, and audit logs to USB or source network drive.

*There are collection profiles available for: 
>- Volatile Artifacts
>- Triage Collection (Volatile, RAM, Pagefile, Triage artifacts)
>- Just RAM
>- RAM & Pagefile
>- or build your own using the RESPONSE CLI options


<h5>
   Prerequisites:
</h5>

>- [MAGNET Response](https://www.magnetforensics.com/resources/magnet-response/)
>- [Encrypted Disk Detector](https://www.magnetforensics.com/resources/encrypted-disk-detector/) 


<h5>
Network Collections:
</h5>

CyberPipe 5 also has the capability to write captures to a network repository. Just un-comment # the Network section and update the `\\server\share` line to reflect your environment.

In this configuration it can be included as part of automation functions like a collection being triggered from an event logged on the EDR.

<h5>
   Prior version (KAPE support):
</h5>

If you're a prior user of CyberPipe and want to use the previous method where KAPE facilitates the collection with the MAGNET tools, or have made other KAPE modifications, use v4.01 `CyberPipe.v4.01.ps1`



> Note: this script was previously titled CSIRT-Collect. Project name and repo updated with version 4.0.

For more information visit [BakerStreetForensics.com](https://bakerstreetforensics.com/2024/02/14/cyberpipe-version-5-0/)

