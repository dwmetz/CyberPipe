<div align="center">
 <img style="padding:0;vertical-align:bottom;" height="158" width="311" src="/images/BSF.png"/>
 <p>
  <h2>
    CyberPipe v5.1
  </h2>
  <h6>
  (formerly CSIRT-Collect)
  </h6>

  <h5>
      An easy to use PowerShell script to collect memory and disk forensics for DFIR investigations.
   </h5>
<p>
<p>
 </div>
<div align="center">
  <img style="padding:0;vertical-align:bottom;" height="340" width="526" src="/images/Screenshot.png"/>
  <div align="left">
  <h5>
   Functions:
  </h5>

- :ram: Capture a memory image with MAGNET DumpIt (supports x86, x64, and ARM64) or MAGNET RAM Capture for legacy systems.
- :computer: Collect triage data using MAGNET Response CLI, with selectable profiles or custom options.
- :closed_lock_with_key: Detect full disk encryption using MAGNET Encrypted Disk Detector.
- :key: Recover the active BitLocker Recovery Key (if accessible).
- :floppy_disk: Store collected data, logs, and memory images to a USB device or a defined network location.

Collection profiles include:
- Volatile Only
- RAM Only
- RAM + Pagefile
- Triage (Volatile + RAM + Pagefile + System Artifacts)
- Custom: pass flags to MAGNET Response CLI


<h5>
   Prerequisites:
</h5>

>- [MAGNET Response](https://www.magnetforensics.com/resources/magnet-response/)
>- [MAGNET Encrypted Disk Detector](https://www.magnetforensics.com/resources/encrypted-disk-detector/) 
>- [MAGNET RAM Capture](https://www.magnetforensics.com/resources/magnet-ram-capture/)


<h5>
Network Collections:
</h5>

CyberPipe 5.1 supports saving output directly to a network share. To enable this, uncomment the `#Network` section in the script and set the appropriate UNC path (e.g., `\\server\share`). This is ideal for automated DFIR workflows triggered by EDR or SOC alerts.


<h5>
New in 5.1:
</h5>

- Improved network storage support
- Custom MAGNET Response profiles
- Enhanced logging and error handling


<h5>
Usage Examples:
</h5>

- **Run full triage (default collection profile) to local USB drive:** (RAM, Pagefile, Volatile, System Files)
  ```powershell
  .\CyberPipe.ps1 
  ```

- **Run RAM & Operating System Files (triage light) capture:**
  ```powershell
  .\CyberPipe.ps1 -CollectionProfile RAMSystem
  ```
- **Run memory-only capture:**
  ```powershell
  .\CyberPipe.ps1 -CollectionProfile RAMOnly
  ```

 
- **Run RAM & Pagefile capture:**
  ```powershell
  .\CyberPipe.ps1 -CollectionProfile RAMPage
  ``` 

- **Run RAM & Operating System Files (triage light) capture:**
  ```powershell
  .\CyberPipe.ps1 -CollectionProfile RAMSystem
  ```
- **Run volatile-only capture:**
  ```powershell
  .\CyberPipe.ps1 -CollectionProfile Volatile
  ```
- _You can modify or create custom profiles by specifying CLI arguments supported by MAGNET Response._

<h5>
Tool Directory Structure:
</h5>

- **USB Collections:** The `Tools` directory should be located alongside the script:
  ```
  E:\Triage\CyberPipe\CyberPipe.ps1
  E:\Triage\CyberPipe\Tools\
  ```

- **Network Collections:** The `Tools` directory should be placed in the root of the network share:
  ```
  \\Server\share\Tools\
  ```

<h5>
   Prior version (KAPE support):
</h5>

If you previously used CyberPipe with KAPE (prior to v5), the older workflow remains available in `CyberPipe.v4.01.ps1`.

> Note: CyberPipe was previously known as CSIRT-Collect. The project was renamed starting with version 4.0.

For more information visit [BakerStreetForensics.com](https://bakerstreetforensics.com/2024/02/14/cyberpipe-version-5-0/)
