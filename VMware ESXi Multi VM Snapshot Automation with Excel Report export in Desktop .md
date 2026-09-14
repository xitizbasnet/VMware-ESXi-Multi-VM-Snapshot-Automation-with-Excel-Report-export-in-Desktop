# VMware ESXi Multi-VM Snapshot Automation + Excel Report

> [!NOTE]
> **Purpose:** Automate VM snapshots on a standalone VMware ESXi host and generate an Excel report containing ESXi, VM, and snapshot information.

---

## 📘 Overview

This automation script performs the following operations:

1. Connects to a configured VMware ESXi host.
2. Authenticates using the ESXi username and a password entered securely at runtime.
3. Processes a predefined list of virtual machines.
4. Collects ESXi host information.
5. Collects VM information, including:

   * VM name
   * Power state
   * CPU
   * Memory
   * Guest OS
   * VM version
   * Datastore
   * VM IP address
6. Creates an automated snapshot for each configured VM.
7. Records successful and failed snapshot operations.
8. Generates an Excel report containing:

   * **ESXi Summary**
   * **VM Details**
   * **Snapshot Details**
9. Saves the Excel report to the current user's Desktop.
10. Automatically opens the generated Excel report.
11. Displays a final execution status in the PowerShell console.

---

## 🏗️ Script Architecture

The script is organized into the following major areas:

| Section                         | Purpose                                               |
| ------------------------------- | ----------------------------------------------------- |
| Backend Configuration           | Defines ESXi host, username, and VM list              |
| Start                           | Displays execution information                        |
| Desktop Location                | Determines where the Excel report will be saved       |
| Excel Module                    | Checks and installs the `ImportExcel` module          |
| Prepare Excel File              | Generates the report filename                         |
| Data Collection Arrays          | Stores ESXi, VM, and snapshot report information      |
| Password                        | Requests the ESXi password securely                   |
| Connect to ESXi                 | Establishes the PowerCLI connection                   |
| Get ESXi Host Information       | Collects ESXi version, build, and product information |
| Process Each VM                 | Finds each VM and processes snapshot creation         |
| Collect VM Information          | Collects VM configuration and network information     |
| Create Snapshot                 | Creates the automated snapshot                        |
| ESXi Summary Report             | Builds the ESXi summary data                          |
| Disconnect from ESXi            | Closes the PowerCLI connection                        |
| Generate Excel Report           | Creates and formats the Excel workbook                |
| Final Status                    | Displays execution results                            |
| Open Excel Report Automatically | Opens the generated workbook                          |

---

# 1. ⚙️ Prerequisites

Before executing the script, verify that the workstation has:

* Windows PowerShell or PowerShell 7.
* VMware PowerCLI installed.
* Network connectivity to the ESXi host.
* Valid ESXi credentials.
* Appropriate permissions to:

  * Connect to the ESXi host.
  * Read VM information.
  * Create VM snapshots.
* Microsoft Excel installed if the generated report needs to be opened automatically.
* Internet/package repository access if the `ImportExcel` PowerShell module needs to be installed automatically.

> [!IMPORTANT]
> The script does **not store the ESXi password in the script**. The password is requested interactively using `Read-Host -AsSecureString`.

---

# 2. 🔐 Authentication and Security

The ESXi username is configured in the backend section:

```powershell
$ESXiUsername = "root"
```

The password is entered interactively:

```powershell
$SecurePassword = Read-Host "Enter ESXi password" -AsSecureString
```

The password is then converted into a PowerShell credential object:

```powershell
$Credential = [System.Management.Automation.PSCredential]::new(
    $ESXiUsername,
    $SecurePassword
)
```

> [!WARNING]
> Avoid hard-coding the ESXi password into the PowerShell script. Keeping credentials out of the script reduces the risk of credential exposure through source control, backups, or accidental sharing.

---

# 3. 🖥️ Backend Configuration

The following variables must be reviewed before deploying the script:

```powershell
$ESXiHost     = "192.168.1.3"
$ESXiUsername = "root"
```

The VMs to be processed are configured in `$VMList`.

> [!IMPORTANT]
> VM names must match the names registered on the ESXi host. A VM name that does not exist is recorded as **VM NOT FOUND** and the script continues processing the remaining VMs.

---

# 4. 📋 Configured Virtual Machines

The current configuration contains the following VMs:

```text
AMS
LMS
UnifiController
DC02
DC01
UAT
SCP
PrintServer
SIEMWAZUH
ITS
```

The script automatically calculates the VM count using:

```powershell
$VMList.Count
```

---

# 5. 📦 Excel Reporting Module

The script uses the PowerShell **ImportExcel** module to create the Excel workbook.

If the module is not installed, the script attempts to:

1. Check for the NuGet package provider.
2. Install NuGet if required.
3. Install `ImportExcel` for the current user.
4. Import the module.

The installation uses:

```powershell
Install-Module `
    ImportExcel `
    -Scope CurrentUser `
    -Force `
    -AllowClobber `
    -ErrorAction Stop
```

> [!TIP]
> Using `-Scope CurrentUser` allows the module to be installed without requiring a system-wide PowerShell module installation.

---

# 6. 📊 Excel Report

The generated report filename follows this pattern:

```text
ESXi_<ESXi-IP>_Snapshot_Report_<timestamp>.xlsx
```

For example:

```text
ESXi_192.168.1.3_Snapshot_Report_2026-09-14_18-00-00.xlsx
```

The exact timestamp is generated by:

```powershell
$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
```

The report is saved to the current user's Desktop.

---

# 7. 📑 Excel Worksheets

The workbook contains three worksheets.

## 7.1 ESXi Summary

The **ESXi Summary** worksheet contains:

* ESXi Host Name
* ESXi IP Address
* ESXi Product
* ESXi Version
* ESXi Build
* ESXi Full Version
* Configured VM Count
* VMs Found
* VMs Not Found
* Successful Snapshots
* Failed Snapshots
* Report Generated

---

## 7.2 VM Details

The **VM Details** worksheet contains:

* VM Name
* Power State
* CPU
* Memory GB
* Guest OS
* VM Version
* Datastore
* VM Host
* VM IP Address
* Snapshot Result
* Error

---

## 7.3 Snapshot Details

The **Snapshot Details** worksheet contains:

* VM Name
* Snapshot Name
* Snapshot Created
* Snapshot Size GB
* Memory Included
* Quiesced
* Description
* Snapshot State
* Result
* Error

---

# 8. 📸 Snapshot Naming Convention

Each automated snapshot receives a dynamically generated name:

```text
AUTO-<VMName>-<YYYY-MM-DD-HH-mm-ss>
```

For example:

```text
AUTO-AMS-2026-09-14-18-00-00
```

The timestamp is generated with:

```powershell
$TimeStamp = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"
```

The final snapshot name is created using:

```powershell
$SnapshotName = "AUTO-$VMName-$TimeStamp"
```

> [!TIP]
> A timestamp-based naming convention makes automated snapshots easier to identify and correlate with execution times.

---

# 9. 🖥️ VM Information Collection

For each VM, the script collects:

```powershell
$PowerState = $VM.PowerState
$NumCPU     = $VM.NumCpu
$MemoryGB   = $VM.MemoryGB
$GuestOS    = $VM.Guest.OSFullName
```

The script also attempts to collect:

* VM hardware version
* Datastore
* VM IP address

If information cannot be obtained, the value is recorded as:

```text
Unavailable
```

---

# 10. 🌐 VM IP Address Detection

The script examines the guest IP address information and filters for IPv4 addresses using:

```powershell
$_ -match '^\d{1,3}(\.\d{1,3}){3}$'
```

Multiple addresses are combined into a comma-separated value.

If no matching address is found, the report records:

```text
Unavailable
```

> [!NOTE]
> VM guest IP information depends on VMware Tools/guest information being available from the VM. An unavailable IP does not necessarily indicate that the VM has no network connectivity.

---

# 11. 💾 Datastore Information

The script retrieves datastore information associated with each VM:

```powershell
Get-Datastore `
    -VM $VM `
    -Server $VIServer `
    -ErrorAction SilentlyContinue
```

If multiple datastores are returned, the names are combined into a comma-separated list.

If no datastore information is available:

```text
Unavailable
```

is recorded.

---

# 12. 📸 Snapshot Creation

Snapshots are created using VMware PowerCLI:

```powershell
$NewSnapshot = New-Snapshot `
    -VM $VM `
    -Name $SnapshotName `
    -Description "Automated snapshot created on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" `
    -Memory:$false `
    -Quiesce:$false `
    -ErrorAction Stop
```

The configured snapshot options are:

| Option         |    Value | Meaning                                     |
| -------------- | -------: | ------------------------------------------- |
| `-Memory`      | `$false` | VM memory is not included in the snapshot   |
| `-Quiesce`     | `$false` | Guest filesystem quiescing is not requested |
| `-ErrorAction` |   `Stop` | Snapshot errors are caught by the script    |

> [!WARNING]
> A VMware snapshot is **not a replacement for a backup**. Snapshots should be treated as temporary recovery points and should not be retained indefinitely.

---

# 13. ✅ Successful Snapshot Processing

When a snapshot is successfully created, the script:

1. Increments the successful snapshot counter.
2. Records the snapshot completion time.
3. Retrieves detailed snapshot information.
4. Records snapshot size.
5. Records whether memory was included.
6. Records whether quiescing was used.
7. Records the snapshot description.
8. Adds the information to the Snapshot Details report.
9. Updates the VM report with `SUCCESS`.

The console displays:

```text
SNAPSHOT CREATED SUCCESSFULLY
```

---

# 14. ❌ Failed Snapshot Processing

If snapshot creation fails, the script:

1. Increments the failed snapshot counter.
2. Captures the error message.
3. Displays the error in the console.
4. Adds the failed snapshot to the Snapshot Details worksheet.
5. Updates the corresponding VM entry with `FAILED`.
6. Continues processing the remaining VMs.

This design prevents one VM failure from stopping the entire automation process.

---

# 15. 🔎 VM Not Found Handling

If a configured VM cannot be found, the script:

* Increments `$VMsNotFound`.
* Increments `$FailedSnapshots`.
* Records the error message.
* Adds the VM to the VM Details worksheet.
* Adds a failed entry to the Snapshot Details worksheet.
* Continues with the next VM.

The VM is recorded with values such as:

```text
Power State      = Unavailable
CPU              = Unavailable
Memory GB        = Unavailable
Guest OS         = Unavailable
VM Version       = Unavailable
Datastore        = Unavailable
VM IP Address    = Unavailable
Snapshot Result  = FAILED
```

---

# 16. 📈 Progress Tracking

The script displays progress while processing the VM list.

The progress percentage is calculated using:

```powershell
$Percent = [math]::Round(
    ($CurrentVM / $TotalVMs) * 100
)
```

The PowerShell progress display identifies:

```text
Processing <VMName> [current of total]
```

This is useful when the script is processing a larger VM inventory.

---

# 17. 🔌 ESXi Connection

The script connects to the configured ESXi host using:

```powershell
$VIServer = Connect-VIServer `
    -Server $ESXiHost `
    -Credential $Credential `
    -ErrorAction Stop
```

If the connection succeeds, the script displays:

```text
Connected successfully.
```

If the connection fails, the script displays the PowerCLI error and terminates the operation.

> [!WARNING]
> Verify ESXi network connectivity, TCP port `443`, credentials, and ESXi permissions before troubleshooting the script itself.

---

# 18. 🖥️ ESXi Host Information

The script retrieves ESXi host information using:

```powershell
Get-VMHost
```

The following information is collected:

* ESXi host name
* ESXi IP address
* ESXi version
* ESXi build
* ESXi full version
* VMware product name

The information is added to the **ESXi Summary** worksheet.

---

# 19. 🔌 Disconnecting from ESXi

After VM processing is complete, the script disconnects from ESXi:

```powershell
Disconnect-VIServer `
    -Server $VIServer `
    -Confirm:$false `
    -ErrorAction SilentlyContinue |
    Out-Null
```

The use of:

```powershell
-Confirm:$false
```

prevents PowerCLI from prompting for confirmation.

---

# 20. 📊 Excel Formatting

The generated workbook is formatted automatically.

The script applies:

* Auto-sizing
* Auto-filtering
* Frozen header rows
* Bold header rows
* Specific column widths
* Centered header alignment

The formatting is applied separately to:

* `ESXi Summary`
* `VM Details`
* `Snapshot Details`

---

# 21. 📁 Report Location

The script determines the current user's Desktop with:

```powershell
$Desktop = [Environment]::GetFolderPath("Desktop")
```

The report is then created using:

```powershell
$ExcelFile = Join-Path `
    $Desktop `
    "ESXi_${ESXiHost}_Snapshot_Report_$DateStamp.xlsx"
```

> [!NOTE]
> The report location is based on the Windows account running the script.

---

# 22. 📋 Final Execution Status

At the end of execution, the script displays:

* ESXi host
* ESXi version
* ESXi build
* Configured VMs
* VMs found
* VMs not found
* Successful snapshots
* Failed snapshots
* Excel report filename
* Report location
* Completion timestamp

If there are no failed snapshots:

```text
Status              : SUCCESS
```

If one or more snapshots fail:

```text
Status              : COMPLETED WITH ERRORS
```

---

# 23. 📖 Operational Procedure

## Step 1 — Open PowerShell

Open PowerShell using an account that has the required VMware administration permissions.

## Step 2 — Verify PowerCLI

Verify that VMware PowerCLI is available:

```powershell
Get-Module VMware.PowerCLI -ListAvailable
```

## Step 3 — Verify ESXi Connectivity

Confirm that the workstation can reach the ESXi management IP address.

Example:

```powershell
Test-Connection 192.168.1.3
```

## Step 4 — Review Backend Configuration

Open the script and verify:

```powershell
$ESXiHost
$ESXiUsername
$VMList
```

## Step 5 — Execute the Script

Run the PowerShell script.

When prompted:

```text
Enter ESXi password:
```

enter the ESXi password.

## Step 6 — Monitor Execution

Review:

* ESXi connection status
* VM discovery
* Snapshot creation
* Errors
* Progress information

## Step 7 — Review Excel Report

After successful processing, the generated Excel report is automatically opened.

Review all three worksheets:

1. **ESXi Summary**
2. **VM Details**
3. **Snapshot Details**

---

# 24. 💡 Best Practices

### VM Configuration

* Verify VM names before adding them to `$VMList`.
* Remove retired VMs from the configuration.
* Confirm that newly created VMs are added when required.
* Avoid duplicate VM entries.

### Snapshot Management

* Use descriptive and consistent snapshot names.
* Monitor snapshot age and size.
* Avoid keeping snapshots for unnecessarily long periods.
* Do not treat snapshots as backups.
* Review failed snapshot operations.

### Credentials

* Never store the ESXi password directly in the script.
* Use secure credential handling.
* Restrict access to the PowerShell script.
* Do not commit credentials to GitHub.

### Reporting

* Retain generated reports according to the organization's IT documentation/retention policy.
* Review failed VMs before considering the execution successful.
* Use the report timestamp to correlate automation runs with operational events.

 
---

# 25. 🛠️ Troubleshooting

## ImportExcel Module Error

If the script reports:

```text
EXCEL MODULE ERROR
```

verify:

```powershell
Get-Module -ListAvailable -Name ImportExcel
```

If necessary, install it manually:

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

---

## ESXi Connection Failed

If the script reports:

```text
ESXi CONNECTION FAILED
```

check:

1. ESXi management IP address.
2. Network connectivity.
3. TCP port 443.
4. ESXi username.
5. ESXi password.
6. VMware PowerCLI installation.
7. ESXi permissions.

---

## VM Not Found

If the report shows:

```text
VM NOT FOUND
```

verify the VM name directly through PowerCLI:

```powershell
Get-VM -Server 192.168.1.13
```

Then compare the actual VM name with the corresponding entry in:

```powershell
$VMList
```

---

## Snapshot Failed

If a snapshot fails:

1. Review the error recorded in the Excel report.
2. Verify that the VM exists.
3. Verify available datastore capacity.
4. Check the current snapshot hierarchy.
5. Check ESXi task/event information.
6. Confirm that the account has snapshot permissions.

---

## Excel Report Was Not Opened

If the report is created but does not open automatically, manually open the file from the Desktop.

The script also displays:

```text
You can manually open:
<Excel file path>
```

---

# 26. 📜 Original PowerShell Script

> [!IMPORTANT]
> The following section preserves the supplied PowerShell script content. The script itself has not been intentionally shortened or functionally rewritten.

```powershell
# =====================================================================
# VMware ESXi Multi-VM Snapshot Automation + Excel Report
# Author      : Xitiz Basnet
# Description : Creates snapshots for configured ESXi VMs and exports
#               ESXi, VM and snapshot information to an Excel report.
#               The Excel report is saved to the current user's Desktop
#               and opened automatically after completion.
# =====================================================================


# =====================================================================
# BACKEND CONFIGURATION
# =====================================================================

$ESXiHost     = "192.168.1.3"
$ESXiUsername = "root"

$VMList = @(
"AMS",
"LMS",
"UnifiController",
"DC02",
"DC01",
"UAT",
"SCP",
"PrintServer",
"SIEMWAZUH",
"ITS"
)


# =====================================================================
# START
# =====================================================================

Clear-Host

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host "          VMware ESXi SNAPSHOT AUTOMATION" -ForegroundColor White
Write-Host "                 + EXCEL REPORT" -ForegroundColor White
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

Write-Host "ESXi Host : $ESXiHost"
Write-Host "Username  : $ESXiUsername"
Write-Host "VM Count  : $($VMList.Count)"
Write-Host ""

Write-Host "VMs configured for snapshot:" -ForegroundColor Cyan

foreach ($VMName in $VMList) {
    Write-Host " - $VMName"
}

Write-Host ""


# =====================================================================
# DESKTOP LOCATION
# =====================================================================

$Desktop = [Environment]::GetFolderPath("Desktop")

if ([string]::IsNullOrWhiteSpace($Desktop)) {

    Write-Host "[ERROR] Could not determine Desktop location." -ForegroundColor Red
    Read-Host "Press Enter to close"
    return

}


# =====================================================================
# EXCEL MODULE
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host " Preparing Excel Reporting" -ForegroundColor Yellow
Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host ""

try {

    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {

        Write-Host "[INFO] ImportExcel module not found." -ForegroundColor Yellow
        Write-Host "[INFO] Installing ImportExcel module..." -ForegroundColor Yellow
        Write-Host ""

        if (-not (Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue)) {

            Install-PackageProvider `
                -Name NuGet `
                -MinimumVersion 2.8.5.201 `
                -Force |
                Out-Null
        }

        Install-Module `
            ImportExcel `
            -Scope CurrentUser `
            -Force `
            -AllowClobber `
            -ErrorAction Stop
    }

    Import-Module ImportExcel -Force -ErrorAction Stop

    Write-Host "[ OK ] Excel reporting module ready." -ForegroundColor Green
    Write-Host ""

}
catch {

    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host " EXCEL MODULE ERROR" -ForegroundColor Red
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host ""

    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""

    Read-Host "Press Enter to close"
    return
}


# =====================================================================
# PREPARE EXCEL FILE
# =====================================================================

$DateStamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"

$ExcelFile = Join-Path `
    $Desktop `
    "ESXi_${ESXiHost}_Snapshot_Report_$DateStamp.xlsx"

Write-Host "[ OK ] Desktop location : $Desktop" -ForegroundColor Green
Write-Host "[ OK ] Excel report     : $ExcelFile" -ForegroundColor Green
Write-Host ""


# =====================================================================
# DATA COLLECTION ARRAYS
# =====================================================================

$ESXiReportRows      = @()
$VMReportRows        = @()
$SnapshotReportRows  = @()


# =====================================================================
# PASSWORD
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host " ESXi Authentication" -ForegroundColor Yellow
Write-Host "==============================================================" -ForegroundColor Yellow
Write-Host ""

$SecurePassword = Read-Host "Enter ESXi password" -AsSecureString

$Credential = [System.Management.Automation.PSCredential]::new(
    $ESXiUsername,
    $SecurePassword
)

Write-Host ""
Write-Host "Password received securely." -ForegroundColor Green
Write-Host ""


# =====================================================================
# CONNECT TO ESXi
# =====================================================================

$VIServer = $null

try {

    Write-Host "Connecting to ESXi $ESXiHost ..." -ForegroundColor Cyan

    $VIServer = Connect-VIServer `
        -Server $ESXiHost `
        -Credential $Credential `
        -ErrorAction Stop

    Write-Host "Connected successfully." -ForegroundColor Green
    Write-Host ""

}
catch {

    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host " ESXi CONNECTION FAILED" -ForegroundColor Red
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host ""

    Write-Host "Error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red

    Write-Host ""
    Read-Host "Press Enter to close"

    return
}


# =====================================================================
# GET ESXi HOST INFORMATION
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Collecting ESXi Information" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

$ESXiVersion    = "Unavailable"
$ESXiBuild      = "Unavailable"
$ESXiFullVersion = "Unavailable"
$ESXiProduct    = "VMware ESXi"
$ESXiHostName   = $ESXiHost

try {

    $VMHost = Get-VMHost `
        -Server $VIServer `
        -ErrorAction Stop |
        Select-Object -First 1

    if ($VMHost) {

        $ESXiHostName = $VMHost.Name

        $ESXiVersion = $VMHost.Version
        $ESXiBuild   = $VMHost.Build

        if ($VMHost.ExtensionData.Config.Product) {

            $ProductInfo = $VMHost.ExtensionData.Config.Product

            if ($ProductInfo.Name) {
                $ESXiProduct = $ProductInfo.Name
            }

            if ($ProductInfo.FullName) {
                $ESXiFullVersion = $ProductInfo.FullName
            }
        }
    }

}
catch {

    Write-Host "[WARN] Could not collect complete ESXi information." -ForegroundColor Yellow
    Write-Host "[WARN] $($_.Exception.Message)" -ForegroundColor Yellow
}


Write-Host "ESXi Host Name : $ESXiHostName"
Write-Host "ESXi IP        : $ESXiHost"
Write-Host "ESXi Version   : $ESXiVersion"
Write-Host "ESXi Build     : $ESXiBuild"
Write-Host "ESXi Full Ver. : $ESXiFullVersion"
Write-Host "Product        : $ESXiProduct"
Write-Host ""


# =====================================================================
# PROCESS EACH VM
# =====================================================================

$SuccessfulSnapshots = 0
$FailedSnapshots     = 0
$VMsFound            = 0
$VMsNotFound         = 0

$CurrentVM = 0
$TotalVMs  = $VMList.Count


foreach ($VMName in $VMList) {

    $CurrentVM++

    $Percent = [math]::Round(
        ($CurrentVM / $TotalVMs) * 100
    )

    Write-Progress `
        -Activity "VMware Snapshot Process" `
        -Status "Processing $VMName [$CurrentVM of $TotalVMs]" `
        -PercentComplete $Percent


    # =================================================================
    # TIMESTAMP
    # =================================================================

    $SnapshotStartTime = Get-Date

    $TimeStamp = Get-Date -Format "yyyy-MM-dd-HH-mm-ss"

    $SnapshotName = "AUTO-$VMName-$TimeStamp"


    # =================================================================
    # PROCESSING HEADER
    # =================================================================

    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host " Processing VM [$CurrentVM/$TotalVMs]" -ForegroundColor Cyan
    Write-Host "==============================================================" -ForegroundColor Cyan
    Write-Host ""

    Write-Host "VM       : $VMName"
    Write-Host "Snapshot : $SnapshotName"
    Write-Host ""


    # =================================================================
    # FIND VM
    # =================================================================

    try {

        Write-Host "Finding VM..." -ForegroundColor Cyan

        $VM = Get-VM `
            -Name $VMName `
            -Server $VIServer `
            -ErrorAction Stop

        Write-Host "VM found: $($VM.Name)" -ForegroundColor Green

        $VMsFound++


    }
    catch {

        $VMsNotFound++
        $FailedSnapshots++

        $ErrorMessage = $_.Exception.Message

        Write-Host ""
        Write-Host "VM NOT FOUND" -ForegroundColor Red
        Write-Host "Error: $ErrorMessage" -ForegroundColor Red
        Write-Host ""


        # -------------------------------------------------------------
        # Add failed VM to VM report
        # -------------------------------------------------------------

        $VMReportRows += [PSCustomObject]@{

            "VM Name"          = $VMName
            "Power State"      = "Unavailable"
            "CPU"              = "Unavailable"
            "Memory GB"        = "Unavailable"
            "Guest OS"         = "Unavailable"
            "VM Version"       = "Unavailable"
            "Datastore"        = "Unavailable"
            "VM Host"          = $ESXiHostName
            "VM IP Address"    = "Unavailable"
            "Snapshot Result"  = "FAILED"
            "Error"            = $ErrorMessage
        }


        # -------------------------------------------------------------
        # Add failed snapshot to snapshot report
        # -------------------------------------------------------------

        $SnapshotReportRows += [PSCustomObject]@{

            "VM Name"             = $VMName
            "Snapshot Name"       = $SnapshotName
            "Snapshot Created"    = ""
            "Snapshot Size GB"    = ""
            "Memory Included"     = $false
            "Quiesced"            = $false
            "Description"         = "Automated snapshot"
            "Snapshot State"      = "FAILED"
            "Result"              = "FAILED"
            "Error"               = $ErrorMessage
        }

        continue
    }


    # =================================================================
    # COLLECT VM INFORMATION
    # =================================================================

    $PowerState = $VM.PowerState
    $NumCPU     = $VM.NumCpu
    $MemoryGB   = $VM.MemoryGB
    $GuestOS    = $VM.Guest.OSFullName

    if ([string]::IsNullOrWhiteSpace($GuestOS)) {
        $GuestOS = "Unavailable"
    }


    # -----------------------------------------------------------------
    # VM Version
    # -----------------------------------------------------------------

    $VMVersion = "Unavailable"

    try {

        if ($VM.ExtensionData.Config.Version) {
            $VMVersion = $VM.ExtensionData.Config.Version
        }

    }
    catch {
        $VMVersion = "Unavailable"
    }


    # -----------------------------------------------------------------
    # Datastore
    # -----------------------------------------------------------------

    $DatastoreName = "Unavailable"

    try {

        $DatastoreName = (
            Get-Datastore `
                -VM $VM `
                -Server $VIServer `
                -ErrorAction SilentlyContinue |
                Select-Object -ExpandProperty Name -Unique
        ) -join ", "

        if ([string]::IsNullOrWhiteSpace($DatastoreName)) {
            $DatastoreName = "Unavailable"
        }

    }
    catch {
        $DatastoreName = "Unavailable"
    }


    # -----------------------------------------------------------------
    # VM IP Address
    # -----------------------------------------------------------------

    $VMIPAddress = "Unavailable"

    try {

        $VMIPAddress = (
            $VM.Guest.IPAddress |
            Where-Object {
                $_ -match '^\d{1,3}(\.\d{1,3}){3}$'
            }
        ) -join ", "

        if ([string]::IsNullOrWhiteSpace($VMIPAddress)) {
            $VMIPAddress = "Unavailable"
        }

    }
    catch {
        $VMIPAddress = "Unavailable"
    }


    # =================================================================
    # CREATE VM REPORT ENTRY
    # =================================================================

    $VMReportRows += [PSCustomObject]@{

        "VM Name"          = $VM.Name
        "Power State"      = $PowerState
        "CPU"              = $NumCPU
        "Memory GB"        = $MemoryGB
        "Guest OS"         = $GuestOS
        "VM Version"       = $VMVersion
        "Datastore"        = $DatastoreName
        "VM Host"          = $ESXiHostName
        "VM IP Address"    = $VMIPAddress
        "Snapshot Result"  = "PENDING"
        "Error"            = ""
    }


    # =================================================================
    # CREATE SNAPSHOT
    # =================================================================

    Write-Host ""
    Write-Host "Creating snapshot..." -ForegroundColor Cyan

    try {

        $NewSnapshot = New-Snapshot `
            -VM $VM `
            -Name $SnapshotName `
            -Description "Automated snapshot created on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" `
            -Memory:$false `
            -Quiesce:$false `
            -ErrorAction Stop


        # =============================================================
        # SNAPSHOT SUCCESS
        # =============================================================

        $SuccessfulSnapshots++

        $SnapshotEndTime = Get-Date

        Write-Host ""
        Write-Host "SNAPSHOT CREATED SUCCESSFULLY" -ForegroundColor Green
        Write-Host ""
        Write-Host "VM       : $VMName"
        Write-Host "Snapshot : $SnapshotName"
        Write-Host "Time     : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-Host ""


        # -------------------------------------------------------------
        # Get complete snapshot details
        # -------------------------------------------------------------

        $SnapshotDetails = $null

        try {

            $SnapshotDetails = Get-Snapshot `
                -VM $VM `
                -Name $SnapshotName `
                -Server $VIServer `
                -ErrorAction Stop |
                Select-Object -First 1

        }
        catch {

            $SnapshotDetails = $NewSnapshot
        }


        # -------------------------------------------------------------
        # Snapshot properties
        # -------------------------------------------------------------

        $SnapshotCreated = ""

        if ($SnapshotDetails.Created) {
            $SnapshotCreated = $SnapshotDetails.Created
        }
        else {
            $SnapshotCreated = $SnapshotEndTime
        }


        $SnapshotSizeGB = ""

        if ($null -ne $SnapshotDetails.SizeGB) {

            $SnapshotSizeGB = [math]::Round(
                [double]$SnapshotDetails.SizeGB,
                3
            )
        }


        $MemoryIncluded = $false

        if ($null -ne $SnapshotDetails.Memory) {
            $MemoryIncluded = $SnapshotDetails.Memory
        }


        $Quiesced = $false

        if ($null -ne $SnapshotDetails.Quiesced) {
            $Quiesced = $SnapshotDetails.Quiesced
        }


        $Description = ""

        if ($SnapshotDetails.Description) {
            $Description = $SnapshotDetails.Description
        }
        else {
            $Description = "Automated snapshot"
        }


        # -------------------------------------------------------------
        # Add snapshot report row
        # -------------------------------------------------------------

        $SnapshotReportRows += [PSCustomObject]@{

            "VM Name"             = $VMName
            "Snapshot Name"       = $SnapshotName
            "Snapshot Created"    = $SnapshotCreated
            "Snapshot Size GB"    = $SnapshotSizeGB
            "Memory Included"     = $MemoryIncluded
            "Quiesced"            = $Quiesced
            "Description"         = $Description
            "Snapshot State"      = "Created"
            "Result"              = "SUCCESS"
            "Error"               = ""
        }


        # -------------------------------------------------------------
        # Update VM report result
        # -------------------------------------------------------------

        $VMReportRows[-1]."Snapshot Result" = "SUCCESS"


    }
    catch {

        # =============================================================
        # SNAPSHOT FAILED
        # =============================================================

        $FailedSnapshots++

        $SnapshotError = $_.Exception.Message

        Write-Host ""
        Write-Host "SNAPSHOT FAILED" -ForegroundColor Red
        Write-Host ""
        Write-Host "VM       : $VMName" -ForegroundColor Red
        Write-Host "Error    : $SnapshotError" -ForegroundColor Red
        Write-Host ""


        # -------------------------------------------------------------
        # Add failed snapshot report row
        # -------------------------------------------------------------

        $SnapshotReportRows += [PSCustomObject]@{

            "VM Name"             = $VMName
            "Snapshot Name"       = $SnapshotName
            "Snapshot Created"    = ""
            "Snapshot Size GB"    = ""
            "Memory Included"     = $false
            "Quiesced"            = $false
            "Description"         = "Automated snapshot"
            "Snapshot State"      = "FAILED"
            "Result"              = "FAILED"
            "Error"               = $SnapshotError
        }


        # -------------------------------------------------------------
        # Update VM report result
        # -------------------------------------------------------------

        $VMReportRows[-1]."Snapshot Result" = "FAILED"
        $VMReportRows[-1]."Error" = $SnapshotError

        continue
    }
}


Write-Progress `
    -Activity "VMware Snapshot Process" `
    -Completed


# =====================================================================
# ESXi SUMMARY REPORT
# =====================================================================

$ReportDate = Get-Date

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi Host Name"
    "Value"    = $ESXiHostName
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi IP Address"
    "Value"    = $ESXiHost
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi Product"
    "Value"    = $ESXiProduct
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi Version"
    "Value"    = $ESXiVersion
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi Build"
    "Value"    = $ESXiBuild
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "ESXi Full Version"
    "Value"    = $ESXiFullVersion
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "Configured VM Count"
    "Value"    = $VMList.Count
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "VMs Found"
    "Value"    = $VMsFound
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "VMs Not Found"
    "Value"    = $VMsNotFound
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "Successful Snapshots"
    "Value"    = $SuccessfulSnapshots
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "Failed Snapshots"
    "Value"    = $FailedSnapshots
}

$ESXiReportRows += [PSCustomObject]@{
    "Property" = "Report Generated"
    "Value"    = $ReportDate.ToString("yyyy-MM-dd HH:mm:ss")
}


# =====================================================================
# DISCONNECT FROM ESXi
# =====================================================================

if ($VIServer) {

    Write-Host ""
    Write-Host "Disconnecting from ESXi..." -ForegroundColor Cyan

    try {

        Disconnect-VIServer `
            -Server $VIServer `
            -Confirm:$false `
            -ErrorAction SilentlyContinue |
            Out-Null

        Write-Host "Disconnected successfully." -ForegroundColor Green

    }
    catch {

        Write-Host "Disconnect completed with a warning." -ForegroundColor Yellow
    }
}


# =====================================================================
# GENERATE EXCEL REPORT
# =====================================================================

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host " Generating Excel Report" -ForegroundColor Cyan
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

try {

    # -----------------------------------------------------------------
    # Remove existing file if somehow present
    # -----------------------------------------------------------------

    if (Test-Path $ExcelFile) {

        Remove-Item `
            $ExcelFile `
            -Force `
            -ErrorAction SilentlyContinue
    }


    # -----------------------------------------------------------------
    # SHEET 1 - ESXi SUMMARY
    # -----------------------------------------------------------------

    $ESXiReportRows |
        Export-Excel `
            -Path $ExcelFile `
            -WorksheetName "ESXi Summary" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow


    # -----------------------------------------------------------------
    # SHEET 2 - VM DETAILS
    # -----------------------------------------------------------------

    $VMReportRows |
        Export-Excel `
            -Path $ExcelFile `
            -WorksheetName "VM Details" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow


    # -----------------------------------------------------------------
    # SHEET 3 - SNAPSHOT DETAILS
    # -----------------------------------------------------------------

    $SnapshotReportRows |
        Export-Excel `
            -Path $ExcelFile `
            -WorksheetName "Snapshot Details" `
            -AutoSize `
            -AutoFilter `
            -FreezeTopRow `
            -BoldTopRow


    # -----------------------------------------------------------------
    # Add Excel title / formatting
    # -----------------------------------------------------------------

    $ExcelPackage = Open-ExcelPackage -Path $ExcelFile


    # ================================================================
    # ESXi SUMMARY FORMATTING
    # ================================================================

    $Worksheet = $ExcelPackage.Workbook.Worksheets["ESXi Summary"]

    if ($Worksheet) {

        $Worksheet.Cells["A1:B1"].Style.Font.Bold = $true
        $Worksheet.Cells["A1:B1"].Style.HorizontalAlignment = "Center"

        $Worksheet.Column(1).Width = 30
        $Worksheet.Column(2).Width = 60
    }


    # ================================================================
    # VM DETAILS FORMATTING
    # ================================================================

    $Worksheet = $ExcelPackage.Workbook.Worksheets["VM Details"]

    if ($Worksheet) {

        $Worksheet.Cells["A1:K1"].Style.Font.Bold = $true
        $Worksheet.Cells["A1:K1"].Style.HorizontalAlignment = "Center"

        $Worksheet.Column(1).Width  = 35
        $Worksheet.Column(2).Width  = 15
        $Worksheet.Column(3).Width  = 10
        $Worksheet.Column(4).Width  = 12
        $Worksheet.Column(5).Width  = 35
        $Worksheet.Column(6).Width  = 15
        $Worksheet.Column(7).Width  = 30
        $Worksheet.Column(8).Width  = 25
        $Worksheet.Column(9).Width  = 25
        $Worksheet.Column(10).Width = 18
        $Worksheet.Column(11).Width = 60
    }


    # ================================================================
    # SNAPSHOT DETAILS FORMATTING
    # ================================================================

    $Worksheet = $ExcelPackage.Workbook.Worksheets["Snapshot Details"]

    if ($Worksheet) {

        $Worksheet.Cells["A1:J1"].Style.Font.Bold = $true
        $Worksheet.Cells["A1:J1"].Style.HorizontalAlignment = "Center"

        $Worksheet.Column(1).Width  = 35
        $Worksheet.Column(2).Width  = 55
        $Worksheet.Column(3).Width  = 25
        $Worksheet.Column(4).Width  = 18
        $Worksheet.Column(5).Width  = 18
        $Worksheet.Column(6).Width  = 15
        $Worksheet.Column(7).Width  = 50
        $Worksheet.Column(8).Width  = 20
        $Worksheet.Column(9).Width  = 15
        $Worksheet.Column(10).Width = 60
    }


    # -----------------------------------------------------------------
    # Save Excel package
    # -----------------------------------------------------------------

    Close-ExcelPackage $ExcelPackage


    Write-Host ""
    Write-Host "[ OK ] Excel report generated successfully." -ForegroundColor Green
    Write-Host ""
    Write-Host "Report file : $(Split-Path $ExcelFile -Leaf)"
    Write-Host "Location    : $ExcelFile"
    Write-Host ""


}
catch {

    Write-Host ""
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host " EXCEL REPORT GENERATION FAILED" -ForegroundColor Red
    Write-Host "==============================================================" -ForegroundColor Red
    Write-Host ""

    Write-Host "Error:" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    Write-Host ""

    Read-Host "Press Enter to close"

    return
}


# =====================================================================
# FINAL STATUS
# =====================================================================

Write-Host "==============================================================" -ForegroundColor Green
Write-Host "              SNAPSHOT PROCESS COMPLETED" -ForegroundColor Green
Write-Host "==============================================================" -ForegroundColor Green
Write-Host ""

if ($FailedSnapshots -eq 0) {

    Write-Host "Status              : SUCCESS" -ForegroundColor Green

}
else {

    Write-Host "Status              : COMPLETED WITH ERRORS" -ForegroundColor Yellow
}

Write-Host "ESXi Host           : $ESXiHost"
Write-Host "ESXi Version        : $ESXiVersion"
Write-Host "ESXi Build          : $ESXiBuild"
Write-Host "Configured VMs      : $($VMList.Count)"
Write-Host "VMs Found           : $VMsFound"
Write-Host "VMs Not Found       : $VMsNotFound"
Write-Host "Snapshots Success   : $SuccessfulSnapshots"
Write-Host "Snapshots Failed    : $FailedSnapshots"
Write-Host "Excel Report        : $(Split-Path $ExcelFile -Leaf)"
Write-Host "Report Location     : $ExcelFile"
Write-Host "Completed           : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Host ""

Write-Host "[ OK ] Snapshot results exported to Excel." -ForegroundColor Green
Write-Host ""


# =====================================================================
# OPEN EXCEL REPORT AUTOMATICALLY
# =====================================================================

if (Test-Path $ExcelFile) {

    Write-Host "Opening Excel report..." -ForegroundColor Cyan

    Start-Sleep -Seconds 2

    try {

        Invoke-Item $ExcelFile

        Write-Host "[ OK ] Excel report opened." -ForegroundColor Green

    }
    catch {

        Write-Host "[WARN] Excel file was created but could not be opened automatically." -ForegroundColor Yellow
        Write-Host "      You can manually open:" -ForegroundColor Yellow
        Write-Host "      $ExcelFile" -ForegroundColor Yellow
    }

}
else {

    Write-Host "[WARN] Excel report could not be found." -ForegroundColor Yellow
}


# =====================================================================
# END
# =====================================================================

Write-Host ""
Write-Host "==============================================================" -ForegroundColor Cyan
Write-Host ""

Read-Host "Press Enter to close"
```

---

# 28. 📌 Quick Reference

| Item                      | Value / Behavior                     |
| ------------------------- | ------------------------------------ |
| ESXi Host                 | Configured through `$ESXiHost`       |
| ESXi Username             | Configured through `$ESXiUsername`   |
| ESXi Password             | Entered interactively                |
| VM Selection              | `$VMList`                            |
| Snapshot Prefix           | `AUTO-`                              |
| Snapshot Memory           | Disabled                             |
| Snapshot Quiescing        | Disabled                             |
| Excel Module              | `ImportExcel`                        |
| Report Format             | `.xlsx`                              |
| Report Location           | Current user's Desktop               |
| Excel Sheets              | 3                                    |
| VM Failure Handling       | Continue processing                  |
| Snapshot Failure Handling | Continue processing                  |
| Final Status              | `SUCCESS` or `COMPLETED WITH ERRORS` |
| Excel Opening             | Automatic                            |

---

## 📝 Document Maintenance

When modifying this automation:

* Update the backend configuration when ESXi hosts or VM inventory changes.
* Test changes against a non-production VM first.
* Validate the generated Excel report after script modifications.
* Record significant changes in Git commit messages.
* Never commit credentials or sensitive authentication information.
* Keep this documentation synchronized with the deployed script version.
