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

