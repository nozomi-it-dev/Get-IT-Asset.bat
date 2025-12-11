# Asset Inventory Report - PowerShell Version
# For modern Windows systems without WMIC

$Host.UI.RawUI.WindowTitle = "Asset Inventory Report"
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8

# Get current date/time
$now = Get-Date
$date_str = $now.ToString("yyyy/MM/dd")
$time_str = $now.ToString("HH:mm:ss.ff")
$filename_date = $now.ToString("yyyy-MM-dd")

# Automatically find "SANDISK" Flash Drive
$output_path = $null
$drives = Get-Volume | Where-Object { $_.DriveType -eq 'Removable' -and $_.FileSystemLabel -like "*SANDISK*" }

if ($drives) {
    $output_path = $drives[0].DriveLetter + ":\"
} else {
    $output_path = [Environment]::GetFolderPath("Desktop") + "\"
    Write-Host "Warning: SANDISK drive not found. Saving to Desktop instead."
    Start-Sleep -Seconds 3
}

$filename = "$env:COMPUTERNAME" + "_" + $filename_date + ".txt"
$fullpath = Join-Path $output_path $filename

# Initialize report file
@"
====================================================
--- IT DEVICE INVENTORY DATA ---
Date/Time: $date_str  $time_str
====================================================

"@ | Out-File -FilePath $fullpath -Encoding UTF8

# [1] Brand
$csProduct = Get-CimInstance -ClassName Win32_ComputerSystemProduct
"[1] Brand: $($csProduct.Vendor)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [2] Model
"[2] Model: $($csProduct.Name)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [3] Spec
"[3] Spec:" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# CPU
$cpu = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
"    CPU: $($cpu.Name)" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# RAM
$cs = Get-CimInstance -ClassName Win32_ComputerSystem
$ramGB = [math]::Round($cs.TotalPhysicalMemory / 1GB)
"    RAM: $ramGB GB" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# HDD/SSD (Filter out USB drives)
$disks = Get-CimInstance -ClassName Win32_DiskDrive | Where-Object { $_.InterfaceType -notlike "*USB*" }
foreach ($disk in $disks) {
    $sizeGB = [math]::Round($disk.Size / 1GB)
    "    HDD/SSD: $($disk.Caption) ($sizeGB GB)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
}
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [4] Operating System
$os = Get-CimInstance -ClassName Win32_OperatingSystem
"[4] Operating System: $($os.Caption) $($os.OSArchitecture)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [5] Device ID (UUID)
"[5] Device ID: $($csProduct.UUID)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [6] Product ID
$productId = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion" -Name ProductId -ErrorAction SilentlyContinue).ProductId
"[6] Product ID: $productId" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [7] Serial No
$bios = Get-CimInstance -ClassName Win32_BIOS
"[7] Serial No: $($bios.SerialNumber)" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [8] Device Name
"[8] Device Name: $env:COMPUTERNAME" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [9] User Email
"[9] User Email: $env:USERNAME" | Out-File -FilePath $fullpath -Append -Encoding UTF8
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# [10] Microsoft Office Information
"[10] Microsoft Office:" | Out-File -FilePath $fullpath -Append -Encoding UTF8
$office_found = $false

# Check Office from Registry (Office 2016, 2019, 2021, 365)
$officeVersions = @("16.0", "15.0", "14.0")
foreach ($ver in $officeVersions) {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Registration"
    if (Test-Path $regPath) {
        $subKeys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
        foreach ($key in $subKeys) {
            $productName = (Get-ItemProperty -Path $key.PSPath -Name ProductName -ErrorAction SilentlyContinue).ProductName
            if ($productName) {
                "    Version: $productName" | Out-File -FilePath $fullpath -Append -Encoding UTF8
                $office_found = $true
                break
            }
        }
        if ($office_found) { break }
    }
}

# Check Product ID
foreach ($ver in $officeVersions) {
    $regPath = "HKLM:\SOFTWARE\Microsoft\Office\$ver\Registration"
    if (Test-Path $regPath) {
        $subKeys = Get-ChildItem -Path $regPath -ErrorAction SilentlyContinue
        foreach ($key in $subKeys) {
            $productID = (Get-ItemProperty -Path $key.PSPath -Name ProductID -ErrorAction SilentlyContinue).ProductID
            if ($productID) {
                "    Product ID: $productID" | Out-File -FilePath $fullpath -Append -Encoding UTF8
                break
            }
        }
    }
}

# Check License Status
try {
    $officeLicense = Get-CimInstance -Query "SELECT * FROM SoftwareLicensingProduct WHERE ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND LicenseStatus=1" -ErrorAction Stop |
        Where-Object { $_.Name -like "*Office*" } |
        Select-Object -First 1
    
    if ($officeLicense) {
        "    License Status: Licensed" | Out-File -FilePath $fullpath -Append -Encoding UTF8
    }
} catch {
    # Silently continue if license check fails
}

# Check Click-to-Run (Office 365)
$c2rPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"
if (Test-Path $c2rPath) {
    $edition = (Get-ItemProperty -Path $c2rPath -Name ProductReleaseIds -ErrorAction SilentlyContinue).ProductReleaseIds
    if ($edition) {
        "    Edition: $edition" | Out-File -FilePath $fullpath -Append -Encoding UTF8
        $office_found = $true
    }
}

if (-not $office_found) {
    "    Status: Not installed or not detected" | Out-File -FilePath $fullpath -Append -Encoding UTF8
}
"" | Out-File -FilePath $fullpath -Append -Encoding UTF8

# Check Panda Dome Status
$panda_status = "Not Installed"
$pandaServices = @("PandaAetherAgent", "NanoServiceMain")
foreach ($service in $pandaServices) {
    $svc = Get-Service -Name $service -ErrorAction SilentlyContinue
    if ($svc -and $svc.Status -eq 'Running') {
        $panda_status = "Installed and Running"
        break
    } elseif ($svc) {
        $panda_status = "Installed"
    }
}

if ($panda_status -eq "Not Installed") {
    if (Test-Path "HKLM:\SOFTWARE\Panda Security") {
        $panda_status = "Installed"
    }
}

# Display success message
Write-Host ""
Write-Host "===================================================="
Write-Host "SUCCESS! Report saved to:"
Write-Host $fullpath
Write-Host ""
Write-Host "Panda Dome Status: $panda_status"
Write-Host "===================================================="
Write-Host ""
Write-Host "Press any key to close this window..."
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")