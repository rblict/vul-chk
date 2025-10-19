# ==========================================
# Windows License + Activation Status Report
# ==========================================

# Output file path
$output = "$env:USERPROFILE\Desktop\Windows_License_Report.txt"

# --- Function: Decode DigitalProductId to readable key ---
function ConvertTo-ProductKey {
    param([byte[]]$DigitalProductId)
    $chars = "BCDFGHJKMPQRTVWXY2346789"
    $key = ""
    $isWin8 = ($DigitalProductId[66] -shr 3) -band 1
    if ($isWin8 -eq 1) {
        $DigitalProductId[66] = ($DigitalProductId[66] -band 0xF7) -bor (($DigitalProductId[66] -band 0x08) -shl 1)
    }
    for ($i = 24; $i -ge 0; $i--) {
        $k = 0
        for ($j = 14; $j -ge 0; $j--) {
            $k = $k * 256 -bxor $DigitalProductId[$j + 52]
            $DigitalProductId[$j + 52] = [math]::Floor($k / 24)
            $k = $k % 24
        }
        $key = $chars[$k] + $key
        if (($i % 5) -eq 0 -and $i -ne 0) { $key = "-" + $key }
    }
    return $key
}

# --- Collect core licensing data ---
$license = Get-CimInstance -ClassName SoftwareLicensingService
$productKey = $license.OA3xOriginalProductKey
if (-not $productKey) {
    try {
        $dpid = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DigitalProductId').DigitalProductId
        $productKey = ConvertTo-ProductKey -DigitalProductId $dpid
    } catch {
        $productKey = "No visible key (Digital License or KMS)"
    }
}

# --- Detect license type ---
$licenseType = if ($license.KeyManagementServiceProductKeyID) {
    "Volume / KMS Activation"
} elseif ($license.OA3xOriginalProductKey) {
    "OEM / BIOS Embedded License"
} else {
    "Retail / Digital Activation"
}

# --- Detect activation status ---
try {
    $status = (Get-CimInstance -ClassName SoftwareLicensingProduct | 
               Where-Object { $_.PartialProductKey } |
               Select-Object -First 1 -ExpandProperty LicenseStatus)
    switch ($status) {
        0 { $activationStatus = "Unlicensed" }
        1 { $activationStatus = "Licensed / Activated" }
        2 { $activationStatus = "Out-of-Box Grace Period" }
        3 { $activationStatus = "Out-of-Tolerance Grace Period" }
        4 { $activationStatus = "Non-Genuine Grace Period" }
        5 { $activationStatus = "Notification Mode (Expired)" }
        Default { $activationStatus = "Unknown" }
    }
} catch {
    $activationStatus = "Unable to determine activation state"
}

# --- Optional: Get activation channel ---
try {
    $channel = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").ProductId
} catch {
    $channel = "Unavailable"
}

# --- Compose final report ---
$info = @"
=====================================
   Windows License & Activation Report
=====================================

Machine Name      : $env:COMPUTERNAME
User Name         : $env:USERNAME
Windows Edition   : $($license.Description)
License Type      : $licenseType
Product Key       : $productKey
Activation Status : $activationStatus
Product ID        : $channel

Report generated  : $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
=====================================
"@

# --- Save and open report ---
$info | Out-File -Encoding UTF8 -FilePath $output
Write-Host "`nReport saved to:`n$output" -ForegroundColor Green
Start-Process notepad.exe $output
