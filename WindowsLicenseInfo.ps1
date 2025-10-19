# ==============================
# Windows License Info Extractor
# ==============================

# Output file path
$output = "$env:USERPROFILE\Desktop\Windows_License_Info.txt"

# Function to decode DigitalProductId
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

# Collect data
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

# Detect activation type
$licenseType = if ($license.KeyManagementServiceProductKeyID) {
    "Volume / KMS Activation"
} elseif ($license.OA3xOriginalProductKey) {
    "OEM / BIOS Embedded License"
} else {
    "Retail / Digital Activation"
}

# Compose report
$info = @"
===============================
 Windows License Information
===============================

Machine Name : $env:COMPUTERNAME
User Name    : $env:USERNAME
Windows SKU  : $($license.Description)
License Type : $licenseType
Product Key  : $productKey

Report generated: $(Get-Date)
"@

# Save and show
$info | Out-File -Encoding UTF8 -FilePath $output
Write-Host "License information saved to:`n$output" -ForegroundColor Green
Start-Process notepad.exe $output
