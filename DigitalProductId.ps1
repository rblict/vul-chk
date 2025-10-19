function ConvertTo-ProductKey {
    param([byte[]]$DigitalProductId)

    $chars = "BCDFGHJKMPQRTVWXY2346789"
    $key = ""

    # Windows 8+ tweak detection & fix
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

# Read DigitalProductId from registry and convert
$dpid = (Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -Name 'DigitalProductId').DigitalProductId
ConvertTo-ProductKey -DigitalProductId $dpid
