# Get system information
$computerSystem = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS

# Create an object with the desired properties
$deviceInfo = [PSCustomObject]@{
    Manufacturer = $computerSystem.Manufacturer
    Model        = $computerSystem.Model
    SerialNumber = $bios.SerialNumber
}

# Export to CSV
$csvPath = "$PSScriptRoot\DeviceInfo.csv"
$deviceInfo | Export-Csv -Path $csvPath -NoTypeInformation

Write-Output "Device information exported to $csvPath"
