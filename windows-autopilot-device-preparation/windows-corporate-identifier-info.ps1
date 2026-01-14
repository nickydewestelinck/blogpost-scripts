# Get system information
$computerSystem = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS

# Create an object with the desired properties
$deviceInfo = [PSCustomObject]@{
    Manufacturer = $computerSystem.Manufacturer
    Model        = $computerSystem.Model
    SerialNumber = $bios.SerialNumber
}

# Define output path
$csvPath = "$PSScriptRoot\DeviceInfo.csv"

# Export to CSV without headers
$deviceInfo |
    ConvertTo-Csv -NoTypeInformation |
    Select-Object -Skip 1 |
    Set-Content -Path $csvPath

# Green success message
Write-Host "Device information exported (without headers) to $csvPath" -ForegroundColor Green
