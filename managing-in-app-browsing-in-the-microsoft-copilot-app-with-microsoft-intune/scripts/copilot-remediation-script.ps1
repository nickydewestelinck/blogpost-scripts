$path = 'HKLM:\SOFTWARE\Policies\Microsoft\Copilot'
$name = 'BrowsingEnabled'

try {
    if (-not (Test-Path $path)) {
        New-Item -Path $path -Force | Out-Null
    }
    New-ItemProperty -Path $path -Name $name -Value 0 -PropertyType DWord -Force | Out-Null
    Write-Output "Set $name to 0"
    exit 0
}
catch {
    Write-Output "Failed: $($_.Exception.Message)"
    exit 1
}
