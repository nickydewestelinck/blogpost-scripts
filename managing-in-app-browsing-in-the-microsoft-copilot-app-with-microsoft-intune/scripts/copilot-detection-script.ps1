$path  = 'HKLM:\SOFTWARE\Policies\Microsoft\Copilot'
$name  = 'BrowsingEnabled'

try {
    $value = (Get-ItemProperty -Path $path -Name $name -ErrorAction Stop).$name
    if ($value -eq 0) {
        Write-Output "Compliant: $name = 0"
        exit 0
    }
    Write-Output "Not compliant: $name = $value"
    exit 1
}
catch {
    Write-Output "Not compliant: $name not set"
    exit 1
}
