$keyPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$propertyName = "HideNewOutlookToggle"

# Check if the key exists
if (Test-Path $keyPath) {
    # Check if the property exists and remove it if it does
    if ((Get-ItemProperty -Path $keyPath).$propertyName -ne $null) {
        Remove-ItemProperty -Path $keyPath -Name $propertyName
        Write-Host "Preview Enabled"
    } else {
        Write-Host "Property does not exist"
    }
} else {
    Write-Host "Key does not exist"
}
