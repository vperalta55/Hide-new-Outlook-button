
$keyPath = "HKCU:\Software\Microsoft\Office\16.0\Outlook\Options\General"
$propertyName = "HideNewOutlookToggle"
 
# Check if the key exists and create it if not
if (-not (Test-Path $keyPath)) {
    New-Item -Path $keyPath -ItemType RegistryKey
}
 
# Check if the property exists and create it if not
if ((Get-ItemProperty -Path $keyPath).$propertyName -eq $null){
    New-ItemProperty -Path $keyPath -Name $propertyName -PropertyType Dword -Value 1
    Write-Host "Preview Disabled"
}
