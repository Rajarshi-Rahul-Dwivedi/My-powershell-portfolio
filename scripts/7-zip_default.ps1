$7zInstallationFolder = 'C:\Program Files\7-Zip'
$reg = [Microsoft.Win32.RegistryKey]::OpenBaseKey([Microsoft.Win32.RegistryHive]::ClassesRoot, [Microsoft.Win32.RegistryView]::Default)
$subKeys = $reg.GetSubKeyNames() | where { $_ -match '7-Zip.' }
foreach ($keyName in $subKeys) {
    $key = $reg.OpenSubKey($keyName + '\shell\open\command', $true)
    $key.SetValue('', '"' + $7zInstallationFolder + '\7zG.exe" x "%1" -o*')
}