# Find all PowerShell processes running clipson.ps1
$targets = Get-CimInstance Win32_Process |
Where-Object {
    $_.Name -match '^(powershell|pwsh)\.exe$' -and
    $_.CommandLine -match 'clipson\.ps1'
}

if ($targets) {
    foreach ($proc in $targets) {
        Write-Host "Killing PID $($proc.ProcessId): $($proc.CommandLine)"
        Stop-Process -Id $proc.ProcessId -Force
    }
} else {
    Write-Host "No running instances of clipson.ps1 found."
}
