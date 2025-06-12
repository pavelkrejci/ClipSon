Add-Type -AssemblyName System.Windows.Forms

# Create the monitor from the class defined in mod-clipboard.ps1
$monitor = New-Object ClipboardMonitor
$handler = {
    # Check for text
    if ([System.Windows.Forms.Clipboard]::ContainsText()) {
        $text = [System.Windows.Forms.Clipboard]::GetText()
        if ($text.Trim() -and $text -ne $script:lastClipboardContent) {
            Write-Host "Clipboard:" -ForegroundColor Green
            Write-Host $text -ForegroundColor White
            $script:lastClipboardContent = $text
        }
    }
    # Check for image
    elseif ([System.Windows.Forms.Clipboard]::ContainsImage()) {
        $image = [System.Windows.Forms.Clipboard]::GetImage()
        if ($image -ne $null) {
            Write-Host "Clipboard: [Image]" -ForegroundColor Yellow
            Write-Host ("Image size: {0}x{1}" -f $image.Width, $image.Height) -ForegroundColor Gray
        }
    }
    # Check for file drop list
    elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
        $files = [System.Windows.Forms.Clipboard]::GetFileDropList()
        if ($files.Count -gt 0) {
            Write-Host "Clipboard: [Files]" -ForegroundColor Cyan
            foreach ($file in $files) {
                Write-Host "  $file" -ForegroundColor White
            }
        }
    }
}

######################################################################################
# MAIN
######################################################################################
$script:lastClipboardContent = ""
$exitLoop = $false

# Replace Import-Module with dot-sourcing
. "$PSScriptRoot/mod-misc.ps1"
# Load configuration from mod-misc
try {
    $global:Config = Get-Configuration
    Write-DebugMsg "Configuration loaded successfully" -ForegroundColor Green    
}
catch {
    Write-Error "Failed to load configuration: $_"
}
. "$PSScriptRoot/mod-clipboard.ps1"
. "$PSScriptRoot/mod-nextcloud.ps1"

# --- Open connection to Nextcloud ---
$global:webdavConnection = New-NextcloudConnection `
    -ServerUrl $global:Config.nextcloud.server_url `
    -Username $global:Config.nextcloud.username `
    -Password $global:Config.nextcloud.password `
    -ProxyUrl $global:Config.proxy.url `
    -ProxyUsername $global:Config.proxy.username `
    -ProxyPassword $global:Config.proxy.password `
    -UseSystemProxy $global:Config.proxy.use_system_proxy

# Start the clipboard monitor
Register-ObjectEvent -InputObject $monitor -EventName ClipboardChanged -Action $handler
$monitor.CreateControl()
$monitor.Show()  # <-- Ensure the hidden window is created and message loop is active

Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { $exitLoop = $true }

Write-Host "ClipSon clipboard monitor started. Press Ctrl+C to stop." -ForegroundColor Cyan




# Create a timer for remote file checking
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000  # Check every 5 seconds
$timer.Add_Tick({
    try {
        Check-AllRemoteFilesForUpdates -Connection $webdavConnection
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Error during remote check: $($_.Exception.Message)"
    }
})
$timer.Start()



while (-not $exitLoop) {
    [System.Windows.Forms.Application]::DoEvents()
    Start-Sleep -Milliseconds 100
}

$monitor.Dispose()
Write-Host "ClipSon stopped." -ForegroundColor Cyan
