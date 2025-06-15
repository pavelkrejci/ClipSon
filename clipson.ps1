# ClipSon - Advanced Clipboard Synchronization Tool (All-in-One)
# If you get execution policy errors, run one of these commands first:
# Option 1 (Recommended): Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Option 2 (Temporary): powershell.exe -ExecutionPolicy Bypass -File "clipson-all.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

# Define Write-DebugMsg at the very beginning to ensure it's available
function global:Write-DebugMsg {
    param([string]$Message)
    # Safe default if Config isn't loaded yet
    if ($global:Config -and $global:Config.app.debug_enabled) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - DEBUG: $Message"
    }
    elseif ($global:EnableDebugMessages) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - DEBUG: $Message"
    }
}

# First, load modules in the correct order
. "$PSScriptRoot/mod-misc.ps1"
# Load configuration
try {
    $global:Config = Get-Configuration
    Write-Host "Configuration loaded successfully" -ForegroundColor Green    
}
catch {
    Write-Error "Failed to load configuration: $_"
    exit 1
}

# Then load clipboard module (which uses functions from misc)
. "$PSScriptRoot/mod-clipboard.ps1"
# Finally load nextcloud module (which may depend on both)
. "$PSScriptRoot/mod-nextcloud.ps1"

# Initialize globals
$global:EnableDebugMessages = $global:Config.app.debug_enabled
$global:skipClipboardCheckUntil = [DateTime]::MinValue
$global:exitLoop = $false

# Output directory - make it global so it's accessible in event handlers
$global:outputDir = ".\captures-$($env:COMPUTERNAME)"
$global:maxEntries = 100  # Add maximum entries configuration

if (!(Test-Path $global:outputDir)) {
    New-Item -ItemType Directory -Path $global:outputDir | Out-Null
}

# Get password if needed
Get-PasswordIfNeeded -NextcloudConfig $global:Config.nextcloud

# Connect to Nextcloud
Write-Host "Connecting to Nextcloud..." -ForegroundColor Green
$global:webdavConnection = New-NextcloudConnection `
    -ServerUrl $global:Config.nextcloud.server_url `
    -Username $global:Config.nextcloud.username `
    -Password $global:Config.nextcloud.password `
    -ProxyUrl $global:Config.proxy.url `
    -ProxyUsername $global:Config.proxy.username `
    -ProxyPassword $global:Config.proxy.password `
    -UseSystemProxy $global:Config.proxy.use_system_proxy

if (-not (Test-NextcloudConnection -Connection $global:webdavConnection)) {
    Write-Error "Failed to connect to Nextcloud. Please check your configuration."
    exit 1
}

# Local file paths - make them global so modules can access them
$hostname = $env:COMPUTERNAME
$global:localSyncFile = ".\clipboard-$hostname.json"
$global:localSyncFileGz = ".\clipboard-$hostname.json.gz"
$global:localUploadPath = $global:Config.nextcloud.remote_folder + "clipboard-$hostname.json.gz"

# Global variables
$global:fileCounter = 0
$global:lastClipboardFileModified = $null
$global:lastRemoteCheck = [DateTime]::MinValue
$global:remotePeerTimestamps = @{}  # Track timestamps for each remote peer
$remoteCheckInterval = [TimeSpan]::FromSeconds($global:Config.app.remote_check_interval_seconds)

# Create the clipboard monitor - uses Win32 API to monitor clipboard changes
$monitor = New-Object ClipboardMonitor
$handler = {
    try {
        # Check if we should skip clipboard monitoring (after setting remote content)
        $currentTime = [DateTime]::Now
        if ($currentTime -lt $global:skipClipboardCheckUntil) {
            Write-DebugMsg "Skipping clipboard check due to recent remote update"
            return
        }
        
        Write-DebugMsg "Clipboard content changed at $(Get-Date -Format 'HH:mm:ss.fff')"
        
        # Check if clipboard has image content first (highest priority)
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $image = [System.Windows.Forms.Clipboard]::GetImage()
            if ($image -ne $null) {
                # Save image to memory stream to calculate hash without creating file
                $memoryStream = New-Object System.IO.MemoryStream
                $image.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)
                $imageBytes = $memoryStream.ToArray()
                $memoryStream.Dispose()
                
                Write-DebugMsg "Processing clipboard image"
                
                # Proceed with saving and uploading
                $fileNumber = Get-NextFileNumber
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                $filename = "$outputDir\clipboard_image_$paddedNumber.png"
                
                Write-DebugMsg "Creating image file: $filename"
                
                # Save the image bytes to file
                [System.IO.File]::WriteAllBytes($filename, $imageBytes)
                Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - Image saved: $filename"
                
                # Create unified JSON format for image
                $imageB64 = [System.Convert]::ToBase64String($imageBytes)
                $uploadContent = @{
                    type = "CLIPBOARD_IMAGE"
                    data = $imageB64
                    format = "png"
                    size = $imageBytes.Length
                }
                $uploadJson = ConvertTo-Json $uploadContent -Depth 10
                
                # Upload to WebDAV
                Upload-ToWebDAV -Content $uploadJson -Connection $global:webdavConnection -LocalSyncFile $localSyncFile -LocalSyncFileGz $localSyncFileGz -RemoteFilePath $localUploadPath
                
                # Show notification for image capture
                Show-ClipboardNotification -Title "ClipSon" -Message "Image captured: $filename" -Icon "Info"
                
                $image.Dispose()
            }
        }
        # Check if clipboard has rich text formats (medium priority)
        elseif (Test-ClipboardRichText) {
            $currentFormats = Get-ClipboardFormats
            $currentFormatData = @{}
            
            # Get all available format content
            if ($currentFormats -contains "HTML") {
                try {
                    $htmlContent = [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Html)
                    if ($htmlContent) {
                        $currentFormatData["text/html"] = $htmlContent
                    }
                }
                catch { }
            }
            
            # TODO: text/rtf support - currently commented out due to issues with RTF handling in Linux
            # if ($currentFormats -contains "RTF") {
            #     try {
            #         $rtfContent = [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Rtf)
            #         if ($rtfContent) {
            #             $currentFormatData["text/rtf"] = $rtfContent
            #         }
            #     }
            #     catch { }
            # }
            
            if ($currentFormats -contains "Text" -or $currentFormats -contains "Unicode") {
                try {
                    $textContent = [System.Windows.Forms.Clipboard]::GetText()
                    if ($textContent) {
                        $currentFormatData["text/plain"] = $textContent
                    }
                }
                catch { }
            }
            
            if ($currentFormatData.Count -gt 0) {
                Write-DebugMsg "Rich content changed, saving..."
                Save-ClipboardRichContentJson -FormatData $currentFormatData
            }
        }
        # Check if clipboard has text content (lowest priority)
        elseif ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $currentContent = [System.Windows.Forms.Clipboard]::GetText()
            
            # Process if content is not empty
            if ($currentContent.Trim() -ne "") {
                
                $fileNumber = Get-NextFileNumber
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                $filename = Join-Path $outputDir "clipboard_text_$paddedNumber.txt"
                
                # Verify the filename is correct before writing
                if ($filename -notmatch "clipboard_text_\d{3}\.txt$") {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss.fff') - ERROR: Invalid filename generated: '$filename'"
                    return
                }
                
                Write-DebugMsg "About to write content (length: $($currentContent.Length)) to file: '$filename'"
                
                # Save to numbered file
                try {
                    [System.IO.File]::WriteAllText($filename, $currentContent, [System.Text.Encoding]::UTF8)
                    Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - Text saved: $filename"
                    
                    # Verify file was created correctly
                    if (Test-Path $filename) {
                        $fileSize = (Get-Item $filename).Length
                        Write-DebugMsg "File created successfully, size: $fileSize bytes"
                    } else {
                        Write-Error "$(Get-Date -Format 'HH:mm:ss.fff') - ERROR: File was not created: '$filename'"
                    }
                } catch {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss.fff') - ERROR: Failed to write file '$filename': $($_.Exception.Message)"
                    return
                }
                
                # Create JSON format for upload
                $uploadContent = @{
                    type = "PLAIN_TEXT"
                    content = $currentContent
                }
                $uploadJson = ConvertTo-Json $uploadContent -Depth 10
                
                # Upload to WebDAV
                Upload-ToWebDAV -Content $uploadJson -Connection $global:webdavConnection -LocalSyncFile $localSyncFile -LocalSyncFileGz $localSyncFileGz -RemoteFilePath $localUploadPath
                
                # Show notification for local clipboard capture
                $preview = if ($currentContent.Length -gt 50) { $currentContent.Substring(0, 50) + "..." } else { $currentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Captured: $preview" -Icon "Info"
            }
        }
    }
    catch {
        Write-host "Error accessing clipboard: $($_.Exception.Message)" -ForegroundColor Red
        if ($global:EnableDebugMessages) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - DEBUG: Exception details: $($_.Exception | Out-String)"
        }
    }
}

# Start simple like in simple.ps1 - avoid using too many variables and simplify the event registration
Write-Host "ClipSon clipboard monitor starting..." -ForegroundColor Cyan

# Register the event handler with the monitor
$monitorEvent = Register-ObjectEvent -InputObject $monitor -EventName ClipboardChanged -Action $handler
$monitor.CreateControl()
$monitor.Show()  # CRITICAL: Must show the window to receive Windows messages

# Register exit handler
$exitEvent = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action { 
    $global:exitLoop = $true
}

# Select remote sync file
$remotePeerFiles = Select-RemoteSyncFile -Connection $global:webdavConnection -RemoteFolder $global:Config.nextcloud.remote_folder

# Initialize peer tracking
Initialize-RemotePeerTracking -PeerFiles $remotePeerFiles

# REMOVE THE TIMER APPROACH - use a time-based check in the main loop instead
$lastCheckTime = [DateTime]::MinValue

Write-Host "ClipSon clipboard monitor started. Press Ctrl+C to stop." -ForegroundColor Cyan

# Main loop - use an approach that doesn't involve timers
try {
    while (-not $global:exitLoop) {
        # Process Windows messages
        [System.Windows.Forms.Application]::DoEvents()
        
        # Check if it's time to check for remote files
        $currentTime = [DateTime]::Now
        if (($currentTime - $lastCheckTime) -ge $remoteCheckInterval) {
            $lastCheckTime = $currentTime
            
            try {
                # Do the remote file check inline (no timer event to get interrupted)
                $result = Check-AllRemoteFilesForUpdates -Connection $global:webdavConnection -RemoteFolder $global:Config.nextcloud.remote_folder -CheckInterval $remoteCheckInterval
                
                if ($result) {
                    Set-ClipboardContentUnified -Content $result.Content
                    
                    # Show notification
                    $preview = if ($result.Content.Length -gt 50) { $result.Content.Substring(0, 50) + "..." } else { $result.Content }
                    Show-ClipboardNotification -Title "ClipSon" -Message "Remote update from $($result.Filename)" -Icon "Info"
                    
                    # Set flag to skip clipboard monitoring briefly
                    $global:skipClipboardCheckUntil = [DateTime]::Now.AddSeconds(2)
                }
            }
            catch [System.Management.Automation.PipelineStoppedException] {
                # Gracefully handle pipeline stopped
                $global:exitLoop = $true
                break
            }
            catch {
                Write-Warning "Remote check error: $($_.Exception.Message)"
            }
        }
        
        # Use a small sleep
        Start-Sleep -Milliseconds 100
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    # Handle Ctrl+C gracefully
    Write-Host "Stopping ClipSon..." -ForegroundColor Yellow
}
catch {
    Write-Host "Error in main loop: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Very simple cleanup - no timer to worry about now!
    Write-Host "ClipSon shutting down..." -ForegroundColor Yellow
    
    # Unregister events before disposing objects
    if ($monitorEvent) {
        Unregister-Event -SubscriptionId $monitorEvent.Id -ErrorAction SilentlyContinue
    }
    
    if ($exitEvent) {
        Unregister-Event -SubscriptionId $exitEvent.Id -ErrorAction SilentlyContinue
    }
    
    # Dispose the monitor object last
    if ($monitor) {
        $monitor.Dispose()
    }
    
    Write-Host "ClipSon stopped." -ForegroundColor Cyan
}

