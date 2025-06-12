# ClipSon - Advanced Clipboard Synchronization Tool (All-in-One)
# If you get execution policy errors, run one of these commands first:
# Option 1 (Recommended): Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Option 2 (Temporary): powershell.exe -ExecutionPolicy Bypass -File "clipson-all.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

# Dot-source helper functions from modules
. "$PSScriptRoot/mod-misc.ps1"
. "$PSScriptRoot/mod-nextcloud.ps1"
. "$PSScriptRoot/mod-clipboard.ps1"


# Configuration (now loaded from config.json)
$NextcloudConfig = @{
    ServerUrl    = $Config.nextcloud.server_url      # Replace with your Nextcloud server URL
    Username     = $Config.nextcloud.username        # Replace with your username
    Password     = $Config.nextcloud.password        # Empty password will prompt for input
    RemoteFolder = $Config.nextcloud.remote_folder   # Folder in Nextcloud where files will be uploaded
    # Proxy Configuration (optional)
    ProxyUrl     = $Config.proxy.url                 # e.g., "http://proxy.company.com:8080" or "" for no proxy
    ProxyUsername= $Config.proxy.username            # Proxy username (if required)
    ProxyPassword= $Config.proxy.password            # Proxy password (if required)
    UseSystemProxy= $Config.proxy.use_system_proxy   # Use system proxy settings if true, manual proxy if false
}

# Normalize RemoteFolder to ensure trailing slash
if (-not $NextcloudConfig.RemoteFolder.EndsWith('/')) {
    $NextcloudConfig.RemoteFolder += '/'
}

# ClipSon Functions
function Get-RemoteClipboardFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RemoteFolder
    )

    try {
        # List all files in the remote clipboard folder
        $files = Get-ChildItem -Path $RemoteFolder -File | Sort-Object LastWriteTime -Descending
        return $files
    }
    catch {
        Write-Error "Failed to list remote clipboard files: $($_.Exception.Message)"
        return @()
    }
}

function Select-RemoteSyncFile {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RemoteFolder
    )

    try {
        # Get the most recent file in the remote folder
        $mostRecentFile = Get-RemoteClipboardFiles -RemoteFolder $RemoteFolder | Select-Object -First 1

        if ($mostRecentFile) {
            return $mostRecentFile.FullName
        }
        else {
            Write-Warning "No files found in the remote folder: $RemoteFolder"
            return $null
        }
    }
    catch {
        Write-Error "Failed to select remote sync file: $($_.Exception.Message)"
        return $null
    }
}

function Initialize-RemotePeerTracking {
    param (
        [Parameter(Mandatory = $true)]
        [string]$RemoteFolder
    )

    try {
        # Initialize or reset the remote peer tracking
        $global:remotePeerFiles = @{}
        $global:remoteCheckInterval = [TimeSpan]::FromSeconds(5)

        # Load existing remote files
        $existingFiles = Get-RemoteClipboardFiles -RemoteFolder $RemoteFolder
        foreach ($file in $existingFiles) {
            $global:remotePeerFiles[$file.Name] = $file
        }

        Write-Host "Remote peer tracking initialized. Found $($existingFiles.Count) file(s)."
    }
    catch {
        Write-Error "Failed to initialize remote peer tracking: $($_.Exception.Message)"
    }
}

function Check-AllRemoteFilesForUpdates {
    param (
        [Parameter(Mandatory = $true)]
        $Connection
    )

    try {
        # Check for updates in all remote files
        $remoteFiles = Get-RemoteClipboardFiles -RemoteFolder $NextcloudConfig.RemoteFolder

        foreach ($remoteFile in $remoteFiles) {
            # Check if we already have this file tracked
            if (-not $global:remotePeerFiles.ContainsKey($remoteFile.Name)) {
                # New file detected, download it
                Write-Host "New remote file detected: $($remoteFile.Name)"
                Download-RemoteFile -File $remoteFile -Connection $Connection
            }
            else {
                # File is already tracked, check for updates
                $localFile = $global:remotePeerFiles[$remoteFile.Name]

                # Compare last write times
                if ($remoteFile.LastWriteTime -gt $localFile.LastWriteTime) {
                    # Remote file is newer, download it
                    Write-Host "Remote file updated: $($remoteFile.Name)"
                    Download-RemoteFile -File $remoteFile -Connection $Connection
                }
            }
        }

        # Update the tracking information
        Initialize-RemotePeerTracking -RemoteFolder $NextcloudConfig.RemoteFolder
    }
    catch {
        Write-Error "Failed to check remote files for updates: $($_.Exception.Message)"
    }
}

# Create and use the clipboard monitor
$clipboardMonitor = New-Object ClipboardMonitor

# Define the clipboard change handler as a script block
$clipboardChangeHandler = {
    try {
        # Check if we should skip clipboard monitoring (after setting remote content)
        $currentTime = [DateTime]::Now
        if ($currentTime -lt $global:skipClipboardCheckUntil) {
            Write-DebugMsg "Skipping clipboard check due to recent remote update"
            return
        }
        
        Write-DebugMsg "Clipboard content changed at $(Get-Date -Format 'HH:mm:ss')"
        
        # Check if clipboard has image content first (highest priority)
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $image = [System.Windows.Forms.Clipboard]::GetImage()
            if ($image -ne $null) {
                # Save image to memory stream to calculate hash without creating file
                $memoryStream = New-Object System.IO.MemoryStream
                $image.Save($memoryStream, [System.Drawing.Imaging.ImageFormat]::Png)
                $imageBytes = $memoryStream.ToArray()
                $memoryStream.Dispose()
                
                # Calculate hash from memory
                $md5 = [System.Security.Cryptography.MD5]::Create()
                $hashBytes = $md5.ComputeHash($imageBytes)
                $currentImageHash = [System.BitConverter]::ToString($hashBytes) -replace '-'
                $md5.Dispose()
                
                # Create comparison format
                $currentImageComparison = ConvertTo-Json @{
                    type = "CLIPBOARD_IMAGE"
                    format = "png"
                    size = $imageBytes.Length
                    hash = $currentImageHash
                } -Compress
                
                Write-DebugMsg "Current image hash: $currentImageHash, Last comparison: $($script:lastClipboardContent)"
                
                if ($currentImageComparison -ne $script:lastClipboardContent) {
                    # Hash is different, proceed with saving
                    Write-DebugMsg "About to get next file number for image"
                    $fileNumber = Get-NextFileNumber
                    $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                    $filename = "$outputDir\clipboard_image_$paddedNumber.png"
                    
                    Write-DebugMsg "Creating image file: $filename"
                    
                    # Save the image bytes to file
                    [System.IO.File]::WriteAllBytes($filename, $imageBytes)
                    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Image saved: $filename"
                    
                    # Update hash tracking
                    $script:lastClipboardImageHash = $currentImageHash
                    
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
                    Upload-ToWebDAV -Content $uploadJson

                    # Show notification for image capture
                    Show-ClipboardNotification -Title "ClipSon" -Message "Image captured: $filename" -Icon "Info"
                    
                    # Update last clipboard content for comparison
                    $script:lastClipboardContent = $currentImageComparison
                }
                elseif ($global:EnableDebugMessages) {
                    Write-DebugMsg "Image unchanged, skipping..."
                }
                
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
            
            if ($currentFormats -contains "RTF") {
                try {
                    $rtfContent = [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Rtf)
                    if ($rtfContent) {
                        $currentFormatData["text/rtf"] = $rtfContent
                    }
                }
                catch { }
            }
            
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
                # Create comparison string using content hashes for dynamic formats
                $comparisonData = @{}
                foreach ($fmt in $currentFormatData.Keys) {
                    if ($fmt -in @('text/html', 'text/rtf')) {
                        # For formats that might have dynamic content, use content hash
                        $md5 = [System.Security.Cryptography.MD5]::Create()
                        $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($currentFormatData[$fmt]))
                        $contentHash = [System.BitConverter]::ToString($hashBytes) -replace '-'
                        $md5.Dispose()
                        $comparisonData[$fmt] = $contentHash
                    }
                    else {
                        # For stable formats, use actual content
                        $comparisonData[$fmt] = $currentFormatData[$fmt]
                    }
                }
                
                $currentComparisonContent = ConvertTo-Json @{ formats = $comparisonData } -Compress
                
                # Check if we have a meaningful change
                $hasMeaningfulChange = $true
                if ($script:lastClipboardContent -and $script:lastClipboardContent.StartsWith('{"formats"')) {
                    try {
                        $lastComparisonData = (ConvertFrom-Json $script:lastClipboardContent).formats
                        
                        # Compare stable text formats for meaningful changes
                        $stableFormats = @('text/plain')
                        $currentStable = @{}
                        $lastStable = @{}
                        
                        foreach ($fmt in $stableFormats) {
                            if ($comparisonData.ContainsKey($fmt)) {
                                $currentStable[$fmt] = $comparisonData[$fmt]
                            }
                            if ($lastComparisonData.PSObject.Properties.Name -contains $fmt) {
                                $lastStable[$fmt] = $lastComparisonData.$fmt
                            }
                        }
                        
                        # Compare as JSON strings
                        $currentStableJson = ConvertTo-Json $currentStable -Compress
                        $lastStableJson = ConvertTo-Json $lastStable -Compress
                        
                        if ($currentStableJson -eq $lastStableJson) {
                            $hasMeaningfulChange = $false
                            Write-DebugMsg "Only dynamic content changed (HTML/RTF timestamps), ignoring..."
                        }
                    }
                    catch {
                        Write-DebugMsg "Error comparing content: $($_.Exception.Message)"
                    }
                }
                
                # Only save if there's a meaningful change
                if ($hasMeaningfulChange) {
                    Write-DebugMsg "Rich content changed, saving..."
                    Save-ClipboardRichContentJson -FormatData $currentFormatData
                    # Store the comparison data for next time
                    $script:lastClipboardContent = $currentComparisonContent
                }
                elseif ($global:EnableDebugMessages) {
                    Write-DebugMsg "Rich content unchanged (ignoring dynamic changes), skipping..."
                }
            }
        }
        # Check if clipboard has text content (lowest priority)
        elseif ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $currentContent = [System.Windows.Forms.Clipboard]::GetText()
            
            # Only save if content has changed
            if ($currentContent -ne $script:lastClipboardContent -and $currentContent.Trim() -ne "") {
                Write-DebugMsg "About to get next file number"
                
                $fileNumber = Get-NextFileNumber
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                $filename = Join-Path $outputDir "clipboard_text_$paddedNumber.txt"
                
                # Verify the filename is correct before writing
                if ($filename -notmatch "clipboard_text_\d{3}\.txt$") {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss') - ERROR: Invalid filename generated: '$filename'"
                    return
                }
                
                Write-DebugMsg "About to write content (length: $($currentContent.Length)) to file: '$filename'"
                
                # Save to numbered file
                try {
                    [System.IO.File]::WriteAllText($filename, $currentContent, [System.Text.Encoding]::UTF8)
                    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Text saved: $filename"
                    
                    # Verify file was created correctly
                    if (Test-Path $filename) {
                        $fileSize = (Get-Item $filename).Length
                        Write-DebugMsg "File created successfully, size: $fileSize bytes"
                    } else {
                        Write-Error "$(Get-Date -Format 'HH:mm:ss') - ERROR: File was not created: '$filename'"
                    }
                } catch {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss') - ERROR: Failed to write file '$filename': $($_.Exception.Message)"
                    return
                }
                
                # Create JSON format for upload
                $uploadContent = @{
                    type = "PLAIN_TEXT"
                    content = $currentContent
                }
                $uploadJson = ConvertTo-Json $uploadContent -Depth 10
                
                # Upload to WebDAV
                Upload-ToWebDAV -Content $uploadJson
                
                # Show notification for local clipboard capture
                $preview = if ($currentContent.Length -gt 50) { $currentContent.Substring(0, 50) + "..." } else { $currentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Captured: $preview" -Icon "Info"
                
                $script:lastClipboardContent = $currentContent
            }
        }
    }
    catch {
        Write-host "Error accessing clipboard: $($_.Exception.Message)"
        if ($global:EnableDebugMessages) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: Exception details: $($_.Exception | Out-String)"
        }
    }
}

# Register event handler for clipboard changes
Register-ObjectEvent -InputObject $clipboardMonitor -EventName "ClipboardChanged" -Action $clipboardChangeHandler

# Start the clipboard monitor
$clipboardMonitor.CreateControl()

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

# Keep the application running with Windows Forms message loop
Write-Host "ClipSon started with real-time clipboard monitoring. Press Ctrl+C to stop."
Write-Host "Captured content will be saved to: $outputDir"
Write-Host "Maximum entries: $maxEntries (older files will be automatically deleted)"
Write-Host "Local sync file: $localSyncFile"
Write-Host "Local upload file (upload): $localUploadPath"
Write-Host "Remote peer files (download): $($remotePeerFiles.Count) peer(s)"
foreach ($peerFile in $remotePeerFiles) {
    Write-Host "  - $($peerFile.Name)" -ForegroundColor Green
}
Write-Host "Remote check interval: $($remoteCheckInterval.TotalSeconds) seconds"
Write-Host "Multiple peer synchronization: Enabled" -ForegroundColor Green

# Add global flag for clean shutdown
$global:shouldExit = $false

# Add Ctrl+C handler
$null = Register-EngineEvent -SourceIdentifier PowerShell.Exiting -Action {
    Write-Host "`nShutting down ClipSon..." -ForegroundColor Yellow
    $global:shouldExit = $true
}

# Also handle Ctrl+C specifically
[console]::TreatControlCAsInput = $false
$ctrlCHandler = {
    Write-Host "`nCtrl+C detected, shutting down ClipSon gracefully..." -ForegroundColor Yellow
    $global:shouldExit = $true
}

# Set up console cancel event handler
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;

public static class ConsoleHelper
{
    public delegate bool ConsoleCtrlDelegate(int dwCtrlType);
    
    [DllImport("kernel32.dll")]
    public static extern bool SetConsoleCtrlHandler(ConsoleCtrlDelegate HandlerRoutine, bool Add);
    
    public const int CTRL_C_EVENT = 0;
    public const int CTRL_BREAK_EVENT = 1;
    public const int CTRL_CLOSE_EVENT = 2;
    public const int CTRL_LOGOFF_EVENT = 5;
    public const int CTRL_SHUTDOWN_EVENT = 6;
}
"@

$consoleHandler = {
    param($ctrlType)
    if ($ctrlType -eq 0 -or $ctrlType -eq 1) {  # CTRL_C_EVENT or CTRL_BREAK_EVENT
        Write-Host "`nCtrl+C detected, shutting down ClipSon gracefully..." -ForegroundColor Yellow
        $global:shouldExit = $true
        return $true  # Indicate we handled the event (prevents default behavior)
    }
    return $false  # Let system handle other events
}

[ConsoleHelper]::SetConsoleCtrlHandler($consoleHandler, $true)

try {
    # Keep the application running with a manual loop instead of Windows Forms message loop
    Write-Host "Press Ctrl+C to stop ClipSon." -ForegroundColor Green
    
    while (-not $global:shouldExit) {
        try {
            # Process Windows messages to keep the clipboard monitor active
            [System.Windows.Forms.Application]::DoEvents()
            
            # Small sleep to prevent high CPU usage
            Start-Sleep -Milliseconds 100
        }
        catch [System.Management.Automation.PipelineStoppedException] {
            # Handle pipeline stopped exception gracefully
            Write-Host "`nPipeline stopped, exiting gracefully..." -ForegroundColor Yellow
            $global:shouldExit = $true
            break
        }
        catch {
            # Handle other exceptions
            if ($global:EnableDebugMessages) {
                Write-DebugMsg "Error in main loop: $($_.Exception.Message)"
            }
        }
    }
}
catch [System.Management.Automation.PipelineStoppedException] {
    # Handle pipeline stopped exception at top level
    Write-Host "`nClipSon stopped gracefully." -ForegroundColor Green
}
catch {
    Write-Host "Unexpected error: $($_.Exception.Message)" -ForegroundColor Red
}
finally {
    # Cleanup - this will always run
    Write-Host "Cleaning up resources..." -ForegroundColor Yellow
    
    try {
        if ($timer) {
            $timer.Stop()
            $timer.Dispose()
        }
    } catch { }
    
    try {
        if ($clipboardMonitor) {
            $clipboardMonitor.Dispose()
        }
    } catch { }
    
    # Remove console handler
    try {
        [ConsoleHelper]::SetConsoleCtrlHandler($consoleHandler, $false)
    } catch { }
    
    Write-Host "ClipSon stopped successfully." -ForegroundColor Green
}
