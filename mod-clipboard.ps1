# Module for local clipboard handling functions
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Define ClipboardMonitor class for monitoring clipboard changes
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

public class ClipboardMonitor : Form
{
    private const int WM_CLIPBOARDUPDATE = 0x031D;
    
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool AddClipboardFormatListener(IntPtr hwnd);
    
    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    public static extern bool RemoveClipboardFormatListener(IntPtr hwnd);
    
    public event Action ClipboardChanged;
    
    public ClipboardMonitor()
    {
        this.WindowState = FormWindowState.Minimized;
        this.ShowInTaskbar = false;
        this.Visible = false;
    }
    
    protected override void OnHandleCreated(EventArgs e)
    {
        base.OnHandleCreated(e);
        AddClipboardFormatListener(this.Handle);
    }
    
    protected override void OnHandleDestroyed(EventArgs e)
    {
        RemoveClipboardFormatListener(this.Handle);
        base.OnHandleDestroyed(e);
    }
    
    protected override void WndProc(ref Message m)
    {
        if (m.Msg == WM_CLIPBOARDUPDATE)
        {        
            if (ClipboardChanged != null)
            {
                ClipboardChanged.Invoke();        
            }
        }   
        base.WndProc(ref m);
    }
}
"@ -ReferencedAssemblies System.Windows.Forms,System.Drawing

# Clipboard detection and format utility functions
function global:Get-ClipboardFormats {
    try {
        $formats = @()
        
        # Check for various clipboard formats
        if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::Html)) {
            $formats += "HTML"
        }
        if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::Rtf)) {
            $formats += "RTF"
        }
        if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::UnicodeText)) {
            $formats += "Unicode"
        }
        if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::Text)) {
            $formats += "Text"
        }
        if ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
            $formats += "Files"
        }
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $formats += "Image"
        }
        
        return $formats
    }
    catch {
        return @()
    }
}

function global:Test-ClipboardRichText {
    try {
        $formats = Get-ClipboardFormats
        $richTextFormats = @("HTML", "RTF", "Files")
        
        # Check if we have rich text formats
        $hasRichFormats = ($formats | Where-Object { $_ -in $richTextFormats }).Count -gt 0
        
        # If "Files" format is present, check if it contains actual file paths
        # If it does, don't consider this as rich text (it should be handled as file copy)
        if ($hasRichFormats -and "Files" -in $formats) {
            $fileList = Test-ClipboardFiles
            if ($fileList -and $fileList.Count -gt 0) {
                # Check if we have other rich text formats besides "Files"
                $otherRichFormats = @("HTML", "RTF")
                return ($formats | Where-Object { $_ -in $otherRichFormats }).Count -gt 0
            }
        }
        
        return $hasRichFormats
    }
    catch {
        return $false
    }
}

function global:Test-ClipboardFiles {
    try {
        $formats = Get-ClipboardFormats
        return "Files" -in $formats
    }
    catch {
        return $false
    }
}

function global:Get-ClipboardFiles {
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
            $fileList = [System.Windows.Forms.Clipboard]::GetFileDropList()
            $files = @()
            foreach ($file in $fileList) {
                $files += $file
            }
            return $files
        }
        return @()
    }
    catch {
        Write-DebugMsg "Error getting clipboard files: $($_.Exception.Message)"
        return @()
    }
}

function global:Show-ClipboardNotification {
    param(
        [string]$Title,
        [string]$Message,
        [string]$Icon = "Info"
    )
    
    try {
        # Create notification object
        $notification = New-Object System.Windows.Forms.NotifyIcon
        $notification.Icon = [System.Drawing.SystemIcons]::Information
        $notification.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info
        $notification.BalloonTipText = $Message
        $notification.BalloonTipTitle = $Title
        $notification.Visible = $true
        
        # Show balloon tip
        $notification.ShowBalloonTip(3000)
        
        # Clean up after delay
        Start-Sleep -Milliseconds 3500
        $notification.Dispose()
    }
    catch {
        # Fallback to console if notifications fail
        Write-Host "NOTIFICATION: $Title - $Message"
    }
}

function global:Save-ClipboardFiles {
    try {
        $fileList = Get-ClipboardFiles
        if (-not $fileList -or $fileList.Count -eq 0) {
            return $false
        }
        
        Write-DebugMsg "Processing $($fileList.Count) clipboard files"
        
        $fileNumber = Get-NextFileNumber
        $filesData = @{}
        $savedFiles = @()
        $maxFileSize = $global:Config.app.max_file_copy_size_mb * 1024 * 1024
        
        foreach ($filePath in $fileList) {
            try {
                # Check if file exists and is accessible
                if (-not (Test-Path $filePath)) {
                    Write-DebugMsg "File not found: $filePath"
                    continue
                }
                
                $fileInfo = Get-Item $filePath
                if ($fileInfo.PSIsContainer) {
                    Write-DebugMsg "Skipping directory: $filePath"
                    continue
                }
                
                # Check file size limit
                if ($fileInfo.Length -gt $maxFileSize) {
                    $sizeMB = [math]::Round($fileInfo.Length / 1024 / 1024, 1)
                    $maxSizeMB = $global:Config.app.max_file_copy_size_mb
                    Write-Host "File too large to copy: $filePath ($sizeMB MB > $maxSizeMB MB)" -ForegroundColor Yellow
                    continue
                }
                
                # Read file content
                $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
                $fileB64 = [System.Convert]::ToBase64String($fileBytes)
                
                # Get file name and MIME type
                $fileName = $fileInfo.Name
                $mimeType = Get-MimeType -FilePath $filePath
                
                # Store file data
                $filesData[$fileName] = @{
                    original_path = $filePath
                    content = $fileB64
                    size = $fileInfo.Length
                    mime_type = $mimeType
                }
                
                # Save file locally for reference
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                $localFile = Join-Path $global:outputDir "clipboard_file_${paddedNumber}_$fileName"
                Copy-Item $filePath $localFile
                $savedFiles += $localFile
                
                Write-DebugMsg "Processed file: $fileName ($($fileInfo.Length) bytes)"
            }
            catch {
                Write-DebugMsg "Error processing file $filePath`: $($_.Exception.Message)"
                continue
            }
        }
        
        if ($filesData.Count -gt 0) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - File(s) saved: $($filesData.Count) file(s)"
            foreach ($fileName in $filesData.Keys) {
                Write-Host "  - $fileName"
            }
            
            # Create file upload content
            $uploadContent = @{
                type = "CLIPBOARD_FILES"
                files = $filesData
            }
            $uploadJson = ConvertTo-Json $uploadContent -Depth 10
            
            # Upload to WebDAV
            Upload-ToWebDAV -Content $uploadJson -Connection $global:webdavConnection -LocalSyncFile $global:localSyncFile -LocalSyncFileGz $global:localSyncFileGz -RemoteFilePath $global:localUploadPath
            
            # Show notification
            $fileNames = $filesData.Keys | Select-Object -First 3
            $fileList = $fileNames -join ', '
            if ($filesData.Count -gt 3) {
                $fileList += " (+$($filesData.Count - 3) more)"
            }
            Show-ClipboardNotification -Title "ClipSon" -Message "Files captured: $fileList"
            
            return $true
        }
        
        return $false
    }
    catch {
        Write-Host "Error saving clipboard files: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

function Get-MimeType {
    param([string]$FilePath)
    
    try {
        $extension = [System.IO.Path]::GetExtension($FilePath).ToLower()
        
        # Common MIME types
        $mimeTypes = @{
            '.txt' = 'text/plain'
            '.html' = 'text/html'
            '.htm' = 'text/html'
            '.css' = 'text/css'
            '.js' = 'text/javascript'
            '.json' = 'application/json'
            '.xml' = 'text/xml'
            '.pdf' = 'application/pdf'
            '.doc' = 'application/msword'
            '.docx' = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            '.xls' = 'application/vnd.ms-excel'
            '.xlsx' = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            '.ppt' = 'application/vnd.ms-powerpoint'
            '.pptx' = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
            '.jpg' = 'image/jpeg'
            '.jpeg' = 'image/jpeg'
            '.png' = 'image/png'
            '.gif' = 'image/gif'
            '.bmp' = 'image/bmp'
            '.tiff' = 'image/tiff'
            '.ico' = 'image/x-icon'
            '.svg' = 'image/svg+xml'
            '.mp3' = 'audio/mpeg'
            '.wav' = 'audio/wav'
            '.mp4' = 'video/mp4'
            '.avi' = 'video/x-msvideo'
            '.zip' = 'application/zip'
            '.rar' = 'application/x-rar-compressed'
            '.7z' = 'application/x-7z-compressed'
            '.exe' = 'application/x-msdownload'
            '.msi' = 'application/x-msi'
        }
        
        if ($mimeTypes.ContainsKey($extension)) {
            return $mimeTypes[$extension]
        }
        
        return 'application/octet-stream'
    }
    catch {
        return 'application/octet-stream'
    }
}

function global:Set-ClipboardContentUnified {
    param([string]$Content)
    
    try {
        # Check if this is JSON content
        if ($Content.Trim().StartsWith('{')) {
            try {
                $data = ConvertFrom-Json $Content
                $contentType = $data.type
                
                if ($contentType -eq "CLIPBOARD_IMAGE") {
                    # Handle image content
                    $imageBytes = [System.Convert]::FromBase64String($data.data)
                    return Set-ClipboardImage -ImageData $imageBytes
                }
                elseif ($contentType -eq "MULTI_FORMAT_CLIPBOARD") {
                    # Handle multi-format text content
                    if ($data.formats) {
                        return Set-ClipboardMultipleFormats -FormatData $data.formats
                    }
                }
                elseif ($contentType -eq "PLAIN_TEXT") {
                    # Handle plain text content
                    if ($data.content) {
                        [System.Windows.Forms.Clipboard]::SetText($data.content)
                        return $true
                    }
                }
                elseif ($contentType -eq "CLIPBOARD_FILES") {
                    # Handle file content
                    if ($data.files) {
                        return Set-ClipboardFiles -FilesData $data.files
                    }
                }
            }
            catch {
                # JSON parse failed, fall back to legacy handling
            }
        }
        
        # Fall back to legacy or plain text handling
        return Set-ClipboardRichContent -Content $Content
    }
    catch {
        Write-Warning "Error setting unified clipboard content: $($_.Exception.Message)"
        [System.Windows.Forms.Clipboard]::SetText($Content)
        return $false
    }
}

function Set-ClipboardImage {
    param([byte[]]$ImageData)
    
    try {
        # Set image to clipboard
        $memoryStream = New-Object System.IO.MemoryStream(,$ImageData)
        $image = [System.Drawing.Image]::FromStream($memoryStream)
        [System.Windows.Forms.Clipboard]::SetImage($image)
        
        $image.Dispose()
        $memoryStream.Dispose()
        
        return $true
    }
    catch {
        Write-Warning "Error setting clipboard image: $($_.Exception.Message)"
        return $false
    }
}

function Set-ClipboardFiles {
    param([PSObject]$FilesData)
    
    try {
        # Create temporary directory for restored files
        $tempDir = ".\temp-restored-files-$($env:COMPUTERNAME)"
        if (!(Test-Path $tempDir)) {
            New-Item -ItemType Directory -Path $tempDir | Out-Null
        }
        
        $restoredFiles = @()
        
        foreach ($fileName in $FilesData.PSObject.Properties.Name) {
            $fileInfo = $FilesData.$fileName
            
            try {
                # Decode file content from base64
                $fileBytes = [System.Convert]::FromBase64String($fileInfo.content)
                
                # Create safe filename
                $safeFileName = Get-SafeFileName -FileName $fileName
                $tempFilePath = Join-Path $tempDir $safeFileName
                
                # Write file to temp location
                [System.IO.File]::WriteAllBytes($tempFilePath, $fileBytes)
                
                $restoredFiles += $tempFilePath
                Write-DebugMsg "Restored file: $safeFileName ($($fileBytes.Length) bytes) to $tempFilePath"
            }
            catch {
                Write-DebugMsg "Error restoring file $fileName`: $($_.Exception.Message)"
                continue
            }
        }
        
        if ($restoredFiles.Count -gt 0) {
            # Set clipboard with file list
            $fileCollection = New-Object System.Collections.Specialized.StringCollection
            foreach ($file in $restoredFiles) {
                $fileCollection.Add($file)
            }
            
            [System.Windows.Forms.Clipboard]::SetFileDropList($fileCollection)
            
            Write-Host "Restored $($restoredFiles.Count) file(s) to clipboard:"
            foreach ($filePath in $restoredFiles) {
                Write-Host "  - $filePath"
            }
            
            return $true
        }
        
        return $false
    }
    catch {
        Write-Warning "Error setting clipboard files: $($_.Exception.Message)"
        return $false
    }
}

function Get-SafeFileName {
    param([string]$FileName)
    
    # Remove or replace problematic characters
    $safeChars = '-_.() abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    $safeFileName = ''
    
    foreach ($char in $FileName.ToCharArray()) {
        if ($safeChars.Contains($char)) {
            $safeFileName += $char
        } else {
            $safeFileName += '_'
        }
    }
    
    # Ensure it's not empty and doesn't start with dot
    if ([string]::IsNullOrEmpty($safeFileName) -or $safeFileName.StartsWith('.')) {
        $safeFileName = 'restored_file_' + $safeFileName
    }
    
    # Limit length
    if ($safeFileName.Length -gt 255) {
        $safeFileName = $safeFileName.Substring(0, 255)
    }
    
    return $safeFileName
}

function Set-ClipboardMultipleFormats {
    param([PSObject]$FormatData)
    
    try {
        Write-DebugMsg "Setting $($FormatData.PSObject.Properties.Count) clipboard formats"        

        $dataObject = New-Object System.Windows.Forms.DataObject
        $successCount = 0
        $formatsSet = @()

        # Priority order for setting formats (most important first)
        $setPriority = @(
            'text/plain',
            'text/html', 
            'text/rtf', 
            'application/rtf', 
            'application/x-rtf'
        )
        
        # Map Linux MIME types to Windows formats
        $formatMap = @{
            'text/plain' = [System.Windows.Forms.TextDataFormat]::Text
            'text/html' = [System.Windows.Forms.TextDataFormat]::Html
            'text/rtf' = [System.Windows.Forms.TextDataFormat]::Rtf
            'application/rtf' = [System.Windows.Forms.TextDataFormat]::Rtf
            'application/x-rtf' = [System.Windows.Forms.TextDataFormat]::Rtf
        }
        
        # Try to set each format in priority order
        foreach ($formatName in $setPriority) {            
            
            if ($FormatData.PSObject.Properties.Name -contains $formatName) {
                try {
                    $content = $FormatData.$formatName
                    $windowsFormat = $formatMap[$formatName]
                    
                    # Enhanced debugging for content issues
                    if ($global:Config -and $global:Config.app.debug_enabled) {
                        Write-DebugMsg "Format $formatName - Content type: $($content.GetType().Name), Length: $($content.Length), WindowsFormat: $windowsFormat"
                        Write-DebugMsg "Content preview: '$($content.Substring(0, [Math]::Min(100, $content.Length)))'..."
                    }
                    
                    # More explicit content checking - avoid boolean context issues
                    $hasContent = ($content -ne $null) -and ($content.Length -gt 0)
                    $hasWindowsFormat = ($windowsFormat -ne $null)
                    
                    if ($hasWindowsFormat -and $hasContent) {
                        # Use SetData method instead of SetText to avoid overwriting
                        if ($formatName -eq 'text/plain') {
                            $dataObject.SetData([System.Windows.Forms.DataFormats]::Text, $content)
                            $dataObject.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $content)
                        }
                        elseif ($formatName -eq 'text/html') {
                            $dataObject.SetData([System.Windows.Forms.DataFormats]::Html, $content)
                        }
                        elseif ($formatName -in @('text/rtf', 'application/rtf', 'application/x-rtf')) {
                            $dataObject.SetData([System.Windows.Forms.DataFormats]::Rtf, $content)
                        }
                        
                        $successCount++
                        $formatsSet += $formatName
                        Write-DebugMsg "Successfully set format $formatName"
                    }
                    else {
                        Write-DebugMsg "Format $formatName - HasContent: $hasContent, HasWindowsFormat: $hasWindowsFormat"
                        if (-not $hasContent) {
                            Write-DebugMsg "Content is null or empty"
                        }
                        if (-not $hasWindowsFormat) {
                            Write-DebugMsg "No Windows format mapping found"
                        }
                    }
                }
                catch {
                    Write-DebugMsg "Error setting format $formatName`: $($_.Exception.Message)"
                }
            }
            else {
                Write-DebugMsg "Format $formatName not found in FormatData"
            }
        }
        
        # Set the data object to clipboard
        if ($successCount -gt 0) {
            [System.Windows.Forms.Clipboard]::SetDataObject($dataObject, $true)
            Write-DebugMsg "Successfully set $successCount formats to clipboard"
            return $true
        }
        else {
            Write-DebugMsg "No formats were successfully set"
            return $false
        }
    }
    catch {
        Write-Warning "Error setting multiple clipboard formats: $($_.Exception.Message)"
        return $false
    }
}

function Update-LastClipboardContentFromRemote {
    param([string]$Content)
    
    try {
        if ($Content.Trim().StartsWith('{')) {
            try {
                $data = ConvertFrom-Json $Content
                $contentType = $data.type
                
                if ($contentType -eq "CLIPBOARD_IMAGE") {
                    # Store comparison format for images (without data to save memory)
                    $imageBytes = [System.Convert]::FromBase64String($data.data)
                    $md5 = [System.Security.Cryptography.MD5]::Create()
                    $hashBytes = $md5.ComputeHash($imageBytes)
                    $imageHash = [System.BitConverter]::ToString($hashBytes) -replace '-'
                    $md5.Dispose()
                    
                    $script:lastClipboardImageHash = $imageHash
                    $script:lastClipboardContent = ConvertTo-Json @{
                        type = "CLIPBOARD_IMAGE"
                        format = if ($data.format) { $data.format } else { "png" }
                        size = if ($data.size) { $data.size } else { $imageBytes.Length }
                        hash = $imageHash
                    } -Compress
                }
                elseif ($contentType -eq "MULTI_FORMAT_CLIPBOARD" -and $data.formats) {
                    # Store in comparison format using content hashes for dynamic formats (same as main loop)
                    $comparisonData = @{}
                    foreach ($fmt in $data.formats.PSObject.Properties.Name) {
                        $content = $data.formats.$fmt
                        if ($fmt -in @('text/html', 'text/rtf')) {
                            # For formats that might have dynamic content, use content hash
                            $md5 = [System.Security.Cryptography.MD5]::Create()
                            $hashBytes = $md5.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($content))
                            $contentHash = [System.BitConverter]::ToString($hashBytes) -replace '-'
                            $md5.Dispose()
                            $comparisonData[$fmt] = $contentHash
                        }
                        else {
                            # For stable formats, use actual content
                            $comparisonData[$fmt] = $content
                        }
                    }
                    $script:lastClipboardContent = ConvertTo-Json @{ formats = $comparisonData } -Compress
                    Write-DebugMsg "Updated lastClipboardContent with remote multi-format comparison data"
                }
                elseif ($contentType -eq "PLAIN_TEXT" -and $data.content) {
                    # Store plain text content for comparison
                    $script:lastClipboardContent = $data.content
                }
                else {
                    $script:lastClipboardContent = $Content
                }
            }
            catch {
                $script:lastClipboardContent = $Content
            }
        }
        else {
            # Regular text or legacy rich content
            $script:lastClipboardContent = $Content
        }
    }
    catch {
        $script:lastClipboardContent = $Content
    }
}

# Export Save-ClipboardRichContentJson as a global function so it is visible in the handler
Set-Alias -Name Save-ClipboardRichContentJson -Value global:Save-ClipboardRichContentJson -Scope Global

function global:Save-ClipboardRichContentJson {
    param([hashtable]$FormatData)
    
    try {
        if ($FormatData.Count -eq 0) {
            return $false
        }
        
        Write-DebugMsg "Available clipboard formats: $($FormatData.Keys -join ', ')"
        
        $fileNumber = Get-NextFileNumber
        $baseFilename = "$outputDir\clipboard_rich_$($fileNumber.ToString().PadLeft(3, '0'))"
        
        $savedFiles = @()
        # Build a new hashtable for processed content
        $formatDataForJson = @{}
        
        # Save each available format to individual files
        foreach ($format in $FormatData.Keys) {
            $content = $FormatData[$format]
            # Strip Windows clipboard HTML header if present
            # TODO not quite right, but works for now
            if ($format -eq 'text/html' -and $content -is [string]) {
                if ($content -match "(?s)StartFragment:(\d+).*EndFragment:(\d+)") {
                    $startFragment = [int]$matches[1]
                    $endFragment = [int]$matches[2]
                    $content = $content.Substring($startFragment, $endFragment - $startFragment)
                    Write-DebugMsg "Stripped clipboard HTML header, using only HTML fragment."
                }
            }
            $extension = switch ($format) {
                'text/html' { '.html' }
                'text/rtf' { '.rtf' }
                'text/plain' { '.txt' }
                default { '.txt' }
            }
            
            $filename = "$baseFilename$extension"
            [System.IO.File]::WriteAllText($filename, $content, $utf8NoBom)
            $savedFiles += $filename
            Write-DebugMsg "Saved $format to $filename"
            # Add processed content to new hashtable
            $formatDataForJson[$format] = $content
        }
        
        if ($savedFiles.Count -gt 0) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss.fff') - Rich content saved: $($savedFiles.Count) unique formats"
            foreach ($file in $savedFiles) {
                Write-Host "  - $file"
            }
            
            # Create multi-format upload content with proper JSON escaping
            $uploadContent = @{
                type = "MULTI_FORMAT_CLIPBOARD"
                formats = $formatDataForJson
            }
            
            # Use ConvertTo-Json with proper depth and ensure ASCII escaping for special characters
            $uploadJson = ConvertTo-Json $uploadContent -Depth 10 -Compress:$false
            
            # Debug: Show JSON structure
            if ($global:Config -and $global:Config.app.debug_enabled) {
                Write-DebugMsg "JSON structure created, length: $($uploadJson.Length)"
                Write-DebugMsg "JSON preview: $($uploadJson.Substring(0, [Math]::Min(200, $uploadJson.Length)))..."
                
                # Validate JSON by parsing it back
                try {
                    $testParse = ConvertFrom-Json $uploadJson
                    Write-DebugMsg "JSON validation successful, formats in parsed JSON: $($testParse.formats.PSObject.Properties.Name -join ', ')"
                    
                    # Check RTF content specifically
                    if ($testParse.formats.'text/rtf') {
                        $rtfContent = $testParse.formats.'text/rtf'
                        $hasRtfTags = $rtfContent -match '\\rtf' -and $rtfContent -match '\\par'
                        Write-DebugMsg "RTF content validation: RTF tags present = $hasRtfTags"
                        Write-DebugMsg "RTF content preview: $($rtfContent.Substring(0, [Math]::Min(100, $rtfContent.Length)))..."
                    }
                }
                catch {
                    Write-DebugMsg "JSON validation failed: $($_.Exception.Message)"
                }
            }
            
            # Upload to WebDAV
            Upload-ToWebDAV -Content $uploadJson -Connection $global:webdavConnection -LocalSyncFile $global:localSyncFile -LocalSyncFileGz $global:localSyncFileGz -RemoteFilePath $global:localUploadPath
            
            # Show notification
            $formatList = ($FormatData.Keys | Select-Object -First 3) -join ', '
            Show-ClipboardNotification -Title "ClipSon" -Message "Rich content captured: $formatList" -Icon "Info"
            
            return $true
        }
        
        return $false
    }
    catch {
        Write-Warning "Error saving rich clipboard content: $($_.Exception.Message)"
        return $false
    }
}

function Set-ClipboardRichContent {
    param([string]$Content)
    
    try {
        # Check if this is rich content
        if ($Content.StartsWith("RICH_CONTENT_FORMATS:")) {
            $lines = $Content -split "`n", 3
            if ($lines.Count -ge 3) {
                $formatsLine = $lines[0]
                $actualContent = $lines[2]
                
                # Extract format information
                $formatStr = $formatsLine -replace "RICH_CONTENT_FORMATS:", "" -replace "^\s+", ""
                $availableFormats = $formatStr -split "," | ForEach-Object { $_.Trim() }
                
                Write-DebugMsg "Setting rich content with formats: $($availableFormats -join ', ')"
                
                # Set content based on available format
                if ($availableFormats -contains "HTML") {
                    try {
                        $dataObject = New-Object System.Windows.Forms.DataObject
                        $dataObject.SetText($actualContent, [System.Windows.Forms.TextDataFormat]::Html)
                        $dataObject.SetText($actualContent, [System.Windows.Forms.TextDataFormat]::Text)  # Fallback
                        [System.Windows.Forms.Clipboard]::SetDataObject($dataObject)
                        Write-DebugMsg "Set HTML content to clipboard"
                        return $true
                    }
                    catch {
                        Write-DebugMsg "Failed to set HTML content: $($_.Exception.Message)"
                    }
                }
                elseif ($availableFormats -contains "RTF") {
                    try {
                        $dataObject = New-Object System.Windows.Forms.DataObject
                        $dataObject.SetText($actualContent, [System.Windows.Forms.TextDataFormat]::Rtf)
                        $dataObject.SetText($actualContent, [System.Windows.Forms.TextDataFormat]::Text)  # Fallback
                        [System.Windows.Forms.Clipboard]::SetDataObject($dataObject)
                        Write-DebugMsg "Set RTF content to clipboard"
                        return $true
                    }
                    catch {
                        Write-DebugMsg "Failed to set RTF content: $($_.Exception.Message)"
                    }
                }
                
                # Fallback to plain text
                [System.Windows.Forms.Clipboard]::SetText($actualContent)
                return $true
            }
        }
        else {
            # Regular text content
            [System.Windows.Forms.Clipboard]::SetText($Content)
            return $true
        }
    }
    catch {
        Write-Warning "Error setting rich clipboard content: $($_.Exception.Message)"
        return $false
    }
}



