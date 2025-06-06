# ClipSon - Advanced Clipboard Synchronization Tool (All-in-One)
# If you get execution policy errors, run one of these commands first:
# Option 1 (Recommended): Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Option 2 (Temporary): powershell.exe -ExecutionPolicy Bypass -File "clipson-all.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Load configuration from JSON file
function Load-Configuration {
    $configPath = ".\config.json"
    if (-not (Test-Path $configPath)) {
        Write-Error "Configuration file not found: $configPath"
        Write-Host "Please create config.json file with your Nextcloud settings."
        exit 1
    }
    
    try {
        $configContent = Get-Content $configPath -Raw -Encoding UTF8
        $config = ConvertFrom-Json $configContent
        return $config
    }
    catch {
        Write-Error "Failed to parse configuration file: $($_.Exception.Message)"
        exit 1
    }
}

# Load configuration
$Config = Load-Configuration

# Global Debug Configuration
$global:EnableDebugMessages = $Config.app.debug_enabled  # Set to $true to enable debug output

# Debug helper function
function Write-DebugMsg {
    param([string]$Message)
    if ($global:EnableDebugMessages) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: $Message"
    }
}

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

# Function to get password if needed
function Get-PasswordIfNeeded {
    if ([string]::IsNullOrWhiteSpace($NextcloudConfig.Password)) {
        Write-Host "No password configured for user: $($NextcloudConfig.Username)" -ForegroundColor Yellow
        
        # Use Read-Host with -AsSecureString for secure password input
        $SecurePassword = Read-Host "Please enter your Nextcloud password" -AsSecureString
        
        # Convert SecureString to plain text (needed for WebDAV authentication)
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
        $PlainPassword = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
        
        if ([string]::IsNullOrWhiteSpace($PlainPassword)) {
            Write-Error "Password cannot be empty. Exiting."
            exit 1
        }
        
        $NextcloudConfig.Password = $PlainPassword
        Write-Host "Password configured successfully." -ForegroundColor Green
    }
}

# WebDAV Functions
function New-NextcloudConnection {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ServerUrl,
        
        [Parameter(Mandatory=$true)]
        [string]$Username,
        
        [Parameter(Mandatory=$true)]
        [string]$Password,
        
        [Parameter(Mandatory=$false)]
        [string]$ProxyUrl = "",
        
        [Parameter(Mandatory=$false)]
        [string]$ProxyUsername = "",
        
        [Parameter(Mandatory=$false)]
        [string]$ProxyPassword = "",
        
        [Parameter(Mandatory=$false)]
        [bool]$UseSystemProxy = $true
    )
    
    # Remove trailing slash if present
    $ServerUrl = $ServerUrl.TrimEnd('/')
    
    # Create WebDAV URL
    $WebDAVUrl = "$ServerUrl/remote.php/dav/files/$Username/"
    
    # Create credentials
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
    
    # Setup proxy configuration
    $ProxyConfig = $null
    if (-not $UseSystemProxy -and $ProxyUrl) {
        $ProxyConfig = @{
            Url = $ProxyUrl
            Username = $ProxyUsername
            Password = $ProxyPassword
        }
    }
    
    # Create connection object
    $Connection = @{
        WebDAVUrl = $WebDAVUrl
        Credential = $Credential
        Username = $Username
        ServerUrl = $ServerUrl
        ProxyConfig = $ProxyConfig
        UseSystemProxy = $UseSystemProxy
    }
    
    return $Connection
}

function Set-WebRequestProxy {
    param(
        [System.Net.WebRequest]$WebRequest,
        [hashtable]$Connection
    )
    
    if ($Connection.UseSystemProxy) {
        # Use system proxy settings
        $WebRequest.Proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $WebRequest.Proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    }
    elseif ($Connection.ProxyConfig -and $Connection.ProxyConfig.Url) {
        # Use manual proxy configuration
        $proxy = New-Object System.Net.WebProxy($Connection.ProxyConfig.Url)
        
        if ($Connection.ProxyConfig.Username -and $Connection.ProxyConfig.Password) {
            $proxySecurePassword = ConvertTo-SecureString $Connection.ProxyConfig.Password -AsPlainText -Force
            $proxyCredential = New-Object System.Management.Automation.PSCredential ($Connection.ProxyConfig.Username, $proxySecurePassword)
            $proxy.Credentials = $proxyCredential.GetNetworkCredential()
        }
        
        $WebRequest.Proxy = $proxy
    }
}

function Send-FileToNextcloud {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Connection,
        
        [Parameter(Mandatory=$true)]
        [string]$LocalFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$RemoteFilePath
    )
    
    try {
        # Check if local file exists
        if (-not (Test-Path $LocalFilePath)) {
            throw "Local file not found: $LocalFilePath"
        }
        
        # Prepare remote path (remove leading slash if present)
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $UploadUrl = $Connection.WebDAVUrl + $RemoteFilePath
        
        # Read file content
        $FileContent = [System.IO.File]::ReadAllBytes($LocalFilePath)
        
        # Create web request
        $WebRequest = [System.Net.WebRequest]::Create($UploadUrl)
        $WebRequest.Method = "PUT"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.ContentLength = $FileContent.Length
        $WebRequest.ContentType = "application/octet-stream"
        
        # Configure proxy
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        
        # Upload file
        $RequestStream = $WebRequest.GetRequestStream()
        $RequestStream.Write($FileContent, 0, $FileContent.Length)
        $RequestStream.Close()
        
        # Get response
        $Response = $WebRequest.GetResponse()
        $StatusCode = $Response.StatusCode
        $Response.Close()
        
        if ($StatusCode -eq "Created" -or $StatusCode -eq "NoContent") {
            Write-DebugMsg "Successfully uploaded: $LocalFilePath -> $RemoteFilePath"
            return $true
        } else {
            Write-Warning "Upload failed with status: $StatusCode"
            return $false
        }
    }
    catch {
        Write-Error "Upload failed: $($_.Exception.Message)"
        return $false
    }
}

function Test-NextcloudConnection {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Connection
    )
    
    try {
        # Test connection by making a PROPFIND request to the root directory
        $WebRequest = [System.Net.WebRequest]::Create($Connection.WebDAVUrl)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "0")
        
        # Configure proxy
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        
        $Response = $WebRequest.GetResponse()
        $StatusCode = [int]$Response.StatusCode
        $Response.Close()
        
        if ($StatusCode -eq 207) {  # 207 = MultiStatus
            Write-Host "Nextcloud connection successful (Status: $StatusCode)"
            return $true
        } else {
            Write-Warning "Connection test failed with status: $StatusCode"
            return $false
        }
    }
    catch {
        Write-Error "Connection test failed: $($_.Exception.Message)"
        return $false
    }
}

function Get-NextcloudFileTimestamp {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Connection,
        
        [Parameter(Mandatory=$true)]
        [string]$RemoteFilePath
    )
    
    try {
        # Prepare remote path (remove leading slash if present)
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $PropFindUrl = $Connection.WebDAVUrl + $RemoteFilePath
        
        # Create PROPFIND request to get file properties
        $WebRequest = [System.Net.WebRequest]::Create($PropFindUrl)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "0")
        $WebRequest.ContentType = "application/xml"
        
        # Configure proxy
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        
        # PROPFIND body to request last modified time
        $PropFindBody = @"
<?xml version="1.0" encoding="utf-8" ?>
<D:propfind xmlns:D="DAV:">
    <D:prop>
        <D:getlastmodified/>
    </D:prop>
</D:propfind>
"@
        
        $RequestBytes = [System.Text.Encoding]::UTF8.GetBytes($PropFindBody)
        $WebRequest.ContentLength = $RequestBytes.Length
        
        # Send the request body
        $RequestStream = $WebRequest.GetRequestStream()
        $RequestStream.Write($RequestBytes, 0, $RequestBytes.Length)
        $RequestStream.Close()
        
        # Get response
        $Response = $WebRequest.GetResponse()
        $ResponseStream = $Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($ResponseStream)
        $ResponseContent = $Reader.ReadToEnd()
        $Reader.Close()
        $Response.Close()
        
        # Parse XML response to extract last modified date
        $XmlDoc = New-Object System.Xml.XmlDocument
        $XmlDoc.LoadXml($ResponseContent)
        
        # Find the getlastmodified element
        $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
        $NamespaceManager.AddNamespace("D", "DAV:")
        
        $LastModifiedNode = $XmlDoc.SelectSingleNode("//D:getlastmodified", $NamespaceManager)
        
        if ($LastModifiedNode -and $LastModifiedNode.InnerText) {
            # Parse the RFC 2822 date format (e.g., "Mon, 12 Jan 1998 09:25:56 GMT")
            $LastModifiedUTC = [DateTime]::ParseExact($LastModifiedNode.InnerText, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
            # Convert GMT/UTC to local time for proper comparison
            $LastModifiedLocal = $LastModifiedUTC.ToLocalTime()
            return $LastModifiedLocal
        } else {
            Write-Warning "Could not extract last modified time from response"
            return $null
        }
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response.StatusCode -eq "NotFound") {
            Write-Host "Remote file does not exist: $RemoteFilePath"
            return $null
        } else {
            Write-Error "Failed to get remote file timestamp: $($_.Exception.Message)"
            return $null
        }
    }
    catch {
        Write-Error "Failed to get remote file timestamp: $($_.Exception.Message)"
        return $null
    }
}

function Get-NextcloudFile {
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$Connection,
        
        [Parameter(Mandatory=$true)]
        [string]$RemoteFilePath,
        
        [Parameter(Mandatory=$true)]
        [string]$LocalFilePath,
        
        [Parameter(Mandatory=$false)]
        [switch]$IncludeTimestamp
    )
    
    try {
        # Prepare remote path (remove leading slash if present)
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $DownloadUrl = $Connection.WebDAVUrl + $RemoteFilePath
        
        # Create GET request to download file
        $WebRequest = [System.Net.WebRequest]::Create($DownloadUrl)
        $WebRequest.Method = "GET"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        
        # Configure proxy
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        
        # Get response
        $Response = $WebRequest.GetResponse()
        $StatusCode = [int]$Response.StatusCode
        
        if ($StatusCode -eq 200) {
            # Create directory if it doesn't exist
            $LocalDir = Split-Path $LocalFilePath -Parent
            if (!(Test-Path $LocalDir)) {
                New-Item -ItemType Directory -Path $LocalDir -Force | Out-Null
            }
            
            # Download file content
            $ResponseStream = $Response.GetResponseStream()
            $FileStream = [System.IO.File]::Create($LocalFilePath)
            $ResponseStream.CopyTo($FileStream)
            $FileStream.Close()
            $ResponseStream.Close()
            
            # Extract last modified time from response headers if requested
            $LastModified = $null
            if ($IncludeTimestamp -and $Response.Headers["Last-Modified"]) {
                try {
                    $LastModifiedUTC = [DateTime]::ParseExact($Response.Headers["Last-Modified"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
                    $LastModified = $LastModifiedUTC.ToLocalTime()
                    
                    # Set the local file's timestamp to match the remote file
                    $FileInfo = Get-Item $LocalFilePath
                    $FileInfo.LastWriteTime = $LastModified
                    $FileInfo.CreationTime = $LastModified
                    
                    Write-DebugMsg "Remote file timestamp is: $LastModified"
                } catch {
                    Write-Warning "Could not parse Last-Modified header: $($Response.Headers["Last-Modified"])"
                }
            }
            
            $Response.Close()
            
            Write-Host "Successfully downloaded: $RemoteFilePath -> $LocalFilePath"
            
            # Return result with optional timestamp
            if ($IncludeTimestamp) {
                return @{
                    Success = $true
                    LocalPath = $LocalFilePath
                    LastModified = $LastModified
                }
            } else {
                return $true
            }
        } else {
            $Response.Close()
            Write-Warning "Download failed with status: $StatusCode"
            return $false
        }
    }
    catch [System.Net.WebException] {
        if ($_.Exception.Response.StatusCode -eq "NotFound") {
            Write-Host "Remote file does not exist: $RemoteFilePath"
            return $false
        } else {
            Write-Error "Failed to download file: $($_.Exception.Message)"
            return $false
        }
    }
    catch {
        Write-Error "Failed to download file: $($_.Exception.Message)"
        return $false
    }
}

# ClipSon Functions
function Show-ClipboardNotification {
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

function Get-RemoteClipboardFiles {
    param(
        [hashtable]$Connection
    )
    
    try {
        # Create PROPFIND request to list directory contents
        $WebRequest = [System.Net.WebRequest]::Create($Connection.WebDAVUrl + $NextcloudConfig.RemoteFolder)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "1")
        $WebRequest.ContentType = "application/xml"
        
        # Configure proxy
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        
        # PROPFIND body to request file names
        $PropFindBody = @"
<?xml version="1.0" encoding="utf-8" ?>
<D:propfind xmlns:D="DAV:">
    <D:prop>
        <D:displayname/>
        <D:getlastmodified/>
    </D:prop>
</D:propfind>
"@
        
        $RequestBytes = [System.Text.Encoding]::UTF8.GetBytes($PropFindBody)
        $WebRequest.ContentLength = $RequestBytes.Length
        
        $RequestStream = $WebRequest.GetRequestStream()
        $RequestStream.Write($RequestBytes, 0, $RequestBytes.Length)
        $RequestStream.Close()
        
        $Response = $WebRequest.GetResponse()
        $ResponseStream = $Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($ResponseStream)
        $ResponseContent = $Reader.ReadToEnd()
        $Reader.Close()
        $Response.Close()
        
        # Parse XML response
        $XmlDoc = New-Object System.Xml.XmlDocument
        $XmlDoc.LoadXml($ResponseContent)
        
        $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
        $NamespaceManager.AddNamespace("D", "DAV:")
        
        $files = @()
        $responseNodes = $XmlDoc.SelectNodes("//D:response", $NamespaceManager)
        
        foreach ($node in $responseNodes) {
            $displayNameNode = $node.SelectSingleNode(".//D:displayname", $NamespaceManager)
            $lastModifiedNode = $node.SelectSingleNode(".//D:getlastmodified", $NamespaceManager)
            
            if ($displayNameNode -and $displayNameNode.InnerText -match "^clipboard-.*\.(txt|png)$") {
                $lastModified = $null
                if ($lastModifiedNode -and $lastModifiedNode.InnerText) {
                    try {
                        $lastModifiedUTC = [DateTime]::ParseExact($lastModifiedNode.InnerText, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
                        $lastModified = $lastModifiedUTC.ToLocalTime()
                    } catch {
                        # Ignore parse errors
                    }
                }
                
                $files += @{
                    Name = $displayNameNode.InnerText
                    LastModified = $lastModified
                }
            }
        }
        
        return $files
    }
    catch {
        Write-Warning "Failed to discover remote clipboard files: $($_.Exception.Message)"
        return @()
    }
}

function Select-RemoteSyncFile {
    param(
        [hashtable]$Connection
    )
    
    Write-Host "Discovering remote clipboard files..." -ForegroundColor Green
    $remoteFiles = Get-RemoteClipboardFiles -Connection $Connection
    
    # Filter out files with current hostname (both text and image)
    $hostname = $env:COMPUTERNAME
    $myFiles = @("clipboard-$hostname.txt", "clipboard-$hostname.png")
    $filteredFiles = $remoteFiles | Where-Object { $_.Name -notin $myFiles }
    
    if ($filteredFiles.Count -eq 0) {
        Write-Host "No remote clipboard files from other machines found." -ForegroundColor Yellow
        Write-Host "Will only upload to: clipboard-$hostname.txt and clipboard-$hostname.png" -ForegroundColor Green
        return @()  # Return empty array instead of null
    }
    
    Write-Host "`nFound $($filteredFiles.Count) remote peer(s):" -ForegroundColor Green
    foreach ($file in $filteredFiles) {
        $timeStr = if ($file.LastModified) { $file.LastModified.ToString("yyyy-MM-dd HH:mm:ss") } else { "Unknown" }
        Write-Host "  - $($file.Name) (Modified: $timeStr)" -ForegroundColor Green
    }
    
    return $filteredFiles
}

function Initialize-RemotePeerTracking {
    param(
        [array]$PeerFiles
    )
    
    $global:remotePeerTimestamps = @{}
    
    foreach ($file in $PeerFiles) {
        $timestamp = if ($file.LastModified) { $file.LastModified } else { [DateTime]::MinValue }
        $global:remotePeerTimestamps[$file.Name] = $timestamp
        Write-DebugMsg "Initialized tracking for peer: $($file.Name) with timestamp: $timestamp"
    }
}

function Check-AllRemoteFilesForUpdates {
    param(
        [hashtable]$Connection
    )
    
    try {
        $currentTime = [DateTime]::Now
        if (($currentTime - $global:lastRemoteCheck) -lt $remoteCheckInterval) {
            return $null
        }
        
        $global:lastRemoteCheck = $currentTime
        
        # Get current list of remote files
        $remoteFiles = Get-RemoteClipboardFiles -Connection $Connection
        $hostname = $env:COMPUTERNAME
        $myFiles = @("clipboard-$hostname.txt", "clipboard-$hostname.png")
        $peerFiles = $remoteFiles | Where-Object { $_.Name -notin $myFiles }
        
        $mostRecentContent = $null
        $mostRecentTimestamp = [DateTime]::MinValue
        $mostRecentFilename = $null
        $tempImageFile = $null
        
        foreach ($fileInfo in $peerFiles) {
            $filename = $fileInfo.Name
            
            # Check if this is a new file or an updated file
            $isNewFile = -not $global:remotePeerTimestamps.ContainsKey($filename)
            
            if ($isNewFile) {
                $timestamp = if ($fileInfo.LastModified) { $fileInfo.LastModified } else { [DateTime]::MinValue }
                $global:remotePeerTimestamps[$filename] = [DateTime]::MinValue  # Set to MinValue so it will be processed as an update
                Write-Host "$(Get-Date -Format 'HH:mm:ss') - New remote peer discovered: $filename" -ForegroundColor Yellow
            }
            
            if (-not $fileInfo.LastModified) {
                continue
            }
            
            $remoteTimestamp = $fileInfo.LastModified
            $lastKnownTimestamp = $global:remotePeerTimestamps[$filename]
            
            # Check if this file has been updated (or is newly discovered)
            if ($remoteTimestamp -gt $lastKnownTimestamp) {
                if (-not $isNewFile) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Remote file updated: $filename" -ForegroundColor Cyan
                }
                
                # Download and check if it's the most recent
                $fileExtension = [System.IO.Path]::GetExtension($filename)
                $tempDownloadFile = ".\temp-remote-download-$($filename.Replace($fileExtension, ''))$fileExtension"
                $remotePath = $NextcloudConfig.RemoteFolder + $filename
                
                $downloadResult = Get-NextcloudFile -Connection $Connection -RemoteFilePath $remotePath -LocalFilePath $tempDownloadFile -IncludeTimestamp
                
                if ($downloadResult -and $downloadResult.Success) {
                    try {
                        $hasContent = $false
                        if ($fileExtension -eq '.txt') {
                            $fileContent = Get-Content $tempDownloadFile -Raw -Encoding UTF8
                            $hasContent = ($fileContent -and $fileContent.Trim() -ne "")
                        } elseif ($fileExtension -eq '.png') {
                            $fileContent = [System.IO.File]::ReadAllBytes($tempDownloadFile)
                            $hasContent = ($fileContent -and $fileContent.Length -gt 0)
                        }
                        
                        if ($hasContent -and $remoteTimestamp -gt $mostRecentTimestamp) {
                            $mostRecentContent = $fileContent
                            $mostRecentTimestamp = $remoteTimestamp
                            $mostRecentFilename = $filename
                            if ($fileExtension -eq '.png') {
                                $tempImageFile = $tempDownloadFile
                            }
                        }
                        
                        # Update timestamp regardless
                        $global:remotePeerTimestamps[$filename] = $remoteTimestamp
                        
                        # Only remove temp file if it's not the most recent image (we need it for hash)
                        if ($fileExtension -ne '.png' -or $filename -ne $mostRecentFilename) {
                            Remove-Item $tempDownloadFile -ErrorAction SilentlyContinue
                        }
                    }
                    catch {
                        Write-Warning "Error processing $filename`: $($_.Exception.Message)"
                        Remove-Item $tempDownloadFile -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        
        # Apply the most recent update if found
        if ($mostRecentContent) {
            if ($mostRecentFilename.EndsWith('.png')) {
                # Set clipboard image
                Add-Type -AssemblyName System.Drawing
                $memoryStream = New-Object System.IO.MemoryStream(,$mostRecentContent)
                $image = [System.Drawing.Image]::FromStream($memoryStream)
                [System.Windows.Forms.Clipboard]::SetImage($image)
                
                # Update last clipboard state to prevent re-capture IMMEDIATELY
                if ($tempImageFile -and (Test-Path $tempImageFile)) {
                    $script:lastClipboardImageHash = (Get-FileHash -Algorithm MD5 -LiteralPath $tempImageFile).Hash
                    Write-DebugMsg "Updated lastClipboardImageHash to prevent re-capture: $($script:lastClipboardImageHash)"
                    Remove-Item $tempImageFile -ErrorAction SilentlyContinue
                }
                
                $image.Dispose()
                $memoryStream.Dispose()
                
                # Show notification
                Show-ClipboardNotification -Title "ClipSon" -Message "Remote image update from $mostRecentFilename" -Icon "Info"
            } else {
                [System.Windows.Forms.Clipboard]::SetText($mostRecentContent)
                
                # Show notification
                $preview = if ($mostRecentContent.Length -gt 50) { $mostRecentContent.Substring(0, 50) + "..." } else { $mostRecentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Remote update from $mostRecentFilename`: $preview" -Icon "Info"
                
                # Update last clipboard state to prevent re-capture
                $script:lastClipboardContent = $mostRecentContent
            }
            
            return $mostRecentContent
        }
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Error checking remote files: $($_.Exception.Message)"
    }
    
    return $null
}

function Upload-ToWebDAV {
    param([string]$Content)
    
    try {
        Write-DebugMsg "Uploading content length: $($Content.Length) to file: $localSyncFile"
        
        # Ensure the content is not null or empty
        if ([string]::IsNullOrEmpty($Content)) {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Content is null or empty, skipping upload"
            return
        }
        
        # Use [System.IO.File]::WriteAllText instead of Out-File for better control
        [System.IO.File]::WriteAllText($localSyncFile, $Content, [System.Text.Encoding]::UTF8)
        
        # Verify the file was created correctly
        if (-not (Test-Path $localSyncFile)) {
            Write-Error "$(Get-Date -Format 'HH:mm:ss') - Failed to create local sync file: $localSyncFile"
            return
        }
        
        # Always upload to the hostname-based file (not the remote sync file we download from)
        $uploadResult = Send-FileToNextcloud -Connection $webdavConnection -LocalFilePath $localSyncFile -RemoteFilePath $localUploadPath
        
        if ($uploadResult) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Uploaded to WebDAV: $localUploadPath"
            
            # Update tracking to avoid downloading our own upload
            if (Test-Path $localSyncFile) {
                $global:lastClipboardFileModified = (Get-Item $localSyncFile).LastWriteTime
            }
        }
    }
    catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Error uploading to WebDAV: $($_.Exception.Message)"
    }
}

function Get-NextFileNumber {
    Write-DebugMsg "Get-NextFileNumber called, current counter: $global:fileCounter"
    
    $files = Get-ChildItem -Path $outputDir -File | Where-Object { $_.Name -match "clipboard_(text|image)_(\d{3})\.(txt|png)" }
    Write-DebugMsg "Found $($files.Count) existing clipboard files"
    
    # Initialize counter if this is the first run
    if ($global:fileCounter -eq 0) {
        if ($files.Count -eq 0) {
            $global:fileCounter = 1
        } else {
            $numbers = @()
            $files | ForEach-Object { 
                if ($_.Name -match "clipboard_(text|image)_(\d{3})\.(txt|png)") {
                    $numbers += [int]$matches[2]
                    Write-DebugMsg "Found file number: $($matches[2]) from file: $($_.Name)"
                }
            }
            if ($numbers.Count -gt 0) {
                $maxNumber = ($numbers | Measure-Object -Maximum).Maximum
                $global:fileCounter = $maxNumber + 1
                Write-DebugMsg "Max existing number: $maxNumber, setting counter to: $global:fileCounter"
            } else {
                $global:fileCounter = 1
            }
        }
    } else {
        $global:fileCounter++
        Write-DebugMsg "Incremented counter to: $global:fileCounter"
    }
    
    # If we've reached max entries, remove the oldest file
    if ($files.Count -ge $maxEntries) {
        $oldestFile = $files | Sort-Object CreationTime | Select-Object -First 1
        Write-DebugMsg "Removing oldest file: $($oldestFile.Name)"
        Remove-Item $oldestFile.FullName -Force
    }
    
    # Cycle back to 1 if counter exceeds 999 (3-digit limit)
    if ($global:fileCounter > 999) {
        Write-DebugMsg "Counter exceeded 999, resetting to 1"
        $global:fileCounter = 1
    }
    
    Write-DebugMsg "Returning file number: $global:fileCounter"
    return $global:fileCounter
}

# Main Script
# Get password if needed
Get-PasswordIfNeeded

# Create output directory if it doesn't exist
$outputDir = ".\clipboard-captures"
if (!(Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Rotation settings
$maxEntries = $Config.app.max_history

# Initialize WebDAV connection
Write-Host "Connecting to Nextcloud..." -ForegroundColor Green
$webdavConnection = New-NextcloudConnection -ServerUrl $NextcloudConfig.ServerUrl -Username $NextcloudConfig.Username -Password $NextcloudConfig.Password -ProxyUrl $NextcloudConfig.ProxyUrl -ProxyUsername $NextcloudConfig.ProxyUsername -ProxyPassword $NextcloudConfig.ProxyPassword -UseSystemProxy $NextcloudConfig.UseSystemProxy

if (-not (Test-NextcloudConnection -Connection $webdavConnection)) {
    Write-Error "Failed to connect to Nextcloud. Please check your configuration."
    exit 1
}

# Select remote sync file
$remotePeerFiles = Select-RemoteSyncFile -Connection $webdavConnection

# Local file paths
$hostname = $env:COMPUTERNAME
$localSyncFile = ".\clipboard-$hostname.txt"
$localUploadPath = $NextcloudConfig.RemoteFolder + "clipboard-$hostname.txt"

# Global variables
$global:fileCounter = 0
$global:lastClipboardFileModified = $null
$global:lastRemoteCheck = [DateTime]::MinValue
$global:remotePeerTimestamps = @{}  # Track timestamps for each remote peer
$remoteCheckInterval = [TimeSpan]::FromSeconds($Config.app.remote_check_interval_seconds)

# Initialize peer tracking
Initialize-RemotePeerTracking -PeerFiles $remotePeerFiles

Write-Host "ClipSon started. Press Ctrl+C to stop."
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

$lastClipboardContent = ""
$lastClipboardImageHash = ""

# Change these to script-level variables for proper access from functions
$script:lastClipboardContent = ""
$script:lastClipboardImageHash = ""

while ($true) {
    try {
        # Check all remote files for updates
        $remoteContent = Check-AllRemoteFilesForUpdates -Connection $webdavConnection
        if ($remoteContent) {
            # Update handled inside the function now
        }
        
        # Check if clipboard has text content
        if ([System.Windows.Forms.Clipboard]::ContainsText()) {
            $currentContent = [System.Windows.Forms.Clipboard]::GetText()
            
            # Only save if content has changed
            if ($currentContent -ne $script:lastClipboardContent -and $currentContent.Trim() -ne "") {
                Write-DebugMsg "About to get next file number"
                Write-DebugMsg "Pre-call fileCounter: '$global:fileCounter'"
                
                $fileNumber = Get-NextFileNumber
                Write-DebugMsg "Got file number: $fileNumber (type: $($fileNumber.GetType().Name))"
                Write-DebugMsg "Post-call fileCounter: '$global:fileCounter'"
                
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                Write-DebugMsg "Padded number: '$paddedNumber'"
                
                $filename = Join-Path $outputDir "clipboard_text_$paddedNumber.txt"
                Write-DebugMsg "Full filename: '$filename'"
                
                # Verify the filename is correct before writing
                if ($filename -notmatch "clipboard_text_\d{3}\.txt$") {
                    Write-Error "$(Get-Date -Format 'HH:mm:ss') - ERROR: Invalid filename generated: '$filename'"
                    continue
                }
                
                Write-DebugMsg "About to write content (length: $($currentContent.Length)) to file: '$filename'"
                
                # Use [System.IO.File]::WriteAllText for consistent file writing
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
                    continue
                }
                
                # Upload to WebDAV
                Upload-ToWebDAV -Content $currentContent
                
                # Show notification for local clipboard capture
                $preview = if ($currentContent.Length -gt 50) { $currentContent.Substring(0, 50) + "..." } else { $currentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Captured: $preview" -Icon "Info"
                
                $script:lastClipboardContent = $currentContent
            }
        }
        # Check if clipboard has image content
        elseif ([System.Windows.Forms.Clipboard]::ContainsImage()) {
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
                
                Write-DebugMsg "Current image hash: $currentImageHash, Last hash: $($script:lastClipboardImageHash)"
                
                if ($currentImageHash -eq $script:lastClipboardImageHash) {
                    Write-DebugMsg "Image hash matches previous - skipping save and upload"
                    $image.Dispose()
                    continue
                }
                
                # Hash is different, proceed with saving
                Write-DebugMsg "About to get next file number for image"
                $fileNumber = Get-NextFileNumber
                $paddedNumber = $fileNumber.ToString().PadLeft(3, '0')
                $filename = "$outputDir\clipboard_image_$paddedNumber.png"
                
                Write-DebugMsg "Creating image file: $filename"
                
                # Save the image bytes to file
                [System.IO.File]::WriteAllBytes($filename, $imageBytes)
                $script:lastClipboardImageHash = $currentImageHash
                
                # Upload image to WebDAV (use hostname-based file name)
                $remoteImagePath = $NextcloudConfig.RemoteFolder + "clipboard-$hostname.png"
                $uploadResult = Send-FileToNextcloud -Connection $webdavConnection -LocalFilePath $filename -RemoteFilePath $remoteImagePath
                if ($uploadResult) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Uploaded image to WebDAV: $remoteImagePath"
                }

                # Show notification for image capture
                Show-ClipboardNotification -Title "ClipSon" -Message "Image saved: $filename" -Icon "Info"
                
                $image.Dispose()
            }
        }
    }
    catch {
        Write-Host "Error accessing clipboard: $($_.Exception.Message)"
        if ($global:EnableDebugMessages) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: Exception details: $($_.Exception | Out-String)"
        }
    }
    
    # Wait 500ms before checking again
    Start-Sleep -Milliseconds 500
}
