# ClipSon - Advanced Clipboard Synchronization Tool (All-in-One)
# If you get execution policy errors, run one of these commands first:
# Option 1 (Recommended): Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Option 2 (Temporary): powershell.exe -ExecutionPolicy Bypass -File "clipson-all.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

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

function Compress-JsonFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )
    
    try {
        $jsonBytes = [System.IO.File]::ReadAllBytes($JsonFile)
        $originalSize = $jsonBytes.Length
        
        $fileStream = [System.IO.File]::Create($GzFile)
        $gzipStream = New-Object System.IO.Compression.GzipStream($fileStream, [System.IO.Compression.CompressionLevel]::Fastest)
        
        $gzipStream.Write($jsonBytes, 0, $jsonBytes.Length)
        $gzipStream.Close()
        $fileStream.Close()
        
        $compressedSize = (Get-Item $GzFile).Length
        if ($global:EnableDebugMessages) {
            $ratio = if ($originalSize -gt 0) { (1 - $compressedSize / $originalSize) * 100 } else { 0 }
            Write-DebugMsg "File compression $originalSize -> $compressedSize bytes ($([Math]::Round($ratio, 1))% saved)"
        }
        
        return $true
    }
    catch {
        Write-DebugMsg "File compression failed: $($_.Exception.Message)"
        return $false
    }
}

function Decompress-GzFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )
    
    try {
        $fileStream = [System.IO.File]::OpenRead($GzFile)
        $gzipStream = New-Object System.IO.Compression.GzipStream($fileStream, [System.IO.Compression.CompressionMode]::Decompress)
        $outputStream = [System.IO.File]::Create($JsonFile)
        
        $gzipStream.CopyTo($outputStream)
        
        $outputStream.Close()
        $gzipStream.Close()
        $fileStream.Close()
        
        return $true
    }
    catch {
        Write-DebugMsg "File decompression failed: $($_.Exception.Message)"
        return $false
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
            
            # Now looking for .json.gz files
            if ($displayNameNode -and $displayNameNode.InnerText -match "^clipboard-.*\.json\.gz$") {
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
    
    # Filter out files with current hostname (compressed json.gz now)
    $hostname = $env:COMPUTERNAME
    $myFile = "clipboard-$hostname.json.gz"
    
    # Use explicit filtering with a loop to ensure proper array handling
    $filteredFiles = @()
    foreach ($file in $remoteFiles) {
        if ($file.Name -ne $myFile) {
            $filteredFiles += $file
        }
    }
    
    if ($global:EnableDebugMessages) {
        Write-DebugMsg "Total remote files found: $($remoteFiles.Count)"
        Write-DebugMsg "My hostname: $hostname"
        Write-DebugMsg "My file: $myFile"
        Write-DebugMsg "All remote files: $($remoteFiles | ForEach-Object { $_.Name } | Out-String)"
        Write-DebugMsg "Files after filtering out my file:"
        foreach ($file in $filteredFiles) {
            Write-DebugMsg "  - $($file.Name)"
        }
        Write-DebugMsg "Filtered files count: $($filteredFiles.Count)"
    }
    
    if ($filteredFiles.Count -eq 0) {
        Write-Host "No remote clipboard files from other machines found." -ForegroundColor Yellow
        Write-Host "Will only upload to: $myFile" -ForegroundColor Green
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
        $myFile = "clipboard-$hostname.json.gz"
        
        # Use explicit filtering with a loop to ensure proper array handling
        $peerFiles = @()
        foreach ($file in $remoteFiles) {
            if ($file.Name -ne $myFile -and $file.Name -match "^clipboard-.*\.json\.gz$") {
                $peerFiles += $file
            }
        }
        
        if ($global:EnableDebugMessages) {
            Write-DebugMsg "Remote check - Total files: $($remoteFiles.Count), Peer files: $($peerFiles.Count), My file: $myFile"            
        }
        
        $mostRecentContent = $null
        $mostRecentTimestamp = [DateTime]::MinValue
        $mostRecentFilename = $null
        
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
                
                # Download and decompress
                $tempDownloadFileGz = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json.gz"
                $tempDownloadFileJson = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json"
                $remotePath = $NextcloudConfig.RemoteFolder + $filename
                
                $downloadResult = Get-NextcloudFile -Connection $Connection -RemoteFilePath $remotePath -LocalFilePath $tempDownloadFileGz -IncludeTimestamp
                
                if ($downloadResult -and $downloadResult.Success) {
                    try {
                        # Decompress the downloaded file
                        if (Decompress-GzFile -GzFile $tempDownloadFileGz -JsonFile $tempDownloadFileJson) {
                            $fileContent = Get-Content $tempDownloadFileJson -Raw -Encoding UTF8
                            $hasContent = ($fileContent -and $fileContent.Trim() -ne "")
                            
                            if ($hasContent -and $remoteTimestamp -gt $mostRecentTimestamp) {
                                $mostRecentContent = $fileContent
                                $mostRecentTimestamp = $remoteTimestamp
                                $mostRecentFilename = $filename
                            }
                        }
                        
                        # Update timestamp regardless
                        $global:remotePeerTimestamps[$filename] = $remoteTimestamp
                        
                        # Clean up temp files
                        Remove-Item $tempDownloadFileGz -ErrorAction SilentlyContinue
                        Remove-Item $tempDownloadFileJson -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Warning "Error processing $filename`: $($_.Exception.Message)"
                        Remove-Item $tempDownloadFileGz -ErrorAction SilentlyContinue
                        Remove-Item $tempDownloadFileJson -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        
        # Apply the most recent update if found
        if ($mostRecentContent) {
            # Set last clipboard content BEFORE setting clipboard to prevent re-capture
            Update-LastClipboardContentFromRemote -Content $mostRecentContent
            
            if (Set-ClipboardContentUnified -Content $mostRecentContent) {
                # Show notification
                $preview = if ($mostRecentContent.Length -gt 50) { $mostRecentContent.Substring(0, 50) + "..." } else { $mostRecentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Remote update from $mostRecentFilename`: $preview" -Icon "Info"
                
                # Set flag to skip clipboard monitoring for 2 seconds to prevent loopback
                $global:skipClipboardCheckUntil = [DateTime]::Now.AddSeconds(2)
                Write-DebugMsg "Set skipClipboardCheckUntil to $($global:skipClipboardCheckUntil) to prevent loopback"
            }
            
            return $mostRecentContent
        }
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Error checking remote files: $($_.Exception.Message)"
    }
    
    return $null
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

function Compress-JsonFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )
    
    try {
        $jsonBytes = [System.IO.File]::ReadAllBytes($JsonFile)
        $originalSize = $jsonBytes.Length
        
        $fileStream = [System.IO.File]::Create($GzFile)
        $gzipStream = New-Object System.IO.Compression.GzipStream($fileStream, [System.IO.Compression.CompressionLevel]::Fastest)
        
        $gzipStream.Write($jsonBytes, 0, $jsonBytes.Length)
        $gzipStream.Close()
        $fileStream.Close()
        
        $compressedSize = (Get-Item $GzFile).Length
        if ($global:EnableDebugMessages) {
            $ratio = if ($originalSize -gt 0) { (1 - $compressedSize / $originalSize) * 100 } else { 0 }
            Write-DebugMsg "File compression $originalSize -> $compressedSize bytes ($([Math]::Round($ratio, 1))% saved)"
        }
        
        return $true
    }
    catch {
        Write-DebugMsg "File compression failed: $($_.Exception.Message)"
        return $false
    }
}

function Decompress-GzFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )
    
    try {
        $fileStream = [System.IO.File]::OpenRead($GzFile)
        $gzipStream = New-Object System.IO.Compression.GzipStream($fileStream, [System.IO.Compression.CompressionMode]::Decompress)
        $outputStream = [System.IO.File]::Create($JsonFile)
        
        $gzipStream.CopyTo($outputStream)
        
        $outputStream.Close()
        $gzipStream.Close()
        $fileStream.Close()
        
        return $true
    }
    catch {
        Write-DebugMsg "File decompression failed: $($_.Exception.Message)"
        return $false
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
            
            # Now looking for .json.gz files
            if ($displayNameNode -and $displayNameNode.InnerText -match "^clipboard-.*\.json\.gz$") {
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
    
    # Filter out files with current hostname (compressed json.gz now)
    $hostname = $env:COMPUTERNAME
    $myFile = "clipboard-$hostname.json.gz"
    
    # Use explicit filtering with a loop to ensure proper array handling
    $filteredFiles = @()
    foreach ($file in $remoteFiles) {
        if ($file.Name -ne $myFile) {
            $filteredFiles += $file
        }
    }
    
    if ($global:EnableDebugMessages) {
        Write-DebugMsg "Total remote files found: $($remoteFiles.Count)"
        Write-DebugMsg "My hostname: $hostname"
        Write-DebugMsg "My file: $myFile"
        Write-DebugMsg "All remote files: $($remoteFiles | ForEach-Object { $_.Name } | Out-String)"
        Write-DebugMsg "Files after filtering out my file:"
        foreach ($file in $filteredFiles) {
            Write-DebugMsg "  - $($file.Name)"
        }
        Write-DebugMsg "Filtered files count: $($filteredFiles.Count)"
    }
    
    if ($filteredFiles.Count -eq 0) {
        Write-Host "No remote clipboard files from other machines found." -ForegroundColor Yellow
        Write-Host "Will only upload to: $myFile" -ForegroundColor Green
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
        $myFile = "clipboard-$hostname.json.gz"
        
        # Use explicit filtering with a loop to ensure proper array handling
        $peerFiles = @()
        foreach ($file in $remoteFiles) {
            if ($file.Name -ne $myFile -and $file.Name -match "^clipboard-.*\.json\.gz$") {
                $peerFiles += $file
            }
        }
        
        if ($global:EnableDebugMessages) {
            Write-DebugMsg "Remote check - Total files: $($remoteFiles.Count), Peer files: $($peerFiles.Count), My file: $myFile"            
        }
        
        $mostRecentContent = $null
        $mostRecentTimestamp = [DateTime]::MinValue
        $mostRecentFilename = $null
        
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
                
                # Download and decompress
                $tempDownloadFileGz = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json.gz"
                $tempDownloadFileJson = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json"
                $remotePath = $NextcloudConfig.RemoteFolder + $filename
                
                $downloadResult = Get-NextcloudFile -Connection $Connection -RemoteFilePath $remotePath -LocalFilePath $tempDownloadFileGz -IncludeTimestamp
                
                if ($downloadResult -and $downloadResult.Success) {
                    try {
                        # Decompress the downloaded file
                        if (Decompress-GzFile -GzFile $tempDownloadFileGz -JsonFile $tempDownloadFileJson) {
                            $fileContent = Get-Content $tempDownloadFileJson -Raw -Encoding UTF8
                            $hasContent = ($fileContent -and $fileContent.Trim() -ne "")
                            
                            if ($hasContent -and $remoteTimestamp -gt $mostRecentTimestamp) {
                                $mostRecentContent = $fileContent
                                $mostRecentTimestamp = $remoteTimestamp
                                $mostRecentFilename = $filename
                            }
                        }
                        
                        # Update timestamp regardless
                        $global:remotePeerTimestamps[$filename] = $remoteTimestamp
                        
                        # Clean up temp files
                        Remove-Item $tempDownloadFileGz -ErrorAction SilentlyContinue
                        Remove-Item $tempDownloadFileJson -ErrorAction SilentlyContinue
                    }
                    catch {
                        Write-Warning "Error processing $filename`: $($_.Exception.Message)"
                        Remove-Item $tempDownloadFileGz -ErrorAction SilentlyContinue
                        Remove-Item $tempDownloadFileJson -ErrorAction SilentlyContinue
                    }
                }
            }
        }
        
        # Apply the most recent update if found
        if ($mostRecentContent) {
            # Set last clipboard content BEFORE setting clipboard to prevent re-capture
            Update-LastClipboardContentFromRemote -Content $mostRecentContent
            
            if (Set-ClipboardContentUnified -Content $mostRecentContent) {
                # Show notification
                $preview = if ($mostRecentContent.Length -gt 50) { $mostRecentContent.Substring(0, 50) + "..." } else { $mostRecentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Remote update from $mostRecentFilename`: $preview" -Icon "Info"
                
                # Set flag to skip clipboard monitoring for 2 seconds to prevent loopback
                $global:skipClipboardCheckUntil = [DateTime]::Now.AddSeconds(2)
                Write-DebugMsg "Set skipClipboardCheckUntil to $($global:skipClipboardCheckUntil) to prevent loopback"
            }
            
            return $mostRecentContent
        }
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Error checking remote files: $($_.Exception.Message)"
    }
    
    return $null
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
$localSyncFile = ".\clipboard-$hostname.json"
$localSyncFileGz = ".\clipboard-$hostname.json.gz"
$localUploadPath = $NextcloudConfig.RemoteFolder + "clipboard-$hostname.json.gz"

# Global variables
$global:fileCounter = 0
$global:lastClipboardFileModified = $null
$global:lastRemoteCheck = [DateTime]::MinValue
$global:remotePeerTimestamps = @{}  # Track timestamps for each remote peer
$global:skipClipboardCheckUntil = [DateTime]::MinValue  # Add flag to temporarily skip clipboard monitoring
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
        
        # Check if we should skip clipboard monitoring (after setting remote content)
        $currentTime = [DateTime]::Now
        if ($currentTime -lt $global:skipClipboardCheckUntil) {
            if ($global:EnableDebugMessages -and ($currentTime.Second % 2 -eq 0)) {  # Only show every 2 seconds to avoid spam
                Write-DebugMsg "Skipping clipboard check until $($global:skipClipboardCheckUntil) to prevent loopback"
            }
            Start-Sleep -Milliseconds 500
            continue
        }
        
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
                    continue
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
                    continue
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
        Write-Host "Error accessing clipboard: $($_.Exception.Message)"
        if ($global:EnableDebugMessages) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: Exception details: $($_.Exception | Out-String)"
        }
    }
    
    # Wait 500ms before checking again
    Start-Sleep -Milliseconds 500
}

function Set-ClipboardContentUnified {
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
        # Update hash tracking immediately to prevent re-capture
        $md5 = [System.Security.Cryptography.MD5]::Create()
        $hashBytes = $md5.ComputeHash($ImageData)
        $script:lastClipboardImageHash = [System.BitConverter]::ToString($hashBytes) -replace '-'
        $md5.Dispose()
        
        Write-DebugMsg "Updated lastClipboardImageHash to prevent re-capture: $($script:lastClipboardImageHash)"
        
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

function Set-ClipboardMultipleFormats {
    param([PSObject]$FormatData)
    
    try {
        Write-DebugMsg "Setting $($FormatData.PSObject.Properties.Count) clipboard formats"
        
        $dataObject = New-Object System.Windows.Forms.DataObject
        $successCount = 0
        
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
                    
                    if ($windowsFormat) {
                        $dataObject.SetText($content, $windowsFormat)
                        $successCount++
                        Write-DebugMsg "Successfully set format $formatName"
                    }
                }
                catch {
                    Write-DebugMsg "Error setting format $formatName`: $($_.Exception.Message)"
                }
            }
        }
        
        # Set the data object to clipboard
        if ($successCount -gt 0) {
            [System.Windows.Forms.Clipboard]::SetDataObject($dataObject)
            Write-DebugMsg "Successfully set $successCount formats to clipboard"
            return $true
        }
        
        return $false
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

function Upload-ToWebDAV {
    param([string]$Content)
    
    try {
        Write-DebugMsg "Uploading content length: $($Content.Length) to file: $localSyncFile"
        
        # Ensure the content is not null or empty
        if ([string]::IsNullOrEmpty($Content)) {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Content is null or empty, skipping upload"
            return
        }
        
        # Save JSON content to local file
        [System.IO.File]::WriteAllText($localSyncFile, $Content, [System.Text.Encoding]::UTF8)
        
        # Verify the file was created correctly
        if (-not (Test-Path $localSyncFile)) {
            Write-Error "$(Get-Date -Format 'HH:mm:ss') - Failed to create local sync file: $localSyncFile"
            return
        }
        
        # Compress the JSON file
        if (Compress-JsonFile -JsonFile $localSyncFile -GzFile $localSyncFileGz) {
            # Upload compressed file
            $uploadResult = Send-FileToNextcloud -Connection $webdavConnection -LocalFilePath $localSyncFileGz -RemoteFilePath $localUploadPath
            
            if ($uploadResult) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss') - Uploaded compressed to WebDAV: $localUploadPath"
                
                # Update tracking to avoid downloading our own upload
                if (Test-Path $localSyncFile) {
                    $global:lastClipboardFileModified = (Get-Item $localSyncFile).LastWriteTime
                }
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

function Get-ClipboardFormats {
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

function Test-ClipboardRichText {
    try {
        $formats = Get-ClipboardFormats
        $richTextFormats = @("HTML", "RTF", "Files")
        return ($formats | Where-Object { $_ -in $richTextFormats }).Count -gt 0
    }
    catch {
        return $false
    }
}

function Save-ClipboardRichContentJson {
    param([hashtable]$FormatData)
    
    try {
        if ($FormatData.Count -eq 0) {
            return $false
        }
        
        Write-DebugMsg "Available clipboard formats: $($FormatData.Keys -join ', ')"
        
        $fileNumber = Get-NextFileNumber
        $baseFilename = "$outputDir\clipboard_rich_$($fileNumber.ToString().PadLeft(3, '0'))"
        
        $savedFiles = @()
        
        # Save each available format to individual files
        foreach ($format in $FormatData.Keys) {
            $content = $FormatData[$format]
            $extension = switch ($format) {
                'text/html' { '.html' }
                'text/rtf' { '.rtf' }
                'text/plain' { '.txt' }
                default { '.txt' }
            }
            
            $filename = "$baseFilename$extension"
            [System.IO.File]::WriteAllText($filename, $content, [System.Text.Encoding]::UTF8)
            $savedFiles += $filename
            Write-DebugMsg "Saved $format to $filename"
        }
        
        if ($savedFiles.Count -gt 0) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - Rich content saved: $($savedFiles.Count) unique formats"
            foreach ($file in $savedFiles) {
                Write-Host "  - $file"
            }
            
            # Create multi-format upload content
            $uploadContent = @{
                type = "MULTI_FORMAT_CLIPBOARD"
                formats = $FormatData
            }
            $uploadJson = ConvertTo-Json $uploadContent -Depth 10
            
            # Upload to WebDAV
            Upload-ToWebDAV -Content $uploadJson
            
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
