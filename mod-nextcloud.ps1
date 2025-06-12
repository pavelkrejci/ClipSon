# ClipSon - Advanced Clipboard Synchronization Tool (All-in-One)
# If you get execution policy errors, run one of these commands first:
# Option 1 (Recommended): Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
# Option 2 (Temporary): powershell.exe -ExecutionPolicy Bypass -File "clipson-all.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.IO.Compression

# Configuration (now loaded from config.json)
$NextcloudConfig = @{
    ServerUrl    = $global:Config.nextcloud.server_url      # Replace with your Nextcloud server URL
    Username     = $global:Config.nextcloud.username        # Replace with your username
    Password     = $global:Config.nextcloud.password        # Empty password will prompt for input
    RemoteFolder = $global:Config.nextcloud.remote_folder   # Folder in Nextcloud where files will be uploaded
    # Proxy Configuration (optional)
    ProxyUrl     = $global:Config.proxy.url                 # e.g., "http://proxy.company.com:8080" or "" for no proxy
    ProxyUsername= $global:Config.proxy.username            # Proxy username (if required)
    ProxyPassword= $global:Config.proxy.password            # Proxy password (if required)
    UseSystemProxy= $global:Config.proxy.use_system_proxy   # Use system proxy settings if true, manual proxy if false
}

# Normalize RemoteFolder to ensure trailing slash
if (-not $NextcloudConfig.RemoteFolder.EndsWith('/')) {
    $NextcloudConfig.RemoteFolder += '/'
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
    $ServerUrl = $ServerUrl.TrimEnd('/')
    $WebDAVUrl = "$ServerUrl/remote.php/dav/files/$Username/"
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
    $ProxyConfig = $null
    if (-not $UseSystemProxy -and $ProxyUrl) {
        $ProxyConfig = @{
            Url = $ProxyUrl
            Username = $ProxyUsername
            Password = $ProxyPassword
        }
    }
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
        $WebRequest.Proxy = [System.Net.WebRequest]::GetSystemWebProxy()
        $WebRequest.Proxy.Credentials = [System.Net.CredentialCache]::DefaultCredentials
    }
    elseif ($Connection.ProxyConfig -and $Connection.ProxyConfig.Url) {
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
        if (-not (Test-Path $LocalFilePath)) {
            throw "Local file not found: $LocalFilePath"
        }
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $UploadUrl = $Connection.WebDAVUrl + $RemoteFilePath
        $FileContent = [System.IO.File]::ReadAllBytes($LocalFilePath)
        $WebRequest = [System.Net.WebRequest]::Create($UploadUrl)
        $WebRequest.Method = "PUT"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.ContentLength = $FileContent.Length
        $WebRequest.ContentType = "application/octet-stream"
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        $RequestStream = $WebRequest.GetRequestStream()
        $RequestStream.Write($FileContent, 0, $FileContent.Length)
        $RequestStream.Close()
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
        $WebRequest = [System.Net.WebRequest]::Create($Connection.WebDAVUrl)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "0")
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        $Response = $WebRequest.GetResponse()
        $StatusCode = [int]$Response.StatusCode
        $Response.Close()
        if ($StatusCode -eq 207) {
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
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $PropFindUrl = $Connection.WebDAVUrl + $RemoteFilePath
        $WebRequest = [System.Net.WebRequest]::Create($PropFindUrl)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "0")
        $WebRequest.ContentType = "application/xml"
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
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
        $RequestStream = $WebRequest.GetRequestStream()
        $RequestStream.Write($RequestBytes, 0, $RequestBytes.Length)
        $RequestStream.Close()
        $Response = $WebRequest.GetResponse()
        $ResponseStream = $Response.GetResponseStream()
        $Reader = New-Object System.IO.StreamReader($ResponseStream)
        $ResponseContent = $Reader.ReadToEnd()
        $Reader.Close()
        $Response.Close()
        $XmlDoc = New-Object System.Xml.XmlDocument
        $XmlDoc.LoadXml($ResponseContent)
        $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
        $NamespaceManager.AddNamespace("D", "DAV:")
        $LastModifiedNode = $XmlDoc.SelectSingleNode("//D:getlastmodified", $NamespaceManager)
        if ($LastModifiedNode -and $LastModifiedNode.InnerText) {
            $LastModifiedUTC = [DateTime]::ParseExact($LastModifiedNode.InnerText, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
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
        $RemoteFilePath = $RemoteFilePath.TrimStart('/')
        $DownloadUrl = $Connection.WebDAVUrl + $RemoteFilePath
        $WebRequest = [System.Net.WebRequest]::Create($DownloadUrl)
        $WebRequest.Method = "GET"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
        $Response = $WebRequest.GetResponse()
        $StatusCode = [int]$Response.StatusCode
        if ($StatusCode -eq 200) {
            $LocalDir = Split-Path $LocalFilePath -Parent
            if (!(Test-Path $LocalDir)) {
                New-Item -ItemType Directory -Path $LocalDir -Force | Out-Null
            }
            $ResponseStream = $Response.GetResponseStream()
            $FileStream = [System.IO.File]::Create($LocalFilePath)
            $ResponseStream.CopyTo($FileStream)
            $FileStream.Close()
            $ResponseStream.Close()
            $LastModified = $null
            if ($IncludeTimestamp -and $Response.Headers["Last-Modified"]) {
                try {
                    $LastModifiedUTC = [DateTime]::ParseExact($Response.Headers["Last-Modified"], "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
                    $LastModified = $LastModifiedUTC.ToLocalTime()
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

function Get-RemoteClipboardFiles {
    param(
        [hashtable]$Connection
    )
    try {
        $WebRequest = [System.Net.WebRequest]::Create($Connection.WebDAVUrl + $NextcloudConfig.RemoteFolder)
        $WebRequest.Method = "PROPFIND"
        $WebRequest.Credentials = $Connection.Credential.GetNetworkCredential()
        $WebRequest.Headers.Add("Depth", "1")
        $WebRequest.ContentType = "application/xml"
        Set-WebRequestProxy -WebRequest $WebRequest -Connection $Connection
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
        $XmlDoc = New-Object System.Xml.XmlDocument
        $XmlDoc.LoadXml($ResponseContent)
        $NamespaceManager = New-Object System.Xml.XmlNamespaceManager($XmlDoc.NameTable)
        $NamespaceManager.AddNamespace("D", "DAV:")
        $files = @()
        $responseNodes = $XmlDoc.SelectNodes("//D:response", $NamespaceManager)
        foreach ($node in $responseNodes) {
            $displayNameNode = $node.SelectSingleNode(".//D:displayname", $NamespaceManager)
            $lastModifiedNode = $node.SelectSingleNode(".//D:getlastmodified", $NamespaceManager)
            if ($displayNameNode -and $displayNameNode.InnerText -match "^clipboard-.*\.json\.gz$") {
                $lastModified = $null
                if ($lastModifiedNode -and $lastModifiedNode.InnerText) {
                    try {
                        $lastModifiedUTC = [DateTime]::ParseExact($lastModifiedNode.InnerText, "ddd, dd MMM yyyy HH:mm:ss 'GMT'", [System.Globalization.CultureInfo]::InvariantCulture)
                        $lastModified = $lastModifiedUTC.ToLocalTime()
                    } catch { }
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
    $hostname = $env:COMPUTERNAME
    $myFile = "clipboard-$hostname.json.gz"
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
        return @()
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
        $remoteFiles = Get-RemoteClipboardFiles -Connection $Connection
        $hostname = $env:COMPUTERNAME
        $myFile = "clipboard-$hostname.json.gz"
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
            $isNewFile = -not $global:remotePeerTimestamps.ContainsKey($filename)
            if ($isNewFile) {
                $timestamp = if ($fileInfo.LastModified) { $fileInfo.LastModified } else { [DateTime]::MinValue }
                $global:remotePeerTimestamps[$filename] = [DateTime]::MinValue
                Write-Host "$(Get-Date -Format 'HH:mm:ss') - New remote peer discovered: $filename" -ForegroundColor Yellow
            }
            if (-not $fileInfo.LastModified) { continue }
            $remoteTimestamp = $fileInfo.LastModified
            $lastKnownTimestamp = $global:remotePeerTimestamps[$filename]
            if ($remoteTimestamp -gt $lastKnownTimestamp) {
                if (-not $isNewFile) {
                    Write-Host "$(Get-Date -Format 'HH:mm:ss') - Remote file updated: $filename" -ForegroundColor Cyan
                }
                $tempDownloadFileGz = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json.gz"
                $tempDownloadFileJson = ".\temp-remote-download-$($filename.Replace('.json.gz', '')).json"
                $remotePath = $NextcloudConfig.RemoteFolder + $filename
                $downloadResult = Get-NextcloudFile -Connection $Connection -RemoteFilePath $remotePath -LocalFilePath $tempDownloadFileGz -IncludeTimestamp
                if ($downloadResult -and $downloadResult.Success) {
                    try {
                        if (Decompress-GzFile -GzFile $tempDownloadFileGz -JsonFile $tempDownloadFileJson) {
                            $fileContent = Get-Content $tempDownloadFileJson -Raw -Encoding UTF8
                            $hasContent = ($fileContent -and $fileContent.Trim() -ne "")
                            if ($hasContent -and $remoteTimestamp -gt $mostRecentTimestamp) {
                                $mostRecentContent = $fileContent
                                $mostRecentTimestamp = $remoteTimestamp
                                $mostRecentFilename = $filename
                            }
                        }
                        $global:remotePeerTimestamps[$filename] = $remoteTimestamp
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
        if ($mostRecentContent) {
            Update-LastClipboardContentFromRemote -Content $mostRecentContent
            if (Set-ClipboardContentUnified -Content $mostRecentContent) {
                $preview = if ($mostRecentContent.Length -gt 50) { $mostRecentContent.Substring(0, 50) + "..." } else { $mostRecentContent }
                Show-ClipboardNotification -Title "ClipSon" -Message "Remote update from $mostRecentFilename`: $preview" -Icon "Info"
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

# Change these to script-level variables for proper access from functions
$script:lastClipboardContent = ""
$script:lastClipboardImageHash = ""

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
        Write-Host "Error accessing clipboard: $($_.Exception.Message)"
        if ($global:EnableDebugMessages) {
            Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: Exception details: $($_.Exception | Out-String)"
        }
    }
}

<# 

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
} #>