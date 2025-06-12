# Module for Nextcloud WebDAV operations

Add-Type -AssemblyName System.IO.Compression

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
        [hashtable]$Connection,
        [string]$RemoteFolder
    )
    
    try {
        $WebRequest = [System.Net.WebRequest]::Create($Connection.WebDAVUrl + $RemoteFolder)
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
        [hashtable]$Connection,
        [string]$RemoteFolder
    )
    
    Write-Host "Discovering remote clipboard files..." -ForegroundColor Green
    $remoteFiles = Get-RemoteClipboardFiles -Connection $Connection -RemoteFolder $RemoteFolder
    
    $hostname = $env:COMPUTERNAME
    $myFile = "clipboard-$hostname.json.gz"
    
    $filteredFiles = @()
    foreach ($file in $remoteFiles) {
        if ($file.Name -ne $myFile) {
            $filteredFiles += $file
        }
    }
    
    if ($global:Config.app.debug_enabled) {
        Write-DebugMsg "Total remote files found: $($remoteFiles.Count)"
        Write-DebugMsg "My hostname: $hostname"
        Write-DebugMsg "My file: $myFile"
        Write-DebugMsg "Files after filtering out my file: $($filteredFiles.Count)"
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
        [hashtable]$Connection,
        [string]$RemoteFolder,
        [TimeSpan]$CheckInterval
    )
    
    try {
        $currentTime = [DateTime]::Now
        if (($currentTime - $global:lastRemoteCheck) -lt $CheckInterval) {
            return $null
        }
        
        $global:lastRemoteCheck = $currentTime
        
        $remoteFiles = Get-RemoteClipboardFiles -Connection $Connection -RemoteFolder $RemoteFolder
        $hostname = $env:COMPUTERNAME
        $myFile = "clipboard-$hostname.json.gz"
        
        $peerFiles = @()
        foreach ($file in $remoteFiles) {
            if ($file.Name -ne $myFile -and $file.Name -match "^clipboard-.*\.json\.gz$") {
                $peerFiles += $file
            }
        }
        
        if ($global:Config.app.debug_enabled) {
            Write-DebugMsg "Remote check - Total files: $($remoteFiles.Count), Peer files: $($peerFiles.Count)"
        }
        
        $mostRecentContent = $null
        $mostRecentTimestamp = [DateTime]::MinValue
        $mostRecentFilename = $null
        
        foreach ($fileInfo in $peerFiles) {
            $filename = $fileInfo.Name
            
            $isNewFile = -not $global:remotePeerTimestamps.ContainsKey($filename)
            
            if ($isNewFile) {
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
                $remotePath = $RemoteFolder + $filename
                
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
            # Return the content along with filename info for the caller to handle
            return @{
                Content = $mostRecentContent
                Filename = $mostRecentFilename
                Timestamp = $mostRecentTimestamp
            }
        }
    }
    catch {
        Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Error checking remote files: $($_.Exception.Message)"
    }
    
    return $null
}

function global:Upload-ToWebDAV {
    param(
        [string]$Content,
        [hashtable]$Connection,
        [string]$LocalSyncFile,
        [string]$LocalSyncFileGz,
        [string]$RemoteFilePath
    )
    
    try {
        Write-DebugMsg "Uploading content length: $($Content.Length) to file: $LocalSyncFile"
        
        if ([string]::IsNullOrEmpty($Content)) {
            Write-Warning "$(Get-Date -Format 'HH:mm:ss') - Content is null or empty, skipping upload"
            return
        }
        
        [System.IO.File]::WriteAllText($LocalSyncFile, $Content, [System.Text.Encoding]::UTF8)
        
        if (-not (Test-Path $LocalSyncFile)) {
            Write-Error "$(Get-Date -Format 'HH:mm:ss') - Failed to create local sync file: $LocalSyncFile"
            return
        }
        
        if (Compress-JsonFile -JsonFile $LocalSyncFile -GzFile $LocalSyncFileGz) {
            $uploadResult = Send-FileToNextcloud -Connection $Connection -LocalFilePath $LocalSyncFileGz -RemoteFilePath $RemoteFilePath
            
            if ($uploadResult) {
                Write-Host "$(Get-Date -Format 'HH:mm:ss') - Uploaded compressed to WebDAV: $RemoteFilePath"
                
                if (Test-Path $LocalSyncFile) {
                    $global:lastClipboardFileModified = (Get-Item $LocalSyncFile).LastWriteTime
                }
                
                return $true
            }
        }
        
        return $false
    }
    catch {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - Error uploading to WebDAV: $($_.Exception.Message)"
        return $false
    }
}
