function Get-Configuration {
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

function Write-DebugMsg {
    param([string]$Message)
    if ($global:Config.app.debug_enabled) {
        Write-Host "$(Get-Date -Format 'HH:mm:ss') - DEBUG: $Message"
    }
}

function Get-PasswordIfNeeded {
    param($NextcloudConfig)
    if ([string]::IsNullOrWhiteSpace($NextcloudConfig.Password)) {
        Write-Host "No password configured for user: $($NextcloudConfig.Username)" -ForegroundColor Yellow
        $SecurePassword = Read-Host "Please enter your Nextcloud password" -AsSecureString
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

# ...add any other helpers as needed...
