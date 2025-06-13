# Module for miscellaneous utility functions

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
        if ($global:Config.app.debug_enabled) {
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

function global:Get-NextFileNumber {    
    
    # Use global output directory variable
    if (-not $global:outputDir) {
        Write-Warning "Output directory is not defined"
        $global:outputDir = ".\clipboard-captures"
    }
    
    $files = Get-ChildItem -Path $global:outputDir -File -ErrorAction SilentlyContinue | 
             Where-Object { $_.Name -match "clipboard_(text|image)_(\d{3})\.(txt|png)" }
        
    
    # Initialize counter if this is the first run
    if ($global:fileCounter -eq 0) {
        if ($files.Count -eq 0) {
            $global:fileCounter = 1
        } else {
            $numbers = @()
            $files | ForEach-Object { 
                if ($_.Name -match "clipboard_(text|image)_(\d{3})\.(txt|png)") {
                    $numbers += [int]$matches[2]             
                }
            }
            if ($numbers.Count -gt 0) {
                $maxNumber = ($numbers | Measure-Object -Maximum).Maximum
                $global:fileCounter = $maxNumber + 1
            } else {
                $global:fileCounter = 1
            }
        }
    } else {
        $global:fileCounter++
    }
    
    # If we've reached max entries, remove the oldest file
    $maxEntries = if ($global:maxEntries) { $global:maxEntries } else { 100 }
    if ($files -and $files.Count -ge $maxEntries) {
        $oldestFile = $files | Sort-Object CreationTime | Select-Object -First 1
        if ($oldestFile) {
            Write-DebugMsg "Removing oldest file: $($oldestFile.Name)"
            Remove-Item $oldestFile.FullName -Force
        } else {
            Write-DebugMsg "Removing oldest file: (none found)"
        }
    }
    
    # Cycle back to 1 if counter exceeds 999 (3-digit limit)
    if ($global:fileCounter -gt 999) {
        Write-DebugMsg "Counter exceeded 999, resetting to 1"
        $global:fileCounter = 1
    }
    
    return $global:fileCounter
}
