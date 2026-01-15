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

function Get-ArchiveExtension {
    $use7z = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_7z_encryption") {
        $use7z = [bool]$global:Config.app.use_7z_encryption
    }

    return $(if ($use7z) { ".cpsn" } else { ".gz" })
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

function Get-7zExecutablePath {
    # Prefer known install locations, then fall back to PATH
    $isWindows = $env:OS -eq 'Windows_NT'

    if ($isWindows) {
        $preferred = "C:\Program Files\7-Zip\7z.exe"
        if (Test-Path $preferred) { return $preferred }
    } else {
        $preferred = "/usr/bin/7z"
        if (Test-Path $preferred) { return $preferred }
    }

    $cmd = Get-Command 7z -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source) { return $cmd.Source }

    throw "7z executable not found (expected $preferred or 7z in PATH)"
}

function Test-HasCpsnHeader {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )
    try {
        if (-not (Test-Path $Path)) { return $false }
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            if ($fs.Length -lt 4) { return $false }
            $buf = New-Object byte[] 4
            $read = $fs.Read($buf, 0, 4)
            if ($read -ne 4) { return $false }
            $sig = [System.Text.Encoding]::ASCII.GetString($buf)
            return ($sig -eq "CPSN")
        }
        finally {
            $fs.Close()
        }
    }
    catch {
        return $false
    }
}

function Add-CpsnHeaderToFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    if (Test-HasCpsnHeader -Path $Path) {
        if ($global:Config -and $global:Config.app.debug_enabled) {
            Write-DebugMsg "CPSN header already present: $Path"
        }
        return
    }

    $dir = Split-Path -Parent $Path
    $tmp = Join-Path $dir ("." + ([System.IO.Path]::GetFileName($Path)) + "." + [System.Guid]::NewGuid().ToString("N") + ".cpsn")

    $prefix = [System.Text.Encoding]::ASCII.GetBytes("CPSN")
    $inStream = [System.IO.File]::OpenRead($Path)
    try {
        $outStream = [System.IO.File]::Create($tmp)
        try {
            $outStream.Write($prefix, 0, $prefix.Length)
            $inStream.CopyTo($outStream)
        }
        finally {
            $outStream.Close()
        }
    }
    finally {
        $inStream.Close()
    }

    Move-Item -LiteralPath $tmp -Destination $Path -Force

    if ($global:Config -and $global:Config.app.debug_enabled) {
        Write-DebugMsg "Prepended CPSN header: $Path"
    }
}

function Strip-CpsnHeaderToFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$InputPath,
        [Parameter(Mandatory=$true)]
        [string]$OutputPath
    )

    $inStream = [System.IO.File]::OpenRead($InputPath)
    try {
        $buf = New-Object byte[] 4
        $read = $inStream.Read($buf, 0, 4)
        $sig = if ($read -eq 4) { [System.Text.Encoding]::ASCII.GetString($buf) } else { "" }

        $outStream = [System.IO.File]::Create($OutputPath)
        try {
            if ($sig -eq "CPSN") {
                # Copy remaining bytes (after header)
                $inStream.CopyTo($outStream)
                if ($global:Config -and $global:Config.app.debug_enabled) {
                    Write-DebugMsg "Detected CPSN header; stripped to: $OutputPath"
                }
            } else {
                # Not CPSN: write back the bytes we already read then rest
                if ($read -gt 0) {
                    $outStream.Write($buf, 0, $read)
                }
                $inStream.CopyTo($outStream)
                if ($global:Config -and $global:Config.app.debug_enabled) {
                    Write-DebugMsg "No CPSN header; copied archive to: $OutputPath"
                }
            }
        }
        finally {
            $outStream.Close()
        }
    }
    finally {
        $inStream.Close()
    }
}

function Get-ArchiveFormat {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Path
    )

    try {
        if (-not (Test-Path $Path)) { return "unknown" }
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $buf = New-Object byte[] 6
            $read = $fs.Read($buf, 0, 6)
            if ($read -ge 4) {
                $sig = [System.Text.Encoding]::ASCII.GetString($buf, 0, [Math]::Min(4, $read))
                if ($sig -eq "CPSN") { return "cpsn" }
            }
            if ($read -ge 2 -and $buf[0] -eq 0x1f -and $buf[1] -eq 0x8b) { return "gzip" }
            if ($read -eq 6 -and $buf[0] -eq 0x37 -and $buf[1] -eq 0x7a -and $buf[2] -eq 0xbc -and $buf[3] -eq 0xaf -and $buf[4] -eq 0x27 -and $buf[5] -eq 0x1c) {
                return "7z"
            }
            return "unknown"
        }
        finally {
            $fs.Close()
        }
    }
    catch {
        return "unknown"
    }
}

function Compress-JsonFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )

    $use7z = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_7z_encryption") {
        $use7z = [bool]$global:Config.app.use_7z_encryption
    }

    if ($use7z) {
        $encOk = Compress-JsonFileWith7z -JsonFile $JsonFile -GzFile $GzFile
        if ($encOk) { return $true }
        Write-DebugMsg "7z compression failed or unavailable; falling back to gzip"
    }

    return Compress-JsonFileWithGzip -JsonFile $JsonFile -GzFile $GzFile
}

function Compress-JsonFileWith7z {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )
    try {
        $sevenZip = Get-7zExecutablePath
        $password = $global:Config.nextcloud.password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw "Nextcloud password is empty; cannot encrypt"
        }

        $originalSize = (Get-Item $JsonFile).Length

        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("clipson-7z-" + [System.Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        try {
            # Put a stable name inside the archive.
            $payloadPath = Join-Path $tempDir "payload.json"
            Copy-Item -LiteralPath $JsonFile -Destination $payloadPath -Force

            if (Test-Path $GzFile) {
                Remove-Item -LiteralPath $GzFile -Force -ErrorAction SilentlyContinue
            }

            # Note: keep legacy .gz naming, but output is a 7z-encrypted archive.
            $args = @(
                "a",
                "-t7z",
                "-mhe=on",
                "-y",
                "-bd",
                ("-p" + $password),
                $GzFile,
                $payloadPath
            )

            $output = & $sevenZip @args 2>&1
            if ($LASTEXITCODE -ne 0) {
                if ($global:Config.app.debug_enabled) {
                    Write-DebugMsg "7z encryption failed (exit $LASTEXITCODE): $output"
                } else {
                    Write-DebugMsg "7z encryption failed (exit $LASTEXITCODE)"
                }
                return $false
            }

            # Prefix the produced archive bytes with CPSN.
            Add-CpsnHeaderToFile -Path $GzFile

            $encryptedSize = (Get-Item $GzFile).Length
            if ($global:Config.app.debug_enabled) {
                $ratio = if ($originalSize -gt 0) { (1 - $encryptedSize / $originalSize) * 100 } else { 0 }
                Write-DebugMsg "File encryption $originalSize -> $encryptedSize bytes ($([Math]::Round($ratio, 1))% smaller)"
            }

            return $true
        }
        finally {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-DebugMsg "File encryption failed: $($_.Exception.Message)"
        return $false
    }
}

function Compress-JsonFileWithGzip {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )
    $tempPath = "$GzFile.$([System.Guid]::NewGuid().ToString('N')).tmp"
    try {
        if (Test-Path $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }

        $input = [System.IO.File]::OpenRead($JsonFile)
        try {
            $output = [System.IO.File]::Create($tempPath)
            try {
                $gzip = New-Object System.IO.Compression.GzipStream($output, [System.IO.Compression.CompressionLevel]::Optimal, $true)
                try {
                    $input.CopyTo($gzip)
                }
                finally {
                    $gzip.Dispose()
                }
            }
            finally {
                $output.Dispose()
            }
        }
        finally {
            $input.Dispose()
        }

        Move-Item -LiteralPath $tempPath -Destination $GzFile -Force

        if ($global:Config -and $global:Config.app.debug_enabled) {
            $originalSize = (Get-Item $JsonFile).Length
            $compressedSize = (Get-Item $GzFile).Length
            $ratio = if ($originalSize -gt 0) { (1 - $compressedSize / $originalSize) * 100 } else { 0 }
            Write-DebugMsg "Gzip compression $originalSize -> $compressedSize bytes ($([Math]::Round($ratio, 1))% smaller)"
        }

        return $true
    }
    catch {
        Write-DebugMsg "Gzip compression failed: $($_.Exception.Message)"
        return $false
    }
    finally {
        if (Test-Path $tempPath) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
    }
}

function Decompress-GzFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )

    $use7z = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_7z_encryption") {
        $use7z = [bool]$global:Config.app.use_7z_encryption
    }

    $format = Get-ArchiveFormat -Path $GzFile
    $preferred = switch ($format) {
        "cpsn" { "7z" }
        "7z" { "7z" }
        "gzip" { "gzip" }
        default { if ($use7z) { "7z" } else { "gzip" } }
    }

    $methods = @($preferred, $(if ($preferred -eq "7z") { "gzip" } else { "7z" }))

    foreach ($method in $methods) {
        if ($method -eq "7z") {
            if (Decompress-JsonWith7z -GzFile $GzFile -JsonFile $JsonFile) {
                return $true
            }
        } else {
            if (Decompress-GzipFile -GzFile $GzFile -JsonFile $JsonFile) {
                return $true
            }
        }
    }

    return $false
}

function Decompress-JsonWith7z {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )
    try {
        $sevenZip = Get-7zExecutablePath
        $password = $global:Config.nextcloud.password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw "Nextcloud password is empty; cannot decrypt"
        }

        $tempDir = Join-Path ([System.IO.Path]::GetTempPath()) ("clipson-7z-" + [System.Guid]::NewGuid().ToString("N"))
        New-Item -ItemType Directory -Path $tempDir -Force | Out-Null

        try {
            $archiveToExtract = $GzFile
            $strippedArchive = Join-Path $tempDir "archive.7z"
            Strip-CpsnHeaderToFile -InputPath $GzFile -OutputPath $strippedArchive
            $archiveToExtract = $strippedArchive

            $args = @(
                "x",
                "-y",
                "-bd",
                ("-p" + $password),
                ("-o" + $tempDir),
                $archiveToExtract
            )

            $output = & $sevenZip @args 2>&1
            if ($LASTEXITCODE -ne 0) {
                if ($global:Config.app.debug_enabled) {
                    Write-DebugMsg "7z decryption failed (exit $LASTEXITCODE): $output"
                } else {
                    Write-DebugMsg "7z decryption failed (exit $LASTEXITCODE)"
                }
                return $false
            }

            $payload = Join-Path $tempDir "payload.json"
            if (-not (Test-Path $payload)) {
                $first = Get-ChildItem -Path $tempDir -File -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1
                if (-not $first) {
                    Write-DebugMsg "7z decryption produced no files"
                    return $false
                }
                Copy-Item -LiteralPath $first.FullName -Destination $JsonFile -Force
                return $true
            }

            Copy-Item -LiteralPath $payload -Destination $JsonFile -Force
            return $true
        }
        finally {
            Remove-Item -LiteralPath $tempDir -Recurse -Force -ErrorAction SilentlyContinue
        }
    }
    catch {
        Write-DebugMsg "File decryption failed: $($_.Exception.Message)"
        return $false
    }
}

function Decompress-GzipFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )
    try {
        $input = [System.IO.File]::OpenRead($GzFile)
        try {
            $gzip = New-Object System.IO.Compression.GzipStream($input, [System.IO.Compression.CompressionMode]::Decompress)
            try {
                $outStream = [System.IO.File]::Create($JsonFile)
                try {
                    $gzip.CopyTo($outStream)
                }
                finally {
                    $outStream.Dispose()
                }
            }
            finally {
                $gzip.Dispose()
            }
        }
        finally {
            $input.Dispose()
        }
        return $true
    }
    catch {
        Write-DebugMsg "Gzip decompression failed: $($_.Exception.Message)"
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
