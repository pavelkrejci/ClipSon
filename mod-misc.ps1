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
    $useEncryption = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_encryption") {
        $useEncryption = [bool]$global:Config.app.use_encryption
    }

    return $(if ($useEncryption) { ".cpsn" } else { ".gz" })
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

# --- Encryption constants ---
$script:CPSN_VERSION = 2
$script:PBKDF2_ITERATIONS = 100000
$script:SALT_SIZE = 32
$script:IV_SIZE = 16
$script:HMAC_SIZE = 32
$script:AES_KEY_SIZE = 32  # AES-256

function Get-DerivedKeys {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Password,
        [Parameter(Mandatory=$true)]
        [byte[]]$Salt
    )
    # PBKDF2-SHA256 → 64 bytes (32 AES key + 32 HMAC key)
    $deriveBytes = New-Object System.Security.Cryptography.Rfc2898DeriveBytes(
        $Password,
        $Salt,
        $script:PBKDF2_ITERATIONS,
        [System.Security.Cryptography.HashAlgorithmName]::SHA256
    )
    $dk = $deriveBytes.GetBytes(64)
    $deriveBytes.Dispose()

    $aesKey = New-Object byte[] 32
    $hmacKey = New-Object byte[] 32
    [Array]::Copy($dk, 0, $aesKey, 0, 32)
    [Array]::Copy($dk, 32, $hmacKey, 0, 32)

    return @{ AesKey = $aesKey; HmacKey = $hmacKey }
}

function Protect-Data {
    param(
        [Parameter(Mandatory=$true)]
        [byte[]]$Plaintext,
        [Parameter(Mandatory=$true)]
        [string]$Password
    )
    # Generate random salt and IV
    $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
    $salt = New-Object byte[] $script:SALT_SIZE
    $iv = New-Object byte[] $script:IV_SIZE
    $rng.GetBytes($salt)
    $rng.GetBytes($iv)
    $rng.Dispose()

    $keys = Get-DerivedKeys -Password $Password -Salt $salt

    # AES-256-CBC encryption with PKCS7 padding
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256
    $aes.BlockSize = 128
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $keys.AesKey
    $aes.IV = $iv

    $encryptor = $aes.CreateEncryptor()
    $ciphertext = $encryptor.TransformFinalBlock($Plaintext, 0, $Plaintext.Length)
    $encryptor.Dispose()
    $aes.Dispose()

    # HMAC-SHA256 over (version + salt + iv + ciphertext)
    $versionByte = [byte[]]@($script:CPSN_VERSION)
    $macData = New-Object byte[] ($versionByte.Length + $salt.Length + $iv.Length + $ciphertext.Length)
    $offset = 0
    [Array]::Copy($versionByte, 0, $macData, $offset, $versionByte.Length); $offset += $versionByte.Length
    [Array]::Copy($salt, 0, $macData, $offset, $salt.Length); $offset += $salt.Length
    [Array]::Copy($iv, 0, $macData, $offset, $iv.Length); $offset += $iv.Length
    [Array]::Copy($ciphertext, 0, $macData, $offset, $ciphertext.Length)

    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = $keys.HmacKey
    $mac = $hmac.ComputeHash($macData)
    $hmac.Dispose()

    # Assemble: CPSN(4) + version(1) + salt(32) + iv(16) + ciphertext + hmac(32)
    $magic = [System.Text.Encoding]::ASCII.GetBytes("CPSN")
    $result = New-Object byte[] ($magic.Length + $versionByte.Length + $salt.Length + $iv.Length + $ciphertext.Length + $mac.Length)
    $offset = 0
    [Array]::Copy($magic, 0, $result, $offset, $magic.Length); $offset += $magic.Length
    [Array]::Copy($versionByte, 0, $result, $offset, $versionByte.Length); $offset += $versionByte.Length
    [Array]::Copy($salt, 0, $result, $offset, $salt.Length); $offset += $salt.Length
    [Array]::Copy($iv, 0, $result, $offset, $iv.Length); $offset += $iv.Length
    [Array]::Copy($ciphertext, 0, $result, $offset, $ciphertext.Length); $offset += $ciphertext.Length
    [Array]::Copy($mac, 0, $result, $offset, $mac.Length)

    return $result
}

function Unprotect-Data {
    param(
        [Parameter(Mandatory=$true)]
        [byte[]]$Data,
        [Parameter(Mandatory=$true)]
        [string]$Password
    )
    $minSize = 4 + 1 + $script:SALT_SIZE + $script:IV_SIZE + $script:HMAC_SIZE + 16
    if ($Data.Length -lt $minSize) {
        throw "Data too short to be a valid CPSN v2 archive"
    }

    $offset = 0
    $magic = [System.Text.Encoding]::ASCII.GetString($Data, $offset, 4); $offset += 4
    if ($magic -ne "CPSN") {
        throw "Invalid CPSN magic header"
    }

    $version = $Data[$offset]; $offset += 1
    if ($version -ne $script:CPSN_VERSION) {
        throw "Unsupported CPSN version: $version"
    }

    $salt = New-Object byte[] $script:SALT_SIZE
    [Array]::Copy($Data, $offset, $salt, 0, $script:SALT_SIZE); $offset += $script:SALT_SIZE

    $iv = New-Object byte[] $script:IV_SIZE
    [Array]::Copy($Data, $offset, $iv, 0, $script:IV_SIZE); $offset += $script:IV_SIZE

    $ciphertextLen = $Data.Length - $offset - $script:HMAC_SIZE
    $ciphertext = New-Object byte[] $ciphertextLen
    [Array]::Copy($Data, $offset, $ciphertext, 0, $ciphertextLen)

    $storedMac = New-Object byte[] $script:HMAC_SIZE
    [Array]::Copy($Data, $Data.Length - $script:HMAC_SIZE, $storedMac, 0, $script:HMAC_SIZE)

    $keys = Get-DerivedKeys -Password $Password -Salt $salt

    # Verify HMAC (Encrypt-then-MAC)
    $versionByte = [byte[]]@($version)
    $macData = New-Object byte[] ($versionByte.Length + $salt.Length + $iv.Length + $ciphertext.Length)
    $macOffset = 0
    [Array]::Copy($versionByte, 0, $macData, $macOffset, $versionByte.Length); $macOffset += $versionByte.Length
    [Array]::Copy($salt, 0, $macData, $macOffset, $salt.Length); $macOffset += $salt.Length
    [Array]::Copy($iv, 0, $macData, $macOffset, $iv.Length); $macOffset += $iv.Length
    [Array]::Copy($ciphertext, 0, $macData, $macOffset, $ciphertext.Length)

    $hmac = New-Object System.Security.Cryptography.HMACSHA256
    $hmac.Key = $keys.HmacKey
    $expectedMac = $hmac.ComputeHash($macData)
    $hmac.Dispose()

    # Constant-time comparison
    $diff = 0
    for ($i = 0; $i -lt $script:HMAC_SIZE; $i++) {
        $diff = $diff -bor ($storedMac[$i] -bxor $expectedMac[$i])
    }
    if ($diff -ne 0) {
        throw "HMAC verification failed - wrong password or corrupted data"
    }

    # Decrypt AES-256-CBC
    $aes = [System.Security.Cryptography.Aes]::Create()
    $aes.KeySize = 256
    $aes.BlockSize = 128
    $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
    $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
    $aes.Key = $keys.AesKey
    $aes.IV = $iv

    $decryptor = $aes.CreateDecryptor()
    $plaintext = $decryptor.TransformFinalBlock($ciphertext, 0, $ciphertext.Length)
    $decryptor.Dispose()
    $aes.Dispose()

    return $plaintext
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
            $buf = New-Object byte[] 4
            $read = $fs.Read($buf, 0, 4)
            if ($read -ge 4) {
                $sig = [System.Text.Encoding]::ASCII.GetString($buf, 0, 4)
                if ($sig -eq "CPSN") { return "cpsn" }
            }
            if ($read -ge 2 -and $buf[0] -eq 0x1f -and $buf[1] -eq 0x8b) { return "gzip" }
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

    $useEncryption = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_encryption") {
        $useEncryption = [bool]$global:Config.app.use_encryption
    }

    if ($useEncryption) {
        $encOk = Compress-JsonFileWithEncryption -JsonFile $JsonFile -GzFile $GzFile
        if ($encOk) { return $true }
        Write-DebugMsg "Encryption failed; falling back to gzip"
    }

    return Compress-JsonFileWithGzip -JsonFile $JsonFile -GzFile $GzFile
}

function Compress-JsonFileWithEncryption {
    param(
        [Parameter(Mandatory=$true)]
        [string]$JsonFile,
        [Parameter(Mandatory=$true)]
        [string]$GzFile
    )
    try {
        $password = $global:Config.nextcloud.password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw "Nextcloud password is empty; cannot encrypt"
        }

        $originalSize = (Get-Item $JsonFile).Length

        # Read and gzip-compress the JSON data
        $rawData = [System.IO.File]::ReadAllBytes($JsonFile)

        $memStream = New-Object System.IO.MemoryStream
        $gzipStream = New-Object System.IO.Compression.GzipStream($memStream, [System.IO.Compression.CompressionLevel]::Optimal, $true)
        $gzipStream.Write($rawData, 0, $rawData.Length)
        $gzipStream.Dispose()
        $compressed = $memStream.ToArray()
        $memStream.Dispose()

        # Encrypt the compressed data
        $encrypted = Protect-Data -Plaintext $compressed -Password $password

        # Write atomically
        $tempPath = "$GzFile.$([System.Guid]::NewGuid().ToString('N')).tmp"
        [System.IO.File]::WriteAllBytes($tempPath, $encrypted)
        Move-Item -LiteralPath $tempPath -Destination $GzFile -Force

        $encryptedSize = (Get-Item $GzFile).Length
        if ($global:Config.app.debug_enabled) {
            $ratio = if ($originalSize -gt 0) { (1 - $encryptedSize / $originalSize) * 100 } else { 0 }
            Write-DebugMsg "File encryption $originalSize -> $encryptedSize bytes ($([Math]::Round($ratio, 1))% smaller)"
        }

        return $true
    }
    catch {
        Write-DebugMsg "File encryption failed: $($_.Exception.Message)"
        if (Test-Path $tempPath -ErrorAction SilentlyContinue) {
            Remove-Item -LiteralPath $tempPath -Force -ErrorAction SilentlyContinue
        }
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

    $useEncryption = $true
    if ($global:Config -and $global:Config.app -and $global:Config.app.PSObject.Properties.Name -contains "use_encryption") {
        $useEncryption = [bool]$global:Config.app.use_encryption
    }

    $format = Get-ArchiveFormat -Path $GzFile
    $preferred = switch ($format) {
        "cpsn" { "encrypted" }
        "gzip" { "gzip" }
        default { if ($useEncryption) { "encrypted" } else { "gzip" } }
    }

    $methods = @($preferred, $(if ($preferred -eq "encrypted") { "gzip" } else { "encrypted" }))

    foreach ($method in $methods) {
        if ($method -eq "encrypted") {
            if (Decompress-EncryptedFile -GzFile $GzFile -JsonFile $JsonFile) {
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

function Decompress-EncryptedFile {
    param(
        [Parameter(Mandatory=$true)]
        [string]$GzFile,
        [Parameter(Mandatory=$true)]
        [string]$JsonFile
    )
    try {
        $password = $global:Config.nextcloud.password
        if ([string]::IsNullOrWhiteSpace($password)) {
            throw "Nextcloud password is empty; cannot decrypt"
        }

        $data = [System.IO.File]::ReadAllBytes($GzFile)

        # Decrypt
        $compressed = Unprotect-Data -Data $data -Password $password

        # Decompress gzip
        $memStream = New-Object System.IO.MemoryStream(, $compressed)
        $gzipStream = New-Object System.IO.Compression.GzipStream($memStream, [System.IO.Compression.CompressionMode]::Decompress)
        $outStream = New-Object System.IO.MemoryStream
        $gzipStream.CopyTo($outStream)
        $gzipStream.Dispose()
        $memStream.Dispose()
        $rawData = $outStream.ToArray()
        $outStream.Dispose()

        [System.IO.File]::WriteAllBytes($JsonFile, $rawData)
        return $true
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
