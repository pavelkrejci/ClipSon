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
        # ClipSon encrypted container format (CSGZ v1): gzip(json) then AES-CBC + HMAC-SHA256.
        # The output file is stored with .cs extension for WebDAV transport,
        # but the content is NOT a raw gzip stream anymore.

        function Get-ClipsonPasswordOrThrow {
            if ($global:Config -and $global:Config.app -and -not [string]::IsNullOrWhiteSpace($global:Config.app.encryption_password)) {
                return $global:Config.app.encryption_password
            }
            if ($global:Config -and $global:Config.nextcloud -and -not [string]::IsNullOrWhiteSpace($global:Config.nextcloud.password)) {
                return $global:Config.nextcloud.password
            }
            throw "Nextcloud password is not configured (needed for encryption)"
        }

        function Get-RandomBytes {
            param([int]$Length)
            $bytes = New-Object byte[] $Length
            $rng = [System.Security.Cryptography.RandomNumberGenerator]::Create()
            $rng.GetBytes($bytes)
            $rng.Dispose()
            return $bytes
        }

        function Get-ClipsonKdfSupport {
            # Returns preferred KDF id: 2 = PBKDF2-HMAC-SHA256 (if supported), else 1 = PBKDF2-HMAC-SHA1
            $ctor = [System.Security.Cryptography.Rfc2898DeriveBytes].GetConstructor(@(
                [string],
                [byte[]],
                [int],
                [System.Security.Cryptography.HashAlgorithmName]
            ))
            if ($ctor) { return 2 }
            return 1
        }

        function Derive-ClipsonKeys {
            param(
                [Parameter(Mandatory=$true)][string]$Password,
                [Parameter(Mandatory=$true)][byte[]]$Salt,
                [Parameter(Mandatory=$true)][int]$Iterations,
                [Parameter(Mandatory=$true)][int]$KdfId
            )

            if ($KdfId -eq 2) {
                $derive = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($Password, $Salt, $Iterations, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
            } else {
                # .NET Framework default constructor uses HMAC-SHA1
                $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, $Iterations)
            }

            $keyMaterial = $derive.GetBytes(64)
            $derive.Dispose()

            $encKey = New-Object byte[] 32
            $macKey = New-Object byte[] 32
            [System.Array]::Copy($keyMaterial, 0, $encKey, 0, 32)
            [System.Array]::Copy($keyMaterial, 32, $macKey, 0, 32)
            return @{ EncKey = $encKey; MacKey = $macKey }
        }

        function FixedTimeEquals {
            param([byte[]]$A, [byte[]]$B)
            if ($null -eq $A -or $null -eq $B) { return $false }
            if ($A.Length -ne $B.Length) { return $false }
            $diff = 0
            for ($i = 0; $i -lt $A.Length; $i++) {
                $diff = $diff -bor ($A[$i] -bxor $B[$i])
            }
            return ($diff -eq 0)
        }

        $password = Get-ClipsonPasswordOrThrow
        $jsonBytes = [System.IO.File]::ReadAllBytes($JsonFile)
        $originalSize = $jsonBytes.Length

        # 1) gzip compress to bytes
        $gzipMs = New-Object System.IO.MemoryStream
        $gzipStream = New-Object System.IO.Compression.GzipStream($gzipMs, [System.IO.Compression.CompressionLevel]::Fastest, $true)
        $gzipStream.Write($jsonBytes, 0, $jsonBytes.Length)
        $gzipStream.Dispose()
        $gzipBytes = $gzipMs.ToArray()
        $gzipMs.Dispose()

        # 2) derive keys from Nextcloud password
        $iterations = 50000
        $salt = Get-RandomBytes -Length 16
        $iv = Get-RandomBytes -Length 16
        $kdfId = Get-ClipsonKdfSupport
        $keys = Derive-ClipsonKeys -Password $password -Salt $salt -Iterations $iterations -KdfId $kdfId

        # 3) AES-CBC encrypt (PKCS7)
        $aes = [System.Security.Cryptography.Aes]::Create()
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.KeySize = 256
        $aes.Key = $keys.EncKey
        $aes.IV = $iv
        $encryptor = $aes.CreateEncryptor()
        $cipher = $encryptor.TransformFinalBlock($gzipBytes, 0, $gzipBytes.Length)
        $encryptor.Dispose()
        $aes.Dispose()

        # 4) Write container and HMAC
        $ms = New-Object System.IO.MemoryStream
        $bw = New-Object System.IO.BinaryWriter($ms)

        $bw.Write([byte[]][System.Text.Encoding]::ASCII.GetBytes('CSGZ'))
        $bw.Write([byte]1)              # version
        $bw.Write([byte]$kdfId)         # kdf id: 1=sha1, 2=sha256
        $bw.Write([UInt32]$iterations)  # little-endian
        $bw.Write([byte]$salt.Length)
        $bw.Write([byte]$iv.Length)
        $bw.Write($salt)
        $bw.Write($iv)
        # Do not write cipher length; ciphertext length is derived from file size.
        # Decryptor accepts both this layout and a legacy layout that included cipher length.
        $bw.Write($cipher)

        $toMac = $ms.ToArray()
        $hmac = New-Object System.Security.Cryptography.HMACSHA256 -ArgumentList (,$keys.MacKey)
        $tag = $hmac.ComputeHash($toMac)
        $hmac.Dispose()
        $bw.Write($tag)
        $bw.Dispose()

        [System.IO.File]::WriteAllBytes($GzFile, $ms.ToArray())
        $ms.Dispose()

        $finalSize = (Get-Item $GzFile).Length
        if ($global:Config.app.debug_enabled) {
            $ratio = if ($originalSize -gt 0) { (1 - $finalSize / $originalSize) * 100 } else { 0 }
            Write-DebugMsg "File compress+encrypt $originalSize -> $finalSize bytes ($([Math]::Round($ratio, 1))% saved)"
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
        function Get-ClipsonPasswordOrThrow {
            if ($global:Config -and $global:Config.app -and -not [string]::IsNullOrWhiteSpace($global:Config.app.encryption_password)) {
                return $global:Config.app.encryption_password
            }
            if ($global:Config -and $global:Config.nextcloud -and -not [string]::IsNullOrWhiteSpace($global:Config.nextcloud.password)) {
                return $global:Config.nextcloud.password
            }
            throw "Nextcloud password is not configured (needed for decryption)"
        }

        function Get-ClipsonKdfSupport {
            $ctor = [System.Security.Cryptography.Rfc2898DeriveBytes].GetConstructor(@(
                [string],
                [byte[]],
                [int],
                [System.Security.Cryptography.HashAlgorithmName]
            ))
            if ($ctor) { return 2 }
            return 1
        }

        function Derive-ClipsonKeys {
            param(
                [Parameter(Mandatory=$true)][string]$Password,
                [Parameter(Mandatory=$true)][byte[]]$Salt,
                [Parameter(Mandatory=$true)][int]$Iterations,
                [Parameter(Mandatory=$true)][int]$KdfId
            )

            if ($KdfId -eq 2) {
                $derive = [System.Security.Cryptography.Rfc2898DeriveBytes]::new($Password, $Salt, $Iterations, [System.Security.Cryptography.HashAlgorithmName]::SHA256)
            } else {
                $derive = New-Object System.Security.Cryptography.Rfc2898DeriveBytes($Password, $Salt, $Iterations)
            }

            $keyMaterial = $derive.GetBytes(64)
            $derive.Dispose()

            $encKey = New-Object byte[] 32
            $macKey = New-Object byte[] 32
            [System.Array]::Copy($keyMaterial, 0, $encKey, 0, 32)
            [System.Array]::Copy($keyMaterial, 32, $macKey, 0, 32)
            return @{ EncKey = $encKey; MacKey = $macKey }
        }

        function FixedTimeEquals {
            param([byte[]]$A, [byte[]]$B)
            if ($null -eq $A -or $null -eq $B) { return $false }
            if ($A.Length -ne $B.Length) { return $false }
            $diff = 0
            for ($i = 0; $i -lt $A.Length; $i++) {
                $diff = $diff -bor ($A[$i] -bxor $B[$i])
            }
            return ($diff -eq 0)
        }

        $password = Get-ClipsonPasswordOrThrow
        $blob = [System.IO.File]::ReadAllBytes($GzFile)
        if ($blob.Length -lt (4 + 1 + 1 + 4 + 1 + 1 + 4 + 32)) {
            throw "Encrypted file is too small"
        }

        # Parse header from raw bytes (tolerant to legacy layout with explicit cipher length)
        $magic = [System.Text.Encoding]::ASCII.GetString($blob, 0, 4)
        if ($magic -ne 'CSGZ') { throw "Invalid file header (expected CSGZ)" }

        $version = [int]$blob[4]
        if ($version -ne 1) { throw "Unsupported encrypted format version: $version" }

        $kdfId = [int]$blob[5]
        $iterations = [int][System.BitConverter]::ToUInt32($blob, 6)
        $saltLen = [int]$blob[10]
        $ivLen = [int]$blob[11]
        if ($saltLen -le 0 -or $ivLen -le 0) { throw "Invalid salt/iv lengths" }

        $offset = 12
        if ($blob.Length -lt ($offset + $saltLen + $ivLen + 32)) { throw "Encrypted file is too small" }

        $salt = New-Object byte[] $saltLen
        [System.Array]::Copy($blob, $offset, $salt, 0, $saltLen)
        $offset += $saltLen

        $iv = New-Object byte[] $ivLen
        [System.Array]::Copy($blob, $offset, $iv, 0, $ivLen)
        $offset += $ivLen

        $cipherStart = $offset
        $cipherLen = $null
        $tagStart = $null

        # Legacy layout check: u32 cipherLen + cipher + tag(32)
        if ($blob.Length -ge ($cipherStart + 4 + 32)) {
            $candidateLen = [int][System.BitConverter]::ToUInt32($blob, $cipherStart)
            if ($candidateLen -ge 0) {
                $candidateTagStart = $cipherStart + 4 + $candidateLen
                if ($candidateTagStart + 32 -eq $blob.Length) {
                    $cipherStart = $cipherStart + 4
                    $cipherLen = $candidateLen
                    $tagStart = $candidateTagStart
                }
            }
        }

        if ($null -eq $cipherLen) {
            # Current layout: cipher + tag(32)
            $cipherLen = $blob.Length - $cipherStart - 32
            if ($cipherLen -lt 0) { throw "Invalid ciphertext length" }
            $tagStart = $cipherStart + $cipherLen
        }

        $cipher = New-Object byte[] $cipherLen
        [System.Array]::Copy($blob, $cipherStart, $cipher, 0, $cipherLen)

        $tag = New-Object byte[] 32
        [System.Array]::Copy($blob, $tagStart, $tag, 0, 32)

        $supported = Get-ClipsonKdfSupport
        if ($kdfId -eq 2 -and $supported -ne 2) {
            throw "This PowerShell/.NET runtime does not support PBKDF2-HMAC-SHA256, cannot decrypt this file"
        }

        $keys = Derive-ClipsonKeys -Password $password -Salt $salt -Iterations $iterations -KdfId $kdfId

        # Verify HMAC over everything except the trailing tag
        $toMac = New-Object byte[] ($blob.Length - 32)
        [System.Array]::Copy($blob, 0, $toMac, 0, $toMac.Length)
        $hmac = New-Object System.Security.Cryptography.HMACSHA256 -ArgumentList (,$keys.MacKey)
        $expected = $hmac.ComputeHash($toMac)
        $hmac.Dispose()

        if (-not (FixedTimeEquals -A $expected -B $tag)) {
            throw "HMAC verification failed (wrong password or corrupted file)"
        }

        # Decrypt
        $aes = [System.Security.Cryptography.Aes]::Create()
        $aes.Mode = [System.Security.Cryptography.CipherMode]::CBC
        $aes.Padding = [System.Security.Cryptography.PaddingMode]::PKCS7
        $aes.KeySize = 256
        $aes.Key = $keys.EncKey
        $aes.IV = $iv
        $decryptor = $aes.CreateDecryptor()
        $gzipBytes = $decryptor.TransformFinalBlock($cipher, 0, $cipher.Length)
        $decryptor.Dispose()
        $aes.Dispose()

        # Decompress gzip bytes to output file
        $inMs = New-Object System.IO.MemoryStream(,$gzipBytes)
        $gzipStream = New-Object System.IO.Compression.GzipStream($inMs, [System.IO.Compression.CompressionMode]::Decompress)
        $out = [System.IO.File]::Create($JsonFile)
        $gzipStream.CopyTo($out)
        $out.Dispose()
        $gzipStream.Dispose()
        $inMs.Dispose()

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
