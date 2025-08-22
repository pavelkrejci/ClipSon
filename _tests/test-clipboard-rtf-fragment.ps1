Add-Type -AssemblyName System.Windows.Forms

# Get RTF content from clipboard
$rtfContent = $null
try {
    if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::Rtf)) {
        $rtfContent = [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Rtf)
    } else {
        Write-Host "No RTF content found in clipboard."
        exit 1
    }
} catch {
    Write-Host "Error accessing clipboard: $($_.Exception.Message)"
    exit 1
}

# There is no header/fragment in RTF clipboard format, so just write as-is
$outFile = ".\clipboard-rtf-raw.txt"
[System.IO.File]::WriteAllText($outFile, $rtfContent, [System.Text.Encoding]::UTF8)
Write-Host "Clipboard RTF content written to $outFile"

# Debug output
Write-Host "First 200 chars: $($rtfContent.Substring(0, [Math]::Min(200, $rtfContent.Length)))"
