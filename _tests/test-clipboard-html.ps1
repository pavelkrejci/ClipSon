Add-Type -AssemblyName System.Windows.Forms

# Get HTML content from clipboard
$htmlContent = $null
try {
    if ([System.Windows.Forms.Clipboard]::ContainsText([System.Windows.Forms.TextDataFormat]::Html)) {
        $htmlContent = [System.Windows.Forms.Clipboard]::GetText([System.Windows.Forms.TextDataFormat]::Html)
    } else {
        Write-Host "No HTML content found in clipboard."
        exit 1
    }
} catch {
    Write-Host "Error accessing clipboard: $($_.Exception.Message)"
    exit 1
}

# Write raw HTML content to file
$outFile = ".\clipboard-html-raw.txt"
[System.IO.File]::WriteAllText($outFile, $htmlContent, [System.Text.Encoding]::UTF8)
Write-Host "Clipboard HTML content written to $outFile"
