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

# Use (?s) to make . match newlines
$fragment = $null
if ($htmlContent -match "(?s)StartFragment:(\d+).*EndFragment:(\d+)") {
    $startFragment = [int]$matches[1]
    $endFragment = [int]$matches[2]
    Write-Host "StartFragment: $startFragment, EndFragment: $endFragment"
    $fragment = $htmlContent.Substring($startFragment, $endFragment - $startFragment)
} else {
    Write-Host "Could not find StartFragment/EndFragment in clipboard HTML."
    $fragment = $htmlContent
}

# Write the extracted HTML fragment to file
$outFile = ".\clipboard-html-fragment.txt"
[System.IO.File]::WriteAllText($outFile, $fragment, [System.Text.Encoding]::UTF8)
Write-Host "Clipboard HTML fragment written to $outFile"

# Debug output
Write-Host "First 200 chars: $($htmlContent.Substring(0, [Math]::Min(200, $htmlContent.Length)))"