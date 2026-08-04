$path = Join-Path $PSScriptRoot 'index.html'
$insertLine = '    <script src="eb-import-fix.js"></script>'
$lines = Get-Content -Path $path
if ($lines -contains $insertLine) {
    exit 0
}
$index = [Array]::LastIndexOf($lines, '</body>')
if ($index -lt 0) {
    throw 'Closing </body> tag not found.'
}
$updated = @()
if ($index -gt 0) {
    $updated += $lines[0..($index - 1)]
}
$updated += $insertLine
$updated += $lines[$index..($lines.Length - 1)]
Set-Content -Path $path -Value $updated -Encoding utf8
