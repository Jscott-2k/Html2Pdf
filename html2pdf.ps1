param(
    [string]$InputHtml,
    [string]$OutputPdf = $null
)

Add-Type -AssemblyName System.Windows.Forms

if (-not $InputHtml -or -not (Test-Path $InputHtml)) {
    [System.Windows.Forms.MessageBox]::Show("Input HTML file not found.","Error",'OK','Error') | Out-Null
    exit 1
}

if (-not $OutputPdf) {
    $OutputPdf = [System.IO.Path]::ChangeExtension($InputHtml, "pdf")
}

# Chrome paths to check
$chromePaths = @(
    "$Env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "$Env:ProgramFiles(x86)\Google\Chrome\Application\chrome.exe"
)
$chrome = $chromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $chrome) {
    Write-Error "Google Chrome not found."
    pause
    exit 2
}

# absolute Windows path
$fileFullPath = (Resolve-Path $InputHtml).Path
$fileUrl = ([uri]$fileFullPath).AbsoluteUri

# Remove any stale PDF
if (Test-Path $OutputPdf) {
    Remove-Item $OutputPdf -Force
}

# Start Chrome and wait for it to finish
$chromeProcess = Start-Process -FilePath $chrome -ArgumentList @(
    "--headless",
    "--disable-gpu",
    "--no-sandbox",
    "--print-to-pdf=""$OutputPdf""",
    $fileUrl
) -Wait -NoNewWindow -PassThru

# Wait up to 10 seconds for the PDF to appear
$timeout = 10
$elapsed = 0

while (-not (Test-Path $OutputPdf) -and $elapsed -lt $timeout) {
    Start-Sleep -Milliseconds 200
    $elapsed += 0.2
    Write-Host -NoNewline "."
}
Write-Host ""  # Newline after dots

# Alert the user
if (Test-Path $OutputPdf) {
    [System.Windows.Forms.MessageBox]::Show(
        "PDF created at `n$OutputPdf",
        "Success",
        'OK',
        'Information'
    ) | Out-Null
} else {
    [System.Windows.Forms.MessageBox]::Show(
        "Failed to create PDF within $timeout seconds.",
        "Error",
        'OK',
        'Error'
    ) | Out-Null
}

pause