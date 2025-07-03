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
    "--no-sandbox",
    "--disable-javascript",
    "--print-to-pdf=$OutputPdf",
    "--disable-background-networking",
    "--disable-google-translate",
    "--disable-features=NetworkService,NetworkServiceInProcess,PushMessaging,VoiceTyping,VoiceTranscription,BackgroundFetch,InterestFeedContentSuggestions,TensorFlowLite,AutofillServerCommunication,OptimizationHints,InstantExtendedAPI,TranslateUI,MediaRouter,Notifications,NetworkPrediction,AppBanners,BackForwardCache,TabGroups,ZeroSuggest,PasswordManager,WebRtcHardwareEncoding,BackgroundSync,PaymentHandler,StorageAccessAPI,Clipboard,WebAuthentication,WebUsb",
    "--disable-sync",
    "--disable-network",
    "--disable-client-side-phishing-detection",
    "--no-default-browser-check",
    "--disable-extensions",
    "--disable-default-apps",
    "--disable-component-update",
    "--disable-breakpad",
    "--disable-domain-reliability",
    "--disable-hang-monitor",
    "--disable-ipc-flooding-protection",
    "--disable-popup-blocking",
    "--disable-prompt-on-repost",
    "--disable-renderer-backgrounding",
    "--disable-web-resources",
    "--metrics-recording-only",
    "--mute-audio",
    "--no-first-run"
    "--no-pings",
    "--disable-translate-new-ux"
    $fileUrl
) -NoNewWindow -PassThru

if ($chromeProcess.ExitCode -ne 0) {
    [System.Windows.Forms.MessageBox]::Show("Chrome exited with code $($chromeProcess.ExitCode). PDF creation failed.","Error",'OK','Error') | Out-Null
    exit 3
}

# Wait up to 10 seconds for the PDF to appear
$timeout = 10
$elapsed = 0

while (-not (Test-Path $OutputPdf) -and $elapsed -lt $timeout) {
    Start-Sleep -Milliseconds 1000
    $elapsed += 1
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