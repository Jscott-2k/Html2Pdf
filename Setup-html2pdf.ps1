param([switch]$Remove)

# Self‑elevate if not running as admin
if (-not ([Security.Principal.WindowsPrincipal] `
          [Security.Principal.WindowsIdentity]::GetCurrent() `
         ).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator))
{
    Write-Host "[DEBUG] Elevating..."
    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy','Bypass',
        '-File', "`"$PSCommandPath`""
    )
    if ($Remove) { $argList += '-Remove' }
    Start-Process powershell -Verb RunAs -ArgumentList $argList
    exit
}

try {
    Write-Host "[DEBUG] Entered elevated context"

    # Ensure per-user mapping of .html to htmlfile
    Write-Host "[DEBUG] Mapping .html to htmlfile in HKCR"
    New-Item -Path "Registry::HKEY_CLASSES_ROOT\.html" -Force | Out-Null
    Set-ItemProperty -Path "Registry::HKEY_CLASSES_ROOT\.html" -Name '(default)' -Value 'htmlfile' -Force
    Write-Host "[DEBUG] Mapping .htm to htmlfile in HKCR"

    # Define paths
    $source     = Join-Path $PSScriptRoot 'html2pdf.ps1'
    $installDir = 'C:\Program Files\Html2Pdf'
    $dest       = Join-Path $installDir  'html2pdf.ps1'

    Write-Host "[DEBUG] source = $source"
    Write-Host "[DEBUG] dest   = $dest"

    # Registry keys under explorer html ProgID
    $progId = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.html\UserChoice').ProgId
    $baseKey  = "Registry::HKEY_CLASSES_ROOT\$progId\shell\Html2Pdf"
    $cmdKey   = "$baseKey\command"

    # Uninstall if requested
    if ($Remove) {
        Write-Host "[DEBUG] Uninstall block"
        if (Test-Path $baseKey) {
            Remove-Item -Path $baseKey -Recurse -Force
            Write-Host "[DEBUG] Removed registry key"
        } else {
            Write-Host "[DEBUG] No registry key to remove"
        }
        if (Test-Path $dest) {
            Remove-Item -Path $dest -Force
            Write-Host "[DEBUG] Removed script file"
        }
        if (Test-Path $installDir -and -not (Get-ChildItem $installDir)) {
            Remove-Item -Path $installDir -Force
            Write-Host "[DEBUG] Removed empty install folder"
        }
        Write-Host "`nUninstallation complete."
        exit
    }

    # Install block
    Write-Host "[DEBUG] Installing Html2Pdf..."

    if (-not (Test-Path $installDir)) {
        New-Item -Path $installDir -ItemType Directory | Out-Null
        Write-Host "[DEBUG] Created folder $installDir"
    } else {
        Write-Host "[DEBUG] Folder already exists: $installDir"
    }

    if ($source -ne $dest) {
        Write-Host "[DEBUG] Copying html2pdf.ps1 to $dest"
        Copy-Item -Path $source -Destination $dest -Force
        Write-Host "[DEBUG] Copy complete"
    } else {
        Write-Host "[DEBUG] Source equals destination; skipping copy"
    }

    if (-not (Test-Path $baseKey)) {
        New-Item -Path $baseKey -Force | Out-Null
        Write-Host "[DEBUG] Created registry key $baseKey"
    }
    if (-not (Test-Path $cmdKey)) {
        New-Item -Path $cmdKey -Force | Out-Null
        Write-Host "[DEBUG] Created registry key $cmdKey"
    }

    Write-Host "[DEBUG] Writing registry values"
    Set-ItemProperty -Path $baseKey -Name '(default)' -Value 'Convert HTML to PDF'
    $cmdValue = '"powershell.exe" -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "' + $dest + '" -InputHtml "%1"'
    Set-ItemProperty -Path $cmdKey -Name '(default)' -Value $cmdValue
    Write-Host "[DEBUG] Registry values set"

    Write-Host "`nInstallation complete.  Restart Explorer to see the menu."
}
catch {
    Write-Error "[ERROR] $($_.Exception.Message)"
    exit 1
}
pause