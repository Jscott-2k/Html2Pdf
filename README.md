# Html2Pdf

A lightweight PowerShell utility that adds a right-click context menu for converting `.html` files to `.pdf` using headless Google Chrome.

## Why?

This tool was built to solve a niche but real-world problem: converting HTML files to PDFs in locked-down environments where traditional PDF libraries are blocked — but **Google Chrome is available**. It’s fast, offline, and has zero external dependencies.

## Features

- Right-click context menu integration for `.html` files.
- Converts using headless Google Chrome.
- Offline-only mode: disables JavaScript, network, and background activity.
- No third-party dependencies — just PowerShell and Chrome.
- Simple install/uninstall process.

## Requirements

- Windows 10 or 11  
- PowerShell 5.1 or later  
- Google Chrome installed (standard desktop version)

## Installation

1. Download or clone this repository.
2. **Right-click** `Setup-html2pdf.ps1` → **Run with PowerShell as Administrator**.
3. Done — right-click any `.html` file to convert it.

## How to Use

1. **Right-click** a `.html` file.
2. Choose **Convert HTML to PDF**.
3. A `.pdf` will be created in the same folder with the same name.

## Uninstall

To remove the context menu integration:

```powershell
.\Setup-html2pdf.ps1 -Remove