<#
.SYNOPSIS
Install Python and MarkItDown prerequisites for Windows PCs.

.DESCRIPTION
This script installs Python 3 using winget (if available), upgrades pip, and installs MarkItDown with broad file-format support.
It also prints the commands needed to convert files with MarkItDown.
#>

[CmdletBinding()]
param()

function Write-Ok($Message) {
    Write-Host "[OK] $Message" -ForegroundColor Green
}

function Write-WarningLine($Message) {
    Write-Host "[WARN] $Message" -ForegroundColor Yellow
}

function Write-ErrorLine($Message) {
    Write-Host "[ERROR] $Message" -ForegroundColor Red
}

function CommandExists($name) {
    return [bool](Get-Command $name -ErrorAction SilentlyContinue)
}

$pyCommand = Get-Command py -ErrorAction SilentlyContinue
if (-not $pyCommand) {
    $pythonCommand = Get-Command python -ErrorAction SilentlyContinue
} else {
    $pythonCommand = $pyCommand
}

if (-not $pythonCommand) {
    Write-ErrorLine "Python was not found on this machine. Please install the latest Python from the Microsoft Store: https://apps.microsoft.com/detail/9nq7512cxl7t?hl=ro-RO&gl=RO"
    throw "Python installation is required for MarkItDown."
}

if (-not $pythonCommand) {
    throw "Python installation is required for MarkItDown."
}

$pythonExe = $pythonCommand.Source
Write-Ok "Using Python interpreter: $pythonExe"

Write-Host "Upgrading pip..."
& $pythonExe -m pip install --upgrade pip
if ($LASTEXITCODE -ne 0) {
    Write-WarningLine "pip upgrade failed, but installation may still work."
}

Write-Host "Installing MarkItDown with broad format support..."
& $pythonExe -m pip install --user "markitdown[all]"
if ($LASTEXITCODE -ne 0) {
    Write-ErrorLine "Failed to install MarkItDown. Please inspect the output and retry."
    exit 1
}
Write-Ok "MarkItDown installed successfully."

$scriptPath = & $pythonExe -c "import sysconfig; print(sysconfig.get_path('scripts'))"
Write-Host "Python scripts path: $scriptPath"

if (-not (Test-Path $scriptPath)) {
    Write-WarningLine "The Python Scripts path was not found. If MarkItDown was installed successfully, add the correct Scripts folder to PATH manually."
}

# Print helpful commands for the user to run after installation.
Write-Host ""
Write-Host "=== MarkItDown usage examples ==="
Write-Host "# Convert a PDF to Markdown in the same folder:"
Write-Host "py -3 -m markitdown \"input\\Vlad Gheorghe Branzei\\GHEORGHE VLAD BRANZEI.pdf\" -o \"input\\Vlad Gheorghe Branzei\\GHEORGHE VLAD BRANZEI.md\""
Write-Host ""
Write-Host "# Convert other document types to Markdown:"
Write-Host "py -3 -m markitdown \"input\\Vlad Gheorghe Branzei\\document.docx\" -o \"input\\Vlad Gheorghe Branzei\\document.md\""
Write-Host "py -3 -m markitdown \"input\\Vlad Gheorghe Branzei\\slides.pptx\" -o \"input\\Vlad Gheorghe Branzei\\slides.md\""
Write-Host "py -3 -m markitdown \"input\\Vlad Gheorghe Branzei\\spreadsheet.xlsx\" -o \"input\\Vlad Gheorghe Branzei\\spreadsheet.md\""
Write-Host ""
Write-Host "If the 'markitdown' command is available in PATH, you can also use:"
Write-Host "markitdown path-to-file.pdf -o path-to-file.md"
Write-Host ""
Write-Host "If the script path is not on PATH, add it using:"
Write-Host "  setx PATH \"$env:PATH;$scriptPath\""
Write-Ok "Prerequisites installation complete."
