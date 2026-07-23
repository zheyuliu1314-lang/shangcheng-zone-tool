param(
    [string]$Message = "Update project"
)

$ErrorActionPreference = "Stop"
$git = "C:\Users\wy98k\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\git\cmd\git.exe"
if (-not (Test-Path -LiteralPath $git)) {
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if ($gitCommand) { $git = $gitCommand.Source }
}
if (-not (Test-Path -LiteralPath $git)) {
    throw "Git was not found. Install Git for Windows first."
}

& $git add .
& $git diff --cached --quiet
if ($LASTEXITCODE -eq 0) {
    Write-Host "No new changes."
    exit 0
}

& $git commit -m $Message
$token = $env:GITHUB_TOKEN
if (-not $token) {
    $secureToken = Read-Host "Enter GitHub fine-grained token (used only for this push)" -AsSecureString
    $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken))
}
if (-not $token) { throw "No GitHub token was provided." }

$env:GIT_TERMINAL_PROMPT = "0"
& $git -c "http.extraheader=AUTHORIZATION: bearer $token" push origin main
if ($LASTEXITCODE -ne 0) {
    throw "GitHub push failed. Check token repository Contents write permission."
}
Write-Host "Push completed. Render will deploy the latest commit."
