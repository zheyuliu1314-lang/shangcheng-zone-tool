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
if ($LASTEXITCODE -ne 0) {
    & $git commit -m $Message
} else {
    Write-Host "No new file changes; checking for unpushed commits."
}
$token = $env:GITHUB_TOKEN
if (-not $token) {
    $env:GIT_TERMINAL_PROMPT = "1"
    & $git push origin main
    if ($LASTEXITCODE -eq 0) {
        Write-Host "Push completed. Render will deploy the latest commit."
        exit 0
    }
    $secureToken = Read-Host "Enter GitHub fine-grained token (used only for this push)" -AsSecureString
    $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken))
}
if (-not $token) { throw "No GitHub token was provided." }

$env:GIT_TERMINAL_PROMPT = "0"
$username = "zheyuliu1314-lang"
$basicBytes = [Text.Encoding]::ASCII.GetBytes("$username`:$token")
$basicToken = [Convert]::ToBase64String($basicBytes)
& $git -c "http.extraheader=AUTHORIZATION: Basic $basicToken" push origin main
if ($LASTEXITCODE -ne 0) {
    throw "GitHub push failed. Check token repository Contents write permission."
}
Write-Host "Push completed. Render will deploy the latest commit."
