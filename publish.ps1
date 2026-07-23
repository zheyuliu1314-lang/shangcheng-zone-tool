param(
    [string]$Message = "更新片区分类工具"
)

$ErrorActionPreference = "Stop"
$git = "C:\Users\wy98k\.cache\codex-runtimes\codex-primary-runtime\dependencies\native\git\cmd\git.exe"
if (-not (Test-Path -LiteralPath $git)) {
    $gitCommand = Get-Command git -ErrorAction SilentlyContinue
    if ($gitCommand) { $git = $gitCommand.Source }
}
if (-not (Test-Path -LiteralPath $git)) {
    throw "未找到 Git，请先安装 Git for Windows。"
}

& $git add app.py templates requirements.txt render.yaml netlify.toml netlify-site scripts zones.json README.md .gitignore publish.ps1
if ((& $git diff --cached --quiet) -and ($LASTEXITCODE -eq 0)) {
    Write-Host "没有新的代码修改。"
    exit 0
}
& $git commit -m $Message
$token = $env:GITHUB_TOKEN
if (-not $token) {
    $secureToken = Read-Host "请输入 GitHub fine-grained token（仅本次使用，不会写入文件）" -AsSecureString
    $token = [Runtime.InteropServices.Marshal]::PtrToStringBSTR([Runtime.InteropServices.Marshal]::SecureStringToBSTR($secureToken))
}
if (-not $token) { throw "未提供 GitHub token，已取消推送。" }
$env:GIT_TERMINAL_PROMPT = "0"
& $git -c "http.extraheader=AUTHORIZATION: bearer $token" push origin main
if ($LASTEXITCODE -ne 0) { throw "GitHub 推送失败，请检查 token 是否拥有该仓库的 Contents: Read and write 权限。" }
Write-Host "代码已推送到 GitHub，Render 将自动部署。"
