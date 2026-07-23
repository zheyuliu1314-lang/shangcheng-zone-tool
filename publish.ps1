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
& $git push origin main
Write-Host "代码已推送到 GitHub，Render 将自动部署。"
