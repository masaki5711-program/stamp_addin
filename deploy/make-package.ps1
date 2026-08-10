# 配布パッケージ作成スクリプト
# 使い方:  .\deploy\make-package.ps1 -HostUrl "https://<ユーザー名>.github.io/<リポジトリ名>"
# 出力:    deploy/stampset-setup.zip(これを配布する)
param(
    [Parameter(Mandatory = $true)]
    [string]$HostUrl
)
$ErrorActionPreference = "Stop"
$deploy = $PSScriptRoot

$HostUrl = $HostUrl.TrimEnd("/")
if ($HostUrl -notmatch "^https://") {
    throw "HostUrl は https:// で始まる必要があります(Officeアドインの要件): $HostUrl"
}

# 1. テンプレートから本番マニフェストを生成
$template = Get-Content (Join-Path $deploy "manifest.prod.template.xml") -Raw
$manifest = $template -replace "\{\{HOST_URL\}\}", $HostUrl
$manifestPath = Join-Path $deploy "manifest.prod.xml"
Set-Content -Path $manifestPath -Value $manifest -Encoding UTF8
Write-Host "生成: $manifestPath(URL: $HostUrl)"

# 2. インストーラーフォルダにコピーして zip 化
Copy-Item $manifestPath (Join-Path $deploy "installer\manifest.prod.xml") -Force
$zipPath = Join-Path $deploy "stampset-setup.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }
Compress-Archive -Path (Join-Path $deploy "installer\*") -DestinationPath $zipPath
Write-Host "作成: $zipPath"

Write-Host ""
Write-Host "配布パッケージができました。stampset-setup.zip を相手に渡してください。" -ForegroundColor Green
Write-Host "相手の手順: zipを展開 → install.cmd をダブルクリック → Excel再起動"
Write-Host ""
Write-Host "※ 事前に src/ の中身が $HostUrl で公開されている必要があります。"
Write-Host "   検証コマンド: npx office-addin-manifest validate deploy/manifest.prod.xml"
