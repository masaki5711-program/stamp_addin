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
Write-Host "作成: $zipPath(利用者に配るインストーラー)"

# 3. サーバー配置用一式(アドイン本体 + ダウンロードページ + インストーラーzip + マニフェスト)
$src = Join-Path (Split-Path $deploy -Parent) "src"
$siteZip = Join-Path $deploy "site.zip"
if (Test-Path $siteZip) { Remove-Item $siteZip -Force }
Compress-Archive -Path (Join-Path $src "*"), $zipPath, $manifestPath -DestinationPath $siteZip
Write-Host "作成: $siteZip(サーバー担当者に渡す一式)"

Write-Host ""
Write-Host "パッケージができました。" -ForegroundColor Green
Write-Host "・site.zip          → サーバー担当者へ。$HostUrl の直下に展開してもらう"
Write-Host "                      (展開後 $HostUrl/index.html がダウンロードページになる)"
Write-Host "・stampset-setup.zip → 利用者へ直接渡す場合はこちら。展開 → install.cmd → Office再起動"
Write-Host ""
Write-Host "検証コマンド: npx office-addin-manifest validate deploy/manifest.prod.xml"
