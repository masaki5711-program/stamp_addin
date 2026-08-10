# デート印アドイン アンインストーラー
$ErrorActionPreference = "Stop"

$AddinId = "a512f6de-9cea-42a8-8f5e-a50782575550"
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"

if (Test-Path $key) {
    Remove-ItemProperty -Path $key -Name $AddinId -ErrorAction SilentlyContinue
}
$dstDir = Join-Path $env:LOCALAPPDATA "StampSet"
if (Test-Path $dstDir) {
    Remove-Item -Recurse -Force $dstDir
}

Write-Host ""
Write-Host "アンインストールしました。Excelを再起動すると「デート印」ボタンが消えます。" -ForegroundColor Green
Write-Host "(作成した印のデータが不要な場合、Excel内のタスクペインのデータは自動的に使われなくなります)"
Write-Host ""
