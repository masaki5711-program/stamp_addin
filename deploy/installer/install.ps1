# デート印アドイン 簡易インストーラー(Windows / Microsoft 365版Excel用)
# レジストリ(HKCU\...\WEF\Developer)へのマニフェスト登録を自動化する。管理者権限は不要。
$ErrorActionPreference = "Stop"

$AddinId = "a512f6de-9cea-42a8-8f5e-a50782575550"
$AddinName = "便利アドインセット(デート印)"

$src = Join-Path $PSScriptRoot "manifest.prod.xml"
if (-not (Test-Path $src)) {
    Write-Host "エラー: manifest.prod.xml が見つかりません。zipを展開したフォルダのまま実行してください。" -ForegroundColor Red
    exit 1
}

# マニフェストを利用者のローカルにコピー(zipフォルダを消しても動くように)
$dstDir = Join-Path $env:LOCALAPPDATA "StampSet"
New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
$dst = Join-Path $dstDir "manifest.prod.xml"
Copy-Item $src $dst -Force

# Excelのアドイン登録(公式のサイドロード方式)
$key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $key)) { New-Item -Path $key -Force | Out-Null }
New-ItemProperty -Path $key -Name $AddinId -Value $dst -PropertyType String -Force | Out-Null

Write-Host ""
Write-Host "$AddinName をインストールしました。" -ForegroundColor Green
Write-Host "Excelを再起動すると、ホームタブの右端に「デート印」ボタンが表示されます。"
Write-Host "(利用にはインターネット接続と Microsoft 365 版の Excel が必要です)"
Write-Host ""
