# 1. 杀 Excel
Get-Process | Where-Object { $_.ProcessName -eq 'excel' } | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2

# 2. 删所有 WEF 目录残留 (含 "hopeOffice" 或 "hopeOffice (2026)")
$wef = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\WEF'
if (Test-Path $wef) {
    Get-ChildItem $wef -Directory -ErrorAction SilentlyContinue | ForEach-Object {
        $n = $_.Name
        if ($n -like '*hope*' -or $n -like '*Hope*') {
            Write-Host "rm WEF: $n"
            Remove-Item -Recurse -Force $_.FullName -ErrorAction SilentlyContinue
        }
    }
}

# 3. 清 Office MediaCache (icon 缓存)
$media = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\16.0\MediaCache'
if (Test-Path $media) {
    Write-Host "rm MediaCache"
    Remove-Item -Recurse -Force $media -ErrorAction SilentlyContinue
}

# 4. 清 Office WEF 注册表残留 (含 TrustedCatalogs)
$wefRoot = 'HKCU:\Software\Microsoft\Office\16.0\WEF'
if (Test-Path $wefRoot) {
    Get-ChildItem $wefRoot -Recurse -ErrorAction SilentlyContinue | ForEach-Object {
        $path = $_.PSPath
        if ($path -like '*hope*' -or $path -like '*Hope*') {
            Write-Host "rm reg: $path"
            Remove-Item -Recurse -Force $path -ErrorAction SilentlyContinue
        }
    }
}

# 5. 装 production manifest
Set-Location E:\OH-workspace\hopeOffice
pwsh installer/install.ps1 -Production

# 6. 启 Excel
Start-Process excel
Write-Host "done"
