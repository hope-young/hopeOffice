for ($i = 1; $i -le 5; $i++) {
    Write-Host "--- attempt $i ---"
    $a = (Invoke-WebRequest -Uri 'https://hope-young.github.io/hopeOffice/manifest.xml' -UseBasicParsing -TimeoutSec 15 -ErrorAction SilentlyContinue).StatusCode
    $b = (Invoke-WebRequest -Uri 'https://hope-young.github.io/hopeOffice/src/taskpane/index.html' -UseBasicParsing -TimeoutSec 15 -ErrorAction SilentlyContinue).StatusCode
    Write-Host "manifest=$a  taskpane=$b"
    Start-Sleep -Milliseconds 800
}
