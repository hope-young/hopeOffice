$url = 'https://hope-young.github.io/hopeOffice/assets/taskpane-CnxcdbFJ.js'
$dst = 'E:\OH-workspace\hopeOffice\_bundle.js'
try {
    $wc = New-Object System.Net.WebClient
    $wc.DownloadFile($url, $dst)
    $size = (Get-Item $dst).Length
    Write-Host "downloaded: $size bytes"
} catch {
    Write-Host "fail: $($_.Exception.Message)"
    exit 1
}
$content = [System.IO.File]::ReadAllText($dst)
$checks = @(
    @{ name = 'word substring';  pat = 'word' },
    @{ name = 'browser (no Office)'; pat = 'browser (no Office)' },
    @{ name = '.includes(';  pat = '.includes(' },
    @{ name = 'hostToLabel'; pat = 'hostToLabel' },
    @{ name = 'mapHostType'; pat = 'mapHostType' }
)
foreach ($c in $checks) {
    $hit = $content.Contains($c.pat)
    Write-Host ("  {0,-30} {1}" -f $c.name, $hit)
}
