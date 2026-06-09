$url = 'https://raw.githubusercontent.com/hope-young/hope-Office/master/manifest.production.xml'
$dst = 'E:\OH-workspace\hopeOffice\_auto_prod.xml'
$wc = New-Object System.Net.WebClient
$wc.DownloadFile($url, $dst)
(Get-Item $dst).Length
