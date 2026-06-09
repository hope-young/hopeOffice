$urls = @(
    'https://hope-young.github.io/hopeOffice/',
    'https://hope-young.github.io/hopeOffice/manifest.xml',
    'https://hope-young.github.io/hopeOffice/src/taskpane/index.html',
    'https://hope-young.github.io/hopeOffice/assets/icon-16.png',
    'https://hope-young.github.io/hopeOffice/assets/icon-32.png',
    'https://hope-young.github.io/hopeOffice/assets/icon-64.png',
    'https://hope-young.github.io/hopeOffice/assets/icon-80.png',
    'https://hope-young.github.io/hopeOffice/src/taskpane/executor/iframe.html'
)
foreach ($u in $urls) {
    try {
        $req = [System.Net.HttpWebRequest]::Create($u)
        $req.Timeout = 12000
        $req.Method = 'GET'
        $req.AllowAutoRedirect = $true
        $req.UserAgent = 'Mozilla/5.0 (Office Online; Excel)'
        $resp = $req.GetResponse()
        $code = [int]$resp.StatusCode
        $len = $resp.ContentLength
        $ct = $resp.ContentType
        '{0,-72} {1}  {2,-30}  {3} bytes' -f $u, $code, $ct, $len
        $resp.Close()
    } catch {
        if ($_.Exception.Response) {
            '{0,-72} {1}' -f $u, [int]$_.Exception.Response.StatusCode
        } else {
            '{0,-72} ERR  {1}' -f $u, $_.Exception.Message
        }
    }
}
