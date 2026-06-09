# Simulate Excel online's validation pass: walk the manifest, hit every URL
# with a GET, and report. The service-side validator does the same on upload.
$ErrorActionPreference = 'Stop'

$manifestUrl = 'https://hope-young.github.io/hopeOffice/manifest.xml'
$out = 'E:\OH-workspace\hopeOffice\_manifest.xml'

$wc = New-Object System.Net.WebClient
$wc.DownloadFile($manifestUrl, $out)
$xml = [xml](Get-Content $out -Raw)
$ns = @{ o = 'http://schemas.microsoft.com/office/appforoffice/1.1'; bt = 'http://schemas.microsoft.com/office/officeappbasictypes/1.0' }

$urls = @()
# Top-level DefaultValue URLs
foreach ($el in @('IconUrl', 'HighResolutionIconUrl', 'SupportUrl')) {
    $u = $xml.OfficeApp.$el.DefaultValue
    if ($u) { $urls += [PSCustomObject]@{ Kind = $el; Url = $u } }
}
# AppDomain
$ad = $xml.OfficeApp.AppDomains.AppDomain
if ($ad) { $urls += [PSCustomObject]@{ Kind = 'AppDomain'; Url = $ad } }
# DefaultSettings SourceLocation
$sl = $xml.OfficeApp.DefaultSettings.SourceLocation.DefaultValue
if ($sl) { $urls += [PSCustomObject]@{ Kind = 'DefaultSettings.SourceLocation'; Url = $sl } }
# bt:Image / bt:Url DefaultValue URLs in Resources
$resources = $xml.OfficeApp.VersionOverrides.Resources
if ($resources -ne $null) {
    foreach ($u in $resources.Images.Image) {
        $urls += [PSCustomObject]@{ Kind = 'bt:Image'; Url = $u.DefaultValue }
    }
    foreach ($u in $resources.Urls.Url) {
        $urls += [PSCustomObject]@{ Kind = 'bt:Url'; Url = $u.DefaultValue }
    }
}

Write-Host "Found $($urls.Count) URLs to verify"
foreach ($u in $urls) {
    try {
        $req = [System.Net.HttpWebRequest]::Create($u.Url)
        $req.Timeout = 12000
        $req.Method = 'GET'
        $req.AllowAutoRedirect = $true
        $req.UserAgent = 'Mozilla/5.0'
        $resp = $req.GetResponse()
        $code = [int]$resp.StatusCode
        $len = $resp.ContentLength
        $ct = $resp.ContentType
        '{0,-40} {1}  ct={2,-20} len={3,7}  {4}' -f $u.Kind, $code, $ct, $len, $u.Url
        $resp.Close()
    } catch {
        if ($_.Exception.Response) {
            '{0,-40} {1}  {2}' -f $u.Kind, [int]$_.Exception.Response.StatusCode, $u.Url
        } else {
            '{0,-40} ERR  {1}  {2}' -f $u.Kind, $_.Exception.Message, $u.Url
        }
    }
}
