$p = 'E:\OH-workspace\hopeOffice\manifest.xml'
$content = [System.IO.File]::ReadAllText($p, [System.Text.Encoding]::UTF8)
$content = $content.Replace('<CustomTab id="hopeOfficeTab">', '<OfficeTab id="TabHome">')
$content = $content.Replace('</CustomTab>', '</OfficeTab>')
[System.IO.File]::WriteAllText($p, $content, (New-Object System.Text.UTF8Encoding $true))
Write-Host "Done"
