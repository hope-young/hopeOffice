$results = Select-String -Path 'E:\OH-workspace\hopeOffice\node_modules\@types\office-js\index.d.ts' -Pattern 'getActive|getSelected|SelectedRange|ActiveChart|ActiveCell' | Select-Object -First 15
foreach ($r in $results) {
    Write-Host $r.Line.Substring(0, [Math]::Min(180, $r.Line.Length))
    Write-Host '---'
}
