# Opens every .pptx passed (or found in the given folders) in PowerPoint via
# COM, invisibly, and reports whether it opens cleanly. A file that PowerPoint
# cannot open (or must repair) throws on automation open, which is exactly the
# spike U1/U2 pass criterion. Desktop-interactive COM use is supported by
# Microsoft; this script is for the local dev loop, never for the server.
param(
    [string[]]$Paths = @("out", "fixtures")
)

$files = @()
foreach ($p in $Paths) {
    if (Test-Path $p -PathType Container) {
        $files += Get-ChildItem -Path $p -Filter *.pptx | Select-Object -ExpandProperty FullName
    } elseif (Test-Path $p) {
        $files += (Resolve-Path $p).Path
    }
}
if ($files.Count -eq 0) { Write-Output "No .pptx files found."; exit 1 }

$app = New-Object -ComObject PowerPoint.Application
$app.DisplayAlerts = 1  # ppAlertsNone: suppress dialogs so a bad file throws instead of hanging

$results = @()
foreach ($f in $files) {
    $status = "OK"; $slides = 0; $err = ""
    try {
        # Open(FileName, ReadOnly=msoTrue(-1), Untitled=msoFalse(0), WithWindow=msoFalse(0))
        $pres = $app.Presentations.Open($f, -1, 0, 0)
        $slides = $pres.Slides.Count
        $pres.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres) | Out-Null
    } catch {
        $status = "FAIL"
        $err = $_.Exception.Message -replace "`r`n", " "
    }
    $results += [PSCustomObject]@{ File = Split-Path $f -Leaf; Status = $status; Slides = $slides; Error = $err }
}

$app.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($app) | Out-Null

$results | Format-Table -AutoSize | Out-String -Width 200 | Write-Output
$failed = @($results | Where-Object { $_.Status -eq "FAIL" })
Write-Output ("SUMMARY: {0}/{1} opened cleanly in PowerPoint desktop." -f ($results.Count - $failed.Count), $results.Count)
if ($failed.Count -gt 0) { exit 1 }
