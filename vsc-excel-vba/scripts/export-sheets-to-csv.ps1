. "$PSScriptRoot\functions.ps1"

$fmtCSV = 6
$fmtXLSM = 52

$filepath = $Args[0]
$distdir = $Args[1]

try {
    $bk = getWorkBooks $filepath $false
    $app = $bk.Application

    $app.ScreenUpdating = $false
    $app.DisplayStatusBar = $false
    $app.EnableEvents = $false
    # $app.Visible = $false

    $app.displayAlerts = $false
    $orgsheet = $bk.ActiveSheet
    $bkPath =  $bk.FullName
    $null = New-Item $distdir -ItemType Directory -Force 
    foreach($ws in $bk.Worksheets){
        $ws.Activate()
        $app.ActiveWindow.DisplayFormulas = $true
        $csvpath = Join-Path $distdir "$($ws.name).csv"
        $ws.SaveAs($csvpath, $fmtCSV)
        $app.ActiveWindow.DisplayFormulas = $false
    }

    $app.ScreenUpdating = $true
    $app.DisplayStatusBar = $true
    $app.EnableEvents = $true
    # $app.Visible = $true

    $orgsheet.Activate()
    $bk.SaveAs($bkPath, $fmtXLSM)

    Set-Location -Path $distdir
    ls -file -filter *.csv | % { (get-content -encoding Default $_.FullName) -join "`r`n" | set-content -encoding Default $_.FullName }
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
