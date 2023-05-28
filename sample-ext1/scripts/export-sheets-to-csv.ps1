. "$PSScriptRoot\functions.ps1"

$fmtCSV = 6
$fmtXLSM = 52

try {    
    $bookname = $Args[0]
    $distdir = $Args[1]

    $bk = getWorkBooks $bookname $false
    $bk.Application.displayAlerts = $false
    $orgsheet = $bk.ActiveSheet
    $bkPath =  $bk.FullName
    $null = New-Item $distdir -ItemType Directory -Force 
    foreach($ws in $bk.Worksheets){
        $ws.Activate()
        $csvpath = Join-Path $distdir "$($ws.name).csv"
        $ws.SaveAs($csvpath, $fmtCSV)
    }
    $orgsheet.Activate()
    $bk.SaveAs($bkPath, $fmtXLSM)

} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
