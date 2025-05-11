. "$PSScriptRoot\functions.ps1"

$filepath = $Args[0]
$sheetname = $Args[1]

try {
    $bk = getWorkBooks $filepath $false
    foreach($ws in $bk.Worksheets){
        if($ws.name -eq $sheetname){
            $ws.Activate()
            break
        }
    }
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
