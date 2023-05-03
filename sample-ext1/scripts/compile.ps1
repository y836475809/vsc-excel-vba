. "$PSScriptRoot\functions.ps1"

$bookName = $Args[0]
try {
    $bk = getWorkBooks($bookName)
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$cm = $bk.Application.VBE.CommandBars(1).Controls(6).Controls(1)
$cm.Execute()
$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
