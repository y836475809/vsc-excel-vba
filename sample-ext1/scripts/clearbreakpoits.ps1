. "$PSScriptRoot\functions.ps1"

$bookname = $Args[0]

try {
    $bk = getWorkBooks($bookname)

    $debugItems = $bk.Application.VBE.CommandBars(1).Controls(6)

    $clearbp = $debugItems.Controls(10)
    $clearbp.Execute()
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr