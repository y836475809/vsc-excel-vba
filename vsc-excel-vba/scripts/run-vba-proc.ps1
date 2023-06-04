. "$PSScriptRoot\functions.ps1"

$bookname = $Args[0]
$procname = $Args[1]

try {
    $bk = getWorkBooks($bookname)
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}
$bk.Application.Run("${bookname}!${procname}")
$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
