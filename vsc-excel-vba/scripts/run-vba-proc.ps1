. "$PSScriptRoot\functions.ps1"

$filepath = $Args[0]
$procname = $Args[1]
$bookname = getFileName $filepath

try {
    $bk = getWorkBooks $filepath
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}
$bk.Application.Run("${bookname}!${procname}")
$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
