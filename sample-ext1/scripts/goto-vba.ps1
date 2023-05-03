. "$PSScriptRoot\functions.ps1"

$bookname = $Args[0]
$moudlename = $Args[1]
$selline = [int]$Args[2]

try {
    $bk = getWorkBooks($bookname)
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$cmp = $bk.VBProject.VBComponents.Item($moudlename)
$cmod = $cmp.CodeModule
$line = $cmod.Lines($selline, 1)
$len = $line.Length + 1
$cmod.CodePane.SetSelection($selline, 1, $selline, $len)
$cmod.CodePane.Show()

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr