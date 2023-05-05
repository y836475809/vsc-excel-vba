. "$PSScriptRoot\functions.ps1"

$bookname = $Args[0]
$moudlename = $Args[1]
$debugline = [int]$Args[2]

try {
    $bk = getWorkBooks($bookname)

    $cmp = $bk.VBProject.VBComponents.Item($moudlename)
    $cmp.CodeModule.CodePane.SetSelection($debugline, 1, $debugline, 1)
    $cmp.CodeModule.CodePane.Show()

    $debugItems = $bk.Application.VBE.CommandBars(1).Controls(6)
    $togglebp = $debugItems.Controls(9)
    $togglebp.Execute()
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr