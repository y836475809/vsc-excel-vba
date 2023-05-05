. "$PSScriptRoot\functions.ps1"

$bookname = $Args[0]
$bplocs = $null
if($Args.Length -gt 1){
    $bplocs = $Args[1..($Args.Length-1)]
}
# bplocs m1:1-2-3 m2:2

try {
    $bk = getWorkBooks($bookname)
    $debugItems = $bk.Application.VBE.CommandBars(1).Controls(6)
    $clearbp = $debugItems.Controls(10)
    $clearbp.Execute()
    
    if($bplocs -eq $null){
        $jsonStr = getResJson "ok" "" ""
        Write-Output $jsonStr
        exit
    }

    foreach($bploc in $bplocs){
        $bp = $bploc -split ":"
        $moudlename = $bp[0]
        $lines = $bp[1] -split "-"

        $cmp = $bk.VBProject.VBComponents.Item($moudlename)
        $cmp.CodeModule.CodePane.Show()

        foreach($line in $lines){
            $cmp.CodeModule.CodePane.SetSelection([int]$line, 1, [int]$line, 1)
            $togglebp = $debugItems.Controls(9)
            $togglebp.Execute()
        }
    } 
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr