. "$PSScriptRoot\functions.ps1"

$BookName = $Args[0]

$vbext_ct_StdModule = 1
$vbext_ct_ClassModule = 2

try {
    $bk = getWorkBooks($bookname)
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$cm = $bk.Application.VBE.ActiveCodePane.CodeModule
$objectName = $cm.Name
$objectType = ""
if($cm.Parent.Type -eq $vbext_ct_StdModule){
    $objectType = "bas"
}
if($cm.Parent.Type -eq $vbext_ct_ClassModule){
    $objectType = "cls"
}
$cmp = $bk.VBProject.VBComponents($objectName)
$startline = -1
$startcol = -1
$endline = -1
$endcol = -1
$cmp.CodeModule.CodePane.GetSelection(
    [ref]$startline, [ref]$startcol, [ref]$endline, [ref]$endcol)

$json = @{
    "objectname" = $objectName;
    "objecttype" = $objectType;
    "startline" = $startline;
    "startcol" = $startcol;
    "endline" = $endline;
    "endcol" = $endcol;
}
$data = ConvertTo-Json $json
$jsonStr = getResJson "ok" "" $data
Write-Output $jsonStr
