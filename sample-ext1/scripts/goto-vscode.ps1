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
$module_name = $cm.Name
$module_type = ""
if($cm.Parent.Type -eq $vbext_ct_StdModule){
    $module_type = "bas"
}
if($cm.Parent.Type -eq $vbext_ct_ClassModule){
    $module_type = "cls"
}
$cmp = $bk.VBProject.VBComponents($module_name)
$startline = -1
$startcol = -1
$endline = -1
$endcol = -1
$cmp.CodeModule.CodePane.GetSelection(
    [ref]$startline, [ref]$startcol, [ref]$endline, [ref]$endcol)

Add-Type -AssemblyName System.Web
$module_name = [System.Web.HttpUtility]::UrlEncode($cm.Name)
$json = @{
    "module_name" = $module_name;
    "module_type" = $module_type;
    "startline" = $startline;
    "startcol" = $startcol;
    "endline" = $endline;
    "endcol" = $endcol;
}
$data = ConvertTo-Json $json
$jsonStr = getResJson "ok" "" $data
Write-Output $jsonStr
