. "$PSScriptRoot\functions.ps1"

$vbext_ct_StdModule = 1
$vbext_ct_ClassModule = 2

$filepath = $Args[0]
$distdir = $Args[1]

try {
    $bk = getWorkBooks $filepath
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}
try {
    $null = New-Item $distdir -ItemType Directory -Force 
    foreach($vbcmp in $bk.VBProject.VBComponents){
        $ext = $null
        if($vbcmp.Type -eq $vbext_ct_StdModule){
            $ext = "bas"
        }
        if($vbcmp.Type -eq $vbext_ct_ClassModule){
            $ext = "cls"
        }
        if($ext -ne $null){
            $cmname = $vbcmp.Name
            $distfilepath = Join-Path $distdir "${cmname}.${ext}"
            $vbcmp.Export($distfilepath)
        }
    }
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}
$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr
