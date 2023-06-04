. "$PSScriptRoot\functions.ps1"

$vbext_vm_Run = 0
$vbext_vm_Break = 1
$vbext_vm_Design = 2

$bookname = $Args[0]
$srcdir = $Args[1]

try {
    $bk = getWorkBooks $bookname
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}

$project = getVBProject $bk $bookname
$mode = $project.Mode
if(($mode -eq $vbext_vm_Run) -Or ($mode -eq $vbext_vm_Break)){
    $msg = "Cannot import. because it is running. please stop running"
    $jsonStr = getResJson "error" $msg ""
    Write-Output $jsonStr
    exit
}

$filelist = Get-ChildItem $srcdir -File -Recurse -Include *.bas,*.cls
foreach($item in $filelist){
    $filepath = $item.FullName
    $filename = Split-Path $filepath -Leaf
    $filename = [IO.Path]::GetFileNameWithoutExtension($filename)
    $module = $bk.VBProject.VBComponents.Item($filename)

    if($module -ne $null){ 
        $bk.VBProject.VBComponents.Remove($module)
        $ret = $bk.VBProject.VBComponents.Import($filepath)
    }else{
        $ret = $bk.VBProject.VBComponents.Import($filepath)
    }
}

$jsonStr = getResJson "ok" "" ""
Write-Output $jsonStr

# pause