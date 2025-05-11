[Console]::OutputEncoding = [System.Text.Encoding]::GetEncoding('utf-8')

function getResJson($code, $message, $data){
    $json = @{
        "code" = $code;
        "message" = $message;
        "data" = $data;
    }
    $jsonStr = ConvertTo-Json $json
    return $jsonStr
}

function getFileName([string]$filepath){
    return (Get-Item $filepath ).Name
}

function showVBEditor($bk, $visible){
    if($visible -And ($bk.Application.VBE.MainWindow.Visible -eq $false)){
        $bk.Application.VBE.MainWindow.Visible = $true
    } 
}

function getWorkBooks {
    param ([string]$filepath, [bool]$showVBE=$true)
    $bookname = getFileName $filepath
    try {   
        $ex = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        foreach ($bk in $ex.WorkBooks){
            if ($bk.Name -eq $bookname){
                showVBEditor $bk $showVBE
                return $bk
            }
        }
    } catch {
        # 
    }
    try {
        $excelApp = New-Object -ComObject "Excel.Application"
        $bk = $excelApp.Workbooks.Open($filepath)
        $excelApp.Visible = $true
        showVBEditor $bk $showVBE
        return $bk
    } catch {
        $excelApp.Quit()
        $excelApp = $null
        $bk = $null
        [GC]::Collect()
        $exp_msg = $_.Exception.Message
        $msg = "Not found excel, target=$bookname, $exp_msg"
        throw $msg
    }
}

function getVBComponent($proj, $name){
    foreach($comp in $proj.VBComponents){
        if($comp.Name -eq $name){
            return $comp
        }
    }
    return $null
}

function getVBProject($book, $bookname) {
    foreach ($project in $book.Application.VBE.VBProjects){
        $filename = Split-Path $project.FileName -Leaf
        if($filename -eq $bookname){
            return $project
        }
    }
    $msg = "Not found VBProject, target=$bookname" 
    throw $msg
}

