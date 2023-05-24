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

function getWorkBooks($bookname) {
    try {
        $ex = [System.Runtime.InteropServices.Marshal]::GetActiveObject("Excel.Application")
        foreach ($bk in $ex.WorkBooks){
            if ($bk.Name -eq $bookname){
                if($bk.Application.VBE.MainWindow.Visible -eq $false){
                    $bk.Application.VBE.MainWindow.Visible = $true
                    # Start-Sleep -m 500
                }
                return $bk
            }
        }
    } catch {
        # 
    }
    $msg = "Not found excel, target=$bookname, please open excel" 
    throw $msg
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

