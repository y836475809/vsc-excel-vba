
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


