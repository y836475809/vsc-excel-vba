. "$PSScriptRoot\functions.ps1"

$BookName = $Args[0]

try {
    $bk = getWorkBooks $bookname $false

    $shtNames = @()
    foreach($sht in $bk.Sheets){
        $name = $sht.Name
        if ($name -ne "") {
            $shtNames += $name
        }
    }
    $json = @{
        "sheetnames" = $shtNames;
    }
    $data = ConvertTo-Json $shtNames
    $jsonStr = getResJson "ok" "" $data
    Write-Output $jsonStr
} catch {
    $jsonStr = getResJson "error" $_.Exception.Message ""
    Write-Output $jsonStr
    exit
}
