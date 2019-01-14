#なんしかセキュリティで実行できないのはつまんないのでこれでよしなに。
Set-ExecutionPolicy Unrestricted

#拡張子のデフォルト（ダブルクリック時）挙動を「開く」に戻す
function RepairRegistryProperty($xlsx){
    Set-ItemProperty $xlsx -name "(default)" -value "open"

    $pa = $xlsx + "\" + "OpenNewW"

    if(Test-Path ($pa)){
        Remove-Item $pa -Recurse
    }
}

function ExistPSDrive($dname){
    $vv = Get-PSDrive
    foreach($aa in $vv){
        if($aa.Name -eq $dname){
            1
            return
        }
    }

    0
    return
}

#HKEY_CLASSES_ROOTをPSDriveのHKCRに割り当てる
#通常はNew-PSDriveだけでいいが
#ISEで2回目以降にエラー出るのがうっとうしいのでいったんHKCRの割り当てを解除してからNew-PSDriveしている
$rt = "HKCR"
if(ExistPSDrive $rt -eq 0){
    
}else{
    New-PSDrive -PSProvider Registry -Name $rt -Root HKEY_CLASSES_ROOT
}

RepairRegistryProperty ($rt + ":\Excel.Sheet.12\shell")
RepairRegistryProperty ($rt + ":\Excel.Sheet.8\shell")