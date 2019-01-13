#なんしかセキュリティで実行できないのはつまんないのでこれでよしなに。
Set-ExecutionPolicy Unrestricted

#拡張子のデフォルト（ダブルクリック時）挙動を「開く」に戻す
#※拡張子の起動選択に「別窓」が残るが害はないので放置
function RepairRegistryProperty($xlsx){
    Set-ItemProperty $xlsx -name "(default)" -value "open"
}

#HKEY_CLASSES_ROOTをPSDriveのHKCRに割り当てる
#通常はNew-PSDriveだけでいいが
#ISEで2回目以降にエラー出るのがうっとうしいのでいったんHKCRの割り当てを解除してからNew-PSDriveしている
$rt = "HKCR"
Remove-PSDrive -Name $rt
$v = New-PSDrive -PSProvider Registry -Name $rt -Root HKEY_CLASSES_ROOT

RepairRegistryProperty ($rt + ":\Excel.Sheet.12\shell")
RepairRegistryProperty ($rt + ":\Excel.Sheet.8\shell")