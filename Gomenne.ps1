#なんしかセキュリティで実行できないのはつまんないのでこれでよしなに。
Set-ExecutionPolicy Unrestricted

#別窓開きレジストリ登録関数
function CreateOpenNewItem($xlsx){
    Set-ItemProperty $xlsx -name "(default)" -value "open"
}

#HKEY_CLASSES_ROOTをPSDriveのHKCRに割り当てる
#通常はNew-PSDriveだけでいいが
#ISEで2回目以降にエラー出るのがうっとうしいのでいったんHKCRの割り当てを解除してからNew-PSDriveしている
$rt = "HKCR"
Remove-PSDrive -Name $rt
$v = New-PSDrive -PSProvider Registry -Name $rt -Root HKEY_CLASSES_ROOT

CreateOpenNewItem ($rt + ":\Excel.Sheet.12\shell")
CreateOpenNewItem ($rt + ":\Excel.Sheet.8\shell")