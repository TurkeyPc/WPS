#なんしかセキュリティで実行できないのはつまんないのでこれでよしなに。
Set-ExecutionPolicy RemoteSigned

#別窓開きレジストリ登録関数
function CreateOpenNewItem($xlsx){

    $newkeyname = "OpenNewW"
    $newitem = $xlsx + "\" + $newkeyname

    #あったらば消しさる
    Remove-Item $newitem -Recurse

    #新しいウインドウに表示するコマンド文字列を生成
    #といってもopenのcommandに"/n"つけるだけ。
    $opencommand = Get-ItemPropertyValue ($xlsx + "\open\command") "(default)"
    $opennewcommand = $opencommand -replace '(^"[^"]*")\s("[^"]*")','$1 "/n" $2'
    #$opennewcommand = $opencommand + ' "/n"'

    #ここは単純にレジストリのキーを作って既定の値（右クリック時の表示名）を設定するだけ
    CreateKey $newitem "別窓(&B)"
    
    #↑でつくったキーの下にcommandキーを作って既定値には
    CreateKey ($newitem + "\command") $opennewcommand

    Set-ItemProperty $xlsx -name "(default)" -value $newkeyname
}

#レジストリのキーをつくり既定値を設定する
function CreateKey($item,$defaultvalue){
    New-Item $item
    Set-ItemProperty $item -name "(default)" -value $defaultvalue
}

#HKEY_CLASSES_ROOTをPSDriveのHKCRに割り当てる
#通常はNew-PSDriveだけでいいが
#ISEで2回目以降にエラー出るのがうっとうしいのでいったんHKCRの割り当てを解除してからNew-PSDriveしている
$rt = "HKCR"
Remove-PSDrive -Name $rt
$v = New-PSDrive -PSProvider Registry -Name $rt -Root HKEY_CLASSES_ROOT

CreateOpenNewItem ($rt + ":\Excel.Sheet.12\shell")
CreateOpenNewItem ($rt + ":\Excel.Sheet.8\shell")