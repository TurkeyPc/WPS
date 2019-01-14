#なんしかセキュリティで実行できないのはつまんないのでこれでよしなに。
Set-ExecutionPolicy Unrestricted

#拡張子の起動選択に「別窓」を追加してデフォルトの挙動に設定する
function CreateOpenNewItem2($xlsx){

    $newname = CreateOpenNewItem1($xlsx)

    #上で追加した「別窓」を拡張子のデフォルト（ダブルクリック時）挙動に設定する
    Set-ItemProperty $xlsx -name "(default)" -value "OpenNewW"
}

#拡張子の起動選択に「別窓」を追加する
function CreateOpenNewItem1($xlsx){

    $newkeyname = "OpenNewW"
    $newitem = $xlsx + "\" + $newkeyname

    #あったらば消しさる
    if(Test-Path $newitem){
        Remove-Item $newitem -Recurse
    }

    #新しいウインドウに表示するコマンド文字列を生成
    #といってもopenのcommandに"/n"つけるだけ。
    $opencommand = Get-ItemPropertyValue ($xlsx + "\open\command") "(default)"
    $opennewcommand = $opencommand -replace '(^"[^"]*")\s("[^"]*")','$1 "/n" $2'
    #$opennewcommand = $opencommand + ' "/n"'

    #ここは単純にレジストリのキーを作って既定の値（右クリック時の表示名）を設定するだけ
    CreateKey $newitem "別窓(&B)"
    
    #↑でつくったキーの下にcommandキーを作って既定値には
    CreateKey ($newitem + "\command") $opennewcommand
}

#レジストリのキーをつくり既定値を設定する
function CreateKey($item,$defaultvalue){
    New-Item $item
    Set-ItemProperty $item -name "(default)" -value $defaultvalue
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

CreateOpenNewItem2 ($rt + ":\Excel.Sheet.12\shell")
CreateOpenNewItem2 ($rt + ":\Excel.Sheet.8\shell")