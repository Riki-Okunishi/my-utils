# pwsh

PowerShell 向け


## 一覧

+ `merge-ppt.ps1`
  + 複数の `pptx` ファイルを1つにマージするスクリプト
  + スライドマスターを保持しつつマージを自動化
+ `Microsoft.PowerShell_profile.ps1`
  + PowerShellのエイリアスを定義する


## 詳細

### `merge-ppt.ps1`

複数のpptxファイルを1つにマージする．

**スライドマスターは保持した上でマージされる**．

マージ後ダイアログが開きファイル名と保存場所を指定する．

#### 使い方

引数でファイルを指定する場合

```powershell
# sample1.pptx, sample2.pptx, ... がマージされる
> .\merge-ppt.ps1 sample1.pptx sample2.pptx ...
```

パイプラインでファイルを指定する場合

```powershell
# ls コマンドには -Name オプションを付ける
> ls -N | .\merge-ppt.ps1
```

ファイル選択ダイアログで指定する場合

```powershell
> .\merge-ppt.ps1
```

### `Microsoft.PowerShell_profile.ps1`

PowerShellのエイリアスを定義する．

配置場所は，`$HOME/Documents/PowerShell/`の中．

記述した関数はエイリアスとしてPowerShellターミナルから呼び出せる

#### エイリアス一覧

+ `wslsd`
  + WSLのシャットダウンのエイリアス
  + `wsl.exe --shutdown`と等価