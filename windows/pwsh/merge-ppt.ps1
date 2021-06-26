# Param([switch]$Param1, [switch]$Param2) 何かフラグが欲しければ

Add-Type -AssemblyName System.Windows.Forms


[string[]]$FileNames = @() # 結合するファイル名のリスト

if ($Args.Length -gt 0) {
  # ファイル名を引数で渡した場合
  foreach ($f in $Args) {
    $FileNames += $f
    }
}elseif(($pipelineArgs = @($input)).Length -gt 0) {
  # ファイル名をパイプラインで渡した場合
  foreach ($f in $pipelineArgs) {
    $FileNames += $f
  }
} else {
  # ファイル名を指定せずファイルダイアログで選択する場合(複数選択)
  $dialog = New-Object System.Windows.Forms.OpenFileDialog
  $dialog.Filter = "PowerPointプレゼンテーション(*.pptx;*.ppt;*.pptm)|*.pptx;*.ppt;*.pptm|すべてのファイル (.)|."
  $dialog.InitialDirectory = [Environment]::GetFolderPath('MyDocuments')
  $dialog.Title = "ファイルを選択してください"
  $dialog.Multiselect = $true # 複数選択を許可したい時は Multiselect を設定する
  $dialog.RestoreDirectory = $true
  

  # ダイアログを表示
  # ダイアログの最前面化 (参照:https://fm-aid.com/bbs2/viewtopic.php?pid=45468#p45468)
  if ($dialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) -eq [System.Windows.Forms.DialogResult]::OK) {
    $FileNames = $dialog.FileNames
  }
}

if ($FileNames.Length -eq 0) {
  exit
}

# ppt の操作
$pptApp = New-Object -ComObject PowerPoint.Application

$baseFileName = $FileNames[0]
$baseSlides = $pptApp.Presentations.Open($baseFileName)

for ($i = 1; $i -lt $FileNames.Length; $i++) {
  $addedSlides = $pptApp.presentations.Open($FileNames[$i])

  $oldBaseCount = $baseSlides.Slides.Count
  $baseSlides.Slides.InsertFromFile($FileNames[$i], $baseSlides.Slides.Count, 1, $addedSlides.Slides.Count) # ファイル名を指定してスライドの挿入
  
  # 結合前のスライドマスターを結合後に適用
  for ($j = 1; $j -le $addedSlides.Slides.Count; $j++) {
    $baseSlides.Slides($oldBaseCount + $j).Design = $addedSlides.Slides($j).Design
  }

  $addedSlides.Close()
}

try{
  $SaveFileDialog = New-Object -TypeName System.Windows.Forms.SaveFileDialog
  $SaveFileDialog.DefaultExt = "pptx"
  $SaveFileDialog.FileName = "Merged_" + (Split-Path -Leaf $baseFileName)
  $SaveFileDialog.Filter = "PowerPointプレゼンテーション(*.pptx;*.ppt;*.pptm)|*.pptx;*.ppt;*.pptm|すべてのファイル (.)|."
  $SaveFileDialog.FilterIndex = 1
  $SaveFileDialog.InitialDirectory = Split-Path -Parent $baseFileName
  $SaveFileDialog.OverwritePrompt = $true
  $SaveFileDialog.ShowHelp = $true
  $SaveFileDialog.Title = "結合ファイルの保存"

  # ダイアログの最前面化 (参照:https://fm-aid.com/bbs2/viewtopic.php?pid=45468#p45468)
  if ($SaveFileDialog.ShowDialog((New-Object System.Windows.Forms.Form -Property @{TopMost = $true })) -eq [System.Windows.Forms.DialogResult]::OK) { 
    $pptApp.ActivePresentation.SaveAs($SaveFileDialog.FileName)
  }

  $pptApp.ActivePresentation.Close()

}finally{
  $pptApp.Quit()

  [gc]::collect()
  [gc]::WaitForPendingFinalizers()
  [gc]::collect()
}