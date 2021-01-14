# 引数の取得
Param( $macroBook, $procedureName )

# 起動するファイルを標準出力
$macroBook

# ファイルパスの取得
$macroBookFullPath = (ls $macroBook).FullName
$macroProcName = (ls $macroBook).Name + "!" + $procedureName

# 起動するマクロを標準出力
$macroProcName

# Excelの起動
$excelApp = New-Object -com "Excel.Application"

try {

  # マクロの起動
  $excelApp.DisplayAlerts = $false
  $book = $excelApp.Workbooks.Open($macroBookFullPath)
  $excelApp.Run($macroProcName)
  
  # 上書き保存
  $book.Save()
  
} finally {

  # Excelの終了
  $excelApp.EnableEvents = $false
  $excelApp.DisplayAlerts = $false
  $excelApp.Visible = $false
  $excelApp.Workbooks | % { $_.Close($false) }
  $excelApp.Quit()
  
}