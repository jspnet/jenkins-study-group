# �����̎擾
Param( $macroBook, $procedureName )

# �N������t�@�C����W���o��
$macroBook

# �t�@�C���p�X�̎擾
$macroBookFullPath = (ls $macroBook).FullName
$macroProcName = (ls $macroBook).Name + "!" + $procedureName

# �N������}�N����W���o��
$macroProcName

# Excel�̋N��
$excelApp = New-Object -com "Excel.Application"

try {

  # �}�N���̋N��
  $excelApp.DisplayAlerts = $false
  $book = $excelApp.Workbooks.Open($macroBookFullPath)
  $excelApp.Run($macroProcName)
  
  # �㏑���ۑ�
  $book.Save()
  
} finally {

  # Excel�̏I��
  $excelApp.EnableEvents = $false
  $excelApp.DisplayAlerts = $false
  $excelApp.Visible = $false
  $excelApp.Workbooks | % { $_.Close($false) }
  $excelApp.Quit()
  
}