' ���R�}���h�v�����v�g�ł̎��s���@��
' �@cscript runMacro.vbs "Excel�t�@�C����_�t���p�X" "���s����}�N����"'

Option Explicit

' �g�p����ϐ����`
Dim excelApp,excel,file,macro

' �������擾
file = WScript.Arguments(0)
macro = WScript.Arguments(1)

'�G���[�������̓G���[�𖳎�����
On Error Resume Next

' Excel�̋N���Ɛݒ�
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False            ' Excel���\���ɂ���'
excelApp.DisplayAlerts = False      ' �|�b�v�A�b�v���b�Z�[�W���\���ɂ���'
excelApp.AutomationSecurity = 1     ' �}�N����L���ɂ���'

' �u�b�N���J��'
Set excel = excelApp.Workbooks.Open(file)

WScript.Echo "---�}�N�����s��---"
WScript.Echo "   �t�@�C���F" & file
WScript.Echo "   �}�N���F" & macro

' �}�N���̎��s
excelApp.Run macro

' �u�b�N�̏㏑���ۑ�
excel.Save

' �G���[����
If Err.Number <> 0 Then
    WScript.Echo "�G���[���������܂����B"
    WScript.Echo "�G���[�ԍ��F" & Err.Number & " " & "�G���[���e�F" & Err.Description
End If

' �G���[�̖����͂����܂�
On Error Goto 0

WScript.Echo "---�}�N���̎��s���������܂���---"

' Excel�̏I��
excelApp.Quit
Set excelApp = Nothing