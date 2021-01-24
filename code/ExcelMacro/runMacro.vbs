' ＜コマンドプロンプトでの実行方法＞
' 　cscript runMacro.vbs "Excelファイル名_フルパス" "実行するマクロ名"'

Option Explicit

' 使用する変数を定義
Dim excelApp,excel,file,macro

' 引数を取得
file = WScript.Arguments(0)
macro = WScript.Arguments(1)

'エラー発生時はエラーを無視する
On Error Resume Next

' Excelの起動と設定
Set excelApp = CreateObject("Excel.Application")
excelApp.Visible = False            ' Excelを非表示にする'
excelApp.DisplayAlerts = False      ' ポップアップメッセージを非表示にする'
excelApp.AutomationSecurity = 1     ' マクロを有効にする'

' ブックを開く'
Set excel = excelApp.Workbooks.Open(file)

WScript.Echo "---マクロ実行中---"
WScript.Echo "   ファイル：" & file
WScript.Echo "   マクロ：" & macro

' マクロの実行
excelApp.Run macro

' ブックの上書き保存
excel.Save

' エラー判定
If Err.Number <> 0 Then
    WScript.Echo "エラーが発生しました。"
    WScript.Echo "エラー番号：" & Err.Number & " " & "エラー内容：" & Err.Description
End If

' エラーの無視はここまで
On Error Goto 0

WScript.Echo "---マクロの実行が完了しました---"

' Excelの終了
excelApp.Quit
Set excelApp = Nothing