'// 前準備
'// 本スクリプトがあるフォルダのパスを取得
Dim objFileSystemObject     '// ファイルシステムオブジェクト
Dim strFolderPath           '// .vbs のあるフォルダのフォルダパス
Dim strMacroPath            '// 実行するマクロのパス
Set objFileSystemObject = WScript.CreateObject("Scripting.FileSystemObject")
strFolderPath = objFileSystemObject.getParentFolderName(WScript.ScriptFullName)
strMacroPath = strFolderPath & "\" & "hcpMakerMacro.xlsm"
'msgbox strMacroPath

'// メイン処理
Dim objExcel

'// テキストファイルのフルパスをセット
filePath = Wscript.Arguments(0)

'// Excelのオブジェクトを作成する
Set objExcel = CreateObject("Excel.Application")

'// Excelを見える形で表示させる。
objExcel.Application.Visible = true

'// Excelのパスとシート名を選択
objExcel.Workbooks.Open(strMacroPath)
objExcel.worksheets("Sheet1").select

'// マクロを実行
'// 引数1：実行するマクロ名（Sheet1のマクロの場合は、"Sheet1."が必要）
'// 引数2：ファイルのフルパスをセット（Sheet1.Mainマクロの引数1）
objExcel.Application.Run "Sheet1.Main", filePath

Set objExcel = Nothing