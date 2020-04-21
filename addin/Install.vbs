' 参考： [VBScript で Excel にアドインを自動でインストール/アンインストールする方法: ある SE のつぶやき](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

On Error Resume Next

Dim installPath
Dim IsJA
Dim addInName
Dim addInFileName
Dim strMessage
Dim objExcel
Dim objAddin
Dim objWshShell
Dim objFileSys

Function IIf(ByVal str, ByVal trueval, ByVal falseval)
    Dim rtn
    If str Then
        rtn = trueval
    Else
        rtn = falseval
    End If
    IIf = rtn
End Function

IsJA = GetLocale() = 1041

'アドイン情報を設定 
addInName = IIf(IsJA, "CSV変換", "Convert CSV To Table")
addInFileName = "ConvertCsvToTable.xlam"

IF MsgBox(IIf(IsJA, "アドインをインストールしますか？", "Do you want to install this add-in ?"), vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成
'(ex)C:\Users\[User]\AppData\Roaming\Microsoft\AddIns\[addInFileName]
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイルコピー(上書き)
objFileSys.CopyFile  addInFileName ,installPath , True

Set objWshShell = Nothing
Set objFileSys = Nothing

'Excel インスタンス化
Set objExcel = CreateObject("Excel.Application")

'アドイン Workbook タイトル設定
Set objWorkbook = objExcel.Workbooks.Open(installPath)
objWorkbook.Title = addInName
objWorkbook.Save
objWorkbook.Close

objExcel.Workbooks.Add

'アドイン登録
Set objAddin = objExcel.AddIns.Add(installPath, True)
objAddin.Installed = True

'Excel 終了
objExcel.Quit

Set objAddin = Nothing
Set objExcel = Nothing

IF Err.Number = 0 THEN
    MsgBox IIf(IsJA, "アドインのインストールが完了しました", "Installation is now complete !"), vbInformation, addInName
ELSE
    MsgBox IIf(IsJA, "エラーが発生しました" & vbCrLF & "実行環境を確認してください", "An error has occurred." & vbCrLF & "Please check your environment."), vbExclamation, addInName
End IF
