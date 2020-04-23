' 参考： [VBScript で Excel にアドインを自動でインストール/アンインストールする方法: ある SE のつぶやき](http://fnya.cocolog-nifty.com/blog/2014/03/vbscript-excel-.html)

On Error Resume Next

Dim installPath
Dim IsJA
Dim addInName
Dim addInFileName
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

'Excel動作中判定
Err.Clear
Set objExcel = GetObject(, "Excel.Application")
If Err.Number = 0 Then
    Set objExcel = Nothing
    MsgBox IIf(IsJA, "Excel を全て閉じてください！", "Please close all Excel applications !"), vbExclamation,addInName
    WScript.Quit
End If
Err.Clear

IF MsgBox(IIf(IsJA, "アドインをインストールしますか？", "Do you want to install this add-in ?"), vbYesNo + vbQuestion, addInName) = vbNo Then
    WScript.Quit
End IF

'Excel インスタンス化
Set objExcel = CreateObject("Excel.Application")
objExcel.Workbooks.Add

'登録済み対策
For i = 1 To objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = addInFileName Then
        objAddin.Installed = False
    End If
Next

Set objWshShell = CreateObject("WScript.Shell")
Set objFileSys = CreateObject("Scripting.FileSystemObject")

'インストール先パスの作成
installPath = objWshShell.SpecialFolders("Appdata") & "\Microsoft\Addins\" & addInFileName

'ファイルコピー(上書き)
objFileSys.CopyFile addInFileName , installPath, True

Set objWshShell = Nothing
Set objFileSys = Nothing

'アドイン Workbook タイトル設定（言語に応じてタイトル変更）
Set objWorkbook = objExcel.Workbooks.Open(installPath)
objExcel.DisplayAlerts = False
objWorkbook.Title = addInName
objWorkbook.Save
objWorkbook.Close
objExcel.DisplayAlerts = True

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
    MsgBox IIf(IsJA, "エラーが発生しました: " & CStr(Err.Number) & vbCrLF & "実行環境を確認してください", "An error has occurred." & vbCrLF & "Please check your environment."), vbExclamation, addInName
End IF
