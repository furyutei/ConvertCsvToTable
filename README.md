[Excel用CSV変換アドイン](https://github.com/furyutei/ConvertCsvToTable)
=======================================================================

- License: The MIT license  
- Copyright (c) 2020 風柳(furyu)  
- 対象Excel: Microsoft® Excel® for Office 365 MSO 32ビット
- 対象OS: Windows 10

CSVファイルをExcelに関連付けているときに、見やすい形（テーブル）に変換して開くためのアドイン。  
CSVファイルの形式がUTF-8(BOM無し)でも、文字化けせずに開ける……かも？  


■ インストール
---
1. 右上の [Clone or download ▽] → 「Download ZIP」でダウンロード
2. Excel が起動している場合、終了させる
3. ZIP ファイルを展開して出てくる addin フォルダ中の Install.vbs をダブルクリックし、指示に従う  


■ 使い方
---
インストールすると、リボンに「CSV変換」という名前のタブが追加され、これをクリックするとメニューが表示される。  

![「CSV変換」タブ](https://github.com/furyutei/ConvertCsvToTable/blob/images/ConvertCsvToTable.Menu.png)

- 「コントロール」では、CSVをダブルクリックで開いた際に自動変換するかどうかを選択可能
- 「操作」では、以下のことが可能
  - 「手動変換」：（自動変換をしていない場合）CSVを開いた後での手動変換
  - 「文字を数値に変換」：文字列として認識されている列を選択して実行すると、数値に変換される

[デモ動画はこちら](https://youtu.be/v0nORRevUjw)。  


■ アンインストール
---
1. Excel が起動している場合、終了させる
2. addin フォルダ中の Uninstall.vbs をダブルクリックし、指示に従う  
