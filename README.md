# 対象CSVファイルを読み込み、Excelへ転記する

## 解決したい問題
Excel で CSV ファイルを直接開くだと、自動変換機能が働いて読み込みと同時にデータが変換されてしまい、

- 先頭の0を消える
- 日付フォーマト変更される
- 桁数多い数字で E+11 が表示される

などなどの問題があります。

<img src="https://github.com/megumiimai/CsvToExcel/blob/master/docs/csv_open_to_excel_%20malfunction.png" width="600">

セルの書式定義を変更することによって解決できますが、大量 OR フォーマット異なるCSV を確認するとき手間かかります。


上記の問題を解消するため、エクセルのマクロよりCSVファイルを読み込み、ワークシートに転記するようにしました。

## 使い方
エクセルファイルをダウンロードする。</br>
マクロのセキュリティ設定は「マクロ有効」となっている。</br>
「参照」ボタンを押して、対象CSVファイルを選択してください。</br>
「区切り文字」でCSVファイルのカラム区切り文字が指定できます。</br>
「ログファイル読み込み」ボタンを押して、新しいシートが作成されて、CSVファイルの内容が転記されています。<br>

<img src="https://github.com/megumiimai/CsvToExcel/blob/master/docs/how_to.gif" width="800">



## 動作確認済み Office
Windows 7  & Office 2013<br>
Windows 10 & Office 2013<br>
Windows 10 & Office 2016<br>


### 要注意
Mac では動作しません。
これから対応する予定です。
