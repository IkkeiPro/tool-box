# Tool Box

日々の作業で使うツール集です。

## Excel Masker

Excel内の個人名を「A社のメンバー1」などの会社別ラベルに一括置換します。対応表の入力・保存、複数ファイル・フォルダ選択に対応しています。HTML画面とPython標準ライブラリで動作し、外部サービスへデータを送信しません。

### 起動方法

Python 3.10以降を用意し、次のコマンドを実行してください。

```sh
cd excel-masker
python app.py
```

Windowsでは `excel-masker/start.bat` のダブルクリックでも起動できます。ブラウザが開いたら対応表とExcelファイルを指定してください。

対応形式は `.xlsx` です。画像内の文字は処理できません。シート名・数式参照などに登録名が残るファイルは出力を停止します。

詳しい使い方と制限は [Excel Maskerの説明書](excel-masker/README.md) を参照してください。

## テスト

```sh
cd excel-masker
python -m unittest -v test_app
```
