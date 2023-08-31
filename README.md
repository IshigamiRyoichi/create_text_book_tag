# create_text_book_tag

## はじめに
バイトの作業を効率化するだけのものだから，気にするな．

## 処理手順
1. 棚札を作成のための教科書一覧を取得

1. 棚札を作成
    1. excelファイルを読み込む
    1. セルを接合する
    1. セルに値を入力する
    1. セルのフォントを設定する
    1. 罫線を引く

1. excelファイルを保存する

## ライブラリーのインストール
* pandas
```shell
$ pip3 install pandas
```

* openpyxl
```shell
$pip3 install openpyxl
```

## run

```python
$ python3 create_text_book_tag.py
```