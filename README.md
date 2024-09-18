# create_text_book_tag

## はじめに
バイトの作業を効率化するだけのものだから，気にするな．

## ライブラリーのインストール
* pandas
```shell
$ pip3 install pandas
```

* openpyxl
```shell
$pip3 install openpyxl
```

## 実行
1. [ここ](https://github.com/IshigamiRyoichi/create_text_book_tag/blob/3b4f216d63f5c838f1f1b7ff6d5c19fd72df118c/create_text_book_tag.py#L79)のExcel名をDLした教科書一覧に変更する。
2. 次のコマンドを実行
    ```python
    $ python3 create_text_book_tag.py
    ```

## 処理手順
1. 棚札を作成のための教科書一覧を取得

1. 棚札を作成
    1. excelファイルを読み込む
    1. セルを接合する
    1. セルに値を入力する
    1. セルのフォントを設定する
    1. 罫線を引く

1. excelファイルを保存する
