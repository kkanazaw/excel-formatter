# excel-formatter

特定のセルの参照先を変更して上書き保存するスクリプト
（上書きするので、実行前に元ファイルバックアップをおねがいします。）

## ビルド、実行環境
- java 1.8.0_72
- gradle 2.1.0

## ビルド方法

```
$ gradle clean
$ gradle build
```

## 実行方法

```
# 1ファイルの場合
$ java -jar  build/libs/excel-formatter-1.0.0.jar ファイルパス

# sample以下の複数ファイルの場合
$ find sample -name *.xlsx | xargs -I{} java -jar  build/libs/excel-formatter-1.0.0.jar {}
```

## 参考資料

- (POIメモ(Hishidama's Apache POI Memo))[http://www.ne.jp/asahi/hishidama/home/tech/apache/poi/]
- (成済みのシートを取得 - シート - Apache POIでExcelを操作)[http://www.javadrive.jp/poi/sheet/index2.html]
- (ubuntuにgradleをインストールする方法)[http://qiita.com/htano/items/31b042a264c3f2983b12]
