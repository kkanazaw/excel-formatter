= excel-formatter

特定のセルの参照先を変更して上書き保存するスクリプト
（上書きするので、実行前に元ファイルバックアップをおねがいします。）

== ビルド、実行環境
- java 1.8.0_72
- gradle 2.1.0

== ビルド方法

```
$ gradle clean
$ gradle build
```

== 実行方法

```
# 1ファイルの場合
$ java -jar  build/libs/jp.designrule.ExcelFormatter.jar {}

# sample以下のファイルの場合
$ find sample -name *.xlsx | xargs -I{} java -jar  build/libs/jp.designrule.ExcelFormatter.jar {}
```
