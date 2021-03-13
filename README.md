# ExcelVBALibrary
ExcelVBALibraryは、ExcelVBAで開発する際に役立つ標準モジュール・クラスモジュールをまとめたライブラリです。

## Features

下記の標準モジュール・クラスモジュールから任意のモジュールを導入し、利用することができます。

◆**標準モジュール一覧**

| モジュール名         | 概要                                                         | 依存モジュール |
| -------------------- | ------------------------------------------------------------ | -------------- |
| ArrayUtils           | 配列操作に関するユーティリティモジュール                     | LangUtils      |
| CellAddressUtils     | セルアドレスに関するユーティリティモジュール                 | -              |
| JapaneseHolidayUtils | 日本の「国民の祝日」、「振替休日」、「国民の休日」に関するユーティリティクラス | -              |
| LangUtils            | ExcelVBAの共通的処理に使用されるユーティリティモジュール     | ArrayUtils     |

◆**クラスモジュール一覧**

| モジュール名          | 概要                     | 依存モジュール |
| --------------------- | ------------------------ | -------------- |
| BusinessDayCalculator | 営業日数を計算するクラス | -              |

## Installation

1. srcフォルダ内のbas/clsファイルを任意の場所にダウンロードする。
2. ライブラリを導入したいExcelブックを開き、Visual Basic Editorを起動する(Alt + F11)。
3. プロジェクトエクスプローラ上のツリーの右クリックメニューから「ファイルのインポート」を実行し、ダウンロードしたファイルの中からインポートしたいファイルを選択する(依存関係がある場合は依存先のファイルもインポートする必要があります)。

## Usage
下記ブログを参照してください。<br>
[[ExcelVBA\][サンプルコード] 動的配列が空かどうか判定する](http://javasampleokiba.blog.fc2.com/blog-entry-6.html)<br>
[[ExcelVBA\][サンプルコード] 配列やオブジェクト型の内容をDebug.Printで出力する](http://javasampleokiba.blog.fc2.com/blog-entry-9.html)<br>
[[ExcelVBA\][サンプルコード] ワークシートの列番号と列名の相互変換](http://javasampleokiba.blog.fc2.com/blog-entry-10.html)<br>
[[ExcelVBA\][サンプルコード] 配列をソート（クイックソート）する](http://javasampleokiba.blog.fc2.com/blog-entry-18.html)<br>
[[ExcelVBA\][サンプルコード] 営業日数を計算する](http://javasampleokiba.blog.fc2.com/blog-entry-22.html)<br>
[[ExcelVBA\][サンプルコード] 配列内の要素の位置を取得する](http://javasampleokiba.blog.fc2.com/blog-entry-47.html)<br>
[[ExcelVBA\][サンプルコード] 配列操作（取得・追加・削除編）](http://javasampleokiba.blog.fc2.com/blog-entry-48.html)

## License
[LICENSE](https://github.com/javasampleokiba/ExcelVBALibrary/blob/main/LICENSE)を参照してください。
