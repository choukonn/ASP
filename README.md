# ASP
ASP開発環境やチュートリアル

## 開発環境整える

### 手順

1. コントロール　パネル
2. プログラム
3. プログラムと機能
4. Windowsの機能の有効化または無効化
5. インターネットインフォメーションサービス（チェック）
    1. Web管理ツール（チェック）
    2. World Wide Webサービス（チェック）
        1. アプリケーション開発機能（チェック）
            1. ASP(チェック)
6. OK
7. ie ====> localhost(test)

### add new web

> サイト右クリック　=>  ウェブサイト追加：
* サイト名（任意"ポート-projectName"推薦）
* 物理パス: local directory
* ポート（80以外）

> double click your new サイト：
* ASP (double click)
    * デバッグプロパティ
        * ブラウザーへのエラー送信（True）
    * 動作
        * 親パスを有効する
* 適用

> ディレクトリの参照(double click)
* 有効にする

> 既定のドキュメント
* 入口ファイル（ex. index.asp）

> アプリケーションプール
* 右クリック作成した新たなサイト
* 詳細設定
    * 32ビットアプリケーションの有効化（True）

## チュートリアル

### [ASP](https://www.w3schools.com/asp/asp_introduction.asp)

### [VBScript](https://docs.microsoft.com/ja-jp/previous-versions/windows/scripting/cc392489(v=msdn.10))

obj.open ==> obj.close  set obj = nothing

### [SQL Server](https://www.sqlshack.com/step-by-step-installation-of-sql-server-2017/)
