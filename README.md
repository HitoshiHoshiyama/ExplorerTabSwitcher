# ExplorerTabSwitcher

## 目的

Windows11でタブ化したエクスプローラのタブを、マウスホイールで切り替えるために開発しました。  
ついでということで、Edge/Windows Terminal/メモ帳/Acrobat Readerのタブ切り替えも可能にしました。  

## ダウンロード

[https://github.com/HitoshiHoshiyama/ExplorerTabSwitcher/releases](https://github.com/HitoshiHoshiyama/ExplorerTabSwitcher/releases)から最新版のZIPファイルをダウンロードしてください。  
執筆時点(2025/08/28)の最新バージョンはAlpha-7です。

## 動作環境

- Windows11 22H2以降
- .NET 8.0

  動作確認はWindows11 23H2(x64)環境でしか行っていません。  
  Windows10は未確認ですが、EdgeとWindows Terminalの切り替えは動く可能性はあります。

## 利用条件

このソフトウェアはMITライセンスの元で公開されたオープンソースソフトウェアです。  
利用する上で対価不要なためフリーソフトウェアとして扱って問題ありませんが、このソフトウェアを使用して発生したいかなる損害についても責任を負いません。  
ライセンスの全文はLICENSEファイルを参照してください。

## 使用方法

- インストール
  1. 適当な場所にフォルダを作成し、ZIPファイルの中身を展開します。
  1. `TaskRegist.cmd` を管理者として実行します。
- アンインストール  
  1. ZIPファイルの中身を展開したフォルダをエクスプローラで開きます。
  1. `TaskRemove.cmd` を管理者として実行します。
  1. ZIPファイルの中身を展開したフォルダを削除します。

## トラブルシューティング

- 起動しない場合(タスクを起動しても準備完了に戻ってしまう)は、まず動作環境を確認してください。  
  .NET 8.0 がインストールされていない場合は、[.NET 8.0 のダウンロード](https://dotnet.microsoft.com/ja-jp/download/dotnet/8.0)からダウンロード・インストールをお願いします。

- 現象に再現性がある場合、以下の手順でログを採ってIssue登録していただけると解決する可能性があります。  
  1. Windowsの「タスクスケジューラ」を開きます。
  1. 「タスク スケジューラ ライブラリ」から「ExplorerTabSwitcher」を探し、右クリックメニューから「終了」をクリックします。
  1. インストールフォルダにある `NLog.config` をメモ帳などで開きます。
  1. `minlevel="Info"` となっている個所を、`minlevel="Trace"` に変更・保存します。  
  
       変更前
       ```xml
        <rules>
          <logger name="*" minlevel="Info" writeTo="logFile" />
          <logger name="*" minlevel="Trace" writeTo="console" />
        </rules>
       ```
       変更後
       ```xml
        <rules>
          <logger name="*" minlevel="Trace" writeTo="logFile" />
          <logger name="*" minlevel="Trace" writeTo="console" />
        </rules>
       ```
  1. 「ExplorerTabSwitcher」を右クリックし、「実行する(R)」をクリックします。
     状態が「実行中」になるのを確認します。
  1. 現象を再現させます。
  1. インストールフォルダにある `logs` フォルダを開きます。
  1. `logs` フォルダにある今日の日付のファイルが、目的のログファイルです。
     これをIssueに添付してください。
  1. 先ほど変更した `NLog.config` を元に戻して保存します。
  1. 「タスクスケジューラ」から「ExplorerTabSwitcher」を「終了」->「実行する(R)」で再起動したら完了です。

## TODO

- UI要素解析のパターン化(対応アプリがもっと増えたら必要になりそう)

## 作者連絡先

Mail : [oss.develop.public@hosiyama.net](<mailto:oss.develop.public@hosiyama.net>)

Copyright © 2023 Hitoshi Hoshiyama All Rights Reserved.  
This project is licensed under the MIT License, see the LICENSE file.
