MakeDatedFolders WebUI (localhost)

■まずこれで起動してください（推奨）
  1) フォルダを展開
  2) Start.bat をダブルクリック
  3) アプリモードのウィンドウで http://localhost:8787/ が開きます（ポートは自動選択）

■注意
- PowerShellの「#requires」や「param(...)」は“スクリプト先頭でのみ有効”なディレクティブです。
  PowerShellコンソールに直接コピペして実行するとエラーになります。
  必ず .ps1 をファイルとして実行してください。

- 実行ポリシー(ExecutionPolicy)で「署名されていないので実行できない」と出る場合は、
  Start.bat が -ExecutionPolicy Bypass と Unblock-File を試行します。

■フォルダ選択ボタンについて
- ブラウザからローカルのフォルダ選択ダイアログを開くため、PowerShell側が STA で起動します。

■構成
- core/ : コア処理（UI非依存）
- host/ : localhostサーバ & API
- ui/   : スタイリッシュなWeb UI
