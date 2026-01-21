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

■テスト手順（安定化・高速化の確認）
1) 起動直後に自動previewしない
   - 起動直後、DevTools の Network で /api/preview が発生しないことを確認
   - 期待結果: /api/preview が自動で走らない
2) BasePath変更しても自動previewしない
   - BasePath を変更し、Network で /api/preview が発生しないことを確認
   - 期待結果: 「変更されました」案内のみ表示
3) Previewボタン押下でのみ /api/preview が走る
   - Preview ボタンを押下して /api/preview が 1 回だけ走ることを確認
   - 期待結果: Previewボタン押下時のみ実行
4) C:\ の大量フォルダでも固まらず、上限/タイムアウトで失敗する場合は理由がUIに出る
   - BasePath をフォルダ数が多い場所（例: C:\Windows）に設定し preview を実行
   - 期待結果: 長時間 Pending にならず、上限/タイムアウトの理由が表示される
5) M:\ 共有フォルダでも “待ち続けない”
   - 共有パスで preview を実行
   - 期待結果: 遅い場合でも上限/タイムアウトで返り、理由が表示される
6) Preview中にRunを押してもbusyで弾かれUIが崩れない
   - preview 連打、preview 中に run を押下
   - 期待結果: 409/busy で抑止され、UI は「処理中」表示＆ボタン無効化
7) Month BasePath の月跨ぎは禁止される
   - 月フォルダを指定し、開始日/終了日が月跨ぎになるように設定
   - 期待結果: preview/run ともに ok:false で停止し、月跨ぎ禁止メッセージ
8) 狭い幅でもヘッダー操作ボタンが2行にならない
   - ウィンドウを狭くしてヘッダー操作ボタンを確認
   - 期待結果: ボタンが折り返さず、必要なら横スクロールできる
9) 文言が「停止」or「終了」に統一されている
   - ヘッダーのサーバー操作ボタンを確認
   - 期待結果: 「終了」の文言で統一されている
