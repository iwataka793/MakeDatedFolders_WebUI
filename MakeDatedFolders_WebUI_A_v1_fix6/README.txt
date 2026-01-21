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
1) ローカル（C:）
   - BasePath をフォルダ数が多い場所（例: C:\Windows）に設定し preview を実行
   - 期待結果: 長時間 Pending にならず、必要なら探索上限のエラーメッセージが返る
2) 共有（M: / UNC）
   - 提示された共有パスで preview を実行
   - 期待結果: 遅い場合でもタイムアウト/上限で明確なメッセージが返る
3) Root 指定
   - BasePath を年/月より上のルート階層にして preview
   - 期待結果: 深掘り探索が暴走せず、上限超過で「絞ってください」メッセージ
4) 月フォルダ BasePath
   - 月フォルダを指定し、開始日/終了日が月跨ぎになるように設定
   - 期待結果: preview/run ともに ok:false で停止し、月跨ぎ禁止メッセージ
5) 多重クリック
   - preview 連打、preview 中に run を押下
   - 期待結果: 409/busy で抑止され、UI は「処理中」表示＆ボタン無効化
6) UI
   - ウィンドウを狭くしてヘッダー操作ボタンを確認
   - 期待結果: ボタンが折り返さず、必要なら横スクロールできる
