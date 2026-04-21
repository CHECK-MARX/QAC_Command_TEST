# Lin_U QAC コマンドテストランナー (Avalonia)

このデスクトップアプリは、手動で行う `Lin_U` コマンドテスト手順を GUI で実行するためのものです。

## できること

- 次のファイルから候補となるプロジェクトペアを読み込みます。
  - `master_settings/all/JISX0208_UTF8_I.txt`
  - `master_settings/all/JISX0208_UTF8_J.txt`
- ペアを選択できます（通常は 3 件）。
- 任意の I/J 入力に対応しています。
  - カスタム行を追加
  - グリッド上で I/J 値を直接編集
  - 行を削除
- 選択したペアを次のファイルに書き込みます。
  - `master_settings/JISX0208_UTF8_I.txt`
  - `master_settings/JISX0208_UTF8_J.txt`
- 元スクリプトを編集しないよう、`Lin_U/.gui_runtime/` 配下に実行用スクリプトを生成します。
  - `testonce.generated.sh`（環境変数の上書き）
  - `testloop.generated.sh`（生成した testonce を呼び出し）
- テストを実行し、出力をリアルタイム表示します。
- `Press [Enter]` プロンプトに対して、設定可能な遅延で自動 Enter を送ります。
- サマリー行（`成功および、N 失敗`）と一般的なエラーマーカーから失敗を検出します。
- GUI 実行ログを `Lin_U/gui_logs/` に保存します。

## ビルド

```bash
export PATH=/home/itoke/.dotnet:$PATH
cd /home/itoke/Cusor_Project_Command_test/LinUCommandTestGui
dotnet restore
dotnet build -c Release
```

## 実行

```bash
export PATH=/home/itoke/.dotnet:$PATH
cd /home/itoke/Cusor_Project_Command_test/LinUCommandTestGui
dotnet run -c Release
```

## 注意事項

- このアプリは Linux のシェルスクリプト（`testloop.sh`, `testonce.sh`）を前提にしています。
- Windows では `testloop.bat` があればそれを使用し、なければ `bash` を試します。
- テスト手順上、QAC 操作の間に待機が必要なため、プロンプト自動確認は意図的に有効化したままにしています。
- 未処理例外は `~/.linu_gui_crash.log` に記録されます。
