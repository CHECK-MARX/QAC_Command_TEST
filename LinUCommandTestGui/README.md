# Lin_U QAC コマンドテストランナー (Avalonia)

このデスクトップアプリは、`Lin_U` の手動コマンドテスト手順を GUI で実行・監視するためのツールです。
実装に基づく現行機能を以下にまとめます。

## 主な機能

### 1. パス設定とプロジェクトペア管理
- コマンドテスト用ディレクトリを指定し、`Lin_U` ルートを自動解決します。
- フォルダピッカーでディレクトリ選択できます（`~` 展開対応）。
- テスト対象 I/J ファイルを読み込み、以下の編集ができます。
  - 行追加
  - I/J 値の直接編集
  - 行削除
  - 選択（チェック）
  - 先頭3件の一括選択
- 選択したペアを以下へ保存します。
  - `master_settings/JISX0208_UTF8_I.txt`
  - `master_settings/JISX0208_UTF8_J.txt`

### 2. 環境変数管理（testonce 連携）
- `testonce.sh` から環境変数を読み込み、GUIに反映します。
  - `QAF_ROOT`, `QACLI_BIN`, `TEST_ROOT`
  - `QAV_SERVER`, `QAV_USER`, `QAV_PASS`
  - `VAL_SERVER`, `VAL_USER`, `VAL_PASS`
- GUI値を `testonce.sh` へ保存できます。

### 3. 実行用スクリプト自動生成
- 元スクリプトを直接編集せず、`Lin_U/.gui_runtime/` 配下に実行用スクリプトを生成します。
  - `testonce.generated.sh`
  - `testloop.generated.sh`
- 生成時に以下を自動適用します。
  - 環境変数の反映
  - `qacli` 存在チェックと `PATH` 補正ガード
  - ループ進捗マーカー注入（`[GUI-LOOP-START/DONE/FAIL]`）
  - `--edit` 利用時のエディタバイパス（必要時）

### 4. 実行制御（開始・一時停止・再開・停止）
- `Start` で実行開始。実行中に `Start` を押すと一時停止状態から再開します。
- `Stop` は実行中は一時停止、停止中（PAUSED）に再度押すと終了要求を送ります。
- Linux/macOS ではプロセスツリーにシグナルを送り一時停止/再開します。
- 実行PID、経過時間、状態をGUIで表示します。

### 5. 自動 Enter・出力監視
- `Press [Enter]` を検知すると遅延付きで自動 Enter を送信します。
- 出力が一定時間止まった場合のアイドル再送も行います。
- リアルタイム出力表示（文字数上限付き）に対応します。
- 自動スクロールON/OFFと、選択範囲コピーが可能です。

### 6. 進捗・判定・エラー検知
- 進捗バーと進捗テキスト（完了数/総数）を表示します。
- 予想終了時刻・残り時間を表示します。
- 監視指標を表示します。
  - ループ開始/完了/失敗
  - サマリー成功/失敗合計
  - ERR行数
  - 最終出力時刻・出力停止経過秒
- 判定（`RUNNING` / `PAUSED` / `STOPPED` / `PASS` / `FAIL`）と理由を表示します。
- 失敗検知はサマリー行・ループ失敗・エラーマーカーを組み合わせて行います。
- `最初のエラーで停止` オプションに対応します。

### 7. ログと例外処理
- 実行ログを `Lin_U/gui_logs/gui_run_YYYYMMDD_HHMMSS.log` に保存します。
- GUI例外はステータスへ反映し、未処理例外はクラッシュログに記録します。
  - `~/.linu_gui_crash.log`

### 8. 表示言語切替
- 日本語 / 英語のUI切替に対応します。

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

## 対応環境と注意

- 本アプリは `Lin_U` 側に `testonce.sh` / `testloop.sh` / `master_settings` がある構成を前提にします。
- Windows では `testloop.bat` が存在すればそれを優先し、無い場合は `bash` 実行を試みます。
- 疑似端末 (`script`) が使える環境では、手動実行に近い表示モードで起動します。
