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

### 2. 環境変数管理（環境スクリプト連携）
- ディレクトリ指定時に環境スクリプトから環境変数を読み込み、GUIに反映します。
  - Linux: `testonce.sh`
  - Windows: `test_loop.bat` / `testloop.bat`
  - `QAF_ROOT`, `QACLI_BIN`, `TEST_ROOT`
  - `COM_QAC`, `COM_QACPP`, `COM_RCMA`, `COM_MTA`, `COM_NAME`, `COM_DATA`
  - `QAV_SERVER`, `QAV_USER`, `QAV_PASS`
  - `VAL_SERVER`, `VAL_USER`, `VAL_PASS`
- GUI値を読み込んだ環境スクリプトへ保存できます（Windowsは `.bat` を更新）。

### 3. 元スクリプトの実行
- Linuxでは `testloop.sh`（または `test_loop.sh`）を疑似端末上で実行します。古い `testloop.generated.sh` は選択しません。
- 通常の `qacli` は引数・端末入出力・終了コードを保持し、1回実行します。出力の一括取得や自動再試行は行いません。
- `project config-file --edit` は別端末に編集画面を表示し、終了するまで待ちます。編集テストは省略しません。
- 実行用の端末中継関数だけを `Lin_U/.gui_runtime/qacli-terminal-bridge.sh` に生成します。元のテストスクリプトは変更しません。
- ログには実行スクリプトのパスとSHA-256を記録します。

### 4. 実行制御（開始・一時停止・再開・停止）
- `Start` で実行開始。実行中に `Start` を押すと一時停止状態から再開します。
- `Stop` は実行中は一時停止、停止中（PAUSED）に再度押すと終了要求を送ります。
- Linux/macOS ではプロセスツリーにシグナルを送り一時停止/再開します。
- 実行PID、経過時間、状態をGUIで表示します。

### 5. Enter入力・出力監視
- 初期状態は手動続行です。`Press [Enter]` が出たら、本体またはリアルタイム出力画面の「次へ（Enter）」をクリックします。
- 自動Enterを選択した場合だけ、検出したプロンプトに対して指定の遅延後に入力します。出力停止を理由に入力を再送しません。
- 同じプロンプトへの二重入力や、自動Enterを無効にした後の予約入力を防ぎます。
- 自動Enterの待ち時間はサーバ処理の完了保証ではありません。Dashboardが前のスナップショットを処理中なら、処理完了を確認してから進めてください。
- コマンドの出力は実行中に表示します。自動スクロールON/OFFと選択範囲コピーに対応します。

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
- 編集端末自身の標準エラーは `[TERMINAL-DIAGNOSTIC]` を付けてログに残し、QACの失敗件数に含めません。端末起動失敗・編集コマンドの非ゼロ終了は別途 `[ERR]` として検出します。
- `最初のエラーで停止` オプションに対応します。

### 7. ログと例外処理
- 実行ログを `Lin_U/gui_logs/gui_run_YYYYMMDD_HHMMSS.log` に保存します。
- GUI例外はステータスへ反映し、未処理例外はクラッシュログに記録します。
  - `~/.linu_gui_crash.log`

### 8. 表示言語切替
- 日本語 / 英語のUI切替に対応します。

### 9. クロスプラットフォーム事前確認
- Validate 認証の事前確認タイムアウトは既定で60秒です。
- 必要な場合は環境変数 `LINU_VAL_AUTH_TIMEOUT_SECONDS` で10～600秒に変更できます。
- Windows版 `qacli` の出力はCP932、Linux版はUTF-8として読み取ります。
- 認証引数はシェル文字列へ連結せず、個別のプロセス引数として渡します。
- ウィンドウタイトルとヘッダーに、アプリのバージョン・Gitコミット短縮SHA・実行RIDを表示します（ローカルビルドは `local` と表示）。

## ビルド

```bash
export PATH=/home/itoke/.dotnet:$PATH
cd /home/itoke/Cusor_Project_Command_test/LinUCommandTestGui
dotnet restore
dotnet build -c Release
```

## CI ビルド (GitHub Actions)

- ワークフロー: `.github/workflows/ci-build.yml`
- トリガー:
  - `main` への push
  - Pull Request
  - 手動実行 (`workflow_dispatch`)
- 生成物:
  - Windows: `LinUCommandTestGui-win-x64-<commit-sha>.zip`
  - Linux: `LinUCommandTestGui-linux-x64-<commit-sha>.tar.gz`

CIはWindowsとLinuxの両方でクロスプラットフォームテストを実行し、成功した場合だけ配布物を生成します。

GitHub の Actions 実行結果画面から、Artifacts としてダウンロードできます。

## 実行

```bash
export PATH=/home/itoke/.dotnet:$PATH
cd /home/itoke/Cusor_Project_Command_test/LinUCommandTestGui
dotnet run -c Release
```

## 対応環境と注意

- Linux は `testonce.sh` / `testloop.sh` / `master_settings` がある構成を前提にします。
- Windows は `test_loop.bat`（または `testloop.bat`）を前提にし、環境値も `.bat` 側を読み書きします。
- 疑似端末 (`script`) が使える環境では、手動実行に近い表示モードで起動します。
- 同じGitコミットSHAのWindows/Linux成果物を組み合わせて使用してください。
