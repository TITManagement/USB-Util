
# USB-util

USB-utilは、接続済みのUSBデバイスをスキャンし、取得した詳細情報をGUIで閲覧できるPythonツールです。PyUSBで収集した生データをJSONとして保存しつつ、`usb.ids`データベースを参照してベンダー名・製品名を補完します。

## 主な機能
- PyUSBを用いたUSBデバイスのスキャンとメタデータ取得
- `usb.ids`と照合したベンダー/プロダクト名の自動解決
- 取得したスナップショットの`usb_devices.json`への保存
- CustomTkinterベースのGUIでのデバイス切り替え表示とJSONプレビュー
- バックエンドや権限不足などで発生したエラーメッセージの明示表示

## 動作要件
- Python 3.9 以降
- libusb 1.0（PyUSBのバックエンドとして利用）
- ネイティブライブラリのビルド/読み込みに必要な環境（OS付属またはパッケージ管理システムでインストール）

### Python依存パッケージ
```
pyusb
customtkinter
```

> **補足**: macOSでHomebrewを利用している場合は `brew install libusb`、Linuxでは各ディストリビューションのパッケージマネージャーから `libusb-1.0` をインストールしてください。WindowsではZadig等でWinUSB/libusbKドライバーを導入する必要があります。

## セットアップ
1. 必要なライブラリをインストールします。
   ```zsh
   python -m venv .venv
   source .venv/bin/activate  # Windowsは .venv\Scripts\activate
   pip install pyusb customtkinter
   ```
2. `usb.ids` をプロジェクトルートに配置するか、環境変数 `USB_IDS_PATH` で外部の`usb.ids`ファイルを指定します。  
   代表的な設置場所（Linux, macOS）は `/usr/share/hwdata/usb.ids` 等です。

## 使い方
1. Python環境をアクティブにした状態で `main.py` を実行します。
   ```zsh
   python main.py
   ```
2. 初回起動時にUSBデバイスをスキャンし、結果を `usb_devices.json` に保存します。
3. GUI上部のコンボボックスからデバイス（`VID:PID`形式）を切り替えると、左側に主要情報、右側に詳細JSONが表示されます。

### USB IDsの検索パス
アプリケーションは以下の優先順位で `usb.ids` を探索します。
1. 環境変数 `USB_IDS_PATH`
2. プロジェクトルートの `usb.ids`
3. 現在の作業ディレクトリの `usb.ids`
4. OS毎の代表パス（例: `/usr/share/hwdata/usb.ids`、`/opt/homebrew/share/hwdata/usb.ids` など）

## トラブルシューティング
- **PyUSBが見つからない**: `pip install pyusb` を実行してください。
- **libusb backend が見つからない**: OSに合わせてlibusb 1.0を導入します。WindowsではZadig、macOSではHomebrew、LinuxではAPT/YUM等が利用できます。
- **USBデバイスが表示されない**: 実行ユーザーにUSBデバイスへアクセスできる権限があるか確認してください。Linuxでは`plugdev`グループへの所属や`udev`ルールの調整が必要な場合があります。
- **GUIフォントが崩れる**: `Meiryo`が存在しない環境では、CustomTkinterが代替フォントを使用します。必要に応じてコード中のフォント設定を変更してください。

## 開発メモ
- 収集したUSBスナップショットは `usb_devices.json` に配列形式で保存されます。内容を直接編集するとGUIにも反映されます。
- `UsbScanner` や `UsbDataStore` などのクラス構造は `main.py` に定義されています。機能拡張やCLI化を行う際は当該ファイルを参照してください。

## ライセンス
MIT License

## 作者
- TITManagement　2025
# USB-Util
