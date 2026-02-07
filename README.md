# USB-util

USB-utilは、接続済みのUSBデバイスをスキャンし、取得した詳細情報をGUIで閲覧できるPythonツールです。macOS/LinuxではPyUSB+libusbを通じてディスクリプタを直接取得し、WindowsではWMI経由でドライバ差し替えなしにCOMポートへ紐付けられるUSB情報を収集します。`usb.ids`データベースでベンダー名・製品名も補完し、スキャン・永続化・表示を責務分離したサービス層＋ViewModel構成を採用しています。

## 主な機能
- macOS/Linux: PyUSBを用いたUSBデバイスのスキャンとメタデータ取得
- Windows: WMI経由でドライバ差し替え不要のUSBメタデータ取得とCOMポート突合
- `usb.ids`と照合したベンダー/プロダクト名の自動解決
- 取得したスナップショットの`usb_devices.json`への保存
- CustomTkinterベースのGUIでのデバイス切り替え表示とJSONプレビュー（ViewとViewModelを分離）
- バックエンドや権限不足などで発生したエラーメッセージの明示表示
- COMポート情報の自動逆引き（`core/com_ports.py` 内のヘルパーを利用）
- WindowsではWMIを用いたハブ/ポート連鎖・USBコントローラ名の自動解析
- Serial / Port Path / BUS/Address を組み合わせたクロスプラットフォームな識別タグ生成

## USB/BLE情報モデル

本ツールは USB と BLE で取得できる情報の性質が異なる点を前提に、共通スナップショットへ統合して扱います。

- USB（実装済み）: VID/PID、Manufacturer/Product、Serial、Bus/Address、Port Path、Descriptor情報
- BLE（実装済み）: BLE Address、デバイス名、RSSI、Service UUIDs

補足:

- USB は物理接続とディスクリプタを中心とした情報体系です。
- BLE は無線広告情報とサービス情報を中心とした情報体系です。
- ベンダー名/製品名の厳密な補完は主に USB（`usb.ids`）で行います。

## 動作要件
- Python 3.9 以降
- macOS/Linux: libusb 1.0（PyUSBのバックエンドとして利用）
- Windows: WMI サービスが有効な環境（WinUSB/Zadig等は不要）
- ネイティブライブラリのビルド/読み込みに必要な環境（OS付属またはパッケージ管理システムでインストール）

### OS別セットアップ

#### macOS
- Homebrewで `brew install libusb`

#### Ubuntu 24.04
- `sudo apt update && sudo apt install libusb-1.0-0`

#### Windows
- Python仮想環境有効化: `venv\Scripts\activate`（PowerShell/コマンドプロンプト）
- 追加パッケージ: `pip install customtkinter wmi`
- シリアル通信やCOMポート列挙を行う場合は `pip install pyserial`

> **注意**: Windowsは管理者権限で実行することを推奨します。

### Python依存パッケージ
```
customtkinter
pyusb      # macOS/LinuxでのUSBスキャン用
wmi        # WindowsでのUSB情報取得用
pyserial   # COMポート列挙/シリアル通信を行う場合
```

> **補足**: macOSでHomebrewを利用している場合は `brew install libusb`、Linuxでは各ディストリビューションのパッケージマネージャーから `libusb-1.0` をインストールしてください。WindowsではWMIサービスが有効であれば追加のドライバ導入は不要です。

## セットアップ
1. 必要なライブラリをインストールします。
   ```zsh
   python -m venv .venv
   source .venv/bin/activate  # Windowsは .venv\Scripts\activate
   pip install customtkinter
   pip install pyusb  # macOS/Linuxの場合
   pip install wmi    # Windowsの場合
   pip install pyserial  # COMポートを扱う場合
   ```
2. `usb.ids` をプロジェクトルートに配置するか、環境変数 `USB_IDS_PATH` で外部の`usb.ids`ファイルを指定します。  
   代表的な設置場所（Linux, macOS）は `/usr/share/hwdata/usb.ids` 等です。

## 使い方
1. Python環境をアクティブにした状態で `usb_util_gui.py` を実行します。
   ```zsh
   python usb_util_gui.py
   ```
2. 初回起動時にUSBデバイスをスキャンし、結果を `usb_devices.json` に保存します。
3. GUI上部のコンボボックスからデバイス（`VID:PID`形式）を切り替えると、左側に主要情報、右側に詳細JSONが表示されます（ViewModelが選択状態と表示内容を管理します）。

### USB IDsの検索パス
アプリケーションは以下の優先順位で `usb.ids` を探索します。
1. 環境変数 `USB_IDS_PATH`
2. プロジェクトルートの `usb.ids`
3. 現在の作業ディレクトリの `usb.ids`
4. OS毎の代表パス（例: `/usr/share/hwdata/usb.ids`、`/opt/homebrew/share/hwdata/usb.ids` など）

## COMポート自動割当・逆引き機能

USB-Serialデバイスは、PCから抜き差しするたびにCOMポート番号（例: COM5, /dev/ttyUSB0 など）が変わることがありますが、
VID（ベンダーID）、PID（プロダクトID）、Serial（シリアル番号）はデバイス固有で基本的に不変です。

本ツールでは、
- `get_com_port_for_device(vid, pid, serial=None)` というAPIで、
- 指定したUSBデバイス（VID/PID/Serial）に一致する現在割り当てられているCOMポート番号を自動で取得できます。

これにより、
- 物理的な抜き差しやPC再起動後でも、同じデバイスを正しく自動制御できる
- LiquidDispenser("COM5") のようなAPIの引数を自動で割り当て可能
- GUI上でもCOMポート情報が自動表示される

### 使い方例
```python
com_port = get_com_port_for_device(vid, pid, serial)
if com_port:
    sender = LiquidDispenser(com_port)
```

### 外部モジュールからの利用例
プロジェクトとは別のスクリプトから直接COMポートを解決したい場合は、`usb_util_gui.py` が提供するヘルパー関数をインポートして使用できます。

```python
# other_module.py からの利用サンプル
from usb_util_gui import get_com_port_for_device

vid = "0x1234"
pid = "0x5678"
serial = "ABCDEF123"

port = get_com_port_for_device(vid, pid, serial, refresh=True)
if port:
    print(f"USBデバイスのCOMポートは {port} です")
else:
    print("一致するCOMポートが見つかりませんでした")
```

`refresh=True` を指定すると呼び出し前にUSBデバイスの再スキャンを行います（既存スナップショットがある場合は省略可能）。戻り値は `serial.Serial(port=...)` に渡せる文字列です。

### メリット
- デバイス固有情報（vid, pid, serial）からCOMポートを逆引きできる
- USBデバイスの抜き差しや環境変化にも柔軟に対応
- GUIでもCOMポート情報を確認可能
- Windowsでは、WMI経由でHub/Port番号とUSBコントローラ名を同時取得
- Serialが取得できないケースでも、Port PathやBUS/Addressを組み合わせた識別タグで同一機種を区別可能

> 実装上は `core/com_ports.py` の `ComPortManager` を介してアクセスします。GUIでは ViewModel からこのクラスを利用し、同期的にCOMポート情報を解決しています。

## コマンドラインからCOMポートを取得する方法

USBデバイスのVID/PID/Serialから、現在割り当てられているCOMポート番号をコマンドラインで取得できます。

### 使い方
`usb_util_gui.py` を以下のように実行します：

```zsh
python usb_util_gui.py <vid> <pid> [serial]
```
- `<vid>`: ベンダーID（例: 0x1234）
- `<pid>`: プロダクトID（例: 0x5678）
- `[serial]`: シリアル番号（省略可能）
- `--refresh` フラグを併用するとコマンド呼び出しの直前に再スキャンを実施します。

#### 実行例
```zsh
python usb_util_gui.py 0x1234 0x5678 ABCDEF123
```
→ 該当するCOMポート名（例: COM5, /dev/tty.usbserial-xxxx）が標準出力に表示されます。
VID/PIDに対するベンダー名・製品名（usb.ids由来）と、実機から取得した生のManufacturer/Product文字列も併記されるため、入力値の検証が容易です。識別タグ・ポートパス・バス/アドレスも併せて出力されるため、同一機種が複数ある場合でも個体を特定できます。

シリアル番号は省略可能です。

> GUIを起動したい場合は引数なしで `python usb_util_gui.py` を実行してください。

#### 代表的な実行例（`usb_devices.json` に基づく実サンプル）
```zsh
python usb_util_gui.py 0x25a4 0x9311
python usb_util_gui.py 0x25a4 0x9311 --send "STATUS" --append-newline --read-until OK --baudrate 115200
python usb_util_gui.py 0x03e7 0x2485 03e72485 --refresh
python usb_util_gui.py --self-test
```

### 自己診断コマンド

環境セットアップが正しく行われているか確認したいときは、以下のコマンドで自己診断を実行できます。

```zsh
python usb_util_gui.py --self-test
```

- 現在保存されているUSBスナップショット（最大5件）と、pyserialが検出したCOMポートを一覧表示します。
- Windowsの場合はWMI経由の取得結果、macOS/Linuxの場合はPyUSB経由の結果が確認できます。
- 失敗時には不足しているライブラリや権限に関するヒントが表示されます。

## アーキテクチャ

```
core/
 ├─ com_ports.py        # USB-Serialポート検出ヘルパー
 ├─ models.py           # UsbDeviceSnapshotデータクラス
 ├─ scanner.py          # OS毎にPyUSB/WMIを使い分けてUSBデバイスをスキャン
 ├─ repository.py       # JSONファイル永続化
 ├─ service.py          # UsbSnapshotService（スキャンと永続化の調停）
 └─ topology_wmi.py     # Windows用Hub/Port/Controller解析
ui/
 └─ view_model.py       # UsbDevicesViewModel（GUIの状態管理）
usb_util_gui.py         # TkベースのView（UsbDevicesApp）とアプリ起動エントリ
```

- **サービス層 (`core/service.py`)**  
  `UsbSnapshotService` が `UsbScanner` と `UsbSnapshotRepository` を協調させ、スキャン結果をJSONへ保存します。またUSB接続状態確認の窓口もここに集約しています。

- **スキャナ (`core/scanner.py`)**  
  macOS/LinuxではPyUSB、WindowsではWMIを利用してUSBデバイスを列挙・詳細取得します。CLI/GUIの双方から共通利用できるよう、依存は最小限に抑えています。

- **トポロジ解析 (`core/topology_wmi.py`)**  
  Windows環境ではWMIを用いてPnPデバイスを解析し、Hub/Port連鎖とUSBホストコントローラ名を抽出します。サービス層から呼び出され、スナップショットへ追加情報として格納されます。

- **ViewModel (`ui/view_model.py`)**  
  `UsbDevicesViewModel` がスナップショットの並び替え、選択状態、派生情報（COMポート・表示用テキスト）を管理し、View から呼び出し可能なメソッドを提供します。

- **View (`usb_util_gui.py`)**  
  `UsbDevicesApp` はCustomTkinterでUI要素を構築し、ViewModelが提供する値に従って描画・再描画します。イベントハンドラはViewModelへの委譲を行うだけに簡素化されています。

この構成により、ビジネスロジック（スキャン／永続化／状態管理）と表示ロジックを明確に分離し、テストや機能拡張を容易にしています。

## トラブルシューティング
- **PyUSBが見つからない**: macOS/Linuxでは `pip install pyusb` を実行してください。WindowsではPyUSB依存なしで動作します。
- **libusb backend が見つからない**: macOS/Linux向けにlibusb 1.0を導入します。Windowsでは不要です。
- **WMIでエラーが発生する**: PowerShellで `Get-Service Winmgmt` が実行できるか、管理者権限でアプリを起動しているか確認してください。
- **USBデバイスが表示されない**: 実行ユーザーにUSBデバイスへアクセスできる権限があるか確認してください。Linuxでは`plugdev`グループへの所属や`udev`ルールの調整が必要な場合があります。
- **GUIフォントが崩れる**: `Meiryo`が存在しない環境では、CustomTkinterが代替フォントを使用します。必要に応じてコード中のフォント設定を変更してください。

## 開発メモ
- 収集したUSBスナップショットは `usb_devices.json` に配列形式で保存されます。内容を直接編集するとGUIにも反映されます。
- スキャン・永続化ロジックは `core/` 配下、GUIはViewModel＋Viewに二分されています。エントリポイントは `usb_util_gui.py` の `main()` です。
- ViewModelはテストしやすいようにPure Pythonで実装されています。UIの結合テストを行う場合もViewModelをモック化することで効率的に検証できます。

## ライセンス
MIT License

## 作者
- TITManagement　2025
