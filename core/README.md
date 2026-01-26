# core モジュール概要

このディレクトリは USB-util の中核ロジックをまとめた場所です。  
各ファイルの役割は以下の通りです。

## ファイル一覧

- `device_models.py`  
  デバイススナップショットのデータモデル、永続化（JSON）、およびサービス層をまとめています。
  - `UsbDeviceSnapshot`: USB/BLE の統合モデル
  - `UsbSnapshotRepository`: JSON 保存/読み込み
  - `UsbSnapshotService`: スキャン・保存・検索・COMポート照合の窓口

- `scanners.py`  
  USB/BLE のスキャンと、Windows 向け USB トポロジー補完を担います。
  - `UsbScanner`: USB の実スキャン（WMI / PyUSB）
  - `DeviceScanner`: USB + BLE の統合スキャン
  - `annotate_windows_topology`: WMI でトポロジー情報を付与

- `com_ports.py`  
  シリアル/COM ポート列挙とフィルタリングのユーティリティです。
  - `ComPortManager`: OS ごとのポート取得と簡易フィルタ

- `__init__.py`  
  core パッケージの公開範囲（`__all__`）を定義します。

## 依存関係（ざっくり）

- GUI層 → `UsbSnapshotService` を利用
- `UsbSnapshotService` → `DeviceScanner`（USB/BLE）＋ `UsbSnapshotRepository`
- `DeviceScanner` → `UsbScanner`（USB）＋ `BleScanner`（BLE, ak_communication）

## 注意点

- BLE スキャンは `ak_communication.ble_scanner` に依存します。  
  依存が無い場合は BLE スキャンのみ失敗します（USB は継続）。

## 役割分離の見解

- `com_ports.py`  
  → OS依存のシリアル列挙で単独責務。分離維持が妥当。
- `device_models.py`  
  → モデル＋永続化＋サービス。ドメイン中心の集約として妥当。
- `scanners.py`  
  → USB/BLEスキャン＋Windowsトポロジー。スキャン周りの集約として妥当。
