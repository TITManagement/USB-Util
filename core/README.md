# core モジュール概要

<!-- README_LEVEL: L3 -->

| 項目 | 内容 |
| --- | --- |
| 文書ID | `AILAB-LAB-AUTOMATION-MODULE-USB-UTIL-CORE-README` |
| 作成日 | 2026-03-08 |
| 作成者 | tinoue |
| 最終更新日 | 2026-03-08 |
| 最終更新者 | tinoue (with Codex) |
| 版数 | v1.0 |
| 状態 | 運用中 |


## 目的

`core/` は USB-util の中核ロジックを集約するディレクトリです。デバイス情報のモデル化、スキャン、保存、COMポート列挙を担当します。

## 含まれる要素

- `device_models.py`: スナップショットモデル・保存/読込・サービス層
- `scanners.py`: USB/BLE スキャンと Windows トポロジー補完
- `com_ports.py`: シリアル/COM ポート列挙とフィルタ
- `__init__.py`: 公開 API (`__all__`) の定義

## 更新ルール

- OS 依存処理は `scanners.py` / `com_ports.py` に閉じ込める
- GUI 層は `core` を直接改変せずサービス経由で利用する
- BLE 依存の変更時は USB 単独動作が維持されることを確認する
