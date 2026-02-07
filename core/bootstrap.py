"""USB-utilのサービス初期化ヘルパー。"""

from __future__ import annotations

import sys
from typing import List, Optional, Tuple

from .device_models import UsbDeviceSnapshot, UsbSnapshotRepository, UsbSnapshotService
from .scanners import DeviceScanner


def setup_services(
    usb_json_path: str,
    *,
    ble_timeout: float = 5.0,
) -> Tuple[UsbSnapshotService, List[UsbDeviceSnapshot], Optional[str]]:
    """スキャナ・リポジトリ・サービスを組み立て、最新スナップショットを取得する。"""
    scanner = DeviceScanner(ble_timeout=ble_timeout)
    repository = UsbSnapshotRepository(usb_json_path)
    service = UsbSnapshotService(scanner, repository)
    snapshots, scan_error = service.refresh()
    print(f"[DEBUG] scan() snapshots count: {len(snapshots)}")
    if not snapshots:
        print("[DEBUG] USBデバイスが1つも取得できませんでした")
    if scan_error:
        print(scan_error, file=sys.stderr)
    return service, snapshots, scan_error
