"""USB/BLEデバイスGUIのビューモデル。"""

from __future__ import annotations

import json
import sys
from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING

from core.com_ports import ComPortManager
from core.device_models import UsbDeviceSnapshot, UsbSnapshotService


if TYPE_CHECKING:  # pragma: no cover - 循環参照回避の型ヒントのみ
    from usb_util_gui import UsbIdsDatabase  # type: ignore


class UsbDevicesViewModel:
    """USBスナップショットとGUIをつなぐ状態管理。"""

    def __init__(self, snapshot_service: UsbSnapshotService, ids_db: "UsbIdsDatabase") -> None:
        self._service = snapshot_service
        self._ids_db = ids_db
        self.snapshots: List[UsbDeviceSnapshot] = []
        self.com_ports: List[Dict[str, Optional[str]]] = []
        self.selected_index: int = 0
        self._last_error: Optional[str] = None

    # ----- データライフサイクル -------------------------------------------
    def load_initial(self, snapshots: List[UsbDeviceSnapshot]) -> None:
        self._update_state(snapshots, preserve_selection=False)

    def refresh(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        snapshots, error = self._service.refresh()
        self._last_error = error
        self._update_state(snapshots, preserve_selection=True)
        return snapshots, error

    def _update_state(self, snapshots: List[UsbDeviceSnapshot], preserve_selection: bool) -> None:
        previous_key = None
        if preserve_selection and self.snapshots and self.selected_index < len(self.snapshots):
            previous_key = self.snapshots[self.selected_index].key()

        self.snapshots = [snap for snap in self._sort_snapshots(snapshots) if not snap.error]
        self.com_ports = ComPortManager.get_com_ports()

        if previous_key:
            self.select_by_key(previous_key)
        else:
            self.selected_index = 0 if self.snapshots else -1

    def error_message(self) -> str:
        """スキャン時のエラーメッセージがあれば返す。"""
        return self._last_error or ""

    # ----- 選択ヘルパー ---------------------------------------------------
    def device_count(self) -> int:
        return len(self.snapshots)

    def get_options(self) -> List[str]:
        return [snapshot.key() for snapshot in self.snapshots]

    def select_by_index(self, index: int) -> None:
        if not self.snapshots:
            self.selected_index = -1
            return
        self.selected_index = max(0, min(index, len(self.snapshots) - 1))

    def select_by_key(self, key: str) -> None:
        if not self.snapshots:
            self.selected_index = -1
            return
        key_lower = key.lower()
        for idx, snapshot in enumerate(self.snapshots):
            if snapshot.key().lower() == key_lower:
                self.selected_index = idx
                return
        self.selected_index = 0

    def current_snapshot(self) -> Optional[UsbDeviceSnapshot]:
        if not self.snapshots or self.selected_index < 0:
            return None
        if self.selected_index >= len(self.snapshots):
            self.selected_index = len(self.snapshots) - 1
        return self.snapshots[self.selected_index]

    # ----- 表示用派生データ ----------------------------------------------
    def info_values(self) -> Dict[str, str]:
        snapshot = self.current_snapshot()
        if snapshot is None:
            return {}

        identity = self._identity_without_vidpid(snapshot)
        if snapshot.device_type == "ble":
            uuids_text = ", ".join(snapshot.ble_uuids) if snapshot.ble_uuids else "―"
            info = {
                "VID": "―",
                "usb.ids Vendor": "―",
                "PID": "―",
                "usb.ids Product": "―",
                "Manufacturer": "―",
                "Product": "―",
                "Serial": "―",
                "Identity": identity,
                "Bus": "―",
                "Address": "―",
                "Port Path": "―",
                "Class Guess": snapshot.class_guess,
                "COMポート": "―",
                "接続経路": "―",
                "LocationInformation": "―",
                "BLE Address": snapshot.ble_address or "―",
                "BLE Name": snapshot.ble_name or "―",
                "BLE RSSI": str(snapshot.ble_rssi) if snapshot.ble_rssi is not None else "―",
                "BLE UUIDs": uuids_text,
            }
            return info

        vendor_label, product_label = snapshot.resolve_names(self._ids_db)
        port_path_text = "-".join(str(p) for p in snapshot.port_path) if snapshot.port_path else "不明"
        bus_text = str(snapshot.bus) if snapshot.bus is not None else "不明"
        address_text = str(snapshot.address) if snapshot.address is not None else "不明"
        com_port_value = self._match_com_port(snapshot)
        hub_path = " -> ".join(snapshot.topology_chain) if snapshot.topology_chain else "未取得"

        info = {
            "VID": snapshot.vid,
            "usb.ids Vendor": vendor_label,
            "PID": snapshot.pid,
            "usb.ids Product": product_label,
            "Manufacturer": snapshot.manufacturer or "―",
            "Product": snapshot.product or "―",
            "Serial": snapshot.serial or "―",
            "Identity": identity,
            "Bus": bus_text,
            "Address": address_text,
            "Port Path": port_path_text,
            "Class Guess": snapshot.class_guess,
            "COMポート": com_port_value or "情報なし",
            "接続経路": hub_path,
            "LocationInformation": snapshot.location_information or snapshot.location_fallback or "―",
            "BLE Address": "―",
            "BLE Name": "―",
            "BLE RSSI": "―",
            "BLE UUIDs": "―",
        }

        return info

    def detail_json(self) -> str:
        snapshot = self.current_snapshot()
        if snapshot is None:
            return "{}"
        return json.dumps(snapshot.to_dict(), ensure_ascii=False, indent=2)

    def list_entries(self) -> List[Tuple[str, bool]]:
        entries: List[Tuple[str, bool]] = []
        for snapshot in self.snapshots:
            if snapshot.device_type == "ble":
                uuids_text = ", ".join(snapshot.ble_uuids) if snapshot.ble_uuids else "―"
                item_text = (
                    f"[BLE] {snapshot.ble_name or '―'}\n"
                    f"Address: {snapshot.ble_address or '―'}\n"
                    f"RSSI: {snapshot.ble_rssi if snapshot.ble_rssi is not None else '―'}\n"
                    f"UUIDs: {uuids_text}"
                )
                entries.append((item_text, False))
                continue
            vendor_label, product_label = snapshot.resolve_names(self._ids_db)
            com_port_value = self._match_com_port(snapshot)
            port_connected = bool(
                com_port_value
                and any(port.get("device") == com_port_value for port in self.com_ports)
            )
            usb_connected = self._service.is_usb_device_connected(snapshot.vid, snapshot.pid, snapshot.serial)
            dimmed = not (port_connected and usb_connected)
            hub_path = " -> ".join(snapshot.topology_chain) if snapshot.topology_chain else "未取得"
            identity = self._identity_without_vidpid(snapshot)
            item_text = (
                f"{product_label or snapshot.product or '―'} / {snapshot.manufacturer or '―'}\n"
                f"  Raw Product: {snapshot.product or '―'}\n"
                f"VID:PID {snapshot.vid}:{snapshot.pid}\n"
                f"Serial: {snapshot.serial or '―'}\n"
                f"ID: {identity}\n"
                f"COM: {com_port_value or '情報なし'}\n"
                f"Class: {snapshot.class_guess}\n"
                f"接続経路: {hub_path}"
            )
            entries.append((item_text, dimmed))
        return entries

    # ----- ユーティリティ ------------------------------------------------
    def _match_com_port(self, snapshot: UsbDeviceSnapshot) -> Optional[str]:
        if snapshot.device_type != "usb":
            return None
        for port in self.com_ports:
            if (
                port.get("vid") == snapshot.vid
                and port.get("pid") == snapshot.pid
                and (
                    not snapshot.serial
                    or port.get("serial_number") == snapshot.serial
                )
            ):
                return port.get("device")
        return None

    @staticmethod
    def _sort_snapshots(snapshots: List[UsbDeviceSnapshot]) -> List[UsbDeviceSnapshot]:
        def sort_key(snap: UsbDeviceSnapshot) -> Tuple[object, object, object, str]:
            if snap.device_type == "ble":
                label = snap.ble_address or snap.ble_name or ""
                return (1, "", "", label)
            return (
                0,
                UsbDevicesViewModel._id_sort_value(snap.vid),
                UsbDevicesViewModel._id_sort_value(snap.pid),
                "",
            )

        return sorted(snapshots, key=sort_key)

    @staticmethod
    def _id_sort_value(value: Any) -> Tuple[int, int, str]:
        if isinstance(value, int):
            return (0, value, format(value, "04x"))
        text = str(value or "").strip().lower()
        if text.startswith("0x"):
            text = text[2:]
        try:
            number = int(text, 16)
            return (0, number, text)
        except ValueError:
            return (1, sys.maxsize, text)

    def selected_option(self) -> str:
        snapshot = self.current_snapshot()
        return snapshot.key() if snapshot else ""

    @staticmethod
    def _identity_without_vidpid(snapshot: UsbDeviceSnapshot) -> str:
        parts = [part for part in snapshot.identity().split(" | ") if not part.startswith("VIDPID:")]
        return " | ".join(parts) if parts else snapshot.identity()
