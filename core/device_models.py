"""USB/BTユーティリティのドメインモデルとサービス。"""

from __future__ import annotations

import json
import os
import sys
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Tuple, TYPE_CHECKING, Union

from .com_ports import ComPortManager

if TYPE_CHECKING:  # pragma: no cover - 循環参照回避の型ヒントのみ
    from usb_util_gui import UsbIdsDatabase  # type: ignore


@dataclass
class UsbDeviceSnapshot:
    """USB/BLEデバイスのシリアライズ可能なスナップショット。"""

    vid: str
    pid: str
    device_type: str = "usb"
    manufacturer: str = ""
    product: str = ""
    serial: str = ""
    bus: Optional[int] = None
    address: Optional[int] = None
    port_path: List[int] = field(default_factory=list)
    device_descriptor: Dict[str, Any] = field(default_factory=dict)
    configurations: List[Dict[str, Any]] = field(default_factory=list)
    class_guess: str = "-"
    error: Optional[str] = None
    topology_chain: List[str] = field(default_factory=list)
    location_information: str = ""
    location_fallback: str = ""
    usb_controllers: List[str] = field(default_factory=list)
    ble_address: str = ""
    ble_name: str = ""
    ble_rssi: Optional[int] = None
    ble_uuids: List[str] = field(default_factory=list)

    def key(self) -> str:
        """VID/PIDの正規キーを返す。"""
        if self.device_type == "ble":
            token = self.ble_address or self.ble_name or "-"
            return f"BLE:{token}"
        return f"{self.vid}:{self.pid}"

    def identity(self) -> str:
        """同一デバイスを区別するための識別子を返す。"""
        if self.device_type == "ble":
            parts: List[str] = []
            if self.ble_address:
                parts.append(f"ADDR:{self.ble_address}")
            if self.ble_name:
                parts.append(f"NAME:{self.ble_name}")
            if self.ble_rssi is not None:
                parts.append(f"RSSI:{self.ble_rssi}")
            if not parts:
                parts.append("UNIDENTIFIED")
            return " | ".join(parts)
        parts = []
        serial = self.serial.strip() if isinstance(self.serial, str) else ""
        if serial and serial not in {"取得不可", "-"}:
            parts.append(f"SER:{serial}")
        elif self.port_path:
            path = "-".join(str(p) for p in self.port_path)
            parts.append(f"PORT:{path}")
        if self.bus is not None:
            parts.append(f"BUS:{self.bus}")
        if self.address is not None:
            parts.append(f"ADDR:{self.address}")
        if not parts:
            parts.append("UNIDENTIFIED")
        parts.append(f"VIDPID:{self.vid}:{self.pid}")
        return " | ".join(parts)

    def resolve_names(self, ids_db: "UsbIdsDatabase") -> Tuple[str, str]:
        """usb.idsからベンダー/製品名を解決する。"""
        if self.device_type == "ble":
            return "―", "―"
        if self.error:
            return "取得不可", "取得不可"
        if (
            isinstance(self.vid, str)
            and isinstance(self.pid, str)
            and self.vid.startswith("0x")
            and self.pid.startswith("0x")
        ):
            vendor_name, product_name = ids_db.lookup(self.vid, self.pid)
            return vendor_name or "不明", product_name or "不明"
        return "不明", "不明"

    def to_dict(self) -> Dict[str, Any]:
        """辞書へシリアライズする。"""
        return {
            "device_type": self.device_type,
            "vid": self.vid,
            "pid": self.pid,
            "manufacturer": self.manufacturer,
            "product": self.product,
            "serial": self.serial,
            "bus": self.bus,
            "address": self.address,
            "port_path": self.port_path,
            "device_descriptor": self.device_descriptor,
            "configurations": self.configurations,
            "class_guess": self.class_guess,
            "error": self.error,
            "topology_chain": self.topology_chain,
            "location_information": self.location_information,
            "location_fallback": self.location_fallback,
            "usb_controllers": self.usb_controllers,
            "ble_address": self.ble_address,
            "ble_name": self.ble_name,
            "ble_rssi": self.ble_rssi,
            "ble_uuids": self.ble_uuids,
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "UsbDeviceSnapshot":
        """辞書からデシリアライズする。"""
        return cls(
            device_type=data.get("device_type", "usb"),
            vid=data.get("vid", "-"),
            pid=data.get("pid", "-"),
            manufacturer=data.get("manufacturer", ""),
            product=data.get("product", ""),
            serial=data.get("serial", ""),
            bus=data.get("bus"),
            address=data.get("address"),
            port_path=data.get("port_path", []),
            device_descriptor=data.get("device_descriptor", {}),
            configurations=data.get("configurations", []),
            class_guess=data.get("class_guess", "-"),
            error=data.get("error"),
            topology_chain=data.get("topology_chain", []),
            location_information=data.get("location_information", ""),
            location_fallback=data.get("location_fallback", ""),
            usb_controllers=data.get("usb_controllers", []),
            ble_address=data.get("ble_address", ""),
            ble_name=data.get("ble_name", ""),
            ble_rssi=data.get("ble_rssi"),
            ble_uuids=data.get("ble_uuids", []),
        )


class UsbSnapshotRepository:
    """USB/BTデバイススナップショットのJSON保存を扱う。"""

    def __init__(self, json_path: str) -> None:
        self.json_path = json_path

    def load(self) -> List[UsbDeviceSnapshot]:
        """JSONストレージからスナップショットを読み込む。"""
        if not os.path.exists(self.json_path):
            return [self.placeholder("USB/BTデバイス情報が存在しません")]
        try:
            with open(self.json_path, "r", encoding="utf-8") as infile:
                data = json.load(infile)
        except (OSError, json.JSONDecodeError):
            return [self.placeholder("USB/BTデバイス情報の読み込みに失敗しました")]
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list) or not data:
            return [self.placeholder("USB/BTデバイス情報が空です")]
        return [UsbDeviceSnapshot.from_dict(item) for item in data]

    def save(self, snapshots: List[UsbDeviceSnapshot]) -> None:
        """JSONストレージへスナップショットを保存する。"""
        try:
            with open(self.json_path, "w", encoding="utf-8") as outfile:
                json.dump(
                    [snapshot.to_dict() for snapshot in snapshots],
                    outfile,
                    ensure_ascii=False,
                    indent=2,
                )
        except OSError as exc:
            print(f"USB/BT情報の書き込みに失敗しました: {exc}", file=sys.stderr)

    @staticmethod
    def placeholder(message: str) -> UsbDeviceSnapshot:
        """エラー情報を含むプレースホルダースナップショットを返す。"""
        return UsbDeviceSnapshot(vid="-", pid="-", error=message)


class UsbSnapshotService:
    """デバイススナップショットのスキャン/保存を統括する。"""

    def __init__(self, scanner: "DeviceScanner", repository: UsbSnapshotRepository) -> None:
        self._scanner = scanner
        self._repository = repository

    def refresh(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        """USB/BLEデバイスをスキャンし、保存して (snapshots, error) を返す。"""
        snapshots, scan_error = self._scanner.scan()
        if not snapshots:
            placeholder_message = scan_error or "USB/BTデバイスが見つかりません"
            snapshots = [self._repository.placeholder(placeholder_message)]
        usb_snapshots = [snap for snap in snapshots if snap.device_type == "usb"]
        if usb_snapshots:
            from .scanners import annotate_windows_topology

            annotate_windows_topology(usb_snapshots)
        self._repository.save(snapshots)
        return snapshots, scan_error

    def load(self) -> List[UsbDeviceSnapshot]:
        """永続化ストレージからスナップショットを読み込む。"""
        return self._repository.load()

    def is_usb_device_connected(
        self, vid: str, pid: str, serial: Optional[str] = None
    ) -> bool:
        """デバイス接続確認を下位スキャナへ委譲する。"""
        if hasattr(self._scanner, "is_usb_device_connected"):
            return self._scanner.is_usb_device_connected(vid, pid, serial)
        return False

    def find_snapshots(
        self,
        vid: str,
        pid: str,
        serial: Optional[str] = None,
        *,
        refresh: bool = False,
    ) -> List[UsbDeviceSnapshot]:
        """指定VID/PID/(任意)シリアルに一致するスナップショットを返す。"""
        if refresh:
            snapshots, _ = self.refresh()
        else:
            snapshots = self.load()

        target_vid = _normalize_hex(vid)
        target_pid = _normalize_hex(pid)
        target_serial = _normalize_serial(serial)

        matches: List[UsbDeviceSnapshot] = []
        for snapshot in snapshots:
            if snapshot.error:
                continue
            if snapshot.device_type != "usb":
                continue
            if _normalize_hex(snapshot.vid) != target_vid:
                continue
            if _normalize_hex(snapshot.pid) != target_pid:
                continue
            serial_value = _normalize_serial(snapshot.serial)
            if target_serial and serial_value != target_serial:
                continue
            matches.append(snapshot)
        return matches

    def find_device_connections(
        self,
        vid: str,
        pid: str,
        serial: Optional[str] = None,
        *,
        refresh: bool = False,
    ) -> List[dict]:
        """デバイス情報に関連COMポートを付与して返す。"""
        snapshots = self.find_snapshots(vid, pid, serial, refresh=refresh)
        ports = ComPortManager.get_com_ports()
        results: List[dict] = []
        for snapshot in snapshots:
            if snapshot.device_type != "usb":
                continue
            com_port = self._match_com_port(snapshot, ports)
            results.append(
                {
                    "snapshot": snapshot,
                    "identity": snapshot.identity(),
                    "serial": snapshot.serial,
                    "com_port": com_port,
                    "port_path": list(snapshot.port_path),
                    "bus": snapshot.bus,
                    "address": snapshot.address,
                }
            )
        return results

    def send_serial_command(
        self,
        vid: str,
        pid: str,
        command: Union[str, bytes],
        *,
        serial: Optional[str] = None,
        refresh: bool = False,
        baudrate: int = 9600,
        timeout: float = 2.0,
        read_bytes: Optional[int] = None,
        read_until: Optional[str] = None,
        encoding: str = "utf-8",
        append_newline: bool = False,
    ) -> Dict[str, Any]:
        """
        VID/PIDに一致するデバイスのCOMポートを特定し、指定コマンドを送受信する。

        Returns:
            dict: 送受信に利用したポート名、書き込みバイト数、受信データなどを格納。

        Raises:
            RuntimeError: 対象デバイス未検出やポート未特定、pyserial未導入、通信失敗時。
        """
        try:
            import serial  # type: ignore
            from serial import SerialException  # type: ignore
        except ImportError as exc:
            raise RuntimeError("pyserialが必要です。pip install pyserial を実行してください。") from exc

        connections = self.find_device_connections(vid, pid, serial, refresh=refresh)
        if not connections:
            raise RuntimeError("指定されたVID/PIDに一致するデバイスが見つかりません。")

        candidates = [entry for entry in connections if entry.get("com_port")]
        if not candidates:
            raise RuntimeError("一致するデバイスは見つかったものの、COMポートを検出できませんでした。")
        if len(candidates) > 1 and serial is None:
            raise RuntimeError(
                "複数の一致するCOMポートが見つかりました。シリアル番号を指定して絞り込んでください。"
            )

        target = candidates[0]
        port_name = target.get("com_port")
        if not port_name:
            raise RuntimeError("COMポート情報を取得できませんでした。")

        if read_bytes is not None and read_bytes <= 0:
            read_bytes = None

        if isinstance(command, bytes):
            payload = command
        else:
            payload = command.encode(encoding)
        if append_newline and not payload.endswith(b"\n"):
            payload += "\n".encode(encoding)

        bytes_written = 0
        response_bytes = b""
        try:
            with serial.Serial(
                port=port_name,
                baudrate=baudrate,
                timeout=timeout,
                write_timeout=timeout,
            ) as ser:
                ser.reset_input_buffer()
                ser.reset_output_buffer()
                bytes_written = ser.write(payload)
                ser.flush()
                if read_until is not None:
                    delimiter_text = read_until if read_until else "\n"
                    delimiter = delimiter_text.encode(encoding)
                    response_bytes = ser.read_until(delimiter)
                elif read_bytes:
                    response_bytes = ser.read(read_bytes)
        except SerialException as exc:
            raise RuntimeError(f"シリアルポート {port_name} へのアクセスに失敗しました: {exc}") from exc

        response_text = ""
        if response_bytes:
            try:
                response_text = response_bytes.decode(encoding)
            except UnicodeDecodeError:
                response_text = ""

        return {
            "port": port_name,
            "bytes_written": bytes_written,
            "response_raw": response_bytes,
            "response_text": response_text,
            "encoding": encoding,
            "device": target,
        }

    def get_com_port_for_device(
        self,
        vid: str,
        pid: str,
        serial: Optional[str] = None,
        *,
        refresh: bool = False,
    ) -> Optional[str]:
        """
        指定VID/PID/(任意)シリアルに一致するCOMポート名を返す。

        Args:
            vid: ベンダーID（0x1234形式推奨）
            pid: プロダクトID
            serial: シリアル番号（任意）
            refresh: Trueの場合は照合前にスキャンを実行

        Returns:
            一致するCOMポート名。検出できない場合はNone。
        """
        connections = self.find_device_connections(vid, pid, serial, refresh=refresh)
        for entry in connections:
            com_port = entry.get("com_port")
            if com_port:
                return com_port
        return None

    @staticmethod
    def _match_com_port(snapshot: UsbDeviceSnapshot, ports: List[dict]) -> Optional[str]:
        for port in ports:
            if (
                _normalize_hex(port.get("vid")) == _normalize_hex(snapshot.vid)
                and _normalize_hex(port.get("pid")) == _normalize_hex(snapshot.pid)
            ):
                serial = _normalize_serial(snapshot.serial)
                port_serial = _normalize_serial(port.get("serial_number"))
                if serial and serial != port_serial:
                    continue
                return port.get("device")
        return None


def _normalize_hex(value: Optional[object]) -> str:
    if value is None:
        return ""
    if isinstance(value, int):
        return format(value, "04x")
    text = str(value).strip().lower()
    if text.startswith("0x"):
        text = text[2:]
    return text.zfill(4)


def _normalize_serial(value: Optional[object]) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text in {"", "取得不可", "-"}:
        return ""
    return text
