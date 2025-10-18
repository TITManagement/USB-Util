"""Service layer that coordinates USB scanning and persistence."""

from __future__ import annotations

from typing import Any, Dict, List, Optional, Tuple, Union

from .com_ports import ComPortManager
from .models import UsbDeviceSnapshot
from .repository import UsbSnapshotRepository
from .topology_wmi import annotate_windows_topology
class UsbSnapshotService:
    """High-level orchestrator for scanning and storing USB snapshots."""

    def __init__(self, scanner: "UsbScanner", repository: UsbSnapshotRepository) -> None:
        self._scanner = scanner
        self._repository = repository

    def refresh(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        """Scan USB devices, persist results, and return (snapshots, error)."""
        snapshots, scan_error = self._scanner.scan()
        if not snapshots:
            placeholder_message = scan_error or "USBデバイスが見つかりません"
            snapshots = [self._repository.placeholder(placeholder_message)]
        annotate_windows_topology(snapshots)
        self._repository.save(snapshots)
        return snapshots, scan_error

    def load(self) -> List[UsbDeviceSnapshot]:
        """Load snapshots from persistence."""
        return self._repository.load()

    def is_usb_device_connected(
        self, vid: str, pid: str, serial: Optional[str] = None
    ) -> bool:
        """Delegate to the underlying scanner to check device presence."""

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
        """Return snapshots matching the given VID/PID/(optional) Serial."""

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
        """Return device metadata plus associated COM port (if any)."""

        snapshots = self.find_snapshots(vid, pid, serial, refresh=refresh)
        ports = ComPortManager.get_com_ports()
        results: List[dict] = []
        for snapshot in snapshots:
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
                # 古いバッファをクリアして直近の応答のみを取得する
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
        Return the first COMポート名 that matches the given VID/PID/(optional) Serial.

        Args:
            vid: ベンダーID（0x1234形式推奨）
            pid: プロダクトID
            serial: シリアル番号（任意）
            refresh: Trueの場合は照合前にスキャンを実行

        Returns:
            一致するCOMポート名。検出できない場合はNone。
        """

        connections = self.find_device_connections(
            vid, pid, serial, refresh=refresh
        )
        for entry in connections:
            com_port = entry.get("com_port")
            if com_port:
                return com_port
        return None
