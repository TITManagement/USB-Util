"""シリアル/COMポートの列挙とフィルタ用ユーティリティ。"""

from __future__ import annotations

import platform
import sys
from typing import Dict, List, Optional


class ComPortManager:
    """USBシリアル/COMポート調査のクロスプラットフォーム補助。"""

    _ports_cache: Optional[List[Dict[str, Optional[str]]]] = None

    @staticmethod
    def is_usb_device_connected(vid: str, pid: str, serial: Optional[str] = None) -> bool:
        """指定識別子のUSBデバイスが接続中ならTrueを返す。"""

        try:
            import usb.core  # type: ignore
            import usb.util  # type: ignore
        except ImportError:
            return False

        devices = usb.core.find(find_all=True)
        if devices is None:
            return False

        for device in devices:
            dev_vid = getattr(device, "idVendor", None)
            dev_pid = getattr(device, "idProduct", None)
            dev_vid_str = (
                hex(dev_vid)
                if isinstance(dev_vid, int)
                else str(dev_vid)
                if dev_vid is not None
                else ""
            )
            dev_pid_str = (
                hex(dev_pid)
                if isinstance(dev_pid, int)
                else str(dev_pid)
                if dev_pid is not None
                else ""
            )
            dev_serial = None
            try:
                import usb.util  # type: ignore

                dev_serial = usb.util.get_string(device, getattr(device, "iSerialNumber", None))
            except Exception:  # pragma: no cover - usb backend dependent
                dev_serial = None
            if str(vid).lower() == dev_vid_str.lower() and str(pid).lower() == dev_pid_str.lower():
                if serial is None or str(serial) == str(dev_serial):
                    return True
        return False

    @staticmethod
    def get_com_ports() -> List[Dict[str, Optional[str]]]:
        """Windows/macOS/Linuxのシリアルポートを列挙する。"""

        system = platform.system()
        com_ports: List[Dict[str, Optional[str]]] = []
        if system == "Windows":
            try:
                import win32com.client  # type: ignore

                wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
                for device in wmi.ConnectServer(".", "root\\cimv2").ExecQuery(
                    "SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"
                ):
                    com_ports.append(
                        {
                            "device": device.Name.split()[-1].replace("(", "").replace(")", "")
                            if device.Name
                            else None,
                            "description": device.Name,
                            "pnp_id": device.PNPDeviceID,
                            "vid": None,
                            "pid": None,
                            "serial_number": None,
                            "manufacturer": None,
                            "product": None,
                        }
                    )
            except Exception as exc:  # pragma: no cover - platform specific
                print(f"COMポート情報取得エラー(Windows): {exc}", file=sys.stderr)
        else:
            try:
                import serial.tools.list_ports as list_ports  # type: ignore
            except ImportError:
                print("pyserialが必要です。pip install pyserial を実行してください。", file=sys.stderr)
                return []
            for port in list_ports.comports():
                if getattr(port, "vid", None) is not None and getattr(port, "pid", None) is not None:
                    com_ports.append(
                        {
                            "device": port.device,
                            "description": port.description,
                            "hwid": port.hwid,
                            "vid": hex(port.vid) if port.vid is not None else None,
                            "pid": hex(port.pid) if port.pid is not None else None,
                            "serial_number": getattr(port, "serial_number", None),
                            "manufacturer": getattr(port, "manufacturer", None),
                            "product": getattr(port, "product", None),
                        }
                    )
        return com_ports

    @classmethod
    def get_com_ports_cached(cls, force_refresh: bool = False) -> List[Dict[str, Optional[str]]]:
        """キャッシュ済みのシリアルポート列挙結果を返す。"""

        if force_refresh or cls._ports_cache is None:
            cls._ports_cache = cls.get_com_ports()
        return cls._ports_cache

    @staticmethod
    def filter_ports(
        ports: List[Dict[str, Optional[str]]],
        vid: Optional[str] = None,
        pid: Optional[str] = None,
        serial: Optional[str] = None,
    ) -> List[Dict[str, Optional[str]]]:
        """VID/PID/シリアル番号でポート一覧を絞り込む。"""

        results: List[Dict[str, Optional[str]]] = []
        for port in ports:
            if vid and port.get("vid") != vid:
                continue
            if pid and port.get("pid") != pid:
                continue
            if serial and port.get("serial_number") != serial:
                continue
            results.append(port)
        return results

    @staticmethod
    def format_port_name(port_name: Optional[str]) -> str:
        """OSに合わせてCOMポート名を正規化する。"""

        if not port_name:
            return ""
        if platform.system() == "Windows":
            return port_name.upper()
        return port_name

    @staticmethod
    def is_port_connected(port_name: Optional[str]) -> bool:
        """対象ポートが現在存在するか確認する。"""

        if not port_name:
            return False
        return any(port.get("device") == port_name for port in ComPortManager.get_com_ports())
