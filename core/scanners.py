"""USB/BLEスキャンとWindowsトポロジー取得のユーティリティ。"""

from __future__ import annotations

import asyncio
import os
import platform
import re
import sys
from ctypes.util import find_library
from typing import Any, Dict, Iterable, List, Optional, Tuple

from .device_models import UsbDeviceSnapshot

AK_COMM_ROOT = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "..", "lab_automation_api", "libs", "ak_communication")
)
if AK_COMM_ROOT not in sys.path and os.path.isdir(AK_COMM_ROOT):
    sys.path.insert(0, AK_COMM_ROOT)

try:  # pragma: no cover - optional dependency
    from ak_communication.ble_scanner import BleScanner, BleDevice  # type: ignore
except Exception:  # pragma: no cover - optional dependency
    BleScanner = None  # type: ignore
    BleDevice = None  # type: ignore


class UsbScanner:
    """プラットフォーム別のUSBスキャンを、フォールバック付きで提供する。"""

    def scan(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        system = platform.system().lower()
        if system == "windows":
            return self._scan_windows()
        return self._scan_pyusb()

    def _scan_windows(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        try:
            import wmi  # type: ignore
        except ImportError:
            message = "WindowsでUSB情報を取得するには wmi パッケージが必要です。pip install wmi を実行してください。"
            return [self._error_snapshot(message)], message

        try:
            client = wmi.WMI()
        except Exception as exc:  # pragma: no cover - platform specific
            message = (
                f"WMI初期化に失敗しました: {exc}. "
                "PowerShellで `Get-Service Winmgmt` が実行できる状態か、"
                "管理者権限でアプリを起動しているか確認してください。"
            )
            return [self._error_snapshot(message)], message

        snapshots: List[UsbDeviceSnapshot] = []

        try:
            devices = list(client.Win32_PnPEntity())
        except Exception as exc:  # pragma: no cover - platform specific
            message = f"WMIからUSBデバイス情報を取得できませんでした: {exc}"
            return [self._error_snapshot(message)], message

        for dev in devices:
            device_id = getattr(dev, "DeviceID", "") or ""
            if not device_id.startswith("USB\\"):
                continue

            vid, pid = self._parse_vid_pid(device_id)
            if not vid or not pid:
                continue

            serial = self._parse_serial_from_pnpid(device_id)
            manufacturer = self._safe_wmi_attr(dev, "Manufacturer") or "取得不可"
            product_name = self._safe_wmi_attr(dev, "Name") or self._safe_wmi_attr(dev, "Caption") or "取得不可"
            class_guess = (
                self._safe_wmi_attr(dev, "PNPClass")
                or self._safe_wmi_attr(dev, "ClassGuid")
                or "-"
            )
            location_information = self._safe_wmi_attr(dev, "LocationInformation")
            descriptor: Dict[str, Any] = {
                "pnp_device_id": device_id,
                "description": self._safe_wmi_attr(dev, "Description"),
                "service": self._safe_wmi_attr(dev, "Service"),
                "status": self._safe_wmi_attr(dev, "Status"),
                "present": getattr(dev, "Present", None),
            }

            snapshot = UsbDeviceSnapshot(
                vid=f"0x{vid.lower()}",
                pid=f"0x{pid.lower()}",
                manufacturer=manufacturer,
                product=product_name,
                serial=serial,
                bus=None,
                address=None,
                port_path=[],
                device_descriptor=descriptor,
                configurations=[],
                class_guess=class_guess,
                error=None,
                topology_chain=[],
                location_information=location_information,
                location_fallback="",
                usb_controllers=[],
            )
            snapshots.append(snapshot)

        if not snapshots:
            message = (
                "WMI経由でUSBデバイス情報が取得できませんでした。"
                "USBの接続状態やアプリ実行権限（管理者権限）を確認してください。"
            )
            return [], message

        return snapshots, None

    def _scan_pyusb(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        try:
            import usb.core  # type: ignore
            import usb.util  # type: ignore
            from usb.core import NoBackendError, USBError  # type: ignore
        except ImportError:
            message = "PyUSBが見つかりません。pip install pyusb でインストールしてください。"
            return [self._error_snapshot(message)], message

        backend = self._resolve_backend()
        find_kwargs: Dict[str, Any] = {"find_all": True}
        if backend is not None:
            find_kwargs["backend"] = backend

        try:
            devices_iter = usb.core.find(**find_kwargs)
        except NoBackendError:
            message = self._no_backend_message()
            return [self._error_snapshot(message)], message
        except USBError as exc:
            message = f"USBデバイスへのアクセスに失敗しました: {exc}"
            return [self._error_snapshot(message)], message

        if devices_iter is None:
            return [], None

        snapshots: List[UsbDeviceSnapshot] = []
        for device in devices_iter:
            snapshots.append(self._snapshot_device(device, usb.util))
        return snapshots, None

    @staticmethod
    def is_usb_device_connected(vid: str, pid: str, serial: Optional[str] = None) -> bool:
        system = platform.system().lower()
        if system == "windows":
            return UsbScanner._is_connected_windows(vid, pid, serial)
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
            dev_vid_str = UsbScanner._normalize_vid_pid(dev_vid)
            dev_pid_str = UsbScanner._normalize_vid_pid(dev_pid)
            dev_serial = None
            try:
                import usb.util  # type: ignore

                dev_serial = usb.util.get_string(device, getattr(device, "iSerialNumber", None))
            except Exception:
                dev_serial = None
            if vid.lower() == dev_vid_str.lower() and pid.lower() == dev_pid_str.lower():
                if serial is None or str(serial) == str(dev_serial):
                    return True
        return False

    @staticmethod
    def _is_connected_windows(vid: str, pid: str, serial: Optional[str]) -> bool:
        try:
            import wmi  # type: ignore
        except ImportError:
            return False

        target_vid = UsbScanner._normalize_vid_token(vid)
        target_pid = UsbScanner._normalize_vid_token(pid)
        target_serial = UsbScanner._normalize_serial(serial)

        client = wmi.WMI()
        for dev in client.Win32_PnPEntity():
            device_id = getattr(dev, "DeviceID", "") or ""
            if not device_id.startswith("USB\\"):
                continue
            dev_vid, dev_pid = UsbScanner._parse_vid_pid(device_id)
            if dev_vid != target_vid or dev_pid != target_pid:
                continue
            dev_serial = UsbScanner._parse_serial_from_pnpid(device_id)
            if target_serial and dev_serial != target_serial:
                continue
            return True
        return False

    @staticmethod
    def _resolve_backend():
        try:
            from usb.backend import libusb1  # type: ignore
        except ImportError:
            return None
        backend = libusb1.get_backend()
        if backend:
            return backend
        candidate_names: List[str] = []
        if sys.platform == "darwin":
            candidate_names.append("libusb-1.0.dylib")
        else:
            candidate_names.extend(["libusb-1.0.so", "libusb.so"])
        lib_from_ctypes = find_library("usb-1.0")
        if lib_from_ctypes:
            candidate_names.append(lib_from_ctypes)
        seen = set()
        for name in candidate_names:
            if not name or name in seen:
                continue
            seen.add(name)
            backend = libusb1.get_backend(find_library=lambda _: name)
            if backend:
                return backend
        return None

    @staticmethod
    def _safe_get(obj: Any, attr: str) -> Any:
        try:
            value = getattr(obj, attr)
            return "取得不可" if value is None else value
        except AttributeError:
            return "取得不可"

    @classmethod
    def _safe_str(cls, usb_util: Any, obj: Any, idx: Any) -> str:
        try:
            return usb_util.get_string(obj, idx)
        except Exception:
            return "取得不可"

    def _snapshot_device(self, device: Any, usb_util: Any) -> UsbDeviceSnapshot:
        vid_val = self._safe_get(device, "idVendor")
        pid_val = self._safe_get(device, "idProduct")
        device_descriptor: Dict[str, Any] = {}
        for attr in [
            "idVendor",
            "idProduct",
            "bcdDevice",
            "bDeviceClass",
            "bDeviceSubClass",
            "bDeviceProtocol",
            "bMaxPacketSize0",
            "iManufacturer",
            "iProduct",
            "iSerialNumber",
            "bNumConfigurations",
        ]:
            device_descriptor[attr] = self._safe_get(device, attr)

        configurations: List[Dict[str, Any]] = []
        for cfg in device:
            cfg_info: Dict[str, Any] = {"configuration_descriptor": {}, "interfaces": []}
            for cfg_attr in [
                "bConfigurationValue",
                "bmAttributes",
                "bMaxPower",
                "iConfiguration",
                "bNumInterfaces",
            ]:
                cfg_info["configuration_descriptor"][cfg_attr] = self._safe_get(cfg, cfg_attr)
            for intf in cfg:
                intf_info: Dict[str, Any] = {"interface_descriptor": {}, "endpoints": []}
                for intf_attr in [
                    "bInterfaceNumber",
                    "bAlternateSetting",
                    "bNumEndpoints",
                    "bInterfaceClass",
                    "bInterfaceSubClass",
                    "bInterfaceProtocol",
                    "iInterface",
                ]:
                    intf_info["interface_descriptor"][intf_attr] = self._safe_get(intf, intf_attr)
                endpoints = getattr(intf, "_endpoints", None)
                if endpoints:
                    for ep in endpoints:
                        endpoint_info: Dict[str, Any] = {}
                        for ep_attr in [
                            "bEndpointAddress",
                            "bmAttributes",
                            "wMaxPacketSize",
                            "bInterval",
                        ]:
                            endpoint_info[ep_attr] = self._safe_get(ep, ep_attr)
                        intf_info["endpoints"].append(endpoint_info)
                cfg_info["interfaces"].append(intf_info)
            configurations.append(cfg_info)

        class_list: List[str] = []
        for cfg in device:
            for intf in cfg:
                cls_value = self._safe_get(intf, "bInterfaceClass")
                if cls_value == "取得不可":
                    continue
                class_list.append(self._class_name(cls_value))

        manufacturer = self._safe_str(usb_util, device, self._safe_get(device, "iManufacturer"))
        product = self._safe_str(usb_util, device, self._safe_get(device, "iProduct"))
        serial = self._safe_str(usb_util, device, self._safe_get(device, "iSerialNumber"))
        bus = self._safe_get(device, "bus")
        address = self._safe_get(device, "address")
        port_numbers: List[int] = []
        try:
            port_numbers = list(device.port_numbers)
        except TypeError:
            try:
                port_numbers = list(device.port_numbers())
            except Exception:
                port_numbers = []

        return UsbDeviceSnapshot(
            vid=hex(vid_val) if isinstance(vid_val, int) else str(vid_val),
            pid=hex(pid_val) if isinstance(pid_val, int) else str(pid_val),
            manufacturer=manufacturer,
            product=product,
            serial=serial,
            bus=bus if isinstance(bus, int) else None,
            address=address if isinstance(address, int) else None,
            port_path=port_numbers,
            device_descriptor=device_descriptor,
            configurations=configurations,
            class_guess=",".join(class_list) if class_list else "-",
        )

    @staticmethod
    def _class_name(value: Any) -> str:
        try:
            cls_value = int(value)
        except (TypeError, ValueError):
            return str(value)
        if cls_value == 0x02:
            return "CDC-ACM"
        if cls_value == 0x03:
            return "HID"
        if cls_value == 0xFE:
            return "USBTMC"
        if cls_value == 0xFF:
            return "Vendor"
        return f"0x{cls_value:02X}"

    @staticmethod
    def _error_snapshot(message: str) -> UsbDeviceSnapshot:
        return UsbDeviceSnapshot(vid="-", pid="-", manufacturer="", product="", error=message)

    @staticmethod
    def _no_backend_message() -> str:
        if sys.platform.startswith("linux"):
            return "libusb backend が見つかりません。libusb-1.0 パッケージをインストールしてください。"
        return "libusb backend が見つかりません。libusb-1.0 をインストールしてください。"

    @staticmethod
    def _normalize_vid_pid(value: Any) -> str:
        if value is None:
            return ""
        if isinstance(value, int):
            return hex(value)
        text = str(value).strip().lower()
        if text.startswith("0x"):
            return text
        try:
            return hex(int(text, 16))
        except ValueError:
            return text

    @staticmethod
    def _normalize_vid_token(value: Any) -> str:
        text = str(value or "").strip().upper()
        if text.startswith("0X"):
            text = text[2:]
        return text.zfill(4)

    @staticmethod
    def _normalize_serial(value: Any) -> str:
        text = (str(value).strip() if value is not None else "").upper()
        if not text or text in {"取得不可", "-", ""}:
            return ""
        return text

    @staticmethod
    def _parse_vid_pid(pnp_device_id: str) -> Tuple[str, str]:
        match = re.search(r"VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})", pnp_device_id or "")
        if not match:
            return "", ""
        return match.group(1).upper(), match.group(2).upper()

    @staticmethod
    def _parse_serial_from_pnpid(pnp_device_id: str) -> str:
        if not pnp_device_id:
            return ""
        parts = re.split(r"[\\#]", pnp_device_id)
        if not parts:
            return ""
        return parts[-1].upper()

    @staticmethod
    def _safe_wmi_attr(obj: Any, attr: str) -> str:
        value = getattr(obj, attr, "")
        if value is None:
            return ""
        return str(value).strip()


class DeviceScanner:
    """USB/BLEデバイスをスキャンし、統合スナップショット一覧を返す。"""

    def __init__(self, *, ble_timeout: float = 10.0, adapter: Optional[str] = None) -> None:
        self._usb = UsbScanner()
        self._ble_timeout = ble_timeout
        self._ble = BleScanner(adapter=adapter) if BleScanner else None

    def scan(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        usb_snapshots, usb_error = self._usb.scan()
        ble_snapshots, ble_error = self._scan_ble()
        snapshots = usb_snapshots + ble_snapshots
        return snapshots, self._combine_errors(usb_error, ble_error)

    def is_usb_device_connected(
        self, vid: str, pid: str, serial: Optional[str] = None
    ) -> bool:
        return self._usb.is_usb_device_connected(vid, pid, serial)

    def _scan_ble(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        if self._ble is None:
            return [], "BLEスキャンには ak_communication の ble_scanner.py が必要です。"
        try:
            devices = asyncio.run(self._ble.scan(timeout=self._ble_timeout))
        except Exception as exc:
            return [], f"BLEスキャンに失敗しました: {exc}"

        snapshots = [self._ble_snapshot(device) for device in devices]
        return snapshots, None

    @staticmethod
    def _ble_snapshot(device: "BleDevice") -> UsbDeviceSnapshot:
        return UsbDeviceSnapshot(
            vid="-",
            pid="-",
            device_type="ble",
            manufacturer="",
            product="",
            serial="",
            bus=None,
            address=None,
            port_path=[],
            device_descriptor={},
            configurations=[],
            class_guess="BLE",
            error=None,
            topology_chain=[],
            location_information="",
            location_fallback="",
            usb_controllers=[],
            ble_address=getattr(device, "address", "") or "",
            ble_name=getattr(device, "name", "") or "",
            ble_rssi=getattr(device, "rssi", None),
            ble_uuids=list(getattr(device, "uuids", []) or []),
        )

    @staticmethod
    def _combine_errors(usb_error: Optional[str], ble_error: Optional[str]) -> Optional[str]:
        if usb_error and ble_error:
            return f"USB: {usb_error} / BLE: {ble_error}"
        return usb_error or ble_error


PORT_TOKEN_RE = re.compile(r"(Port_#\d+|Hub_#\d+)", re.IGNORECASE)


def annotate_windows_topology(snapshots: Iterable[UsbDeviceSnapshot]) -> None:
    """WindowsでUSBトポロジー情報を付与する。"""
    if not snapshots:
        return
    if platform.system().lower() != "windows":
        return
    try:
        import wmi  # type: ignore
    except ImportError:
        return

    resolver = _TopologyResolver(wmi.WMI())
    mapping = resolver.build_mapping()
    if not mapping:
        return

    for snapshot in snapshots:
        key_with_serial = _topology_snapshot_key(snapshot, include_serial=True)
        key_without_serial = _topology_snapshot_key(snapshot, include_serial=False)
        entry = mapping.get(key_with_serial) or mapping.get(key_without_serial)
        if not entry:
            continue
        snapshot.topology_chain = entry.get("port_hub_chain", [])
        snapshot.location_information = entry.get("location_information", "")
        snapshot.location_fallback = entry.get("location_fallback", "")
        snapshot.usb_controllers = entry.get("usb_controllers", [])


class _TopologyResolver:
    def __init__(self, wmi_client: "wmi.WMI") -> None:  # type: ignore
        self._wmi = wmi_client

    def build_mapping(self) -> Dict[Tuple[str, str, str], Dict[str, List[str]]]:
        dep_to_ctrl = self._map_entity_to_controller()
        ctrl_names = self._controller_names()
        mapping: Dict[Tuple[str, str, str], Dict[str, List[str]]] = {}

        for dev in self._wmi.Win32_PnPEntity():
            device_id = _topology_norm(getattr(dev, "DeviceID", ""))
            if not device_id.startswith("USB\\"):
                continue
            vid, pid = _topology_parse_vid_pid(device_id)
            if not vid or not pid:
                continue
            serial = _topology_parse_serial_from_pnpid(device_id)
            location_info = _topology_norm(getattr(dev, "LocationInformation", ""))
            chain = _topology_parse_location_chain(location_info)
            controllers = [
                ctrl_names.get(ctrl_id, ctrl_id)
                for ctrl_id in dep_to_ctrl.get(device_id, [])
            ]

            entry = {
                "pnp_device_id": device_id,
                "location_information": location_info,
                "location_fallback": "",
                "port_hub_chain": chain,
                "usb_controllers": controllers,
            }

            key_with_serial = (vid, pid, serial)
            key_without_serial = (vid, pid, "")
            mapping[key_with_serial] = entry
            mapping.setdefault(key_without_serial, entry)

        return mapping

    def _map_entity_to_controller(self) -> Dict[str, List[str]]:
        dep_to_ctrl: Dict[str, List[str]] = {}

        def extract_deviceid(relpath: Optional[str]) -> Optional[str]:
            if not relpath:
                return None
            match = re.search(r'DeviceID="([^\"]+)"', relpath)
            return match.group(1) if match else None

        for rel in self._wmi.Win32_USBControllerDevice():
            dep_path = getattr(rel, "Dependent", None)
            ant_path = getattr(rel, "Antecedent", None)
            dep_id = extract_deviceid(dep_path)
            ant_id = extract_deviceid(ant_path)
            if dep_id and ant_id:
                dep_to_ctrl.setdefault(dep_id, []).append(ant_id)

        return dep_to_ctrl

    def _controller_names(self) -> Dict[str, str]:
        names: Dict[str, str] = {}
        for ctrl in self._wmi.Win32_USBController():
            device_id = _topology_norm(getattr(ctrl, "DeviceID", ""))
            if device_id:
                names[device_id] = _topology_norm(getattr(ctrl, "Name", "")) or device_id
        return names


def _topology_snapshot_key(snapshot: UsbDeviceSnapshot, *, include_serial: bool) -> Tuple[str, str, str]:
    vid = _topology_normalize_vid_pid(snapshot.vid)
    pid = _topology_normalize_vid_pid(snapshot.pid)
    serial = _topology_normalize_serial(snapshot.serial) if include_serial else ""
    return vid, pid, serial


def _topology_parse_location_chain(location_info: str) -> List[str]:
    if not location_info:
        return []
    return [token for token in PORT_TOKEN_RE.findall(location_info)]


def _topology_parse_vid_pid(pnp_device_id: str) -> Tuple[str, str]:
    match = re.search(r"VID_([0-9A-Fa-f]{4}).*PID_([0-9A-Fa-f]{4})", pnp_device_id or "")
    if match:
        return match.group(1).upper(), match.group(2).upper()
    return "", ""


def _topology_parse_serial_from_pnpid(pnp_device_id: str) -> str:
    if not pnp_device_id:
        return ""
    parts = re.split(r"[\\#]", pnp_device_id)
    return parts[-1].upper() if parts else ""


def _topology_normalize_vid_pid(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().upper()
    if text.startswith("0X"):
        text = text[2:]
    return text.zfill(4)


def _topology_normalize_serial(value: Any) -> str:
    text = (str(value).strip() if value is not None else "").upper()
    if not text or text in {"取得不可", "-", ""}:
        return ""
    return text


def _topology_norm(value: Optional[str]) -> str:
    return (value or "").strip()
