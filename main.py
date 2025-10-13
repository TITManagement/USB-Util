# USB-util: USBデバイス情報のスキャン・表示・保存・補完を行うPython GUIツール

# 主な機能
# - PyUSBでUSBデバイスの詳細情報取得
# - usb.idsによるベンダー名・製品名補完
# - JSON保存・読み込み
# - CustomTkinter GUIで情報表示
# - 権限不足/バックエンド未導入時のエラー表示
#
# クラス構成
# - UsbScanner: USBデバイスのスキャン
# - UsbIdsDatabase: usb.idsパースと名称解決
# - UsbDeviceSnapshot: デバイス情報構造体
# - UsbDataStore: JSON保存・読込管理
# - UsbDevicesApp: GUI表示
import json
import os
import sys
from dataclasses import dataclass, field
from ctypes.util import find_library
from typing import Any, Dict, List, Optional, Tuple

import customtkinter as ctk

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
USB_JSON_PATH = os.path.join(BASE_DIR, "usb_devices.json")


def find_usb_ids_path() -> str:
    """Return the first existing usb.ids path across supported platforms."""
    candidates: List[Optional[str]] = []
    env_path = os.environ.get("USB_IDS_PATH")
    if env_path:
        candidates.append(env_path)
    candidates.append(os.path.join(BASE_DIR, "usb.ids"))
    candidates.append(os.path.join(os.getcwd(), "usb.ids"))
    candidates.extend(
        [
            "/usr/share/hwdata/usb.ids",
            "/usr/share/misc/usb.ids",
            "/var/lib/usbutils/usb.ids",
            "/opt/homebrew/share/hwdata/usb.ids",
            "/opt/local/share/hwdata/usb.ids",
        ]
    )
    if sys.platform.startswith("win"):
        program_data = os.environ.get("ProgramData")
        if program_data:
            candidates.append(os.path.join(program_data, "usb.ids"))
    for path in candidates:
        if path and os.path.exists(path):
            return os.path.abspath(path)
    return os.path.join(BASE_DIR, "usb.ids")


def _normalize_usb_id(value: Any) -> Optional[str]:
    if value is None:
        return None
    if isinstance(value, int):
        return format(value, "04x")
    text = str(value).strip().lower()
    if text.startswith("0x"):
        text = text[2:]
    return text.zfill(4)


class UsbIdsDatabase:
    # usb.idsファイルを遅延パースし、ベンダーID・プロダクトIDから名称を解決するクラス。
 
    def __init__(self, ids_path: Optional[str] = None) -> None:
        self.ids_path = ids_path or find_usb_ids_path()
        self._cache: Optional[Dict[str, Dict[str, Any]]] = None

    def reload(self) -> None:
        self._cache = None

    def lookup(self, vid: Any, pid: Any) -> Tuple[Optional[str], Optional[str]]:
        vendors = self._ensure_cache()
        norm_vid = _normalize_usb_id(vid)
        norm_pid = _normalize_usb_id(pid)
        if norm_vid is None or norm_pid is None:
            return None, None
        vendor_entry = vendors.get(norm_vid)
        if not vendor_entry:
            return None, None
        vendor_name = vendor_entry.get("name")
        product_entry = vendor_entry["products"].get(norm_pid)
        product_name = product_entry.get("name") if product_entry else None
        return vendor_name, product_name

    def _ensure_cache(self) -> Dict[str, Dict[str, Any]]:
        if self._cache is None:
            self._cache = self._parse_usb_ids(self.ids_path)
        return self._cache

    @staticmethod
    def _parse_usb_ids(ids_path: str) -> Dict[str, Dict[str, Any]]:
        vendors: Dict[str, Dict[str, Any]] = {}
        current_vendor: Optional[str] = None
        current_product: Optional[str] = None
        try:
            with open(ids_path, "r", encoding="utf-8") as infile:
                for raw_line in infile:
                    line = raw_line.rstrip("\n")
                    if not line or line.startswith("#"):
                        continue
                    if line.startswith("\t\t"):
                        if current_vendor is None or current_product is None:
                            continue
                        parts = line.strip().split(None, 1)
                        if not parts:
                            continue
                        interface_code = parts[0].lower()
                        interface_name = parts[1].strip() if len(parts) > 1 else ""
                        product_entry = vendors[current_vendor]["products"].setdefault(
                            current_product, {"name": "", "interfaces": []}
                        )
                        product_entry["interfaces"].append(
                            {"code": interface_code, "name": interface_name}
                        )
                        continue
                    if line.startswith("\t"):
                        if current_vendor is None:
                            continue
                        parts = line.strip().split(None, 1)
                        if not parts:
                            continue
                        product_id = _normalize_usb_id(parts[0])
                        product_name = parts[1].strip() if len(parts) > 1 else ""
                        vendors[current_vendor]["products"][product_id] = {
                            "name": product_name,
                            "interfaces": [],
                        }
                        current_product = product_id
                        continue
                    parts = line.split(None, 1)
                    if not parts:
                        continue
                    vendor_id = _normalize_usb_id(parts[0])
                    vendor_name = parts[1].strip() if len(parts) > 1 else ""
                    vendors.setdefault(vendor_id, {"name": vendor_name, "products": {}})
                    vendors[vendor_id]["name"] = vendor_name
                    current_vendor = vendor_id
                    current_product = None
        except OSError:
            return {}
        return vendors


@dataclass
class UsbDeviceSnapshot:
    vid: str
    pid: str
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

    def key(self) -> str:
        return f"{self.vid}:{self.pid}"

    def resolve_names(self, ids_db: UsbIdsDatabase) -> Tuple[str, str]:
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
        return {
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
        }

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> "UsbDeviceSnapshot":
        return cls(
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
        )


class UsbScanner:
    # PyUSBを用いてUSBデバイスをスキャンし、詳細情報を取得するクラス。
    """
    Encapsulate PyUSB scanning with backend resolution and error handling.
    """

    def scan(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        try:
            import usb.core
            import usb.util
            from usb.core import NoBackendError, USBError
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
    def _resolve_backend():
        try:
            from usb.backend import libusb1
        except ImportError:
            return None
        backend = libusb1.get_backend()
        if backend:
            return backend
        candidate_names: List[str] = []
        if sys.platform.startswith("win"):
            candidate_names.extend(["libusb-1.0.dll", "libusb0.dll"])
        elif sys.platform == "darwin":
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
        device_descriptor = {}
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
                if cls_value == 0x02:
                    class_name = "CDC-ACM"
                elif cls_value == 0x03:
                    class_name = "HID"
                elif cls_value == 0xFE:
                    class_name = "USBTMC"
                elif cls_value == 0xFF:
                    class_name = "Vendor"
                else:
                    class_name = f"0x{cls_value:02X}"
                class_list.append(class_name)

        manufacturer = self._safe_str(usb_util, device, self._safe_get(device, "iManufacturer"))
        product = self._safe_str(usb_util, device, self._safe_get(device, "iProduct"))
        serial = self._safe_str(usb_util, device, self._safe_get(device, "iSerialNumber"))
        bus = self._safe_get(device, "bus")
        address = self._safe_get(device, "address")
        port_numbers = []
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
    def _error_snapshot(message: str) -> UsbDeviceSnapshot:
        return UsbDeviceSnapshot(
            vid="-",
            pid="-",
            manufacturer="",
            product="",
            error=message,
        )

    @staticmethod
    def _no_backend_message() -> str:
        if sys.platform.startswith("win"):
            return "libusb backend が見つかりません。Zadig などで WinUSB/libusbK を導入してください。"
        if sys.platform.startswith("linux"):
            return "libusb backend が見つかりません。libusb-1.0 パッケージをインストールしてください。"
        return "libusb backend が見つかりません。libusb-1.0 をインストールしてください。"


class UsbDataStore:
    # USBデバイススナップショットのJSON保存・読み込みを管理するクラス。
    # USBデバイススナップショットのJSON保存・読み込みを管理するクラス。

    def __init__(self, json_path: str, scanner: UsbScanner) -> None:
        self.json_path = json_path
        self.scanner = scanner

    def refresh(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        snapshots, scan_error = self.scanner.scan()
        if not snapshots:
            if scan_error:
                snapshots = [self._placeholder(scan_error)]
            else:
                snapshots = [self._placeholder("USBデバイスが見つかりません")]
        self._write_json(snapshots)
        return snapshots, scan_error

    def load(self) -> List[UsbDeviceSnapshot]:
        if not os.path.exists(self.json_path):
            return [self._placeholder("USBデバイス情報が存在しません")]
        try:
            with open(self.json_path, "r", encoding="utf-8") as infile:
                data = json.load(infile)
        except (OSError, json.JSONDecodeError):
            return [self._placeholder("USBデバイス情報の読み込みに失敗しました")]
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list) or not data:
            return [self._placeholder("USBデバイス情報が空です")]
        return [UsbDeviceSnapshot.from_dict(item) for item in data]

    def _write_json(self, snapshots: List[UsbDeviceSnapshot]) -> None:
        try:
            with open(self.json_path, "w", encoding="utf-8") as outfile:
                json.dump([snap.to_dict() for snap in snapshots], outfile, ensure_ascii=False, indent=2)
        except OSError as exc:
            print(f"USB情報の書き込みに失敗しました: {exc}", file=sys.stderr)

    @staticmethod
    def _placeholder(message: str) -> UsbDeviceSnapshot:
        return UsbDeviceSnapshot(vid="-", pid="-", error=message)


class UsbDevicesApp:
    # GUI application wrapper for displaying USB device snapshots.

    def __init__(self, snapshots: List[UsbDeviceSnapshot], ids_db: UsbIdsDatabase) -> None:
        self.snapshots = snapshots
        self.ids_db = ids_db
        self.app = ctk.CTk()
        self.app.title("USB Devices Viewer")
        self.app.geometry("900x600")
        self.info_labels: Dict[str, ctk.CTkLabel] = {}
        self.error_label: Optional[ctk.CTkLabel] = None
        self.detail_box: Optional[ctk.CTkTextbox] = None
        self.combo: Optional[ctk.CTkComboBox] = None
        self._build_layout()

    def run(self) -> None:
        self.app.mainloop()

    def _build_layout(self) -> None:
        main_frame = ctk.CTkFrame(self.app)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        top_frame = ctk.CTkFrame(main_frame)
        top_frame.pack(fill="x", padx=10, pady=(10, 5))

        title = ctk.CTkLabel(top_frame, text="USBデバイス選択", font=("Meiryo", 20, "bold"))
        title.pack(side="left", padx=10)

        options = [snapshot.key() for snapshot in self.snapshots]
        self.combo = ctk.CTkComboBox(top_frame, values=options, width=300)
        self.combo.pack(side="left", padx=20)
        if options:
            self.combo.set(options[0])
        self.combo.configure(command=self._on_selection_change)

        bottom_frame = ctk.CTkFrame(main_frame)
        bottom_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))
        bottom_frame.grid_columnconfigure(0, weight=1)
        bottom_frame.grid_columnconfigure(1, weight=1)
        bottom_frame.grid_rowconfigure(0, weight=1)

        left_frame = ctk.CTkFrame(bottom_frame)
        left_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        right_frame = ctk.CTkFrame(bottom_frame)
        right_frame.grid(row=0, column=1, sticky="nsew", padx=(10, 0))

        left_label = ctk.CTkLabel(left_frame, text="デバイス情報", font=("Meiryo", 14, "bold"))
        left_label.pack(anchor="w", padx=10, pady=(10, 0))

        info_frame = ctk.CTkScrollableFrame(left_frame)
        info_frame.pack(fill="both", expand=True, padx=10, pady=10)

        info_frame.grid_columnconfigure(0, weight=0)
        info_frame.grid_columnconfigure(1, weight=1)

        fields = [
            "VID",
            "usb.ids Vendor",
            "PID",
            "usb.ids Product",
            "Manufacturer",
            "Product",
            "Serial",
            "Bus",
            "Address",
            "Port Path",
            "Class Guess",
        ]
        for row, label in enumerate(fields):
            key_label = ctk.CTkLabel(info_frame, text=f"{label}:", font=("Meiryo", 12, "bold"))
            key_label.grid(row=row, column=0, sticky="w", padx=(10, 6), pady=4)
            value_label = ctk.CTkLabel(info_frame, text="―", font=("Meiryo", 12))
            value_label.grid(row=row, column=1, sticky="w", padx=(0, 10), pady=4)
            self.info_labels[label] = value_label

        self.error_label = ctk.CTkLabel(
            info_frame, text="", font=("Meiryo", 11), text_color="red"
        )
        self.error_label.grid(row=len(fields), column=0, columnspan=2, sticky="w", padx=10, pady=(8, 0))

        right_label = ctk.CTkLabel(right_frame, text="デバイス詳細 (JSON)", font=("Meiryo", 14, "bold"))
        right_label.pack(anchor="w", padx=10, pady=(10, 0))

        self.detail_box = ctk.CTkTextbox(right_frame, font=("Meiryo", 12))
        self.detail_box.pack(fill="both", expand=True, padx=10, pady=10)

        self._update_detail(0)

    def _on_selection_change(self, _: Optional[str] = None) -> None:
        index = self._selected_index()
        self._update_detail(index)

    def _selected_index(self) -> int:
        selected = self.combo.get() if self.combo else ""
        for idx, snapshot in enumerate(self.snapshots):
            if snapshot.key().lower() == selected.lower():
                return idx
        return 0

    def _update_detail(self, index: int) -> None:
        if not (0 <= index < len(self.snapshots)):
            index = 0
        snapshot = self.snapshots[index]
        vendor_label, product_label = snapshot.resolve_names(self.ids_db)
        port_text = (
            "-".join(str(p) for p in snapshot.port_path) if snapshot.port_path else "不明"
        )
        bus_text = str(snapshot.bus) if snapshot.bus is not None else "不明"
        address_text = str(snapshot.address) if snapshot.address is not None else "不明"
        info_values = {
            "VID": snapshot.vid,
            "usb.ids Vendor": vendor_label,
            "PID": snapshot.pid,
            "usb.ids Product": product_label,
            "Manufacturer": snapshot.manufacturer or "―",
            "Product": snapshot.product or "―",
            "Serial": snapshot.serial or "―",
            "Bus": bus_text,
            "Address": address_text,
            "Port Path": port_text,
            "Class Guess": snapshot.class_guess,
        }

        for key, value in info_values.items():
            label = self.info_labels.get(key)
            if label:
                label.configure(text=value)

        if self.error_label:
            if snapshot.error:
                self.error_label.configure(text=f"Error: {snapshot.error}")
            else:
                self.error_label.configure(text="")

        if self.detail_box:
            self.detail_box.configure(state="normal")
            self.detail_box.delete("1.0", "end")
            self.detail_box.insert("end", json.dumps(snapshot.to_dict(), ensure_ascii=False, indent=2))
            self.detail_box.configure(state="disabled")


def main() -> None:
    ids_db = UsbIdsDatabase()
    scanner = UsbScanner()
    data_store = UsbDataStore(USB_JSON_PATH, scanner)
    snapshots, scan_error = data_store.refresh()
    if scan_error:
        print(scan_error, file=sys.stderr)
    app = UsbDevicesApp(snapshots, ids_db)
    app.run()


if __name__ == "__main__":
    main()
