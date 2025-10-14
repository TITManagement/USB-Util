"""
USB-util: USBデバイス情報のスキャン・表示・保存・COMポート逆引きを行うPython GUI/CLIツール

- USBデバイスの詳細情報取得（PyUSB）
- usb.idsによるベンダー名・製品名補完
- JSON保存・読み込み
- CustomTkinter GUIで情報表示
- COMポート情報の取得・逆引き（pyserial/win32com）
- コマンドラインからCOMポート取得も可能
"""

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
import platform

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
    """
    usb.idsファイルをパースし、ベンダーID・プロダクトIDから名称を解決するクラス。
    - reload(): キャッシュ再構築
    - lookup(vid, pid): ベンダー名・製品名取得
    """

    # usb.idsのパスを受け取り、ベンダー/プロダクト名の辞書を構築
    def __init__(self, ids_path: Optional[str] = None) -> None:
        # 初期化。ids_pathはusb.idsファイルのパス
        self.ids_path = ids_path or find_usb_ids_path()
        self._cache: Optional[Dict[str, Dict[str, Any]]] = None

    def reload(self) -> None:
        # キャッシュをクリアして再パース可能にする
        self._cache = None

    def lookup(self, vid: Any, pid: Any) -> Tuple[Optional[str], Optional[str]]:
        # ベンダーID・プロダクトIDから名称を取得
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
        # キャッシュがなければパースして構築
        if self._cache is None:
            self._cache = self._parse_usb_ids(self.ids_path)
        return self._cache

    @staticmethod
    def _parse_usb_ids(ids_path: str) -> Dict[str, Dict[str, Any]]:
        # usb.idsファイルをパースして辞書化
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
                        # インターフェース情報
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
                        # プロダクト情報
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
                    # ベンダー情報
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
    """
    USBデバイスのスナップショット情報構造体。
    - resolve_names(ids_db): usb.idsから名称解決
    - to_dict()/from_dict(): 辞書⇔構造体変換
    """

    # USBデバイスの属性を保持
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
        # デバイス一意識別子（VID:PID）
        return f"{self.vid}:{self.pid}"

    def resolve_names(self, ids_db: UsbIdsDatabase) -> Tuple[str, str]:
        # usb.idsデータベースからベンダー名・製品名を解決
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
        # 辞書形式に変換
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
        # 辞書から構造体へ変換
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
    """
    PyUSBを用いてUSBデバイスをスキャンし、詳細情報を取得するクラス。
    - scan(): USBデバイス一覧取得
    """

    # USBデバイスをスキャンして詳細情報を取得
    def scan(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        # USBデバイス一覧とエラー情報を返す
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
        # libusbバックエンドを解決
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
        # 属性取得の安全ラッパー
        try:
            value = getattr(obj, attr)
            return "取得不可" if value is None else value
        except AttributeError:
            return "取得不可"

    @classmethod
    def _safe_str(cls, usb_util: Any, obj: Any, idx: Any) -> str:
        # USBデバイスの文字列属性取得
        try:
            return usb_util.get_string(obj, idx)
        except Exception:
            return "取得不可"

    def _snapshot_device(self, device: Any, usb_util: Any) -> UsbDeviceSnapshot:
        # 1デバイスの詳細情報をスナップショット化
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
        # エラー時のスナップショット生成
        return UsbDeviceSnapshot(
            vid="-",
            pid="-",
            manufacturer="",
            product="",
            error=message,
        )

    @staticmethod
    def _no_backend_message() -> str:
        # libusbバックエンド未検出時のメッセージ
        if sys.platform.startswith("win"):
            return "libusb backend が見つかりません。Zadig などで WinUSB/libusbK を導入してください。"
        if sys.platform.startswith("linux"):
            return "libusb backend が見つかりません。libusb-1.0 パッケージをインストールしてください。"
        return "libusb backend が見つかりません。libusb-1.0 をインストールしてください。"


class UsbDataStore:
    """
    USBデバイススナップショットのJSON保存・読み込み管理クラス。
    - refresh(): スキャン＆保存
    - load(): JSONから復元
    """

    # USBデバイス情報の保存・復元を管理
    def __init__(self, json_path: str, scanner: UsbScanner) -> None:
        # 初期化。保存パスとスキャナを受け取る
        self.json_path = json_path
        self.scanner = scanner

    def refresh(self) -> Tuple[List[UsbDeviceSnapshot], Optional[str]]:
        # USBデバイスをスキャンし、JSON保存。エラーも返す
        snapshots, scan_error = self.scanner.scan()
        if not snapshots:
            if scan_error:
                snapshots = [self._placeholder(scan_error)]
            else:
                snapshots = [self._placeholder("USBデバイスが見つかりません")]
        self._write_json(snapshots)
        return snapshots, scan_error

    def load(self) -> List[UsbDeviceSnapshot]:
        # JSONからUSBデバイス情報を復元
        print("[DEBUG] load() called")
        if not os.path.exists(self.json_path):
            print("[DEBUG] json_path not found:", self.json_path)
            return [self._placeholder("USBデバイス情報が存在しません")]
        try:
            with open(self.json_path, "r", encoding="utf-8") as infile:
                data = json.load(infile)
        except (OSError, json.JSONDecodeError):
            print("[DEBUG] JSON load error")
            return [self._placeholder("USBデバイス情報の読み込みに失敗しました")]
        if isinstance(data, dict):
            data = [data]
        if not isinstance(data, list) or not data:
            print("[DEBUG] data is empty or not a list")
            return [self._placeholder("USBデバイス情報が空です")]
        devices = [UsbDeviceSnapshot.from_dict(item) for item in data]
        def sort_key(dev: UsbDeviceSnapshot):
            def to_int(val):
                s = str(val).lower()
                if s.startswith("0x"):
                    s = s[2:]
                try:
                    return int(s, 16)
                except Exception:
                    return 0
            return (to_int(dev.vid), to_int(dev.pid))
        devices.sort(key=sort_key)
        print("[DEBUG] ソート後のデバイス順:", [dev.key() for dev in devices])
        return devices

    def _write_json(self, snapshots: List[UsbDeviceSnapshot]) -> None:
        # USBデバイス情報をJSON保存
        try:
            with open(self.json_path, "w", encoding="utf-8") as outfile:
                json.dump([snap.to_dict() for snap in snapshots], outfile, ensure_ascii=False, indent=2)
        except OSError as exc:
            print(f"USB情報の書き込みに失敗しました: {exc}", file=sys.stderr)

    @staticmethod
    def _placeholder(message: str) -> UsbDeviceSnapshot:
        # エラー時のダミー情報生成
        return UsbDeviceSnapshot(vid="-", pid="-", error=message)


class UsbDevicesApp:
    """
    USBデバイス情報をGUIで表示するアプリケーション。
    - 左: USBデバイス一覧
    - 中: デバイス詳細
    - 右: JSON表示
    """

    # GUIの初期化・レイアウト構築
    def __init__(self, snapshots: List[UsbDeviceSnapshot], ids_db: UsbIdsDatabase) -> None:
        # USBデバイス情報・idsデータベースを受け取り、GUIを構築
        self.snapshots = self._sort_snapshots(snapshots)
        self.ids_db = ids_db
        self.com_ports = get_com_ports()  # COMポート情報取得
        self.app = ctk.CTk()
        self.app.title("USB Devices Viewer")
        self.app.geometry("900x600")
        self.info_labels: Dict[str, ctk.CTkLabel] = {}
        self.error_label: Optional[ctk.CTkLabel] = None
        self.detail_box: Optional[ctk.CTkTextbox] = None
        self.combo: Optional[ctk.CTkComboBox] = None
        self._build_layout()

    def run(self) -> None:
        # GUIメインループ開始
        self.app.mainloop()

    def _build_layout(self) -> None:
        # GUIレイアウト構築（左:一覧, 中:詳細, 右:JSON）
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
        bottom_frame.grid_columnconfigure(0, weight=0)  # 左パネル
        bottom_frame.grid_columnconfigure(1, weight=1)  # 中央
        bottom_frame.grid_columnconfigure(2, weight=1)  # 右
        bottom_frame.grid_rowconfigure(0, weight=1)

        # --- 左：USBデバイス一覧パネル ---
        device_list_frame = ctk.CTkFrame(bottom_frame)
        device_list_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        device_list_label = ctk.CTkLabel(device_list_frame, text="USBデバイス一覧", font=("Meiryo", 14, "bold"))
        device_list_label.pack(anchor="w", padx=10, pady=(10, 0))
        self.device_listbox = ctk.CTkScrollableFrame(device_list_frame, width=260)
        self.device_listbox.pack(fill="both", expand=True, padx=10, pady=10)
        self.device_list_items = []
        for idx, snap in enumerate(self.snapshots):
            vendor_label, product_label = snap.resolve_names(self.ids_db)
            com_port_value = "―"
            for cp in getattr(self, "com_ports", []):
                if (
                    cp.get("vid") == snap.vid and
                    cp.get("pid") == snap.pid and
                    (not snap.serial or cp.get("serial_number") == snap.serial)
                ):
                    com_port_value = cp.get("device")
                    break
            # 元の色（デフォルト）でCOMポート表示
            item_text = f"{product_label or snap.product or '―'} / {snap.manufacturer or '―'}\nVID:PID {snap.vid}:{snap.pid}\nSerial: {snap.serial or '―'}\nCOM: {com_port_value}\nClass: {snap.class_guess}"
            item_btn = ctk.CTkButton(self.device_listbox, text=item_text, width=240, height=60, command=lambda i=idx: self._on_device_list_select(i))
            item_btn.pack(fill="x", padx=4, pady=4)
            self.device_list_items.append(item_btn)

        # --- 中：デバイス情報 ---
        left_frame = ctk.CTkFrame(bottom_frame)
        left_frame.grid(row=0, column=1, sticky="nsew", padx=(0, 10))

        right_frame = ctk.CTkFrame(bottom_frame)
        right_frame.grid(row=0, column=2, sticky="nsew", padx=(10, 0))

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
            "COMポート",
        ]
        for row, label in enumerate(fields):
            if label == "COMポート":
                key_label = ctk.CTkLabel(info_frame, text=f"{label}:", font=("Meiryo", 12, "bold"), text_color="red")
            else:
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

    def _on_device_list_select(self, index: int) -> None:
        if self.combo:
            self.combo.set(self.snapshots[index].key())
        self._update_detail(index)

    @staticmethod
    def _sort_snapshots(snapshots: List[UsbDeviceSnapshot]) -> List[UsbDeviceSnapshot]:
        # USBデバイス情報をVID/PID順でソート
        return sorted(
            snapshots,
            key=lambda snap: (
                UsbDevicesApp._id_sort_value(snap.vid),
                UsbDevicesApp._id_sort_value(snap.pid),
            ),
        )

    @staticmethod
    def _id_sort_value(value: Any) -> Tuple[int, int, str]:
        # ソート用ID値変換
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

    def _on_selection_change(self, _: Optional[str] = None) -> None:
        # コンボボックス選択変更時の処理
        index = self._selected_index()
        self._update_detail(index)

    def _selected_index(self) -> int:
        # 現在選択中のデバイスインデックス取得
        selected = self.combo.get() if self.combo else ""
        for idx, snapshot in enumerate(self.snapshots):
            if snapshot.key().lower() == selected.lower():
                return idx
        return 0

    def _update_detail(self, index: int) -> None:
        # 選択デバイスの詳細情報をGUIに反映
        if not (0 <= index < len(self.snapshots)):
            index = 0
        snapshot = self.snapshots[index]
        vendor_label, product_label = snapshot.resolve_names(self.ids_db)
        port_text = (
            "-".join(str(p) for p in snapshot.port_path) if snapshot.port_path else "不明"
        )
        bus_text = str(snapshot.bus) if snapshot.bus is not None else "不明"
        address_text = str(snapshot.address) if snapshot.address is not None else "不明"
        # COMポート照合
        com_port_value = "―"
        for cp in getattr(self, "com_ports", []):
            # VID/PID/Serialで照合（完全一致優先）
            if (
                cp.get("vid") == snapshot.vid and
                cp.get("pid") == snapshot.pid and
                (not snapshot.serial or cp.get("serial_number") == snapshot.serial)
            ):
                com_port_value = cp.get("device")
                break
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
            "COMポート": com_port_value,  # 追加
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


# --- COMポート取得（Windows/Mac/Linux共通） ---
def get_com_ports():
    """
    Windows/Mac/LinuxでUSB-SerialデバイスのCOMポート情報を取得する。
    Windowsはwin32com、Mac/Linuxはpyserialを利用。
    戻り値: COMポート情報の辞書リスト
    """

    # OSごとにCOMポート情報を取得
    system = platform.system()
    com_ports = []
    if system == "Windows":
        try:
            import win32com.client
            wmi = win32com.client.Dispatch("WbemScripting.SWbemLocator")
            for device in wmi.ConnectServer(".", "root\\cimv2").ExecQuery("SELECT * FROM Win32_PnPEntity WHERE Name LIKE '%(COM%'"):
                com_ports.append({
                    "device": device.Name.split()[-1].replace("(","").replace(")","") if device.Name else None,
                    "description": device.Name,
                    "pnp_id": device.PNPDeviceID,
                    # WindowsはVID/PID/Serialの抽出が難しいため、pnp_idから部分抽出
                    "vid": None,
                    "pid": None,
                    "serial_number": None,
                    "manufacturer": None,
                    "product": None,
                })
        except Exception as e:
            print("COMポート情報取得エラー(Windows):", e, file=sys.stderr)
    else:
        try:
            import serial.tools.list_ports as list_ports
        except ImportError:
            print("pyserialが必要です。pip install pyserial を実行してください。", file=sys.stderr)
            return []
        for p in list_ports.comports():
            if getattr(p, 'vid', None) is not None and getattr(p, 'pid', None) is not None:
                com_ports.append({
                    "device": p.device,
                    "description": p.description,
                    "hwid": p.hwid,
                    "vid": hex(p.vid) if p.vid is not None else None,
                    "pid": hex(p.pid) if p.pid is not None else None,
                    "serial_number": getattr(p, 'serial_number', None),
                    "manufacturer": getattr(p, 'manufacturer', None),
                    "product": getattr(p, 'product', None),
                })
    return com_ports


def get_com_port_for_device(vid: str, pid: str, serial: str = None) -> str:
    """
    指定したVID/PID/Serialに一致するUSB-SerialデバイスのCOMポート名（例: 'COM5', '/dev/tty.usbserial-xxxx'）を返す。
    一致しない場合は空文字を返す。
    引数:
        vid: ベンダーID（16進文字列）
        pid: プロダクトID（16進文字列）
        serial: シリアル番号（省略可）
    戻り値:
        COMポート名（str）
    """

    # COMポート情報リストから一致するものを検索
    ports = get_com_ports()
    for cp in ports:
        if (
            cp.get("vid") == vid and
            cp.get("pid") == pid and
            (serial is None or cp.get("serial_number") == serial)
        ):
            return cp.get("device", "")
    return ""


def get_current_com_port(vid: str, pid: str, serial: str = None) -> str:
    """
    GUIを使わず、指定したVID/PID/Serialに一致するUSB-Serialデバイスの現在割り当てられているCOMポート名（例: 'COM5', '/dev/tty.usbserial-xxxx'）を返す。
    一致しない場合は空文字を返す。
    引数:
        vid: ベンダーID（16進文字列）
        pid: プロダクトID（16進文字列）
        serial: シリアル番号（省略可）
    戻り値:
        COMポート名（str）
    """

    # get_com_port_for_deviceをラップ
    return get_com_port_for_device(vid, pid, serial)


def main() -> None:
    """
    USB-utilのメインエントリ。
    - USBデバイススキャン＆保存
    - GUI起動
    """

    # USBデバイス情報をスキャンし、GUIを起動
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
