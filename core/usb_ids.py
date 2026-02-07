from __future__ import annotations

import os
import sys
from typing import Any, Dict, List, Optional, Tuple

# USB-util project root
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


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
    """値をUSB ID表記(4桁の16進文字列)に正規化する。"""
    if value is None:
        return None
    if isinstance(value, int):
        return format(value, "04x")
    text = str(value).strip().lower()
    if text.startswith("0x"):
        text = text[2:]
    return text.zfill(4)


class UsbIdsDatabase:
    """usb.idsファイルをパースし、ベンダーID・プロダクトIDから名称を解決する。"""

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
